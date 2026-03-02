import os
import requests
import msal
import urllib.parse
import logging
import uuid
import time
import json

from robust_session import RobustSession
from sharepoint_constants import SharePointConstants
from sharepoint_lists import SharePointListWriter, get_dss_type
from dss_constants import DSSConstants
from common import (
    get_value_from_path, parse_url,
    get_value_from_paths,
    is_empty_path, get_lnt_path,
    format_private_key, format_certificate_thumbprint
)
from safe_logger import SafeLogger


logger = SafeLogger("sharepoint-online plugin", DSSConstants.SECRET_PARAMETERS_KEYS)


class SharePointClientError(ValueError):
    pass


class SharePointClient():

    def __init__(self, config, root_name_overwrite_legacy_mode=False):
        self.config = config
        self.root_name_overwrite_legacy_mode = root_name_overwrite_legacy_mode
        self.sharepoint_root = None
        self.sharepoint_url = None
        self.sharepoint_origin = None
        self.allow_string_recasting = config.get("advanced_parameters", False) and config.get("allow_string_recasting", False)
        attempt_session_reset_on_403 = config.get("advanced_parameters", False) and config.get("attempt_session_reset_on_403", False)
        self.session = RobustSession(status_codes_to_retry=[429, 503], attempt_session_reset_on_403=attempt_session_reset_on_403)
        self.number_dumped_logs = 0

        self.dss_column_name = {}
        self.column_ids = {}
        self.column_names = {}
        self.column_entity_property_name = {}
        self.columns_to_format = []
        self.column_sharepoint_type = {}

        # Graph API IDs resolved after auth
        self.site_id = None
        self._drive_path_prefix = ""
        self._drive_id_cache = {}
        self._list_id_cache = {}

        if config.get('auth_type') == DSSConstants.AUTH_OAUTH:
            logger.info("SharePointClient:sharepoint_oauth")
            login_details = config.get('sharepoint_oauth')
            self.assert_login_details(DSSConstants.OAUTH_DETAILS, login_details)
            self.setup_login_details(login_details)
            self.apply_paths_overwrite(config)
            self.setup_sharepoint_online_url(login_details)
            access_token = login_details['sharepoint_oauth']
            self.session.update_settings(
                session=GraphSession(access_token),
                max_retries=SharePointConstants.MAX_RETRIES,
                base_retry_timer_sec=SharePointConstants.WAIT_TIME_BEFORE_RETRY_SEC
            )
        elif config.get('auth_type') == DSSConstants.AUTH_LOGIN:
            raise SharePointClientError(
                "AUTH_LOGIN (sharepy) has been removed. "
                "Please use 'app-username-password', 'app-certificate', 'site-app-permissions', or 'oauth' instead."
            )
        elif config.get('auth_type') == DSSConstants.AUTH_SITE_APP:
            logger.info("SharePointClient:site_app_permissions")
            login_details = config.get('site_app_permissions')
            self.assert_login_details(DSSConstants.SITE_APP_DETAILS, login_details)
            self.setup_sharepoint_online_url(login_details)
            self.setup_login_details(login_details)
            self.apply_paths_overwrite(config)
            self.tenant_id = login_details.get("tenant_id")
            self.client_secret = login_details.get("client_secret")
            self.client_id = login_details.get("client_id")
            access_token = self._get_site_app_access_token()
            self.session.update_settings(
                session=GraphSession(access_token),
                max_retries=SharePointConstants.MAX_RETRIES,
                base_retry_timer_sec=SharePointConstants.WAIT_TIME_BEFORE_RETRY_SEC
            )
        elif config.get('auth_type') == DSSConstants.AUTH_APP_CERTIFICATE:
            logger.info("SharePointClient:app-certificate")
            login_details = config.get('app_certificate')
            self.assert_login_details(DSSConstants.APP_CERTIFICATE_DETAILS, login_details)
            self.setup_sharepoint_online_url(login_details)
            self.setup_login_details(login_details)
            self.apply_paths_overwrite(config)
            self.tenant_id = login_details.get("tenant_id")
            self.client_certificate = format_private_key(login_details.get("client_certificate"))
            self.client_certificate_thumbprint = format_certificate_thumbprint(login_details.get("client_certificate_thumbprint"))
            self.passphrase = login_details.get("passphrase")
            self.client_id = login_details.get("client_id")
            access_token = self._get_certificate_app_access_token()
            self.session.update_settings(
                session=GraphSession(access_token),
                max_retries=SharePointConstants.MAX_RETRIES,
                base_retry_timer_sec=SharePointConstants.WAIT_TIME_BEFORE_RETRY_SEC
            )
        elif config.get('auth_type') == DSSConstants.AUTH_APP_USERNAME_PASSWORD:
            logger.info("SharePointClient:app-username-password")
            login_details = config.get('app_username_password')
            self.assert_login_details(DSSConstants.APP_USERNAME_PASSWORD_DETAILS, login_details)
            self.setup_sharepoint_online_url(login_details)
            self.setup_login_details(login_details)
            self.apply_paths_overwrite(config)
            self.tenant_id = login_details.get("tenant_id")
            self.client_id = login_details.get("client_id")
            self.sharepoint_tenant = login_details.get("sharepoint_tenant")
            username = login_details.get("username")
            password = login_details.get("password")
            access_token = self._get_username_password_access_token(username, password)
            self.session.update_settings(
                session=GraphSession(access_token),
                max_retries=SharePointConstants.MAX_RETRIES,
                base_retry_timer_sec=SharePointConstants.WAIT_TIME_BEFORE_RETRY_SEC
            )
        else:
            raise SharePointClientError("The type of authentication is not selected")

        self.sharepoint_list_title = config.get("sharepoint_list_title")
        logging.getLogger("urllib3").setLevel(logging.WARNING)
        try:
            from urllib3.connectionpool import log
            log.addFilter(SuppressFilter())
        except Exception as err:
            logging.warning("Error while adding filter to urllib3.connectionpool logs: {}".format(err))

        # Resolve Graph API site ID
        self._resolve_site_id()

    # ---- Setup methods ----

    def setup_login_details(self, login_details):
        self.sharepoint_site = login_details.get('sharepoint_site', "").strip("/")
        logger.info("SharePointClient:sharepoint_site={}".format(self.sharepoint_site))
        if 'sharepoint_root' in login_details:
            self.sharepoint_root = login_details['sharepoint_root'].strip("/")
        else:
            self.sharepoint_root = "Shared Documents"
        logger.info("SharePointClient:sharepoint_root={}".format(self.sharepoint_root))

    def apply_paths_overwrite(self, config):
        advanced_parameters = config.get("advanced_parameters", False)
        sharepoint_root_overwrite = config.get("sharepoint_root_overwrite", "").strip("/")
        if self.root_name_overwrite_legacy_mode:
            sharepoint_root_overwrite = sharepoint_root_overwrite.replace("%20", " ")
        sharepoint_site_overwrite = config.get("sharepoint_site_overwrite", "").strip("/")
        if advanced_parameters and sharepoint_root_overwrite:
            self.sharepoint_root = sharepoint_root_overwrite
        if advanced_parameters and sharepoint_site_overwrite:
            self.sharepoint_site = sharepoint_site_overwrite

    def setup_sharepoint_online_url(self, login_details):
        scheme, domain, tenant = parse_url(login_details['sharepoint_tenant'])
        if scheme:
            self.sharepoint_url = domain
            self.sharepoint_origin = scheme + "://" + domain
        elif tenant.endswith(".sharepoint.com"):
            self.sharepoint_url = tenant
            self.sharepoint_origin = "https://" + tenant
        else:
            self.sharepoint_url = tenant + ".sharepoint.com"
            self.sharepoint_origin = "https://" + self.sharepoint_url
        logger.info("SharePointClient:sharepoint_tenant={}, url={}, origin={}".format(
                login_details['sharepoint_tenant'],
                self.sharepoint_url,
                self.sharepoint_origin
            )
        )

    # ---- Graph API resolution methods ----

    def _resolve_site_id(self):
        if self.sharepoint_site:
            url = "{}/sites/{}:/{}".format(
                SharePointConstants.GRAPH_API_BASE_URL,
                self.sharepoint_url,
                self.sharepoint_site
            )
        else:
            url = "{}/sites/{}".format(
                SharePointConstants.GRAPH_API_BASE_URL,
                self.sharepoint_url
            )
        response = self.session.get(url)
        self.assert_response_ok(response, calling_method="_resolve_site_id")
        self.site_id = response.json()["id"]
        logger.info("Resolved site_id={}".format(self.site_id))

    def _resolve_drive_id(self, path=""):
        if self.sharepoint_root:
            root_parts = self.sharepoint_root.split("/", 1)
            self._drive_path_prefix = root_parts[1] if len(root_parts) > 1 else ""
        else:
            root_parts = [path.strip("/").split("/", 1)[0]]
        root_library_name = root_parts[0]

        # Search for drive in Cache
        if root_library_name.lower() in self._drive_id_cache:
            drive_id = self._drive_id_cache[root_library_name.lower()]
            logger.debug("Resolved drive_id={}, path={}, path_prefix={}".format(drive_id, root_library_name, self._drive_path_prefix))
            return drive_id

        # List all drives to add to cache
        url = "{}/sites/{}/drives".format(SharePointConstants.GRAPH_API_BASE_URL, self.site_id)
        response = self.session.get(url)
        self.assert_response_ok(response, calling_method="_resolve_drive_id")
        drives = response.json().get("value", [])

        for drive in drives:
            if drive.get("name"):
                self._drive_id_cache[drive["name"].lower()] = drive["id"]
            if drive.get("webUrl"):
                decoded_web_url = urllib.parse.unquote(drive["webUrl"])
                self._drive_id_cache[decoded_web_url.rsplit("/", 1)[-1].lower()] = drive["id"]
        
        # Search for drive in Cache after updating it with the list of drives from Graph API
        if root_library_name.lower() in self._drive_id_cache:
            drive_id = self._drive_id_cache[root_library_name.lower()]
            logger.info("Resolved drive_id={}, path={}, path_prefix={}".format(drive_id, root_library_name, self._drive_path_prefix))
            return drive_id

        # No match: raise to be explicit about missing root path
        raise SharePointClientError("Drive with root path '{}' not found".format(root_library_name))

    def _resolve_list_id(self, list_title):
        if list_title in self._list_id_cache:
            return self._list_id_cache[list_title]
        url = "{}/sites/{}/lists".format(SharePointConstants.GRAPH_API_BASE_URL, self.site_id)
        response = self.session.get(url)
        self.assert_response_ok(response, calling_method="_resolve_list_id")
        lists = response.json().get("value", [])
        for l in lists:
            if l["list"]["template"] == "webTemplateExtensionList":
                continue
            self._list_id_cache[l.get("displayName", "")] = l["id"]
            self._list_id_cache[l.get("name", "")] = l["id"]
        try:
            list_id = self._list_id_cache[list_title]
        except KeyError:
            raise SharePointClientError("List '{}' not found".format(list_title))
        return list_id

    # ---- URL/path helpers ----

    def _build_drive_path(self, path):
        parts = []
        if self._drive_path_prefix:
            parts.append(self._drive_path_prefix.strip("/"))
        if path:
            clean_path = path.strip("/")
            if clean_path:
                parts.append(clean_path)
        return "/".join(parts)

    def _get_drive_item_url(self, path, suffix=""):
        full_path = self._build_drive_path(path)
        drive_id = self._resolve_drive_id(full_path)
        # if sharepoint_root is not set, ignore the first path segment to be compatible with old configs
        if not self.sharepoint_root:
            full_path = full_path.split("/", 1)[-1] if "/" in full_path else ""
        if full_path:
            encoded_path = urllib.parse.quote(full_path, safe="/")
            if suffix:
                return "{}/drives/{}/root:/{}:{}".format(
                    SharePointConstants.GRAPH_API_BASE_URL, drive_id, encoded_path, suffix)
            else:
                return "{}/drives/{}/root:/{}".format(
                    SharePointConstants.GRAPH_API_BASE_URL, drive_id, encoded_path)
        else:
            return "{}/drives/{}/root{}".format(
                SharePointConstants.GRAPH_API_BASE_URL, drive_id, suffix)

    def _get_list_url(self, list_id=None):
        if list_id:
            return "{}/sites/{}/lists/{}".format(
                SharePointConstants.GRAPH_API_BASE_URL, self.site_id, list_id)
        return "{}/sites/{}/lists".format(
            SharePointConstants.GRAPH_API_BASE_URL, self.site_id)

    # ---- File/folder operations ----

    def _get_children(self, path):
        url = self._get_drive_item_url(path, suffix="/children")
        all_items = []
        while url:
            response = self.session.get(url)
            if response.status_code == 404:
                return []
            self.assert_response_ok(response, calling_method="_get_all_children")
            json_response = response.json()
            all_items.extend(json_response.get("value", []))
            url = json_response.get("@odata.nextLink")
        return all_items

    def list_folder_items(self, path):
        files = []
        folders = []
        for i in self._get_children(path):
            if "file" in i:
                files.append(i)
            elif "folder" in i:
                folders.append(i)
        return files, folders
    
    def get_item(self, path):
        url = self._get_drive_item_url(path)
        response = self.session.get(url)
        if response.status_code == 404:
            return False
        self.assert_response_ok(response, calling_method="is_file")
        return response.json()

    def get_file_content(self, full_path):
        url = self._get_drive_item_url(full_path, suffix="/content")
        response = self.session.get(url)
        self.assert_response_ok(response, no_json=True, calling_method="get_file_content")
        return response

    def write_file_content(self, full_path, data):
        data.seek(0, 2)
        file_size = data.tell()
        data.seek(0)
        self.check_out_file(full_path)
        if file_size <= SharePointConstants.GRAPH_MAX_SIMPLE_UPLOAD_SIZE:
            self._write_simple_upload(full_path, data, file_size)
        else:
            self._write_upload_session(full_path, data, file_size)
        self.check_in_file(full_path)

    def _write_simple_upload(self, full_path, data, file_size):
        url = self._get_drive_item_url(full_path, suffix="/content")
        headers = {
            "Content-Type": "application/octet-stream",
            "Content-Length": str(file_size)
        }
        response = self.session.put(url, headers=headers, data=data)
        self.assert_response_ok(response, calling_method="write_file_content")
        return response

    def _write_upload_session(self, full_path, data, file_size):
        url = self._get_drive_item_url(full_path, suffix="/createUploadSession")
        session_response = self.session.post(url, json={
            "item": {"@microsoft.graph.conflictBehavior": "replace"}
        })
        self.assert_response_ok(session_response, calling_method="create_upload_session")
        upload_url = session_response.json()["uploadUrl"]

        chunk_size = SharePointConstants.GRAPH_UPLOAD_CHUNK_SIZE
        offset = 0
        while offset < file_size:
            chunk_data = data.read(chunk_size)
            if not chunk_data:
                break
            chunk_end = offset + len(chunk_data)
            content_range = "bytes {}-{}/{}".format(offset, chunk_end - 1, file_size)
            headers = {
                "Content-Length": str(len(chunk_data)),
                "Content-Range": content_range
            }
            logger.info("Uploading chunk: {}".format(content_range))
            # Upload chunks go directly to the pre-signed upload URL (no Bearer token needed)
            chunk_response = requests.put(upload_url, headers=headers, data=chunk_data,
                                          timeout=SharePointConstants.TIMEOUT_SEC)
            if chunk_response.status_code not in [200, 201, 202]:
                raise SharePointClientError("Upload chunk failed with status {}: {}".format(
                    chunk_response.status_code, chunk_response.content))
            offset = chunk_end

    def create_folder(self, full_path):
        if is_empty_path(full_path) and is_empty_path(self.sharepoint_root):
            return None
        parent_path, folder_name = os.path.split(full_path.rstrip("/"))
        if not folder_name:
            return None
        url = self._get_drive_item_url(parent_path, suffix="/children")
        body = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail"
        }
        response = self.session.post(url, json=body)
        return response

    def create_path(self, file_full_path):
        """
        Ensure the path of folders that will contain the file specified in file_full_path exists.
         I.e. create all missing folders along the path.
         Does not create the last element in the end of the path as that is the file we will add later
         (unless it ends in / in which case a folder will be created for that also).
        """
        full_path, filename = os.path.split(file_full_path)
        tokens = full_path.strip("/").split("/")
        path = ""
        previous_status = None
        for token in tokens:
            previous_path = path
            path = get_lnt_path(path + "/" + token)
            response = self.create_folder(path)
            if response is not None:
                status_code = response.status_code
                if previous_status == 403 and status_code == 404:
                    logger.error("Could not create folder for '{}'. Check your write permission for the folder {}.".format(path, previous_path))
                previous_status = status_code

    def move_file(self, full_from_path, full_to_path):
        to_parent, to_name = os.path.split(full_to_path.rstrip("/"))
        url = self._get_drive_item_url(full_from_path)
        body = {"name": to_name}

        from_parent, _ = os.path.split(full_from_path.rstrip("/"))
        if from_parent != to_parent:
            self.create_path(full_to_path)
            parent_url = self._get_drive_item_url(to_parent)
            parent_response = self.session.get(parent_url)
            self.assert_response_ok(parent_response, calling_method="move_file:get_parent")
            parent_id = parent_response.json()["id"]
            body["parentReference"] = {"id": parent_id}

        response = self.session.patch(url, json=body)
        self.assert_response_ok(response, calling_method="move_file")
        return response.json()

    def check_in_file(self, full_path):
        logger.info("Checking in {}.".format(full_path))
        url = self._get_drive_item_url(full_path, suffix="/checkin")
        self.session.post(url, json={"comment": "", "checkInAs": "major"})

    def check_out_file(self, full_path):
        logger.info("Checking out {}.".format(full_path))
        url = self._get_drive_item_url(full_path, suffix="/checkout")
        self.session.post(url)

    def recycle(self, full_path):
        url = self._get_drive_item_url(full_path)
        response = self.session.delete(url)
        self.assert_response_ok(response, no_json=True, calling_method="recycle")

    # ---- List operations ----

    def get_list_fields(self, list_title):
        list_id = self._resolve_list_id(list_title)
        url = "{}/columns".format(self._get_list_url(list_id))
        response = self.session.get(url)
        self.assert_response_ok(response, calling_method="get_list_fields")
        columns = response.json().get("value", [])
        if not columns:
            return None
        return self._transform_columns_to_sp_format(columns)

    def _transform_columns_to_sp_format(self, graph_columns):
        result = []
        for col in graph_columns:
            sp_type = self._get_sp_type_from_graph_column(col)
            result.append({
                SharePointConstants.TITLE_COLUMN: col.get("displayName", ""),
                SharePointConstants.TYPE_AS_STRING: sp_type,
                SharePointConstants.STATIC_NAME: col.get("name", ""),
                SharePointConstants.INTERNAL_NAME: col.get("name", ""),
                SharePointConstants.ENTITY_PROPERTY_NAME: col.get("name", ""),
                SharePointConstants.HIDDEN_COLUMN: col.get("hidden", False) or False,
                SharePointConstants.READ_ONLY_FIELD: col.get("readOnly", False) or False,
            })
        return result

    @staticmethod
    def _get_sp_type_from_graph_column(column):
        for graph_type, sp_type in SharePointConstants.GRAPH_TO_SP_TYPE_MAP.items():
            if graph_type in column:
                if graph_type == "text" and column.get("text", {}).get("allowMultipleLines", False):
                    return "Note"
                return sp_type
        return SharePointConstants.FALLBACK_TYPE

    def get_list_items(self, list_title, records_limit=-1):
        list_id = self._resolve_list_id(list_title)
        url = "{}/items".format(self._get_list_url(list_id))
        graph_params = {
            "$expand": f"fields($select={','.join(self.column_entity_property_name.keys())})",
            "$top": "5000" if records_limit < 1 else str(records_limit)
        }
        record_count = 0
        while url:
            response = self.session.get(url, params=graph_params)
            self.assert_response_ok(response, calling_method="get_list_items")
            json_response = response.json()
            for item in json_response.get("value", []):
                if records_limit > 0 and record_count >= records_limit:
                    return
                record_count += 1
                yield item.get("fields", {})
            url = json_response.get("@odata.nextLink")
            graph_params = None  # nextLink already encodes all query params

    def create_list(self, list_name):
        url = self._get_list_url()
        body = {
            "displayName": list_name,
            "list": {"template": "genericList"}
        }
        response = self.session.post(url, json=body)
        self.assert_response_ok(response, calling_method="create_list")
        json_response = response.json()

        list_id = json_response.get("id", "")
        display_name = json_response.get("displayName", list_name)
        name = json_response.get("name", display_name)

        self._list_id_cache[list_name] = list_id
        self._list_id_cache[name] = list_id

        return {
            "EntityTypeName": name,
            "ListItemEntityTypeFullName": "SP.Data.{}ListItem".format(name),
            "Id": list_id
        }

    def recycle_list(self, list_name):
        try:
            list_id = self._resolve_list_id(list_name)
        except SharePointClientError:
            return None
        url = self._get_list_url(list_id)
        response = self.session.delete(url)
        self._list_id_cache.pop(list_name, None)
        return response

    def get_list_metadata(self, list_name):
        list_id = self._resolve_list_id(list_name)
        url = self._get_list_url(list_id)
        response = self.session.get(url)
        self.assert_response_ok(response, calling_method="get_list_metadata")
        json_response = response.json()

        name = json_response.get("name", list_name)
        self._list_id_cache[name] = list_id

        return {
            "EntityTypeName": name,
            "ListItemEntityTypeFullName": "SP.Data.{}ListItem".format(name),
            "Id": json_response.get("id", ""),
        }

    def get_web_name(self, created_list):
        return created_list.get("EntityTypeName", "")

    def create_custom_field_via_id(self, list_id, field_title, field_type=None):
        field_type = SharePointConstants.FALLBACK_TYPE if field_type is None else field_type
        url = "{}/columns".format(self._get_list_url(list_id))
        body = {
            "name": field_title,
            "displayName": field_title,
        }
        type_config = SharePointConstants.GRAPH_COLUMN_TYPE_MAP.get(field_type, {"text": {}})
        body.update(type_config)
        response = self.session.post(url, json=body)
        self.assert_response_ok(response, calling_method="create_custom_field_via_id")
        return response

    def update_column_type(self, list_id, field, column_name, new_field_type="SP.FieldMultiLineText"):
        logger.info("updating field {}/{} to type {}".format(field, column_name, new_field_type))
        if not new_field_type:
            return None

        url = "{}/columns?$select=id,name,displayName".format(self._get_list_url(list_id))
        response = self.session.get(url)
        self.assert_response_ok(response, calling_method="update_column_type:get_columns")
        columns = response.json().get("value", [])

        column_id = None
        for col in columns:
            if col.get("name") == field or col.get("displayName") == column_name:
                column_id = col.get("id")
                break

        if not column_id:
            raise SharePointClientError("Column '{}' not found in list".format(field))

        patch_url = "{}/columns/{}".format(self._get_list_url(list_id), column_id)
        body = {"text": {"allowMultipleLines": True}}
        response = self.session.patch(patch_url, json=body)
        self.assert_response_ok(response, calling_method="update_column_type")
        return response

    def add_column_to_list_default_view(self, column_name, list_name):
        # No-op for Graph API: columns are visible by default
        return None

    def get_list_default_view(self, list_name):
        return []

    def get_view_id(self, list_title, view_title):
        if not list_title:
            return None
        logger.warning("List view filtering is not supported with Graph API. View '{}' will be ignored.".format(view_title))
        return None

    # ---- Batch operations ----

    def get_add_list_item_kwargs(self, list_title, item):
        if list_title in self._list_id_cache:
            list_id = self._list_id_cache[list_title]
        else:
            list_id = self._resolve_list_id(list_title)

        return {
            "method": "POST",
            "url": "/sites/{}/lists/{}/items".format(self.site_id, list_id),
            "body": {"fields": item},
            "headers": {"Content-Type": "application/json"}
        }

    def process_batch(self, kwargs_array):
        all_responses = []
        for chunk_start in range(0, len(kwargs_array), SharePointConstants.GRAPH_BATCH_LIMIT):
            chunk = kwargs_array[chunk_start:chunk_start + SharePointConstants.GRAPH_BATCH_LIMIT]
            batch_requests = []
            for idx, kwargs in enumerate(chunk):
                batch_request = {
                    "id": str(chunk_start + idx),
                    "method": kwargs["method"],
                    "url": kwargs["url"],
                    "headers": kwargs["headers"]
                }
                if "body" in kwargs:
                    batch_request["body"] = kwargs["body"]
                batch_requests.append(batch_request)

            batch_body = {"requests": batch_requests}
            successful_post = False
            attempt_number = 0
            response = None
            while not successful_post and attempt_number <= SharePointConstants.MAX_RETRIES:
                try:
                    attempt_number += 1
                    logger.info("Posting batch of {} items (chunk starting at {})".format(len(chunk), chunk_start))
                    response = self.session.post(
                        SharePointConstants.GRAPH_BATCH_URL,
                        dku_rs_off=True,
                        json=batch_body
                    )
                    logger.info("Batch post status: {}".format(response.status_code))
                    if response.status_code >= 400:
                        logger.error("Response={}".format(response.content))
                    successful_post = True
                except requests.exceptions.Timeout as err:
                    logger.error("Timeout error:{}".format(err))
                    raise SharePointClientError("Timeout error: {}".format(err))
                except Exception as err:
                    logger.warning("ERROR:{}".format(err))
                    logger.warning("on attempt #{}".format(attempt_number))
                    if attempt_number == SharePointConstants.MAX_RETRIES:
                        raise SharePointClientError("Error in batch processing on attempt #{}: {}".format(attempt_number, err))
                    time.sleep(SharePointConstants.WAIT_TIME_BEFORE_RETRY_SEC)

            if response is not None:
                self._log_batch_errors(response)
                all_responses.append(response)

        return all_responses[-1] if all_responses else None

    def _log_batch_errors(self, response):
        logger.info("Batch error analysis")
        try:
            json_response = response.json()
            responses = json_response.get("responses", [])
            has_errors = False
            for resp in responses:
                status = resp.get("status", 0)
                if status >= 400:
                    has_errors = True
                    body = resp.get("body", {})
                    error = body.get("error", {})
                    logger.warning("Batch item {} error {}: {}".format(
                        resp.get("id"), status, error.get("message", "Unknown error")))
            if has_errors:
                if self.number_dumped_logs == 0:
                    logger.warning("response.content={}".format(response.content))
                else:
                    logger.warning("Batch error analysis KO ({})".format(self.number_dumped_logs))
                self.number_dumped_logs += 1
            else:
                logger.info("Batch error analysis OK")
        except Exception as err:
            logger.warning("Error parsing batch response: {}".format(err))


    # ---- Authentication methods ----

    def _get_site_app_access_token(self):
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority="https://login.microsoftonline.com/{}".format(self.tenant_id),
            client_credential=self.client_secret,
        )
        result = app.acquire_token_for_client(scopes=[SharePointConstants.GRAPH_API_DEFAULT_SCOPE])
        access_token = result.get("access_token")
        if not access_token:
            error_description = result.get("error_description", "Unknown error")
            raise SharePointClientError("Failed to acquire token for site app: {}".format(error_description))
        return access_token

    def _get_certificate_app_access_token(self):
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority="https://login.microsoftonline.com/{}".format(self.tenant_id),
            client_credential={
                "thumbprint": self.client_certificate_thumbprint,
                "private_key": self.client_certificate,
                "passphrase": self.passphrase,
            },
        )
        result = app.acquire_token_for_client(scopes=[SharePointConstants.GRAPH_API_DEFAULT_SCOPE])
        access_token = result.get("access_token")
        if not access_token:
            error_description = result.get("error_description", "Unknown error")
            raise SharePointClientError("Failed to acquire token for certificate app: {}".format(error_description))
        return access_token

    def _get_username_password_access_token(self, username, password):
        authority_url = 'https://login.microsoftonline.com/{}'.format(self.tenant_id)
        app = msal.PublicClientApplication(
            authority=authority_url,
            client_id=self.client_id,
            client_credential=None
        )
        result = app.acquire_token_by_username_password(
            username,
            password,
            scopes=[SharePointConstants.GRAPH_API_DEFAULT_SCOPE]
        )
        access_token = result.get("access_token")
        error_description = result.get("error_description")
        if error_description:
            logger.error("Dumping: {}".format(result))
            raise SharePointClientError("Error: {}".format(error_description))
        return access_token

    # ---- Schema / read operations ----

    def get_writer(self, dataset_schema, dataset_partitioning,
                   partition_id, max_workers, batch_size, write_mode):
        return SharePointListWriter(
            self.config,
            self,
            dataset_schema,
            dataset_partitioning,
            partition_id,
            max_workers=max_workers,
            batch_size=batch_size,
            write_mode=write_mode,
            allow_string_recasting=self.allow_string_recasting
        )

    def get_read_schema(self, display_metadata=False, metadata_to_retrieve=[], write_mode=None):
        logger.info('get_read_schema')
        sharepoint_columns = self.get_list_fields(self.sharepoint_list_title)
        dss_columns = []
        self.column_ids = {}
        self.column_names = {}
        self.column_entity_property_name = {}
        self.columns_to_format = []
        for column in sharepoint_columns:
            if column[SharePointConstants.INTERNAL_NAME] in SharePointConstants.COLUMNS_TO_IGNORE_BY_INTERNAL_NAME:
                continue
            logger.info("get_read_schema:{}/{}/{}/{}/{}/{}".format(
                column[SharePointConstants.TITLE_COLUMN],
                column[SharePointConstants.TYPE_AS_STRING],
                column[SharePointConstants.STATIC_NAME],
                column[SharePointConstants.INTERNAL_NAME],
                column[SharePointConstants.ENTITY_PROPERTY_NAME],
                self.is_column_displayable(column, display_metadata, metadata_to_retrieve)
            ))
            if self.is_column_displayable(column, display_metadata, metadata_to_retrieve):
                sharepoint_type = get_dss_type(column[SharePointConstants.TYPE_AS_STRING])
                self.column_sharepoint_type[column[SharePointConstants.STATIC_NAME]] = column[SharePointConstants.TYPE_AS_STRING]
                if sharepoint_type is not None:
                    dss_columns.append({
                        SharePointConstants.NAME_COLUMN: column[SharePointConstants.TITLE_COLUMN],
                        SharePointConstants.TYPE_COLUMN: sharepoint_type
                    })
                    self.column_ids[column[SharePointConstants.STATIC_NAME]] = sharepoint_type
                    self.column_names[column[SharePointConstants.STATIC_NAME]] = column[SharePointConstants.TITLE_COLUMN]
                    self.column_entity_property_name[column[SharePointConstants.STATIC_NAME]] = column[SharePointConstants.ENTITY_PROPERTY_NAME]
                    self.dss_column_name[column[SharePointConstants.STATIC_NAME]] = column[SharePointConstants.TITLE_COLUMN]
                    self.dss_column_name[column[SharePointConstants.ENTITY_PROPERTY_NAME]] = column[SharePointConstants.TITLE_COLUMN]
                if sharepoint_type == "date":
                    self.columns_to_format.append((column[SharePointConstants.STATIC_NAME], sharepoint_type))
                if column[SharePointConstants.TYPE_AS_STRING] == SharePointConstants.TYPE_NOTE:
                    if write_mode == SharePointConstants.WRITE_MODE_CREATE:
                        self.columns_to_format.append((column[SharePointConstants.COLUMN_TITLE], SharePointConstants.TYPE_NOTE))
                    else:
                        self.columns_to_format.append((column[SharePointConstants.STATIC_NAME], SharePointConstants.TYPE_NOTE))
        logger.info("get_read_schema: Schema updated with {}".format(dss_columns))
        return {
            SharePointConstants.COLUMNS: dss_columns
        }

    def is_column_displayable(self, column, display_metadata=False, metadata_to_retrieve=[]):
        if display_metadata and (column[SharePointConstants.STATIC_NAME] in metadata_to_retrieve):
            return True
        return (not column[SharePointConstants.HIDDEN_COLUMN])

    # ---- Error handling ----

    @staticmethod
    def assert_login_details(required_keys, login_details):
        if login_details is None or login_details == {}:
            raise SharePointClientError("Login details are empty")
        for key in required_keys:
            if key not in login_details.keys():
                raise SharePointClientError(required_keys[key])

    def assert_response_ok(self, response, no_json=False, calling_method=""):
        status_code = response.status_code
        if status_code >= 400:
            logger.error("Error {} in method {}".format(status_code, calling_method))
            logger.error("when calling {}".format(response.url))
            logger.error("dump={}".format(response.content))
            enriched_error_message = self.get_enriched_error_message(response)
            if enriched_error_message is not None:
                raise SharePointClientError("Error {} ({}): {}".format(status_code, calling_method, enriched_error_message))
            if status_code == 400:
                raise SharePointClientError("({}){}".format(calling_method, response.text))
            if status_code == 404:
                raise SharePointClientError("Not found. Please check tenant, site type or site name. ({})".format(calling_method))
            if status_code == 403:
                raise SharePointClientError("403 Forbidden. Please check your account credentials and API permissions. ({})".format(calling_method))
            raise SharePointClientError("Error {} ({})".format(status_code, calling_method))
        if not no_json:
            self.assert_no_error_in_json(response, calling_method=calling_method)

    @staticmethod
    def get_enriched_error_message(response):
        try:
            json_response = response.json()
            error_message = get_value_from_paths(
                json_response,
                [
                    ["error", "message"],
                    ["error_description"],
                    ["error", "message", "value"],
                    ["odata.error", "message", "value"]
                ]
            )
            if error_message:
                return "{}".format(error_message)
        except Exception as error:
            logger.info("Error trying to extract error message: {}".format(error))
            logger.info("Response.content={}".format(response.content))
            return None
        return None

    @staticmethod
    def assert_no_error_in_json(response, calling_method=""):
        if len(response.content) == 0:
            raise SharePointClientError("Empty response from SharePoint ({}). Please check user credentials.".format(calling_method))
        json_response = response.json()
        if "error" in json_response:
            error = json_response["error"]
            message = error.get("message", str(error))
            raise SharePointClientError("Error ({}): {}".format(calling_method, message))

    @staticmethod
    def get_random_guid():
        return str(uuid.uuid4())

    @staticmethod
    def escape_path(path):
        return path.replace("'", "''")


class GraphSession(requests.Session):
    """HTTP session wrapper that adds Bearer token auth for Microsoft Graph API."""

    def __init__(self, access_token):
        super().__init__()
        self.headers.update({
            "Authorization": "Bearer {}".format(access_token),
            "Accept": "application/json",
        })

    def get(self, url, **kwargs):
        kwargs.setdefault("timeout", SharePointConstants.TIMEOUT_SEC)
        return super().get(url, **kwargs)

    def post(self, url, **kwargs):
        kwargs.setdefault("timeout", SharePointConstants.TIMEOUT_SEC)
        return super().post(url, **kwargs)

    def put(self, url, **kwargs):
        kwargs.setdefault("timeout", SharePointConstants.TIMEOUT_SEC)
        return super().put(url, **kwargs)

    def patch(self, url, **kwargs):
        kwargs.setdefault("timeout", SharePointConstants.TIMEOUT_SEC)
        return super().patch(url, **kwargs)

    def delete(self, url, **kwargs):
        kwargs.setdefault("timeout", SharePointConstants.TIMEOUT_SEC)
        return super().delete(url, **kwargs)

    def request(self, method, url, **kwargs):
        kwargs.setdefault("timeout", SharePointConstants.TIMEOUT_SEC)
        return super().request(method, url, **kwargs)


class SuppressFilter(logging.Filter):
    def filter(self, record):
        return 'Failed to parse headers' not in record.getMessage()
