import os
import shutil

from common import (
    assert_no_percent_in_path,
    assert_valid_sharepoint_path,
    get_lnt_path,
    get_rel_path,
)
from dataiku.fsprovider import FSProvider
from dss_constants import DSSConstants
from safe_logger import SafeLogger
from sharepoint_client import SharePointClient
from sharepoint_items import assert_path_is_not_root, get_name, build_dss_item

try:
    from BytesIO import BytesIO  # for Python 2
except ImportError:
    from io import BytesIO  # for Python 3

logger = SafeLogger("sharepoint-online plugin", DSSConstants.SECRET_PARAMETERS_KEYS)


# based on https://docs.microsoft.com/fr-fr/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest

class SharePointFSProvider(FSProvider):
    def __init__(self, root, config, plugin_config):
        """
        :param root: the root path for this provider
        :param config: the dict of the configuration of the object
        :param plugin_config: contains the plugin settings
        """
        if len(root) > 0 and root[0] == '/':
            root = root[1:]
        self.root = root
        self.provider_root = "/"
        logger.info('SharePoint Online plugin fs v{}'.format(DSSConstants.PLUGIN_VERSION))
        logger.info('init:root={}'.format(self.root))
        root_name_overwrite_legacy_mode = plugin_config.get("root_name_overwrite_legacy_mode", False)
        self.client = SharePointClient(
            config,
            root_name_overwrite_legacy_mode=root_name_overwrite_legacy_mode
        )

    # util methods
    def get_full_path(self, path):
        path_elts = [self.provider_root, get_rel_path(self.root), get_rel_path(path)]
        path_elts = [e for e in path_elts if len(e) > 0]
        return os.path.join(*path_elts)

    def close(self):
        logger.info('close')

    def stat(self, path):
        assert_valid_sharepoint_path(path)
        full_path = get_lnt_path(self.get_full_path(path))
        logger.info('stat:path="{}", full_path="{}"'.format(path, full_path))
        item = self.client.get_item(full_path)
        if item:
            return build_dss_item(path, item)
        return None

    def set_last_modified(self, path, last_modified):
        full_path = self.get_full_path(path)
        logger.info('set_last_modified: path="{}", full_path="{}"'.format(path, full_path))
        return False
    

    def browse(self, path):
        assert_valid_sharepoint_path(path)
        path = get_rel_path(path)
        full_path = get_lnt_path(self.get_full_path(path))
        logger.info('browse:path="{}", full_path="{}"'.format(path, full_path))

        files, folders = self.client.list_folder_items(full_path)
        children = []

        for file in files:
            children.append(build_dss_item(path, file))
        for folder in folders:
            children.append(build_dss_item(path, folder))

        if len(children) > 0:
            return {
                DSSConstants.FULL_PATH: get_lnt_path(path),
                DSSConstants.EXISTS: True,
                DSSConstants.DIRECTORY: True,
                DSSConstants.CHILDREN: children
            }
        
        item = self.client.get_item(full_path)
        if item:
            return {
                **build_dss_item(path, item),
                DSSConstants.FULL_PATH: get_lnt_path(path),
            }
        else:
            return {
                DSSConstants.FULL_PATH: get_lnt_path(path),
                DSSConstants.EXISTS: False
            }

    def enumerate(self, path, first_non_empty):
        assert_valid_sharepoint_path(path)
        full_path = get_lnt_path(self.get_full_path(path))
        logger.info('enumerate:path="{}",fullpath="{}", first_non_empty="{}"'.format(path, full_path, first_non_empty))
        path_to_item, item_name = os.path.split(full_path)
        item = self.client.get_item(full_path)
        if item and "file" in item:
            return [build_dss_item(path, item)]
        ret = self.list_recursive(path, full_path, first_non_empty)
        return ret

    def list_recursive(self, path, full_path, first_non_empty):
        paths = []
        files, folders = self.client.list_folder_items(full_path)
        for file in files:
            paths.append(build_dss_item(path, file))
            if first_non_empty:
                return paths
        for folder in folders:
            paths.extend(
                self.list_recursive(
                    get_lnt_path(os.path.join(path, get_name(folder))),
                    get_lnt_path(os.path.join(full_path, get_name(folder))),
                    first_non_empty
                )
            )
        return paths

    def delete_recursive(self, path):
        assert_valid_sharepoint_path(path)
        full_path = self.get_full_path(path)
        logger.info('delete_recursive:path={},fullpath={}'.format(path, full_path))
        assert_path_is_not_root(full_path)
        item = self.client.get_item(get_lnt_path(full_path))
        if item:
            self.client.recycle(get_lnt_path(full_path))
            return 1
        return 0


    def move(self, from_path, to_path):
        assert_valid_sharepoint_path(from_path)
        assert_no_percent_in_path(from_path)
        assert_valid_sharepoint_path(to_path)
        assert_no_percent_in_path(to_path)
        full_from_path = self.get_full_path(from_path)
        full_to_path = self.get_full_path(to_path)
        logger.info('move:from={},to={}'.format(full_from_path, full_to_path))

        self.client.move_file(full_from_path, full_to_path)
        return True

    def read(self, path, stream, limit):
        assert_valid_sharepoint_path(path)
        full_path = self.get_full_path(path)
        logger.info('read:full_path={}'.format(full_path))
        response = self.client.get_file_content(full_path)
        bio = BytesIO(response.content)
        shutil.copyfileobj(bio, stream)

    def write(self, path, stream):
        assert_valid_sharepoint_path(path)
        full_path = self.get_full_path(path)
        logger.info('write:path="{}", full_path="{}"'.format(path, full_path))
        bio = BytesIO()
        shutil.copyfileobj(stream, bio)
        bio.seek(0)
        self.client.create_path(full_path)
        response = self.client.write_file_content(full_path, bio)
        logger.info("write:response={}".format(response))
