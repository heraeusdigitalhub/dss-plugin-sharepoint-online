import dataiku
from dataiku.customrecipe import get_input_names_for_role, get_output_names_for_role, get_recipe_config
from safe_logger import SafeLogger
from dss_constants import DSSConstants
from sharepoint_client import SharePointClient
from common import get_lnt_path

logger = SafeLogger("sharepoint-online plugin", DSSConstants.SECRET_PARAMETERS_KEYS)
logger.info('SharePoint Online list folder details recipe v{}'.format(DSSConstants.PLUGIN_VERSION))

config = get_recipe_config()
output_dataset = dataiku.Dataset(get_output_names_for_role('output_dataset')[0])
folder_path = get_lnt_path(config.get("folder_path") or "")
logger.info('init:auth_type={}, folder_path={}'.format(config.get("auth_type"), folder_path))


def assert_same_config_as_folder(recipe_config, folder_name):
    """Fail if the recipe connection settings differ from those of the input managed folder."""
    folder = dataiku.Folder(folder_name)
    settings = dataiku.api_client().get_project(folder.project_key).get_managed_folder(folder.get_id()).get_settings().get_raw()
    folder_type = settings.get("type", "")
    if "sharepoint-online-graph" not in folder_type:
        raise ValueError("The input folder is not a SharePoint Online folder (type: {})".format(folder_type))
    params = settings.get("params", {})
    folder_config = params.get("config", {})

    mismatches = []
    if recipe_config.get("auth_type") != folder_config.get("auth_type"):
        mismatches.append("auth_type: recipe={}, folder={}".format(recipe_config.get("auth_type"), folder_config.get("auth_type")))
    for key in ["sharepoint_oauth", "site_app_permissions", "app_certificate", "app_username_password"]:
        recipe_preset, folder_preset = recipe_config.get(key), folder_config.get(key)
        # the recipe gets its preset resolved (no name), the folder only stores a reference: compare names only if both have one
        if isinstance(recipe_preset, dict) and isinstance(folder_preset, dict) and recipe_preset.get("name") and folder_preset.get("name") and recipe_preset.get("name") != folder_preset.get("name"):
            mismatches.append("{} preset: recipe={}, folder={}".format(key, recipe_preset.get("name"), folder_preset.get("name")))
    for key in ["sharepoint_site_overwrite", "sharepoint_root_overwrite"]:
        if (recipe_config.get(key) or "") != (folder_config.get(key) or ""):
            mismatches.append("{}: recipe={!r}, folder={!r}".format(key, recipe_config.get(key) or "", folder_config.get(key) or ""))
    if mismatches:
        raise ValueError("The recipe connection settings differ from the input folder's: " + "; ".join(mismatches))


assert_same_config_as_folder(config, get_input_names_for_role('input_folder')[0])

client = SharePointClient(config)

output_dataset.write_schema([
    {"name": "path", "type": "string"},
    {"name": "size", "type": "bigint"},
    {"name": "created", "type": "string"},
    {"name": "last_modified", "type": "string"},
    {"name": "quickxorhash", "type": "string"},
    {"name": "web_url", "type": "string"},
])
with output_dataset.get_writer() as writer:
    for item in client.list_enumerate_items(folder_path):
        if item.get("folder"):
            continue
        writer.write_row_dict({
            "path": item.get("name"),
            "size": item.get("size"),
            "created": item.get("createdDateTime"),
            "last_modified": item.get("lastModifiedDateTime"),
            "quickxorhash": item.get("file_quickxorhash"),
            "web_url": item.get("webUrl"),
        })
