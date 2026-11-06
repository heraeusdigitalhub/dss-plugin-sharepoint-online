import dataiku
from dataiku.customrecipe import get_output_names_for_role, get_recipe_config
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
