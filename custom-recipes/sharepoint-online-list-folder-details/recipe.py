import re
from pathlib import Path
import dataiku
from dataiku.customrecipe import get_input_names_for_role, get_output_names_for_role, get_recipe_config
from safe_logger import SafeLogger
from dss_constants import DSSConstants
from sharepoint_client import SharePointClient
from sharepoint_lists import sharepoint_to_dss_date
from common import get_lnt_path

logger = SafeLogger("sharepoint-online plugin", DSSConstants.SECRET_PARAMETERS_KEYS)
logger.info('SharePoint Online list folder details recipe v{}'.format(DSSConstants.PLUGIN_VERSION))

config = get_recipe_config()
output_dataset = dataiku.Dataset(get_output_names_for_role('output_dataset')[0])
folder_path = get_lnt_path(config.get("folder_path") or "")
logger.info('init:auth_type={}, folder_path={}'.format(config.get("auth_type"), folder_path))


AUTH_TYPE_TO_PRESET = {
    "oauth": "sharepoint_oauth",
    "site-app-permissions": "site_app_permissions",
    "app-certificate": "app_certificate",
    "app-username-password": "app_username_password",
}
PRESET_TO_PARAMETER_SET = {
    "sharepoint_oauth": "oauth-login",
    "site_app_permissions": "site-app-permissions",
    "app_certificate": "app-certificate",
    "app_username_password": "app-username-password",
}


def get_preset_target(preset):
    """Return the normalized tenant / site / root a resolved preset points to."""
    preset = preset or {}
    return {
        "sharepoint_tenant": (preset.get("sharepoint_tenant") or "").strip().lower(),
        "sharepoint_site": (preset.get("sharepoint_site") or "").strip("/ ").lower(),
        "sharepoint_root": (preset.get("sharepoint_root") or "").strip("/ ").lower(),
    }


def resolve_folder_preset(folder_preset, preset_key):
    """The folder settings only store a reference to the preset: look up its content in the plugin settings."""
    if not isinstance(folder_preset, dict):
        raise ValueError("The input folder has no {} preset configured".format(preset_key))
    if folder_preset.get("mode") == "INLINE":
        return folder_preset.get("config") or folder_preset
    preset_name = folder_preset.get("name")
    if not preset_name:
        raise ValueError("Cannot determine the {} preset of the input folder".format(preset_key))
    plugin_settings = dataiku.api_client().get_plugin("sharepoint-online-graph").get_settings().get_raw()
    parameter_set_id = PRESET_TO_PARAMETER_SET[preset_key]
    for preset in plugin_settings.get("presets", []):
        if preset.get("name") == preset_name and parameter_set_id in (preset.get("type") or parameter_set_id):
            return preset.get("config", {})
    raise ValueError("Preset '{}' of the input folder could not be found in the plugin settings".format(preset_name))


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
    preset_key = AUTH_TYPE_TO_PRESET.get(recipe_config.get("auth_type"))
    if preset_key:
        recipe_target = get_preset_target(recipe_config.get(preset_key))
        folder_target = get_preset_target(resolve_folder_preset(folder_config.get(preset_key), preset_key))
        for key in ["sharepoint_tenant", "sharepoint_site", "sharepoint_root"]:
            if recipe_target.get(key) != folder_target.get(key):
                mismatches.append("{} preset {}: recipe={!r}, folder={!r}".format(preset_key, key, recipe_target.get(key), folder_target.get(key)))
    for key in ["sharepoint_site_overwrite", "sharepoint_root_overwrite"]:
        if (recipe_config.get(key) or "") != (folder_config.get(key) or ""):
            mismatches.append("{}: recipe={!r}, folder={!r}".format(key, recipe_config.get(key) or "", folder_config.get(key) or ""))
    # the folder root is stored next to the provider config, not inside it
    folder_root = get_lnt_path(params.get("path") or "")
    recipe_root = get_lnt_path(recipe_config.get("folder_path") or "")
    if recipe_root.lower() != folder_root.lower():
        mismatches.append("folder_path: recipe={!r}, folder={!r}".format(recipe_root, folder_root))
    if mismatches:
        raise ValueError("The recipe connection settings differ from the input folder's: " + "; ".join(mismatches))


assert_same_config_as_folder(config, get_input_names_for_role('input_folder')[0])

client = SharePointClient(config)

RE_FILENAME = re.compile(r'\.(?=[a-zA-Z0-9]+\.[a-zA-Z]+$|[a-zA-Z0-9]+$)')

output_dataset.write_schema([
    # DSS list_folder_contents default columns - start
    {"name": "path", "type": "string"},
    {"name": "basename", "type": "string"},
    {"name": "extension", "type": "string"},
    {"name": "last_modified", "type": "date"},
    {"name": "size", "type": "bigint"},
    # DSS list_folder_contents default columns - end
    {"name": "created", "type": "date"},
    {"name": "quickxorhash", "type": "string"},
    {"name": "web_url", "type": "string"},
])
with output_dataset.get_writer() as writer:
    for item in client.list_enumerate_items(folder_path):
        if item.get("folder"):
            continue
        # special basename/ext handling to align with Dataiku's list_folder_contents recipe
        item_as_path = Path(item.get("name"))
        parts = RE_FILENAME.split(item_as_path.name, maxsplit=1)
        if len(parts) == 2:
            basename, extension = parts
        else:
            basename, extension = item_as_path.stem, item_as_path.suffix.lstrip(".")
        writer.write_row_dict({
            "path": item.get("name"),
            "basename": basename,
            "extension": extension,
            "last_modified": sharepoint_to_dss_date(item.get("lastModifiedDateTime")),
            "size": item.get("size"),
            "created": sharepoint_to_dss_date(item.get("createdDateTime")),
            "quickxorhash": item.get("file_quickxorhash"),
            "web_url": item.get("webUrl"),
        })
