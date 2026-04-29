import os.path

from datetime import datetime
from common import get_lnt_path, get_rel_path
from dss_constants import DSSConstants


def build_dss_item(path: str, item: dict):
    path = path.replace(item["name"], "")
    lnt_path = get_lnt_path(os.path.join(path, item["name"]))
    return {
        DSSConstants.PATH: lnt_path,  # for enumerate + stat
        DSSConstants.FULL_PATH: lnt_path,  # for browse
        DSSConstants.EXISTS: True,  # for all
        DSSConstants.DIRECTORY: "folder" in item,  # for browse
        DSSConstants.IS_DIRECTORY: "folder" in item,  # for stat
        DSSConstants.SIZE: get_size(item),  # for browse
        DSSConstants.LAST_MODIFIED: get_last_modified(item), # for browse
    }


def get_last_modified(item):
    if "lastModifiedDateTime" in item:
        return int(format_date(item["lastModifiedDateTime"]))
    return None


def format_date(date):
    if date is not None:
        # Graph API returns ISO 8601 dates like "2024-01-15T10:30:00Z"
        utc_time = datetime.strptime(date, "%Y-%m-%dT%H:%M:%SZ")
        epoch_time = (utc_time - datetime(1970, 1, 1)).total_seconds()
        return int(epoch_time) * 1000
    else:
        return None


def get_size(item):
    return int(item.get("size", 0))


def get_name(item):
    return item.get("name")


def assert_path_is_not_root(path):
    if path is None:
        raise ValueError("Cannot delete root path")
    path = get_rel_path(path)
    if path == "" or path == "/":
        raise ValueError("Cannot delete root path")
