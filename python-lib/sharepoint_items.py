import os.path

from datetime import datetime
from common import get_lnt_path, get_rel_path


# Key used by SharePointClient methods to wrap Graph driveItem arrays
ITEMS_KEY = "items"


def loop_sharepoint_items(items):
    if ITEMS_KEY not in items or not items[ITEMS_KEY]:
        return
    for item in items[ITEMS_KEY]:
        yield item


def extract_item_from(item_name, items):
    for item in loop_sharepoint_items(items):
        if item and "name" in item and item["name"] == item_name:
            return item
    return None


def has_sharepoint_items(items):
    if ITEMS_KEY not in items or not items[ITEMS_KEY]:
        return False
    return len(items[ITEMS_KEY]) > 0


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
