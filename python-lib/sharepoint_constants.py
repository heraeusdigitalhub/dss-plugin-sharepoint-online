class SharePointConstants(object):
    # Certificate key markers
    CLEAR_KEY_END = "-----END PRIVATE KEY-----"
    CLEAR_KEY_START = "-----BEGIN PRIVATE KEY-----"
    ENCRYPTED_KEY_END = "-----END ENCRYPTED PRIVATE KEY-----"
    ENCRYPTED_KEY_START = "-----BEGIN ENCRYPTED PRIVATE KEY-----"

    # Column/schema constants
    COLUMNS = 'columns'
    COLUMN_TITLE = 'Title'
    COMMENT_COLUMN = 'comment'
    ENTITY_PROPERTY_NAME = 'EntityPropertyName'
    HIDDEN_COLUMN = 'Hidden'
    INTERNAL_NAME = 'InternalName'
    NAME_COLUMN = 'name'
    READ_ONLY_FIELD = 'ReadOnlyField'
    STATIC_NAME = 'StaticName'
    TITLE_COLUMN = 'Title'
    TYPE_AS_STRING = 'TypeAsString'
    TYPE_COLUMN = 'type'
    TYPE_NOTE = 'Note'
    VALUE = 'value'
    FALLBACK_TYPE = "Text"

    # Columns to ingore
    COLUMNS_TO_IGNORE_BY_INTERNAL_NAME = [
        "LinkTitle",
        "_ColorTag",
        "ComplianceAssetId",
        "ContentType",
        "_UIVersionString",
        "Attachments",
        "Edit",
        "LinkTitleNoMenu",
        "DocIcon",
        "ItemChildCount",
        "FolderChildCount",
        "_ComplianceFlags",
        "_ComplianceTag",
        "_ComplianceTagWrittenTime",
        "_ComplianceTagUserId",
        "_IsRecord",
        "AppAuthor",
        "AppEditor"
    ]

    # DriveItem property names (Graph API format)
    NAME = 'name'
    LENGTH = 'size'
    TIME_LAST_MODIFIED = 'lastModifiedDateTime'

    # Date formats
    DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
    TIME_FORMAT = "%Y-%m-%dT%H:%M:%SZ"

    # Error/response keys
    ERROR_CONTAINER = 'error'
    MESSAGE = 'message'

    # File/folder constants
    FORBIDDEN_PATH_CHARS = ['"', '*', ':', '<', '>', '?', '\\', '|']

    # Expendable lookup fields
    EXPENDABLES_FIELDS = {"Author": "Title", "Editor": "Title"}

    # SharePoint type mappings (SharePoint type -> DSS type)
    TYPES = {
        "Text": "string",
        "Number": "double",
        "DateTime": "date",
        "Boolean": "string",
        "URL": "object",
        "Location": "object",
        "Computed": None,
        "Attachments": None,
        "Calculated": "string",
        "User": "array",
        "Thumbnail": "object",
        "Note": "string"
    }

    # Write mode
    WRITE_MODE_CREATE = "create"

    # Retry / timeout
    DEFAULT_WAIT_BEFORE_RETRY = 60
    MAX_RETRIES = 5
    WAIT_TIME_BEFORE_RETRY_SEC = 2
    TIMEOUT_SEC = 300

    # --- Microsoft Graph API constants ---

    GRAPH_API_BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_API_DEFAULT_SCOPE = "https://graph.microsoft.com/.default"

    # Graph batch settings
    GRAPH_BATCH_URL = "https://graph.microsoft.com/v1.0/$batch"
    GRAPH_BATCH_LIMIT = 20

    # Graph upload thresholds
    GRAPH_MAX_SIMPLE_UPLOAD_SIZE = 4 * 1024 * 1024        # 4MB
    GRAPH_UPLOAD_CHUNK_SIZE = 10 * 320 * 1024              # 3.2MB (must be 320KB multiple)

    # Graph column type map: SP field type -> Graph columnDefinition typed property
    GRAPH_COLUMN_TYPE_MAP = {
        "Text": {"text": {"allowMultipleLines": False, "maxLength": 255}},
        "Note": {"text": {"allowMultipleLines": True}},
        "Number": {"number": {}},
        "Integer": {"number": {"decimalPlaces": "none"}},
        "DateTime": {"dateTime": {}},
        "Boolean": {"boolean": {}},
        "URL": {"hyperlinkOrPicture": {}},
    }

    # Reverse map: Graph column type key -> SP TypeAsString
    GRAPH_TO_SP_TYPE_MAP = {
        "text": "Text",
        "number": "Number",
        "dateTime": "DateTime",
        "boolean": "Boolean",
        "hyperlinkOrPicture": "URL",
        "calculated": "Calculated",
        "lookup": "User",
        "thumbnail": "Thumbnail",
        "personOrGroup": "User",
        "choice": "Text",
        "currency": "Number",
    }
