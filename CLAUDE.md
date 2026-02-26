# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Dataiku DSS plugin providing read/write connectivity to SharePoint Online. Plugin ID: `sharepoint-online-graph`, version defined in `plugin.json`. Written in Python 3 (3.5–3.11 compatible). Runtime dependencies: `msal==1.23.0`, `requests`. The plugin uses the **Microsoft Graph API v1.0** exclusively.

## Build & Test Commands

```bash
# Package plugin into distributable .zip in dist/
make plugin

# Run unit tests (creates venv, installs deps, runs pytest)
make unit-tests

# Run integration tests (requires DSS instance)
make integration-tests

# Run both
make tests

# Run a single unit test
PYTHONPATH=$PWD/python-lib pytest tests/python/unit/test_common.py -k "test_name"

# Clean dist
make dist-clean
```

The Makefile's test targets create a fresh `env/` virtualenv each run. Unit tests use `pytest` with `allure-pytest` for reporting. PYTHONPATH must include `python-lib/` for imports to resolve.

## Architecture

### Plugin Entry Points (Dataiku DSS framework)

- **Lists Connector** (`python-connectors/sharepoint-online_lists/connector.py`): Inherits `dataiku.connector.Connector`. Implements `get_read_schema()`, `generate_rows()`, and `get_writer()` for reading/writing SharePoint lists as DSS datasets. Configuration in `connector.json`.
- **FS Provider** (`python-fs-providers/sharepoint-online_shared-documents/fs-provider.py`): Inherits `dataiku.fsprovider.FSProvider`. Provides file/folder browsing, upload, and download against SharePoint Shared Documents. Configuration in `fs-provider.json`.
- **Append Recipe** (`custom-recipes/sharepoint-online-append-list/recipe.py`): Custom DSS recipe that appends rows from a Dataiku dataset to an existing SharePoint list.

### Core Library (`python-lib/`)

- **`sharepoint_client.py`** — Central class `SharePointClient`. Handles all authentication flows, resolves Graph API site/drive IDs, and provides methods for list/file/folder operations. Also contains `GraphSession` (Bearer token HTTP wrapper) and `SharePointClientError`.
- **`sharepoint_lists.py`** — `SharePointListWriter` handles write operations: list creation, column creation, batch item insertion via `ThreadPoolExecutor` with configurable `max_workers` and `batch_size`. Includes DSS↔SharePoint type mapping (`get_dss_type`, `get_sharepoint_type`).
- **`sharepoint_items.py`** — Item-level utilities for converting Graph driveItem responses to DSS format (date→epoch, size, name).
- **`robust_session.py`** — `RobustSession` wraps HTTP calls with automatic retry on 429/503 status codes, respects `Retry-After` headers, and optionally resets sessions on 403 errors.
- **`common.py`** — Shared utilities: URL parsing, path validation (forbidden chars, 400-char limit), date format conversion, query string parsing.
- **`safe_logger.py`** — `SafeLogger` wraps Python logging, masking sensitive keys (auth tokens, passwords, certificates) in output. Masked values appear as `HASHED_SECRET:{type}:{length}`.
- **`dss_constants.py`** / **`sharepoint_constants.py`** — Constants for DSS config keys and SharePoint/Graph API values respectively.

### Authentication (4 active methods, configured via `parameter-sets/`)

Each method has its own parameter set definition in `parameter-sets/<name>/parameter-set.json`. All flows produce a Bearer token used by `GraphSession`.

1. **OAuth** (`oauth-login`) — Azure SSO, bearer token passed directly
2. **Site App Permissions** (`site-app-permissions`) — Client ID + secret, `client_credentials` grant via MSAL
3. **App Certificate** (`app-certificate`) — Certificate-based auth via MSAL `ConfidentialClientApplication`
4. **App Username/Password** (`app-username-password`) — MSAL `acquire_token_by_username_password`

> Note: `sharepoint-login` (sharepy-based) has been removed. It raises `SharePointClientError` if selected.

Auth type is selected via `config['auth_type']` and dispatched in `SharePointClient.__init__`.

### Graph API Resource Resolution

`SharePointClient.__init__` resolves and caches Graph API IDs after authentication:
- `site_id`: resolved via `_resolve_site_id()` from tenant + site path
- `drive_id`: resolved via `_resolve_drive_id()` from site_id + drive/library name
- `_drive_path_prefix`: non-empty when the library root is a subfolder within a drive
- `_list_id_cache`: per-list cache populated lazily by `_resolve_list_id()`

All Graph API calls use `https://graph.microsoft.com/v1.0` as base URL.

### Write Flow

Writes go through `SharePointListWriter` which: recycles existing list (on OVERWRITE mode) → creates new list → creates columns via Graph API column definitions → inserts items in batches using the Graph `$batch` endpoint (max 20 requests per batch). Batch processing lives in `SharePointClient.process_batch()`. Parallel batches are run via `ThreadPoolExecutor` with configurable `max_workers` (1–5).

### Key Conventions

- All imports in `python-lib/` use flat (non-package) imports — files are added to PYTHONPATH, not treated as a package.
- Column identity uses `StaticName` internally; display uses `Title`. The mapping is tracked in `SharePointClient.column_ids`, `column_names`, and `column_entity_property_name` dicts.
- Graph column type keys (`text`, `number`, `dateTime`, etc.) are mapped to SharePoint `TypeAsString` values via `SharePointConstants.GRAPH_TO_SP_TYPE_MAP`.
- File uploads under 4MB use a single POST; larger files use Graph's upload session protocol (3.2MB chunks, must be multiples of 320KB).
- String fields have a 255-char limit in SharePoint. If `allow_string_recasting=True`, columns are auto-upgraded to `Note` (multi-line) type on overflow; otherwise data is truncated.
