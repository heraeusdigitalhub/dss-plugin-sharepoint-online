# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Dataiku DSS plugin providing read/write connectivity to SharePoint Online. Plugin ID: `sharepoint-online`, version defined in `plugin.json`. Written in Python 3 (3.5–3.11 compatible). Runtime dependencies: `sharepy==1.3.0`, `msal==1.23.0`.

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

- **`sharepoint_client.py`** — Central class `SharePointClient`. Handles all 5 authentication flows, constructs SharePoint REST API URLs, manages sessions, and provides methods for list/file/folder operations. Also contains `SharePointSession` (token-based session wrapper) and `get_form_digest_value()`.
- **`sharepoint_lists.py`** — `SharePointListWriter` handles write operations: list creation, column creation, batch item insertion via `ThreadPoolExecutor` with configurable `max_workers` and `batch_size`. Includes DSS↔SharePoint type mapping (`get_dss_type`, `get_sharepoint_type`).
- **`sharepoint_items.py`** — Item-level utilities for converting row data to SharePoint format.
- **`robust_session.py`** — `RobustSession` wraps HTTP calls with automatic retry on 429/503 status codes, respects `Retry-After` headers, and optionally resets sessions on 403 errors.
- **`common.py`** — Shared utilities: URL parsing, path validation (forbidden chars, 400-char limit), date format conversion, query string parsing.
- **`safe_logger.py`** — `SafeLogger` wraps Python logging, masking sensitive keys (auth tokens, passwords, certificates) in output.
- **`dss_constants.py`** / **`sharepoint_constants.py`** — Constants for DSS config keys and SharePoint API values respectively.

### Authentication (5 methods, configured via `parameter-sets/`)

Each method has its own parameter set definition in `parameter-sets/<name>/parameter-set.json`:

1. **OAuth** (`oauth-login`) — Azure SSO, bearer token
2. **SharePoint Login** (`sharepoint-login`) — Direct username/password via `sharepy` (legacy, harsher throttling)
3. **Site App Permissions** (`site-app-permissions`) — Client ID + secret, client_credentials grant
4. **App Certificate** (`app-certificate`) — Certificate-based auth via MSAL
5. **App Username/Password** (`app-username-password`) — MSAL `acquire_token_by_username_password`

Auth type is selected via `config['auth_type']` and dispatched in `SharePointClient.__init__`.

### Write Flow

Writes go through `SharePointListWriter` which: recycles existing list (on OVERWRITE mode) → creates new list → creates columns via schema XML → inserts items in batches using `$batch` API endpoint with multipart/mixed encoding. The batch processing lives in `SharePointClient.process_batch()`.

### Key Conventions

- All imports in `python-lib/` use flat (non-package) imports — files are added to PYTHONPATH, not treated as a package.
- SharePoint REST API responses use `d` (v2) as the results container key, accessed via `SharePointConstants.RESULTS_CONTAINER_V2`.
- Column identity uses `StaticName` internally; display uses `Title`. The mapping between them is tracked in `SharePointClient.column_ids`, `column_names`, and `column_entity_property_name` dicts.
- File uploads under 262MB use single POST; larger files use chunked upload (start/continue/finish pattern).
