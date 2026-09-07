# App Server Index

## Purpose
<!-- MANUAL:BEGIN -->
App-server domain map for FastAPI routers, schemas, and services.
<!-- MANUAL:END -->

## Entry Points
<!-- AUTO-GEN:BEGIN app_server.index.entry_points -->
| Domain | Router | Route Count | Domain Index |
| --- | --- | --- | --- |
| `app_control` | [`app_server/api/app_control_router.py`](../../app_server/api/app_control_router.py) | 4 | [`app_control.md`](domains/app_control.md) |
| `arcrho` | [`app_server/api/arcrho_router.py`](../../app_server/api/arcrho_router.py) | 9 | [`arcrho.md`](domains/arcrho.md) |
| `audit_log` | [`app_server/api/audit_log_router.py`](../../app_server/api/audit_log_router.py) | 2 | [`audit_log.md`](domains/audit_log.md) |
| `book` | [`app_server/api/book_router.py`](../../app_server/api/book_router.py) | 3 | [`book.md`](domains/book.md) |
| `bootstrap` | [`app_server/api/bootstrap_router.py`](../../app_server/api/bootstrap_router.py) | 4 | [`bootstrap.md`](domains/bootstrap.md) |
| `bornhuetter_ferguson` | [`app_server/api/bornhuetter_ferguson_router.py`](../../app_server/api/bornhuetter_ferguson_router.py) | 4 | [`bornhuetter_ferguson.md`](domains/bornhuetter_ferguson.md) |
| `cape_cod` | [`app_server/api/cape_cod_router.py`](../../app_server/api/cape_cod_router.py) | 4 | [`cape_cod.md`](domains/cape_cod.md) |
| `data_processing_rules` | [`app_server/api/data_processing_rules_router.py`](../../app_server/api/data_processing_rules_router.py) | 5 | [`data_processing_rules.md`](domains/data_processing_rules.md) |
| `dataset` | [`app_server/api/dataset_router.py`](../../app_server/api/dataset_router.py) | 18 | [`dataset.md`](domains/dataset.md) |
| `dataset_types` | [`app_server/api/dataset_types_router.py`](../../app_server/api/dataset_types_router.py) | 4 | [`dataset_types.md`](domains/dataset_types.md) |
| `dependent_propagation` | [`app_server/api/dependent_propagation_router.py`](../../app_server/api/dependent_propagation_router.py) | 4 | [`dependent_propagation.md`](domains/dependent_propagation.md) |
| `excel` | [`app_server/api/excel_router.py`](../../app_server/api/excel_router.py) | 5 | [`excel.md`](domains/excel.md) |
| `excel_link` | [`app_server/api/excel_link_router.py`](../../app_server/api/excel_link_router.py) | 3 | [`excel_link.md`](domains/excel_link.md) |
| `field_mapping` | [`app_server/api/field_mapping_router.py`](../../app_server/api/field_mapping_router.py) | 2 | [`field_mapping.md`](domains/field_mapping.md) |
| `object_change_watch` | [`app_server/api/object_change_watch_router.py`](../../app_server/api/object_change_watch_router.py) | 2 | [`object_change_watch.md`](domains/object_change_watch.md) |
| `project_settings` | [`app_server/api/project_settings_router.py`](../../app_server/api/project_settings_router.py) | 14 | [`project_settings.md`](domains/project_settings.md) |
| `reserving_class` | [`app_server/api/reserving_class_router.py`](../../app_server/api/reserving_class_router.py) | 11 | [`reserving_class.md`](domains/reserving_class.md) |
| `result_selection` | [`app_server/api/result_selection_router.py`](../../app_server/api/result_selection_router.py) | 3 | [`result_selection.md`](domains/result_selection.md) |
| `snowflake` | [`app_server/api/snowflake_router.py`](../../app_server/api/snowflake_router.py) | 6 | [`snowflake.md`](domains/snowflake.md) |
| `source_table` | [`app_server/api/source_table_router.py`](../../app_server/api/source_table_router.py) | 12 | [`source_table.md`](domains/source_table.md) |
| `sql_formatting` | [`app_server/api/sql_formatting_router.py`](../../app_server/api/sql_formatting_router.py) | 1 | [`sql_formatting.md`](domains/sql_formatting.md) |
| `sql_server` | [`app_server/api/sql_server_router.py`](../../app_server/api/sql_server_router.py) | 6 | [`sql_server.md`](domains/sql_server.md) |
| `table_summary` | [`app_server/api/table_summary_router.py`](../../app_server/api/table_summary_router.py) | 2 | [`table_summary.md`](domains/table_summary.md) |
| `ui_automation` | [`app_server/api/ui_automation_router.py`](../../app_server/api/ui_automation_router.py) | 6 | [`ui_automation.md`](domains/ui_automation.md) |
| `workflow` | [`app_server/api/workflow_router.py`](../../app_server/api/workflow_router.py) | 5 | [`workflow.md`](domains/workflow.md) |
| `workspace_paths` | [`app_server/api/workspace_paths_router.py`](../../app_server/api/workspace_paths_router.py) | 2 | [`workspace_paths.md`](domains/workspace_paths.md) |
<!-- AUTO-GEN:END -->

## Key Files
<!-- AUTO-GEN:BEGIN app_server.index.key_files -->
- [`app_server/main.py`](../../app_server/main.py) - FastAPI app creation, router registration, static mount.
- [`app_server/api/__init__.py`](../../app_server/api/__init__.py) - Router exports consumed by app startup.
- [`app_server/config.py`](../../app_server/config.py) - Runtime path/config constants and helpers.
- [`app_server/helpers.py`](../../app_server/helpers.py) - Cross-domain utility helpers.
<!-- AUTO-GEN:END -->

## Non-Negotiable Contracts
<!-- MANUAL:BEGIN -->
Mandatory before app-server logic/API/architecture changes:
1. [`../contracts/business_logic_contract.md`](../contracts/business_logic_contract.md)
2. [`../architecture/architecture_guardrails.md`](../architecture/architecture_guardrails.md)
3. [`../contracts/frontend_behavior_contract.md`](../contracts/frontend_behavior_contract.md) for cross-frame/API behavior coupling

High-risk files that must follow contracts:
- `app_server/api/*.py`
- `app_server/services/*.py`
- `app_server/config.py`
<!-- MANUAL:END -->

## External Interfaces
<!-- MANUAL:BEGIN -->
- Public interface is HTTP routes mounted by `app_server/main.py`; the frontend shell is served under `/ui` and shared icon assets under `/icons`.
- Every static mount in both apps uses `RevalidatedStaticFiles` from `app_server/ui_static.py`, which answers with `cache-control: no-cache`. The UI is one ES-module graph resolved by URL, so a module the browser reuses without asking the server can disagree with a freshly fetched importer and break module linking outright. Assets stay cacheable but must be revalidated, which costs one `304` per unchanged file over loopback. Do not mount UI assets with a bare `StaticFiles`.
- Internal interface is router -> service -> filesystem/state helpers.
- Packaged builds include the `arcrho_api` Python package in the frozen app server for Arcode scripting imports, ship a pip-installable wheel under app resources `python_packages/`, and publish the same wheel to the shared Server packages folder as `arcrho_api-latest.whl` for external notebook environments. API-only releases can publish that shared wheel from `python-api/tools` without rebuilding the desktop app.
- Both full ArcRho and standalone Arcode expose local-only `POST /scripting/run-in-arcrho` with the same source-buffer request contract. Full ArcRho captures/executes/applies against its live DFM UI; standalone Arcode verifies and proxies to the running ArcRho desktop server at its resolved local endpoint (default port 28765, or the per-user discovered fallback port).
<!-- MANUAL:END -->

## Data/State/Caches
<!-- MANUAL:BEGIN -->
- Path and cache constants are centralized in `app_server/config.py`.
- Several domains persist JSON caches under project folders or AppData.
- On-disk JSON text is owned by `arcrho_api/io.py`. Every service that persists project, reserving-class, method, sidecar, index, or project-data cache JSON writes `persisted_json_text(payload)` instead of `json.dump(..., indent=2)`, so all producers -- app server, public Python API, ResQ migration, macros, bridge -- emit the same bytes for the same payload. Two-dimensional arrays (triangles, ratio values, exclusion masks, table rows) are written one row per line; everything else keeps the two-space layout and is byte-identical to the previous text. See the root `AGENTS.md` "Persisted JSON Text Format" rule.
- Scripting notebook persistence is file-based under `~/Documents/ArcRho/scripts`; save writes `.ipynb` with code-cell outputs/execution counts and load accepts `.ipynb`, legacy `.arcnb`, and `.py` scripting files from that directory. ArcRho Macro window files are loaded only from `~/Documents/ArcRho/macros`; the app does not seed or overwrite user macro files. `/scripting/macros` parses optional macro metadata `Scope:` values (`DFM`, `Result Selection`, and `Reserving Class`, including comma-separated combinations), defaults unscoped macros to `DFM`, marks generated Task Designer wrapper macros, and exposes their child task labels for immediate progress-table seeding. `/scripting/run-macro` returns an updated DFM payload when a DFM-backed macro modifies active DFM state, and UI-only macros can run without an active DFM target. Macro runs receive `task_window_id`, `task_session_id`, `task_mode`, an injected `task_designer` helper for live result rows, and `run_task_macro(...)` for wrapper macros that execute child macros sequentially while capturing stdout into each row; wrapper macros can call `task_designer.open(...)` themselves when launched from the Macro window. `/scripting/save-task-wrapper` writes generated Task Designer wrapper `.py` macros to the user macros folder. Macro results may include optional preview metadata, such as a Notes diff that the shell can show before applying a returned payload. `/scripting/rename-macro` and `/scripting/delete-macro` rename or remove selected user-created macro `.py` files resolved through the macros directory path guard.
- The shared macro library is a deployer-managed read-only folder at `<workspace root>\shared\macros` (override with `ARCRHO_MACRO_LIBRARY_DIR`; resolved in `app_server/config.py`, never created by the app). `GET /scripting/macro-library` enumerates the folder once and reads macro headers with bounded parallel I/O, reusing the canonical `<arcrho-macro>` metadata parser (now including `Version:` and `Release Note:`), and reports each macro's install status against the local macros folder (`not_installed`, `up_to_date`, `update_available`, `local_differs`); an unreachable library returns `available: false` without failing local macro features. `POST /scripting/macro-library/install` copies a library macro byte-for-byte into the local macros folder with a temp-file + atomic-replace write; when a differing local copy exists it returns `needs_confirmation` until the caller retries with `overwrite: true`. Macros still execute only from the local macros folder. Deployers publish to the library with `python-api/macros/publish_macro_library.py`, which validates version/release-note headers, archives replaced versions under `archive/<stem>/<version>/`, and replaces files atomically.
- UI automation commands use `/ui_automation/commands` plus shell polling/completion endpoints so Python macros and scripts can ask the running local shell to show message boxes, show/update/close progress windows, read active Project Instance context, update Task Designer validation rows, or perform typed window operations without writing ArcRho Server request files.
- Registered `/scripting/run-macro` files and unregistered Arcode source buffers delegate to the same canonical macro-source executor. Source-buffer execution uses the editor text as authoritative input, preserves the source path only for tracebacks, relative resources, and sibling imports, and never copies the file into the user macro directory. Execution is cancelled after 120 seconds without a `report_macro_activity()` heartbeat; long-running macros can report useful progress without disabling runaway-loop protection. Cancellation line tracing is limited to macro source files so imported data-processing libraries do not pay per-line tracing overhead. Unchanged DFM payloads are omitted so inspection-only scripts do not dirty the page, and UI-only scripts can run with `active_dfm=None` when no DFM is active.
- Macro worker threads initialize and release a Windows STA COM apartment around source execution when `pythoncom` is available. COM-backed macros such as ResQ imports therefore observe the same evaluated automation state as standalone scripts, while non-Windows runtimes continue without COM initialization.
- Scripting execution interrupt uses per-session cancellation with trace checks and an interruptible `time.sleep(...)` import hook so `/scripting/interrupt` can stop active cells promptly; `/scripting/run-stream` emits NDJSON stdout/stderr events for live output during long-running cells.
<!-- MANUAL:END -->

## Common Change Tasks
<!-- MANUAL:BEGIN -->
1. Add route: update one router file under `app_server/api`, schema under `app_server/schemas`, and service under `app_server/services`.
2. Change payload contract: update schema first, then router/service.
3. Change project path behavior: sync with [`../runtime/config_paths.md`](../runtime/config_paths.md).
<!-- MANUAL:END -->

## Known Risks
<!-- MANUAL:BEGIN -->
- File-based persistence and path assumptions are sensitive to environment setup.
- Domain cross-calls (for example, table summary -> reserving class refresh) can add side effects.
<!-- MANUAL:END -->
