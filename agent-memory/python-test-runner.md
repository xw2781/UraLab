---
name: python-test-runner
description: No interpreter on this dev PC has pytest; install it into a repo-local --target dir to run python-api tests
metadata: 
  node_type: memory
  type: feedback
  originSessionId: 49fe14c4-94ee-453f-a0ac-d42ea1b6f43e
  modified: 2026-09-07T00:00:00.000Z
---

No Python interpreter on this machine has `pytest` installed — not `C:\Program Files\Python310`, not `Python314`, and none of the `server-components/venvs/*` environments. To run `python-api/tests`, install it into a throwaway directory **inside the repo** and put that on `PYTHONPATH`:

```
python -m pip install --quiet --target e:/XWSpace/Repos/ArcRho/.pytest-tools "pytest>=8"
PYTHONPATH="e:/.../.pytest-tools;e:/.../python-api/src" python -m pytest python-api/tests -q
```

**Why:** AGENTS.md forbids validation commands from writing to the C drive, so a user-site or venv install is out; a repo-local `--target` keeps everything on E: and is trivially removable.

**How to apply:** Delete `.pytest-tools` when finished so it never reaches a commit. Three pre-existing conditions are not your fault: `tests/test_arcrho_api.py` fails collection (a helper named `test_log` takes a non-fixture argument), so exclude it; 6 tests in `test_resq_data_migration_graph.py` / `test_validate_engine_resq_parity.py` already fail on a clean `main`; and `frontend/tests/test_result_selection_cross_producer_contract.py` fails at HEAD too (bridge payload lacks the `_sidecar_status` key the migration payload carries) — confirmed 2026-08-12 via a detached `git worktree` at HEAD.

`frontend/tests/*.py` run under plain `unittest` with the bridge venv python and `PYTHONPATH=frontend;python-api/src;server-components/src`; no pytest needed. `unittest discover` refuses `-s frontend/tests` (not an importable package) — `cd` into the tests folder and use `-s .` with absolute paths on `PYTHONPATH`.

On the **Client PC clone** (`C:\Users\xwei\Repos\ArcRho`, 2026-08-14) the bridge venv lacks `openpyxl`, so use `server-components/venvs/arcrho_engine/Scripts/python.exe` instead. Baseline there: 588 frontend tests with 4 failures that are not yours — `test_sql_formatting_service` (no `sqlfluff`), `test_dataset_number_format_defaults` (needs an "Example Project" folder), `test_result_selection_cross_producer_contract` (known), and `test_engine_dataset_sidecar_contract`, which fails only on a machine whose real `username_index.json` maps the login, because the runtime writer resolves a full name while the migration fixture expects `tester`. `test_class_folder_scan_cache` is timing-flaky. `server-components/tests` is clean except `test_multi_user_instances` (no `psutil`). A worktree baseline also needs `mklink /J <worktree>\frontend\node-portable` to the real one, since node-portable is gitignored and some tests shell out to it. Related: [[shared-macro-library-deploy]], [[frontend-node-test-suite]], [[app-server-route-smoke-test]].

On the `E:\XWSpace\Repos\ArcRho` clone under user `xwei.PRCINS` (2026-08-27): there is no `server-components/venvs` folder at all, and the sandbox denies a Bash `ls` of that path, so do not look for a venv. `py -3.10` (3.10.6) runs the python-api and server-components suites directly: `cd python-api\tests; py -3.10 -m unittest test_resq_sync_plan test_resq_sync_session ...` — each test file inserts its own roots. Run it through the PowerShell tool; the `.....` progress dots come back as a NativeCommandError line that is not a failure, read the `Ran N tests ... OK` tail.

Better on the Client PC (2026-08-17): the user-scoped interpreter `py -3.10` (`C:\Users\xwei\AppData\Local\Programs\Python\Python310`) has fastapi, pydantic, uvicorn, sqlfluff, snowflake-connector, **and pyodbc**, so `cd frontend && py -3.10 -m unittest discover -s tests -p "test_*.py"` runs the whole frontend suite with no PYTHONPATH juggling — each test file inserts the repo roots itself. Baseline that day: 667 tests, 3 failures (`test_dataset_number_format_defaults`, `test_engine_dataset_sidecar_contract`, `test_result_selection_cross_producer_contract`), all reproduced in a detached HEAD worktree. Re-checked 2026-09-05 on the Client PC clone: `C:\Program Files\Python310` does **not** exist there (only `Python313`), so the Bash-tool path is `/c/Users/xwei/AppData/Local/Programs/Python/Python310/python.exe`; `frontend/tests/test_resq_sync_queue_service` also fails at HEAD that day (the registered `resq_sync_queue` mutation lists `selected_names`/`direction` the test does not expect), and `project_settings_source_data.test.mjs` has 3 table-summary regex failures at HEAD.

On `E:\XWSpace\Repos\ArcRho` under `xwei.PRCINS` (2026-08-27, corrected): the system `C:\Program Files\Python310\python.exe` (3.10.6) does have fastapi, pydantic, pandas, numpy and pywin32 — only pytest is missing — so frontend service tests run with `cd frontend/tests && "C:\Program Files\Python310\python.exe" -m unittest test_result_selection_service` (module names, from the tests folder so the workspace stub imports). `python` (3.14) plus the repo-local `.pytest-tools` on `PYTHONPATH` runs the python-api suites. Frontend test runs can leave `tmp*` folders at the repo root; delete them before committing.

Client PC clone, 2026-09-06: neither `C:\Program Files\Python310` nor `py -3.10` was on the Bash-tool PATH (only `Python313`, no pandas), but `server-components/venvs/arcrho_engine/Scripts/python.exe` (3.10.11, pandas + fastapi) runs everything: `cd python-api && PYTHONPATH=../.pytest-tools <engine python> -m pytest tests/...` for python-api (the repo-local `.pytest-tools` is present there), and `cd frontend/tests && <engine python> -m unittest test_x` for frontend. Pre-existing failures that day, confirmed in a detached HEAD worktree: `test_resq_strict_extraction` ×2 (Cape Cod strict read of `PercentageDevelopedValues`), `test_resq_data_migration_graph::test_refresh_preserves_result_selection_precedent_strings` (precedents are dicts now), and `frontend/tests/test_bootstrap_service::test_saved_method_embeds_the_dfm_snapshot_and_a_simulation_summary` (rounding). `test_dfm_service` can throw a one-off `WinError 5` on a sidecar rename; rerun before blaming a change. The Edit tool can leave a file at LF while the rest of the working copy is CRLF (git warns "LF will be replaced by CRLF"); convert it back before finishing.

Client PC, 2026-09-07: `py -3.10` is back on the Bash-tool PATH
(`C:\Users\xwei\AppData\Local\Programs\Python\Python310`, with fastapi + pandas), and there is no
`server-components/venvs` folder in this clone at all - a Glob for `**/Scripts/python.exe` finds nothing,
and the sandbox denies `ls` of that path, so do not hunt for a venv. Frontend service tests:
`cd frontend/tests && py -3.10 -m unittest test_x test_y` (module names, from the tests folder).

**Module order used to pollute `python-api/tests` (2026-08-29, resolved 2026-09-01).** `test_import_resq_reserving_classes_macro` once failed 3 of 13 tests when run after other `test_resq_*` modules, because an earlier module put `python-api/migration` on `sys.path` and the macro's `import resq_data_migration` fallback then found the real 17-entry `RC_PATH`. Since macro v1.6.0 (2026-09-01) the batch macro carries its own hard-coded `RC_PATHS` list and never imports the migration module, so that order dependence is gone (16/16 alone with `py -3.10 -m unittest`). `server-components/tests/test_bridge_import_request_protocol.test_client_delegates_full_import_to_the_canonical_runner` still fails at HEAD (the client now passes `resq_credentials` the test does not expect).
