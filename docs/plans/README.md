# Plans

Cross-component plans live here. A plan whose work has fully landed moves to
[completed/](completed/) and is kept as the record of the decisions behind the
shipped behaviour rather than deleted. Per-page and per-method plans live in
[frontend/docs/plans/](../../frontend/docs/plans/) instead.

Give every plan a `Status:` line under its title and keep it current; this
index is a summary of those lines, not a second source of truth.

## Open

| Plan | Status |
| :--- | :--- |
| [build_new_methods.md](build_new_methods.md) | Four of five ResQ methods shipped; Bootstrap Consolidation is still open. Also holds the ResQ API and sample-project reference notes agents rely on. |
| [generated_formula_dependency_refresh.md](generated_formula_dependency_refresh.md) | Investigated 2026-09-05; scoped Engine refresh and missing generated-formula dependency links explained; implementation proposed. |
| [hosted_save_http_transport.md](hosted_save_http_transport.md) | Implemented for every hosted-save kind; TLS, traffic limits, and retiring the SMB path remain. |
| [hosted_workspace_http_transport.md](hosted_workspace_http_transport.md) | Phase 1 reads and Phase 2 engine calculations implemented; the bounded-server foundation, SSE, and small writes remain. |
| [local_runtime_log_retention_plan.md](local_runtime_log_retention_plan.md) | Audit complete (2026-08-09); remediation not started. |

## Completed

| Plan | Landed |
| :--- | :--- |
| [completed/custom_data_processing.md](completed/custom_data_processing.md) | 2026-07-16 — custom data processing rules and their Project Settings editor. |
| [completed/engine_dependent_propagation_plan.md](completed/engine_dependent_propagation_plan.md) | 2026-08-06 — both phases of the Engine-hosted dependent propagation job. |
| [completed/hosted_rpc_bridge_transport.md](completed/hosted_rpc_bridge_transport.md) | 2026-08-19 — DFM and Result Selection sync moved off SMB, except the deliberately deferred `apply`. |
| [completed/persisted_json_contract_v4.md](completed/persisted_json_contract_v4.md) | 2026-08-23 — every stored JSON file moved to one `snake_case` convention with fewer fields and one audit-log policy, converted in place for `NJ_Annual_Prod_202605_Fake` (the other projects are re-imported from ResQ instead) and released as ArcRho 1.3.3; closed 2026-09-06 with the ultra review dropped and the release left unforced, since every user is assumed to run the latest version. |
| [completed/manual_input_period_rollup.md](completed/manual_input_period_rollup.md) | 2026-09-05 — a hand-entered triangle now records the shape it is really stored at, a coarser view of it is added up along the calendar on every open instead of being cached beside it, and a yearly method can take a monthly or quarterly one as its input. Includes the [backfill run report](completed/manual_input_period_rollup_backfill_report.md) for the 10,997 records already on the server. |
| [completed/manual_input_stored_length_resq_alignment.md](completed/manual_input_stored_length_resq_alignment.md) | 2026-09-07 — a hand-entered triangle can now be stored finer than it is shown, a coarser development view can be typed, pasted and linked into with each figure written at its column's age and the periods between cleared, and the export writes the triangle to ResQ at ArcRho's stored shape. Records the ResQ stored-length rules established against its COM API. |
