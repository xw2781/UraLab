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
| [hosted_save_http_transport.md](hosted_save_http_transport.md) | Dataset-sidecar and DFM-method pilot implemented; the broader rollout is still proposed. |
| [hosted_workspace_http_transport.md](hosted_workspace_http_transport.md) | Phase 1 reads and Phase 2 engine calculations implemented; the bounded-server foundation, SSE, and small writes remain. |
| [local_runtime_log_retention_plan.md](local_runtime_log_retention_plan.md) | Audit complete (2026-08-09); remediation not started. |
| [manual_input_stored_length_resq_alignment.md](manual_input_stored_length_resq_alignment.md) | Investigation notes (2026-09-06): what ResQ's GUI allows for a manual triangle's stored length and coarse-display paste, the gap to ArcRho, and eight questions to answer against the ResQ API before planning. |
| [persisted_json_contract_v4.md](persisted_json_contract_v4.md) | In progress: steps 1-5 landed, step 6 (the conversion script) is next. |

## Completed

| Plan | Landed |
| :--- | :--- |
| [completed/custom_data_processing.md](completed/custom_data_processing.md) | 2026-07-16 — custom data processing rules and their Project Settings editor. |
| [completed/engine_dependent_propagation_plan.md](completed/engine_dependent_propagation_plan.md) | 2026-08-06 — both phases of the Engine-hosted dependent propagation job. |
| [completed/hosted_rpc_bridge_transport.md](completed/hosted_rpc_bridge_transport.md) | 2026-08-19 — DFM and Result Selection sync moved off SMB, except the deliberately deferred `apply`. |
| [completed/manual_input_period_rollup.md](completed/manual_input_period_rollup.md) | 2026-09-05 — a hand-entered triangle now records the shape it is really stored at, a coarser view of it is added up along the calendar on every open instead of being cached beside it, and a yearly method can take a monthly or quarterly one as its input. Includes the [backfill run report](completed/manual_input_period_rollup_backfill_report.md) for the 10,997 records already on the server. |
