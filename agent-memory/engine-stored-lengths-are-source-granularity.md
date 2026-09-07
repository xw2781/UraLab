---
name: engine-stored-lengths-are-source-granularity
description: "An Engine-generated sidecar's stored_* pair is the project's source-table granularity, not the shape of its own CSV, so every reader that opens csv_file needs an engine carve-out"
metadata: 
  node_type: memory
  type: project
  originSessionId: 094180bc-d138-41c9-b156-157b52f8a8fc
  modified: 2026-09-07T17:22:33.456Z
---

Since 2026-09-05 (`docs/plans/completed/manual_input_period_rollup.md`, step 5) an
`engine` sidecar's `stored_origin_length` / `stored_development_length` /
`stored_period_length` record **how fine the project's source table is**, taken from
`field_mapping.json`'s `source_period_months`. In a monthly-source project such as
`NJ_Annual_Prod_2026 Q3-Aug` every generated dataset carries stored 1/1 while its
`csv_file` is `…@12@12@cum@dev.csv`.

That contradicts `arcrho_api.sidecar_core_contract.stored_length_fields`, whose docstring
says the pair is the months per period of the CSV the sidecar names — true for `input`,
`calculated` and method-output sidecars, false for `engine`.

**Why:** step 3 of the same plan moved every reader that opens `csv_file` onto the stored
pair, one step before step 5 gave the pair a second meaning. Any such reader without an
engine carve-out then rejects or misreads a perfectly good annual cache. It cost the
2026-09-07 diagnosis of "Net Loss--Paid is not an annual dataset", which failed the B&S
Case Reserve Adequacy Adjustment on every save and cascaded into Severity--Adjusted,
H 02 and F 91 in the NJ Legacy HOL class.

**How to apply:**

- Before trusting `stored_lengths(sidecar)` as "what this CSV holds", check
  `source_kind == "engine"` first.
- The carve-out already exists in `dfm_service._dfm_precedent…` and
  `precedent_cache_service.precedent_source` / `precedent_csv_path`: rebuild through
  `materialize_engine_source` at the period wanted. Berquist Sherman gained one on
  2026-09-07 (`_read_source_values` looks the annual cache up by name and materializes
  only when it is missing).
- Still unaudited for the same hole: `calculated_dataset_service`'s dependency
  provenance record (recorded, not enforced) and the `dataset_service` stored-shape reads.
- Related: [[origin-length-is-not-row-count]], [[hosted-save-fix-needs-engine-deploy]].
