# DFM JSON Format

## Canonical Method JSON

Current DFM methods use `json_format = arcrho-dfm-v4`. The payload is a complete, location-independent snapshot that can render every DFM tab without reopening an Input Triangle, Ratio Basis dataset, CSV, sidecar, or reserving-class index.

v4 is the shared contract described in `docs/plans/completed/persisted_json_contract_v4.md`: every key is `snake_case` at every depth, every timestamp is UTC with millisecond precision and a `Z`, and every fingerprint is `sha256:` plus sixteen hex characters. There is no legacy fallback. A file written before v4 — `arcrho-dfm-method-by-tab-v2` or anything else — is rejected outright and must be converted by `tools/migrate_persisted_json_v4.py` before a v4 build opens it.

A v4 open reads only:

1. `methods/DFM@<details_tab.name>.json`
2. `sidecars/<details_tab.output_dataset>.json`

Project Instance supplies both identities so the reads can run in parallel. A method-name-only caller reads the method first and then reads the sidecar identity declared by that method.

## Identity and Ownership

- `details_tab.name` is the method identity and owns the `DFM@<name>.json` filename.
- `details_tab.output_dataset` is the output CSV/sidecar identity. A new GUI method defaults it to its method name, but migrated methods may keep a different output name.
- `details_tab.output_type` is the output Vector Dataset Type.
- `details_tab.input_triangle`, period lengths, decimal places, ratio exclusions, average definitions/order/selections/inputs, literal User Entry values, each row's `- Ult` tail factor, stored values for any Excel- or dataset-linked formula, the Curves tab choices (`curves_tab` minus `selected_values`), Ratio Basis selection, ultimate-ratio decimals, and ratio-cell notes are DFM-owned state.
- Input/basis snapshots, ratio values, standard-average values, non-Excel formula results, the Curves tab's `selected_values`, and ultimates are derived state.
- Method Notes, Audit, status, and `precedents`/`dependents` live only in the output sidecar. The canonical sidecar precedent projection combines Input Triangle, Ratio Basis, and every case-insensitively unique dataset parsed from Ratios User Entry `inputs`; each source sidecar receives the reverse dependent edge. Ratio-cell notes live only in method JSON.

The output sidecar registers both the Input Triangle and configured Ratio Basis as precedents. A method save cannot silently reuse an output sidecar owned by another method.

A method file carries no audit log of its own (v4 rule 8). A method's history lives in the sidecar of the dataset its output writes, which is where the app already reads it.

## Stored Sections

`details_tab` stores:

- `name`
- `output_type`
- `output_dataset`
- `output_category`
- `input_triangle`
- `origin_length`
- `development_length`
- `decimal_places`

`data_tab` stores:

- exact `origin_labels` and `development_labels`
- `input_data_triangle_values`, with trailing nulls trimmed from each row so a row ends at its last populated development period
- `data_format`, `number_format`, `decimal_places`, and `source_revision`

The persisted file does not store `input_data_triangle_mask`. A cell is inside the triangle if and only if it holds a value, so the mask can only restate the values beside it; loading derives it and refits every row back to the full development geometry. A null *inside* a row still marks a value missing inside the triangle, exactly as `ratio_values` and `excluded` already store their rows. The in-memory canonical payload keeps the mask and its rectangular geometry, so revisions and calculations are unaffected.

`ratios_tab.ratio_triangle` stores aligned `origin_labels` and `development_labels`, calculated `ratio_values`, and DFM-owned `excluded` cells. `ratios_tab.average_formulas` remains the columnar object with `label`, `custom_average_formula_settings`, `selected`, `values`, aligned User Entry `inputs`, and aligned display-only `display_inputs`. A `display_inputs` cell stores the same formula with dataset coordinate positions replaced by the labels returned when that formula was resolved; calculation, dependency parsing, and editing continue to use `inputs`, and display metadata never creates a graph edge. `ratios_tab.cell_notes` remains keyed by visible row label and visible development-column label.

The ratio triangle keeps its own axis labels on purpose, even though the contract forces them to agree with `data_tab` (Decision 6 of the v4 plan): a method file is also read as raw text by a person or by ArcBot, and a ratio triangle without headings is not a complete table. The validation that they match `data_tab.origin_labels` and the labels derived from `data_tab.development_labels` still runs, so the two can never drift into carrying different information.

Each `excluded` row must be exactly as long as the `ratio_values` row beside it; a payload that breaks that alignment is rejected wherever the method is validated as complete, which includes the macro and ArcBot handoffs. Both rows drop their trailing empty cells, so they stay aligned only while they agree on which cells are empty: a cell whose ratio cannot be calculated -- a zero or missing left value, most often a zero origin row -- is null in `ratio_values` and 2 in `excluded`, never a calculated ratio of 0.

The last `average_formulas.values` column, `<age> - Ult`, is each row's tail factor: an entered value on a User Entry or benchmark row (ResQ's `CustomAverages(i).TailFactor`) and `1.0` on a computed average row. Formulas in `inputs` never apply there.

`curves_tab` stores the Curves tab, owned by `arcrho_api/dfm_curves.py`:

- `fitting_method`: `log_regression` (the only method ArcRho fits by) or `least_squares` when a ResQ import carried that setting
- `future_development_periods`, `free_fit_c`
- `included`: one `0`/`1` per development period, whether that period's Initial Selection takes part in the curve fits
- `user_columns`: one object per user value column with `label`, `column_type` (`user_entry`, `prior_analysis`, `pattern`, or `benchmark`), `values` (one per development period), and `tail`
- `selected_estimates`: one column number per development period (1 = Initial Selection, 2-5 = Exponential Decay, Inverse Power, Power, Weibull, 6 onward = user columns)
- `selected_tail_factor` and `selected_tail_curve`: the column the tail factor and the tail pattern are taken from
- the derived `selected_values`: the selected factor per development period followed by the selected tail, at six decimals, which is the chain `ultimate_vector` and the percentage developed use

A method file written before the Curves tab existed normalizes to the default tab (the Initial Selection selected everywhere, ResQ's default inclusion thresholds of 1.00001 and 2), and a default tab is left out of the revision fingerprints, so such a file keeps its stored revisions and its factors.

`results_tab` stores:

- `ratio_basis_dataset`
- `ratio_basis_data_format`, `ratio_basis_number_format`, `ratio_basis_decimal_places`
- `ratio_basis_values`, aligned to the DFM origins
- `ratio_basis_source_revision`
- `ultimate_ratio_decimal_places`
- the calculated `ultimate_vector`

`results_tab.ratio_basis_origin_labels` is **not** stored. It was a forced copy of `data_tab.origin_labels`, so it could never carry information; loading re-derives it.

`method_metadata` stores:

- `last_modified`: changed only by an owned user save
- `data_refreshed`: changed when embedded precedent data is refreshed
- `owned_revision`
- `derived_revision`
- `publication_revision`

Revisions are deterministic hashes over their canonical projections. They are separate so a dirty window can rebase an owned patch over a newer derived-only disk refresh, while a concurrent owned change produces a conflict.

The hash vocabulary is independent of the persisted key spelling, so renaming a stored field cannot shift a stored revision or mark a method Review Needed. `arcrho_api.fingerprints` is the one producer and truncates every value to `sha256:` plus sixteen hex characters, so both sides of every comparison shorten together.

The file never stores absolute input or output CSV paths.

## On-Disk Text Format

`arcrho_api/io.py::persisted_json_text` owns the file's text, so every producer writes the same bytes for the same payload.

Two-dimensional arrays are stored one row per line: `input_data_triangle_values`, `ratio_values`, `excluded`, and the `average_formulas` row arrays each read as a triangle rather than one scalar per line. A 40-origin method drops from roughly 1,900 lines to 170, and from about 30 KB to 13 KB, which is what a network-drive read pays for. Every other node keeps the two-space layout.

Layout never reaches a revision: `owned`, `derived`, and `publication` hash a `separators=(",", ":")` encoding of the canonical projection, so reformatting a file cannot shift a stored revision.

## Calculation and Numeric Rules

Persisted numeric values use half-away-from-zero normalization at six decimals, with one exception: the observed input triangle (`input_data_triangle_values`) is stored at the full precision it was read with and is never rounded. Six decimals is enough for a derived figure a reader checks by eye, and any fixed decimal place is the wrong rule for the operands every ratio and every average divides, because how much of a number it keeps depends only on how large that number happens to be — ten decimals was generous for a loss figure and still too coarse for a near-zero "% of" figure, where the trimmed tail moved a ratio read at four. A JSON number round-trips a double exactly and the shortest text that reads back as the same value is what lands on disk, so the file stays as readable as it was. Ratio triangle values, average formula values, the ultimate vector, and the Ratio Basis stay at six decimals, and the input source revision is still fingerprinted at six so a stored revision does not shift under this rule. The canonical Python contract owns ratio calculation, average calculation, internal User Entry formula evaluation, ultimate calculation, field projection, and revisions. The frontend calls the local preview endpoint for canonical interactive derivation.

A ratio exists only where both values are present and non-zero. A zero later value stores `null`, not `0`, so the cell renders as the muted placeholder, takes no place in a "last N" window, and enters no sum and no divisor. `Ex hi/lo` then trims its pair from the ratios that remain, which is how ResQ reads the same column.

A formula containing any Excel reference freezes its complete stored result during automatic upstream refresh, including mixed Excel/internal formulas. A non-Excel User Entry formula is recalculated. Literal User Entry values, formula definitions, selections, and exclusions are preserved.

Opaque migrated average rows such as ResQ `Benchmark` keep their persisted values. They are not reinterpreted as a standard Simple or Volume average by either the canonical contract or the frontend renderer.

Origin changes remap owned state by exact label; new origins default to included. Development geometry must remain compatible. Positional remapping is forbidden: an incompatible geometry leaves the prior publication intact and marks the DFM Review Needed.

Ratio Basis values are aligned by exact origin label. A missing or duplicate required label is a refresh error rather than a positional fallback. The saved labels must equal the DFM origins exactly, so the method window re-reads the Ratio Basis dataset at the new Origin Length before it builds a payload on a changed origin basis; the embedded snapshot is never carried across bases.

## Refresh and Publication

ArcRho-managed durable precedent saves refresh affected DFM snapshots and calculations. In the desktop app-server workflow, DFM refresh runs before calculated datasets and Result Selection descendants. After the refresh wave, the DFM remains Review Needed until its own explicit Save; Refresh alone does not acknowledge the alert. Every explicit Save starts downstream propagation even when the DFM publication values are unchanged.

Standalone public-Python and ResQ-migration execution refreshes DFM descendants but does not host the app server's calculated-dataset or Result Selection evaluators. If propagation reaches either method type, that branch is marked Review Needed and a warning is returned so it can be recalculated through the app-server workflow.

Publication runs under the reserving-class lock and uses staged files, revision checks, rollback, unchanged-file suppression, and sidecar-last replacement. A failed branch retains its last valid publication and blocks only its descendants. The upstream save remains successful and reports the propagation warning separately from the human-review status.

An automatic refresh that rewrites the method file — a new input snapshot, ratio basis, or published ultimate — stamps the output sidecar's `updated_at`/`modified_by` (the dataset table's Last Modified) and adds an `Auto Refresh` audit record, whether or not the ultimate publication changed. Output CSVs are rewritten only when the publication changes, a refresh that leaves the method byte-identical touches neither stamp nor audit (it only restores a Review Needed status), and an automatic refresh never clears Review Needed. Every action is kept, including `Auto Refresh`; consecutive automatic records collapse to the most recent, and the log is capped at 200 records.

## Excel Freshness

Excel links are derived from User Entry `inputs`; there is no separate persisted Links section. Excel-only DFM hydration never refreshes workbook values. After Ready, one abortable check-only task per applied method revision reads saved workbook values in a deduplicated batch, compares canonical results, and reports stale/unverified counts without changing method state, caches, rendering, JSON, or dirty state. Dataset-backed formulas are different: after clean hydration the DFM becomes dirty immediately, resolves their current dataset values in memory, and reports the automatic evaluation in the status bar; only explicit Save persists the refreshed `values` and publication.

Manual refresh from the existing Links or Ratios controls remains mutating. Changed values mark the method dirty and require Save; ignoring a warning keeps the stored values.

## Producer Parity

The app server, public Python API, ResQ migration, and bridge-owned-patch flow delegate to the canonical v4 contract, and both app-server writers go through `persisted_projection`, so every producer emits one shape. Sparse RPC, ArcBot, macro, and template payloads are treated as owned-setting patches followed by canonical local calculation; a sparse payload is never persisted directly. The owned-patch carrier is stamped `arcrho-dfm-owned-patch-v4`.

The output sidecar is the same schema as any other dataset sidecar (v4 rule 9): the shared core — labels, notes, number formatting, `precedents`, `dependents`, and `audit_log` last — plus the method-only fields `method_name` and `publication_revision`. `arcrho_api.sidecar_core_contract` owns that shape and every sidecar writer validates against it.
