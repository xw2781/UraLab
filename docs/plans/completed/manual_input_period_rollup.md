# Manual Input Triangles at a Coarser Method Period

Status: Complete 2026-09-05 — all ten steps landed and the change is live on the server.
Last updated: 2026-09-05 — the server's Engine, Gateway and Bridge were rebuilt, so everyone now gets the new behaviour.

## Progress

Plain‑language tracking. The agent that finishes a step ticks its box, fills in the date and how long it took, and leaves one short line on what a user would notice. Nothing technical goes here.

| # | Step | Done | Date | Est. | Took | What changed for the user |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| 1 | Coarser views of a hand‑entered triangle add up correctly | [x] | 2026-09-05 | 2 h | 15 min | Opening a monthly or quarterly triangle at a yearly view now adds the figures up along the calendar, so the numbers on screen are the ones a method would use. |
| 2 | Every dataset records the shape its data is really stored at | [x] | 2026-09-05 | 1 h 45 | 30 min | Every dataset's record now says how fine its figures really are, separately from how coarsely it is shown, so a yearly view of a monthly triangle can no longer be mistaken for monthly data. |
| 3 | The parts of the app that read a dataset's shape use the right one | [x] | 2026-09-05 | 2 h | 25 min | Anything that opens a dataset's figures now asks how fine they really are, so a yearly view of a monthly triangle is no longer mistaken for yearly data. |
| 4 | Existing projects on the server get the new shape record | [x] | 2026-09-05 | 1 h 15 | 2 h | Every dataset already on the server now records how fine its figures really are, so nothing has to guess at it, and every project's dataset list was rebuilt to match. |
| 5 | Generated datasets know how fine their source data is | [x] | 2026-09-05 | 45 min | 20 min | A project now records how fine its source data is, so a dataset the app builds from that data says it can be rebuilt monthly even when you last looked at it a year at a time. |
| 6 | Saving a hand‑entered dataset can no longer lose its detail | [x] | 2026-09-05 | 45 min | 15 min | Looking at a hand‑entered triangle a year at a time and saving no longer turns it into yearly data: the figures stay in the file you typed them into, and the app refuses to write values back at a coarser view. |
| 7 | The Dataset Viewer shows the stored shape and only offers valid views | [x] | 2026-09-05 | 1 h 45 | 15 min | The dataset window now says how fine each triangle's figures really are, offers only the coarser views that add up honestly, reopens at the view you last saved, and will not let you type over a rolled‑up figure. |
| 8 | A method can use a finer hand‑entered triangle as its input | [x] | 2026-09-05 | 45 min | 15 min | A yearly method can now take a monthly or quarterly hand‑entered triangle as its input: the figures are added up along the calendar every time the method reads them, and an old rolled‑up copy left on disk is never used. |
| 9 | Old rolled‑up copies of a hand‑entered dataset are never trusted | [x] | 2026-09-05 | 30 min | 10 min | Looking at a hand‑entered triangle at a coarser period now works the figures out from the ones you typed every time the window opens, so an edit to them always shows, and the coarser copy that used to be left beside the file is neither written nor read. |
| 10 | The change is live on the server | [x] | 2026-09-05 | 45 min | 15 min | The server now runs the new behaviour, so everyone sees a coarser view of a hand‑entered triangle added up along the calendar and can point a yearly method at a monthly or quarterly one. |

Overall: 10 of 10 steps done.

Time remaining: none — every step is done. The ten steps cost 4 h 40 against 12 h 15 of estimate, running at a little under two fifths of it, so even the recalibrated second guess was about two and a half times too slow. The one overrun was Step 4's 2 h against 1 h 15, the share handing back 10,997 files one round trip at a time; the deploy that was expected to overrun did not, because the build machine's virtual environments were already warm and three components rebuilt in two minutes.

The **Est.** column is what the step is expected to cost and **Took** is the wall‑clock time it really ran, filled in once by the agent that lands it. Keeping both is the point: a column of estimates nobody ever checks is worthless.

**Estimates were recalibrated once, on 2026-09-05, after the first step and a half.** The original figures came from the "Rough size" section at the foot of this plan, which was written in human working days and totalled 21 h. Two measurements said that was roughly twice too slow for an agent: Step 1 landed in 15 minutes against 2 h, and Step 2 reached the half‑way point in 52 minutes against 3 h. The code steps were therefore roughly halved. Steps 4 and 10 were cut by less, because most of what they cost is waiting — the share is slow per operation, and a component rebuild takes the time it takes however fast the agent thinks.

Do not recalibrate again. From here the **Est.** column is fixed, so the gap between it and **Took** stays an honest record of how good the second guess was.

The **Time remaining** line is recomputed in every step commit: add up the estimates of the steps still unticked, then scale that total by how the finished steps have really gone — the sum of their **Took** figures divided by the sum of their **Est.** figures, counting every finished step including Step 1. Say which way the scaling went when it is not close to one, as in "about 6 h, running at roughly two thirds of the estimate so far". Round to the nearest quarter hour.

While a step is running its agent also keeps a live note at `temp/plan_eta.json`, outside the repository's tracked files, holding the step number, when it started, what it is doing in one plain phrase, and how many more minutes it expects to need. It is rewritten whenever a checkbox in the step is finished or a test run ends, and at least every ten minutes of work, so a reader can see progress without waiting for the commit. The file is scratch: it is never committed, and the last one left behind after the plan finishes can simply be deleted.

## How agents work this plan

- Take the first unticked step in the Progress table. One step is one context (a session or one workflow subagent), one commit.
- Read the sections between here and the Plan before starting, then only the files the step names. Do not read ahead into later steps.
- Note the wall‑clock time when you start, and keep `temp/plan_eta.json` current as "How the time figures are kept" describes.
- A step is done when its "Done when" list holds, its tests pass, and the commit is in. In that same commit: tick the Progress row, write the date, the time the step took and the one‑line user note, update the "Overall" count and the "Time remaining" line, and update the `Status:` line at the top and this plan's row in [docs/plans/README.md](../README.md).
- If a step turns out to need a decision that is not in "Open decisions", stop, record the question there, commit that note alone, and report it rather than guessing.
- Do not start a step while the previous one is uncommitted.
- The whole plan runs unattended as a workflow of one subagent per step; the `arcrho-plan-breakdown` skill describes how. Step 10 (deploy) runs only when the user's request included it.

## The question that started this

A user enters a 120×120 monthly triangle by hand (an Excel range link). Can an annual DFM (10×10) use it as its input triangle?

Today the answer is "it looks like yes, then no":

1. The DFM Data tab asks the runtime for the dataset at the method's own period with `AllowDerived: true` and `WriteSidecar: false` ([data_tab_request_controller.js:586-602](../../../frontend/ui/shared/tabs/data/data_tab_request_controller.js#L586-L602)). The runtime finds the 1‑month cache, derives a 12‑month cache beside it, and the grid shows a 10×10 triangle.
2. Every server‑side recompute of the DFM (first save, a save that changes the input name, dependent propagation, refresh, Result Selection restatement) reads the input's sidecar and compares its stored period with the method's. Only an Engine‑generated dataset is rebuilt at the method's period; anything else is refused with `422 … incompatible origin period length (1; expected 12)` ([dfm_service.py:272-306](../../../frontend/app_server/services/dfm_service.py#L272-L306)). The refusal is pinned by `test_missing_sidecar_labels_do_not_hide_row_or_period_mismatches` in [test_dfm_service.py](../../../frontend/tests/test_dfm_service.py).

Extending the Engine‑only branch to hand‑entered inputs is a few lines. It must not be done on its own, for the two reasons below.

## Why the obvious fix is unsafe

### 1. The roll‑up arithmetic is wrong for a normal triangle

`_derive_triangle_cache` ([arcrho_runtime_service.py:895-930](../../../frontend/app_server/services/arcrho_runtime_service.py#L895-L930)) builds each coarse cell by summing the **same development column** across the 12 monthly origin rows of a block (cumulative), or the 12×12 square sub‑block (incremental). Those cells are valued at 12 different dates, and past the diagonal most of them are blank. An annual cell should sum each month's value at the **same valuation date**, i.e. along the calendar diagonal, which is how the Engine builds an annual triangle from source tables and how ResQ defines one.

Measured 2026-09-04 by calling the function on a synthetic 24‑month cumulative, dev‑aligned triangle in which every cell equals 100 × age in months (row *i* holds 24 − *i* cells):

| Annual cell | Current roll‑up | Calendar‑consistent |
| :--- | :--- | :--- |
| Year 1, dev 1 | 14,400 | 7,800 |
| Year 1, dev 2 | 2,400 | 22,200 |
| Year 2, dev 1 | 1,200 | 7,800 |

Every cell is off, not only the last diagonal. This already affects the Data tab view of any hand‑entered triangle opened at a coarser period; a DFM wired to it would inherit the numbers silently.

### 2. A derived cache is trusted forever

For a sidecar whose `source_kind` is `input`, `_processing_config_matches` returns `True` without looking at the file ([arcrho_runtime_service.py:562-575](../../../frontend/app_server/services/arcrho_runtime_service.py#L562-L575)), and `arcrho_tri_cache_matches` never compares the sidecar's `csv_file` with the cache being validated. So once a `@12@12` variant exists it is reported as `cache_exact` on every later load. A dataset save rewrites only the sidecar's own CSV ([dataset_service.py:927-975](../../../frontend/app_server/services/dataset_service.py#L927-L975)) and leaves sibling period variants in place. After an Excel refresh changes the monthly numbers, every 12‑month load keeps serving the old roll‑up.

The shared precedent resolver used by Result Selection has the same exposure: for a non‑Engine dataset it looks up a same‑period cache file by name on disk ([precedent_cache_service.py:82-127](../../../frontend/app_server/services/precedent_cache_service.py#L82-L127)).

## Stored length: what ResQ does, and what ArcRho already has

Investigated 2026-09-05 against the decompiled ResQ COM help (`E:\XWSpace\ResQ API Doc\resq_com_api_html_package\topics\stored*length.html`).

### ResQ's rule

ResQ keeps two lengths per dataset. `OriginLength` / `DevelopmentLength` (`PeriodLength` on a vector) are the **displayed** period. `StoredOriginLength` / `StoredDevelopmentLength` (`StoredPeriodLength`) are the **granularity of the data itself**. The rules the help states:

- The displayed length must be a whole multiple of the stored length. Any coarser view is a roll‑up.
- The stored length can be changed only while the dataset holds no data, is not calculated, and has no attached method (and, for a vector, is not an origin vector).
- Values can be entered only when the displayed length equals the stored length. Our own export macro depends on this: it sets `PeriodLength` back to `StoredPeriodLength` before `SetValuesByIndex` and restores the display afterwards ([export_reserving_class_to_resq.py:431-450](../../../python-api/macros/export_reserving_class_to_resq.py#L431-L450)).
- Origin and development are independent. ResQ's own example is a 12/12 display over `StoredDevelopmentLength = 3`, so annual data reported at September can be entered. That is the plan's "non‑square" case (1‑month or 3‑month development under coarser origins), and `_can_derive_cache` already checks the two factors separately.

### ArcRho today stores only one length, and it means two things

For a hand‑entered dataset (`source_kind: "input"`), the sidecar's `origin_length` / `development_length` (`period_length` for a vector) are the granularity of `csv_file`; the file name carries the same numbers (`<Name>@1@1@cum@dev.csv`), and every server consumer reads them that way: the DFM compatibility check, the precedent resolver, the patch mask ([dataset_service.py:1152-1170](../../../frontend/app_server/services/dataset_service.py#L1152-L1170)), and the `index.json` row.

For an Engine‑generated dataset the same fields are the *last generated* period, because a Dataset Viewer save writes the chosen lengths into them and the next open regenerates at them. The finest period available there is the source data's date granularity, which the Engine detects on every request from the first value of the Origin Date column (four digits = annual, six = monthly, [data_processing.py:154-160](../../../server-components/src/arcrho_engine/data_processing.py#L154-L160)) and nobody records.

The *displayed* length must be persisted too: the Dataset Viewer reopens at the shape the user last saved, not at the stored shape. For a hand‑entered dataset the two must therefore come apart: the stored shape is fixed by the data, the displayed shape is a saved preference on top of it.

**Decided 2026-09-05: ArcRho takes ResQ's naming.** `origin_length` / `development_length` / `period_length` become the **display** shape, the counterpart of ResQ's `OriginLength` family, and a new `stored_origin_length` / `stored_development_length` / `stored_period_length` pair carries the granularity of `csv_file`, the counterpart of `StoredOriginLength`. Every sidecar of every source kind carries the stored fields, the `index.json` row projects them, and a one‑time script backfills the share so readers require the field and never guess. The price is that every reader which today treats `origin_length` as the CSV's granularity moves to `stored_*`; Step 3 lists them.

Where the stored length comes from, by `source_kind`:

| Dataset | Stored length |
| :--- | :--- |
| `input` (hand‑entered, Excel/ArcRho‑linked) | the shape the values were entered at |
| `engine`, regenerable | the source data's date granularity per axis (1 or 12 months), recorded in the project's field mapping (Step 5) |
| `engine`, imported ResQ snapshot | its own shape (it cannot be regenerated) |
| `calculated`, method outputs | the period it was produced at |

(The `StoredPeriodLength` pair in the `/arcrho/headers` request is unrelated: a ResQ‑era leftover, always `-1`, and nothing in the Engine reads it.)

### Where ArcRho breaks the rule today

The Data tab lets a manual dataset's lengths be **raised** while it holds values; only lowering is blocked ([data_tab_persistence_controller.js:288-301](../../../frontend/ui/shared/tabs/data/data_tab_persistence_controller.js#L288-L301), pinned by `dataset_length_lock.test.mjs`). The coarser grid the runtime derives (with the wrong arithmetic from section 1) stays fully editable, because `canEditDisplayCell` checks only `source_kind` and the mask ([dataset_grid_interactions.js:482-495](../../../frontend/ui/shared/tabs/data/dataset_grid_interactions.js#L482-L495)). Save then does one of two things, both wrong:

- **With edited cells:** `save_dataset_sidecar` writes the rolled‑up‑plus‑edited grid to a new `@12@12` CSV, points `csv_file` at it and sets `origin_length` to 12 ([dataset_service.py:2016-2065](../../../frontend/app_server/services/dataset_service.py#L2016-L2065)). The monthly file is orphaned on disk and the stored granularity is silently lost.
- **Settings only:** `csv_file` keeps naming the `@1@1` file while `origin_length` becomes 12. The sidecar now contradicts its own CSV, and the DFM check in section 2 of the opening is fooled.

Fixing this is what makes the rest of the plan safe: once the stored lengths cannot drift, "the sidecar's own CSV at the sidecar's own lengths" is always the finest data, and every coarser view or method input is a derivation from it and nothing else.

### The agreed user experience

The Data tab keeps showing the **displayed** lengths as the existing adjustable Origin/Development Length controls, and shows the **stored** lengths beside them as read‑only information. The displayed shape is a saved setting: Save persists it, and the window reopens at the last saved shape rather than the stored one. The stored length is never edited directly: it is fixed by the source data for a generated dataset and by the shape the values were entered at for a manual one, and the only way it moves is a manual dataset that is still empty being saved at a new shape.

## Plan

Ten steps, each sized for one agent context (a 1M session or one workflow subagent): one goal, a short list of files, its own tests, one commit. Steps 2 to 7 are the stored‑length work and must run in order; Step 1 is independent of them; Steps 8 and 9 need Steps 1 and 7; Step 10 needs everything.

### Step 1 — Correct, shared roll‑up helper

**Goal.** One function that aggregates a finer dev‑aligned triangle to a coarser one along calendar diagonals, and the runtime's derived view uses it.

**Read first.** Section "Why the obvious fix is unsafe" above; `_can_derive_cache` and `_derive_triangle_cache` in [arcrho_runtime_service.py:868-930](../../../frontend/app_server/services/arcrho_runtime_service.py#L868-L930); the existing triangle helpers in `python-api/src/arcrho_api/`; `_empty_dataset_geometry_from_general_settings` in [dataset_service.py:1043-1100](../../../frontend/app_server/services/dataset_service.py#L1043-L1100), which is where the origin anchor is applied.

**Do.**
- [x] Add the helper in `python-api/src/arcrho_api/` next to the triangle helpers, so both the app server and the Engine bundle carry it. Cumulative and incremental input, blank cells beyond the diagonal honoured, origin and development factors independent, and a clear error when the target period is not a whole multiple of the source.
- [x] Confirm the origin alignment assumption: both geometries come from the project's General Settings origin start month, so blocks line up.
- [x] Make `_derive_triangle_cache` call the helper, so the Data tab shows what the DFM will compute.

**Tests.** A unit test pinned to a hand‑built expectation: the 24‑month synthetic above, an incremental case, and a non‑square case such as 1‑month origins with 3‑month development. The existing runtime tests still pass.

**Done when.** The 12/12 view of the synthetic triangle matches the calendar‑consistent column of the table above, cell for cell.

### Step 2 — Stored‑length fields on every sidecar

**Goal.** Every sidecar producer writes `stored_origin_length` / `stored_development_length` (triangle) or `stored_period_length` (vector), and the existing length fields are documented as the display shape. No reader changes yet.

**Read first.** The `arcrho-json-contract` skill; `engine_dataset_sidecar_contract.py`, `sidecar_core_contract.py` and `dataset_index_contract.py` in `python-api/src/arcrho_api/`; the producers: `save_dataset_sidecar` and `_create_empty_cached_dataset_impl` in `dataset_service.py`, the engine sidecar write in `arcrho_runtime_service.py` (around line 1189), `resq_migration/extractors.py`, and the calculated and method‑output writers; the existing producer parity test.

**Do.**
- [ ] Add the fields to the contract builders and validation. Values: an `input`, imported, calculated or method‑output sidecar writes its own shape; an Engine sidecar writes the field‑mapping granularity when Step 5 has landed and, until then, the request shape (a note in the code, removed in Step 5).
- [ ] Add the two stored fields to `INDEX_ROW_FIELDS` and the row projection.
- [ ] Document both meanings in [dataset.md](../../../frontend/docs/app_server/domains/dataset.md) and the contract docstrings, and that any view coarser than the stored shape is derived and never written back.
- [ ] Do not add a missing‑field fallback in readers; Step 4 backfills the share.

**Tests.** The producer parity test and the index cross‑component contract test extended with the new fields; the sidecar core validation test covers a sidecar without them being rejected once Step 4 is done (mark that assertion for Step 4 if it cannot hold yet).

**Done when.** Every producer emits the stored fields for the same logical inputs, and the parity test proves it.

### Step 3 — Readers use the right length

**Goal.** Every reader that needs the CSV's granularity reads `stored_*`; every reader that asks for a regeneration or a view keeps reading the display fields.

**Read first.** The list below and each linked range. The runtime's variant candidates parse the lengths from the CSV file name and are unaffected; the Engine's source refresh regenerates at `origin_length`, which is the display shape it should produce.

**Do.**
- [x] Move to `stored_*`: the DFM compatibility check ([dfm_service.py:272-306](../../../frontend/app_server/services/dfm_service.py#L272-L306)), the precedent resolver ([precedent_cache_service.py:82-127](../../../frontend/app_server/services/precedent_cache_service.py#L82-L127)), the patch mask ([dataset_service.py:1152-1170](../../../frontend/app_server/services/dataset_service.py#L1152-L1170)), the Excel‑link geometry ([excel_link_service.py:961-979](../../../frontend/app_server/services/excel_link_service.py#L961-L979)), and the sidecar‑load response the Data tab reads (`load_dataset_sidecar`, which must return both shapes).
- [x] Review each other service that reads `origin_length` (Bornhuetter‑Ferguson, Cape Cod, Berquist‑Sherman, Result Selection, Bootstrap, calculated datasets) and decide per read: stored when it opens the CSV, display when it asks for a regeneration. Record the decision as a one‑line comment at each read.

**Tests.** Each moved reader's existing test gains a case where display and stored differ and the reader picks the stored one.

**Done when.** No reader that opens a CSV consults the display fields to learn its shape.

### Step 4 — Backfill the share

**Goal.** Every existing sidecar on the server carries the stored fields, so readers can require them.

**Read first.** The data access restrictions in `AGENT_GUIDELINES.md`; the memory notes on the offline dependent‑walk replay (patch the project directory helpers, probe paths first); `tools/` for the existing script conventions.

**Do.**
- [x] One script under `tools/` that walks every project's sidecars and writes the stored fields once: `input`, imported, calculated and method outputs copy their current lengths (which equal the CSV's, as the file name confirms); regenerable Engine sidecars take the field‑mapping granularity when Step 5 has landed, else the current lengths with a report line. Dry‑run by default; rebuild every reserving‑class index afterwards. Landed as `tools/backfill_stored_period_lengths.py`.
- [x] Make the sidecar core validation require the fields (the Step 2 assertion). Step 2 already landed it in `_validate_period_lengths`; Step 4 only made that half of the check separately callable, as `validate_period_lengths`, so the backfill can hold a file to the period rule without also holding it to core fields another conversion fills in.

**Tests.** The script on a scratch project tree copied under the scratchpad, covering each source kind and the dry‑run report.

**Done when.** The script has been run for real against the share in dry‑run and then for real, with the report saved beside this plan, and the required‑field validation passes on a live index rebuild.

**Run 2026-09-05.** [manual_input_period_rollup_backfill_report.md](manual_input_period_rollup_backfill_report.md) records it: 10,997 sidecars across 38 projects and 116 reserving classes, every one of them without a stored shape beforehand, and not one whose recorded lengths disagreed with the `@origin@development@` in its own `csv_file` — so the "copy the current lengths" rule above held everywhere. All 116 indexes rebuilt afterwards and their rows carry the stored pair. The one thing the dry run turned up that this plan had not accounted for: 4,034 of those sidecars fall short of the shared core for older reasons, 2,927 of them with no `json_format` stamp at all, because the share has never been through `tools/migrate_persisted_json_v4.py` (step 6 of [persisted_json_contract_v4.md](persisted_json_contract_v4.md), which by decision converted only `NJ_Annual_Prod_202605_Fake`). The backfill therefore checks only the period‑length rule it owns, counts the rest, and leaves them to that conversion.

### Step 5 — Source granularity for generated datasets

**Goal.** A project records how fine its source data is, and the Engine and the app server both read that record.

**Read first.** `field_mapping_service.load_date_role_fields` and [field_mapping.md](../../../frontend/docs/app_server/domains/field_mapping.md); `table_summary` (it already bins date columns by year); the Engine detection at [data_processing.py:154-160](../../../server-components/src/arcrho_engine/data_processing.py#L154-L160); the engine sidecar write in `arcrho_runtime_service.py`; also `arcrho_engine/bundled_sources.py` and `dependent_propagation.configure_canonical_runtime`, which settle whether the Engine's calculation path can import `arcrho_api` at all — it cannot, because the canonical bundle only reaches `sys.path` inside a durable job — and `arcrho_engine/runtime_log.py`, which owns the request log.

**Do.**
- [x] When the field mapping is saved, record months‑per‑period for each date‑role column (`Origin Date`, `Development Date`) in `field_mapping.json`, by the same four‑ or six‑digit test the Engine runs. Expose it next to `load_date_role_fields`. Landed as a `source_period_months` object keyed by significance, with the rule, the date‑role names and the field name owned by the new `arcrho_api.field_mapping_contract`; `load_source_period_months` is the read side.
- [x] Backfill an existing project by running the detection once when the value is missing on load.
- [x] Engine: prefer the recorded value and log a disagreement with its own detection to `engine_requests.log`. The Engine's calculation path runs before the canonical bundle is importable, so it mirrors the rule the way it already mirrors the source‑table paths, pinned by a parity test.
- [x] Engine sidecar write: `stored_*` now comes from the field mapping; remove the Step 2 note. A sidecar of another source kind that this write restamps keeps the requested shape, because only a dataset the Engine rebuilds from the source table can claim the source's granularity.

**Tests.** A field‑mapping test for the recorded granularity and its backfill; an Engine test for the preference and the disagreement log line. Landed as `frontend/tests/test_field_mapping_source_granularity.py` (which also covers the generated sidecar's stored shape) and `server-components/tests/test_engine_source_granularity.py` (which also pins the Engine's mirror against the canonical rule).

**Done when.** A regenerable Engine sidecar in a monthly‑source project carries stored 1/1 whatever period it was last generated at.

### Step 6 — The save rule for hand‑entered datasets

**Goal.** A save can never move a populated manual dataset's stored shape or orphan its CSV.

**Read first.** `save_dataset_sidecar` ([dataset_service.py:2016-2065](../../../frontend/app_server/services/dataset_service.py#L2016-L2065)) and `_write_dataset_csv_and_sidecar`; `_empty_dataset_geometry_from_general_settings` and `_empty_dataset_values`, which build the grid a relabelled empty dataset is rewritten at; "Where ArcRho breaks the rule today" above; open decisions 2 and 3.

**Do.**
- [x] The request's lengths keep going into `origin_length` / `development_length` (the display shape). On an `input` sidecar `stored_*` and `csv_file` stay as they are. A save that names `csv_file` itself is publishing a file and states its shape, so it stays outside the rule.
- [x] Exception: when the stored CSV holds no non‑zero value, a save at a new shape moves `stored_*` to it, rewrites the CSV at that shape and deletes the old `csv_file`. Without `values` the rewritten grid is the empty geometry the project's General Settings give at the new shape, the same one a new dataset is created with.
- [x] A save that carries `values` for a populated `input` dataset at a shape other than `stored_*` is refused with `422`. It is the backstop behind Step 7 and should never be reachable from the UI.

**Tests.** Beside `test_dataset_external_links.py`: `stored_*` surviving a display‑only save, the `422` on a values save at the wrong shape, the empty‑dataset relabel, and the deleted old CSV. Landed as `frontend/tests/test_dataset_stored_shape_save.py`.

**Done when.** Both wrong outcomes listed under "Where ArcRho breaks the rule today" are impossible through the save endpoint.

**For Step 7.** An Excel-link refresh republishes the dataset at its stored lengths ([excel_link_service.py:961-985](../../../frontend/app_server/services/excel_link_service.py#L961-L985)), so it resets the saved display shape to the stored one. Harmless today; if the display shape is to survive a link refresh, that save has to carry the sidecar's display fields rather than the stored ones.

### Step 7 — Dataset Viewer: readout, dropdowns, read‑only view

**Goal.** The Data tab shows the stored shape, offers only valid display shapes, opens at the saved display shape, and makes a coarser view read‑only.

**Read first.** The `arcrho-ui-design` skill; `LEN_CHOICES` and `fillLenDropdowns` ([data_tab_request_controller.js:475-500](../../../frontend/ui/shared/tabs/data/data_tab_request_controller.js#L475-L500)); `validateManualDatasetLengthChange` ([data_tab_persistence_controller.js:288-301](../../../frontend/ui/shared/tabs/data/data_tab_persistence_controller.js#L288-L301)); `canEditDisplayCell` ([dataset_grid_interactions.js:482-495](../../../frontend/ui/shared/tabs/data/dataset_grid_interactions.js#L482-L495)); the Links‑tab read‑only rule in the persistence controller; `dataset_length_lock.test.mjs`; the memory notes on the dev UI cache and version pins.

**Do.**
- [ ] Readout: a muted caption beside each length control, "stored 1", from the sidecar‑load response. While a manual dataset is empty the caption follows the control and reads "will be stored at 12 on first save"; after the first non‑zero save it goes static.
- [ ] Dropdowns offer only whole multiples of the stored length, origin and development independently, for every source kind.
- [ ] Opening: initial lengths come from the sidecar's display fields once the sidecar loads, ahead of the lengths Project Instance passes in the launch query.
- [ ] Coarser view: grid, Links tab and paste read‑only, with a status along the lines of "Values can be entered only at the stored period (Origin 1, Development 1). Set the lengths back to edit." A length change is a dirty display setting, so Save lights up and persists it, but never writes values at a coarser shape.
- [ ] Update [dataset.md](../../../frontend/docs/ui/dataset.md) (UI doc) for the new rules and bump the `?v=` pins the memory notes describe.

**Tests.** `dataset_length_lock.test.mjs` extended for the filtered dropdowns, the read‑only coarser view, the opening shape, and a save that persists the display shape without values.

**Done when.** The manual check in "Verification" for the Dataset Viewer passes on a test project.

### Step 8 — Roll up in memory for the DFM and Result Selection

**Goal.** A method whose period is a whole multiple of a hand‑entered input's stored period rolls the input up in memory on every load.

**Read first.** Step 1's helper; `_load_source_snapshot` and the compatibility check in `dfm_service.py`; `precedent_cache_service.precedent_csv_path`; `result_selection_service._dependency_values`; `test_dfm_service.py` and `test_result_selection_service.py`.

**Do.**
- [x] In `_load_source_snapshot`, when a non‑Engine input's stored period is finer than the method's and a whole multiple of it, read the sidecar's own CSV and roll it up through the helper. No variant file is written.
- [x] Keep the `422` refusal for every other mismatch and keep the Engine branch as it is.
- [x] Do the same for Result Selection in `precedent_cache_service`, or route it through a shared loader with the DFM.

**Tests.** `test_dfm_service.py`: a monthly manual input feeding an annual DFM through save, refresh_dependents and save_propagation_roots; a stale on‑disk `@12@12` variant that must be ignored. The Result Selection equivalent.

**Done when.** The manual check in "Verification" for the DFM passes on a test project, including an Excel refresh flowing through.

### Step 9 — Stop trusting on‑disk variants of hand‑entered datasets

**Goal.** No rolled‑up copy of an `input` dataset on disk is ever served as current.

**Read first.** "A derived cache is trusted forever" above; `resolve_local_triangle_cache`, `_processing_config_matches` and `arcrho_tri_cache_matches` in `arcrho_runtime_service.py`.

**Do.**
- [x] Option A: no variant is materialised for `source_kind == "input"`. The roll‑up is registered as a recipe against the sidecar's own CSV and rebuilt on every grid read, and a coarser copy already on disk is no longer accepted as the current cache.

**Tests.** Change the monthly CSV, reload at 12 months, expect the new numbers.

**Done when.** The test above passes and no `@12@12` file is created for an `input` dataset by a Data‑tab open.

### Step 10 — Deploy and verify

**Goal.** The change is live for every user.

**Read first.** The memory notes on remote component deploy, deploy staleness, and hosted‑save fixes needing an Engine deploy.

**Do.**
- [x] Confirm the Step 4 backfill has run against the share; the redeployed readers have no fallback.
- [x] Redeploy Engine and Gateway via `server-components/deploy.py`; redeploy the Bridge if the migration extractors changed.
- [x] Verify against the deployed `arcrho_canonical` copy.
- [x] Move this plan to `docs/plans/completed/` and update [docs/plans/README.md](../README.md).

**Done when.** Both manual checks in "Verification" pass against the deployed components from the Client PC.

**Run 2026-09-05.** `deploy.py` with no arguments found bridge, engine and gateway stale and built all three on the server's own clone in two minutes, exit code 0. The 28-file payload held nothing but this plan's work. Every app‑server and shared‑library file this plan touched is byte‑identical in the deployed `arcrho_canonical` copy of both the Engine and the Gateway, all three components report fresh heartbeats, the Gateway's health endpoint answers, and nothing is stale any more. The Bridge freezes its Python rather than shipping a canonical copy, so its evidence is the rebuild itself and its heartbeat. The two manual checks in "Verification" are left for the user at the app.

## Open decisions

Recorded 2026-09-05, each with the recommended answer. Decisions 1 to 3 and 5 stand as recommended unless the user says otherwise; 4 is settled.

1. **Home of the generated stored length.** Recommended: the project's field mapping (Step 5), so the app server, the Data tab and the Engine read one small cached file. Alternative: the Engine writes a companion record beside each generated CSV, one more file per dataset on the share.
2. **When a manual dataset may change its stored length.** Recommended: while every stored cell is blank or zero, which is the rule the Data tab already applies to lowering. Alternative: only before the first save with values, which is stricter and needs a new flag.
3. **The old CSV after an empty dataset changes length.** Recommended: delete it, so the sidecar's `csv_file` is the only file that exists for the dataset. Alternative: leave it, which is what Step 9 exists to stop trusting.
4. **Naming of the persisted shapes — decided 2026-09-05.** ResQ's naming: `origin_length` / `development_length` / `period_length` are the display shape, and new `stored_origin_length` / `stored_development_length` / `stored_period_length` are the CSV's granularity. Every sidecar of every source kind carries the stored fields, the `index.json` row projects them, and a one‑time `tools/` script backfills the share so readers need no missing‑field fallback. The rejected alternative was a new `display_*` pair with the existing fields left as stored; it was cheaper but left ArcRho's names the mirror image of ResQ's.
5. **Existing projects with no recorded granularity.** Recommended: detect and backfill on the first field‑mapping load, exactly as the Engine already detects it. Alternative: assume monthly, which is wrong for an annual‑source project until someone re‑saves the mapping.
6. **ResQ import granularity.** Today the import brings a dataset in at ResQ's displayed length. If ResQ's `StoredOriginLength` is finer than the display, that detail is lost at import. Recommended: leave it for now and note it; importing at the stored length is a separate change to the Bridge and the migration extractors, and imported snapshots use their own shape as the stored length either way.

## Verification

- `frontend/tests/test_dfm_service.py`, the new roll‑up unit test, and the extended Data‑tab tests pass locally.
- Dataset Viewer manual check on a test project: enter a monthly triangle by hand, open it at 12 months and compare a cell with a hand sum along the calendar diagonal; confirm the caption reads "stored 1", the dropdowns offer only multiples, the grid is read‑only at 12, and reopening the window returns to the saved 12.
- DFM manual check: create an annual DFM on that dataset, save, refresh the Excel link with changed numbers, and confirm the DFM refreshes to the new values.

## Reproduction script

The measurement above came from a throwaway script that imports `_derive_triangle_cache` from `app_server.services.arcrho_runtime_service` (run from `frontend/` with the `arcrho_engine` venv Python, which has pandas and fastapi), writes the synthetic monthly CSV to a temp folder, derives at 12/12, and prints both frames. Rebuild it from the description when needed; it was not kept.

## Rough size

Superseded on 2026-09-05 by the per‑step estimates in the Progress table, which are measured against what the steps actually cost. This section originally read "about two and a half days of agent time across the ten sessions", which was a human‑scale figure and roughly twice the real cost; the shape it described still holds — Steps 2, 3 and 7 are the larger ones, 5, 6, 8 and 9 the smaller, 4 is mostly waiting on the share, and 10 is a deploy.
