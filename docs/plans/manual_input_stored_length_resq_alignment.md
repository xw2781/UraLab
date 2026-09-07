# Manual Input Triangles: Matching ResQ's Stored-Length Editing

Status: Broken into 7 session-sized steps on 2026-09-06; steps 1 to 4 landed the same day (4 of 7 done). The ResQ rules are established against the COM API and both design decisions are taken, so every step can run unattended.
Last updated: 2026-09-06.

## Progress

Plain-language tracking. The agent that finishes a step ticks its box, fills in the date, and leaves one short line on what a user would notice. Nothing technical goes here.

| # | Step | Done | Date | What changed for the user |
| :--- | :--- | :--- | :--- | :--- |
| 1 | An empty triangle can be told to store its development periods finer than it shows | [x] | 2026-09-06 | Saving a dataset that is still empty can now keep its figures monthly underneath a yearly view; the control to ask for that comes with the next step. |
| 2 | The Data tab shows "Stored at" beside each length, and an empty triangle's development store can be lowered there | [x] | 2026-09-06 | Each length in the Data tab now has a "Stored at" value beside it, and while a hand-entered dataset is still empty you can lower the development one so a yearly view keeps monthly figures underneath. |
| 3 | Values saved at a coarser development view land in the stored cells at their ages | [x] | 2026-09-06 | Figures saved while a yearly view of a monthly dataset is on screen now land in the monthly cells underneath, at the dates those columns stand for, and the rest of the triangle is cleared, exactly as ResQ does it. |
| 4 | Typing, paste and links work when only the development view is coarser than the store | [x] | 2026-09-06 | A hand-entered dataset shown a year at a time over finer figures can now be typed into, pasted into and linked; only a view that groups the rows is still read-only, and the status line says a save will write each figure at its own column date and clear the periods between. |
| 5 | The export macro writes a hand-entered triangle to ResQ at its stored shape | [ ] | | |
| 6 | A yearly view of a monthly triangle is pinned to ResQ's own numbers | [ ] | | |
| 7 | The server components carry the change | [ ] | | |

Overall: 4 of 7 steps done.

## How agents work this plan

- Take the first unticked step in the Progress table. One step is one context (a session or one workflow subagent), one commit.
- Read the sections between here and the Plan before starting, then only the files the step names. Do not read ahead into later steps.
- A step is done when its "Done when" list holds, its tests pass, and the commit is in. In that same commit: tick the Progress row, write the date and the one-line user note, update the "Overall" count, and update the `Status:` line at the top and this plan's row in [README.md](README.md).
- If a step turns out to need a decision that is not in "Decisions", stop, record the question there, commit that note alone, and report it rather than guessing.
- Do not start a step while the previous one is uncommitted.

## Background

Follows on from [completed/manual_input_period_rollup.md](completed/manual_input_period_rollup.md), which gave every dataset a stored shape separate from its displayed shape and made a coarser view a read-only roll-up. Three things ResQ's GUI allows on top of that are not yet in ArcRho. This note records what ResQ does (GUI observation first, then the API probe that pinned the rules down), what ArcRho does today, the gap between them, the decisions taken, and the plan.

The two production cases this has to serve:

1. **Monthly entry, annual display.** The user types or pastes a monthly triangle (its rows and columns follow the project's origin dates and Development End Date) and it is saved and shown as an annual 10×10.
2. **Annual paste into a monthly store.** The user copies or links a 10×10 annual triangle from Excel into a triangle stored at 10×120 (origin 12, development 1) and shown as 10×10. The monthly store is needed because a triangle stored at 12/12 carries the wrong development labels (see ResQ rule 1 below).

In both cases the export macro must leave the ResQ triangle identical to ArcRho's.

## What ArcRho does today

- **One control per axis.** The Data tab's Origin Length and Development Length are the displayed shape. The stored shape is read off the length list itself, where the lengths it rules out are muted, and is never edited directly ([dataset.md:106-109](../../frontend/docs/ui/dataset.md#L106-L109)).
- **The first save fixes the stored shape.** While a manual dataset holds nothing but blanks and zeros, the hint reads `This dataset is still empty: its first save stores it at 12` and follows the control; the save of a still-empty dataset is the one thing that moves the stored pair ([dataset_service.py:2117-2140](../../frontend/app_server/services/dataset_service.py#L2117-L2140), [dataset_service.py:2355-2374](../../frontend/app_server/services/dataset_service.py#L2355-L2374)).
- **Each control offers only whole multiples of its own stored length**, origin and development independently, narrowed out of the one `LEN_CHOICES` ladder ([data_tab_request_controller.js:522-560](../../frontend/ui/shared/tabs/data/data_tab_request_controller.js#L522-L560)).
- **A coarser view is read-only.** When the display is coarser than the stored shape on either axis, the grid, the Links tab and paste all refuse, with the status line `Values can be entered only at the stored period (Origin 1, Development 1). Set the lengths back to edit.` ([data_tab_preferences_controller.js:367-387](../../frontend/ui/shared/tabs/data/data_tab_preferences_controller.js#L367-L387), [data_tab_persistence_controller.js:323-340](../../frontend/ui/shared/tabs/data/data_tab_persistence_controller.js#L323-L340)).
- **The roll-up is anchored on the Development End Date** on the development axis and reads the calendar diagonal on the origin axis: a coarse cell sums, over the finer origin rows of its block, the cell each row holds at the coarse cell's valuation date ([triangle_rollup.py](../../python-api/src/arcrho_api/triangle_rollup.py), [dataset_service.py:1097-1105](../../frontend/app_server/services/dataset_service.py#L1097-L1105)). Landed 2026-09-06 (commit 46726a2a). Both axes match what ResQ's API returns (rules 2 and 8 below).
- **The CSV is at the stored shape**, cumulative or incremental as the sidecar says, and the sidecar carries the stored pair beside the display pair ([sidecar_core_contract.py:35-40](../../python-api/src/arcrho_api/sidecar_core_contract.py#L35-L40)). The ResQ import shows a dataset at its stored lengths before reading it ([extractors.py:249-282](../../python-api/migration/resq_migration/extractors.py#L249-L282)).

## What ResQ does

### Observed in the GUI (2026-09-06)

On `HPPREF\HO+DF\NJ\Legacy\HOL\Net Loss--Incurred Adjusted` and a `*5` copy of it, in a project valued at August 2026 with half-year and annual origins.

- The Edit Triangle dialog has a `Stored at` spinner beside each length. On an all-zero half-year triangle displayed at Origin 6 / Development 3, the development `Stored at` could be lowered from 3 to 1 while the display stayed at 3; the origin `Stored at` was dimmed.
- With stored 3 the columns read 3m, 6m, 9m, …; with stored 1 the same display of 3 read 2m, 5m, 8m, …. (That project's development length is 3, so its period grid ends at August; see rule 2.)
- A 10×10 annual cumulative triangle pasted at Origin 12 / Development 12 into a triangle stored at Development 1 was accepted. Shown at Development 4 the values sit only in the 8m, 20m, 32m, … columns; shown incremental each value appears at 8m with its negative at 12m.
- Copying and pasting requires the displayed origin length to equal the stored origin length. Only the development axis is relaxed.

### Established against the COM API (2026-09-06)

Probed on the Server PC in `NJ_Annual_Prod_202605_Fake`, class `HPPREF\HO+DF\NJ\Legacy\HOL`, with throwaway triangles of the non-unique `Net Loss - ad hoc` type and one `F 00 - Ultimate Net Loss ` vector, all deleted afterwards. The project has annual origins 2017–2026 (origin start 2017-01-01) and a Development End Date of 2026-05-31, so its newest cell is 113 months old and a yearly view is valued at 5, 17, 29, … 113 months. The probe is [resq_stored_length_probe.py](../../tools/resq_stored_length_probe.py) and re-creates every result below; every generated triangle in the class is stored 1/1 and shown 12/12 with those 5m…113m labels, and the one hand-made triangle stored 12/12 shows 12m…120m.

1. **The stored grid runs forward from each origin period's start** in steps of the stored length, and its last period may be partial. Stored 12 in this May-valued project labels the newest cell 120m (the Jun 2025–May 2026 period, cut short at May); stored 3 labels it 114m; stored 1 labels it 113m and has no short period. This is the ResQ limitation behind case 2: a triangle stored at 12/12 can only ever read 12m, 24m, …, whatever the Development End Date.
2. **A coarser display groups stored cells from the newest cell backwards** ("ResQ always crunches in the development dimension from the end", per its own help on non-standard triangles). Stored 1 shown at 12 reads 5m, 17m, … 113m; stored 3 shown at 12 reads 6m, 18m, … 114m. A cumulative display column reads the stored cell at the column's age and nothing else: a monthly store filled with `100000·row + age` read 100005, 100017, 100029, … at the annual display.
3. **On an empty triangle a display put moves the stored length with it**, on both axes, saved or not, with no multiple check: setting `DevelopmentLength` to 4 on a triangle stored at 3 simply makes the store 4. `StoredDevelopmentLength` may then be lowered or raised to any factor of the display length (`The stored development length must be a factor of development length.` otherwise). `StoredOriginLength` has no setter in the type library (`Invalid number of parameters.`): the only way to choose the origin store is to set `OriginLength` while the triangle is empty. On a never-saved triangle an `OriginLength` put also resets the development length and its store to 1, so create in the order `OriginLength`, `DevelopmentLength`, `StoredDevelopmentLength`, `Save`.
4. **Once a triangle holds saved data, the display must be a whole multiple of the store** (`The stored origin length must be a factor of the origin length.` / `…development…`) and `StoredDevelopmentLength` is refused (`The stored development length may not be set in this triangle.`). "Holds data" is judged on saved non-zero values: an unsaved `SetValuesByIndex` does not lock the store, a triangle saved with explicit zeros in every cell counts as empty, and `ClearData` unlocks it at once without a `Save`.
5. **A write at a development display coarser than the store is accepted and rebuilds the whole triangle from the display grid, at the moment of the write.** `SetValuesByIndex` at an annual display over a filled monthly store: the written display cell's cumulative value lands in the stored cell at that column's age, every other stored cell of the whole triangle — every row, written or not — becomes cumulative 0, and the other display-age cells keep their values (one cell written over 590 filled stored cells: 535 zeroed, 54 kept, 1 changed; visible before `Save`). Shown incremental, each value appears at its age with its negative at the next stored age. An incremental display behaves the same way: the running sum of the entered increments is stored as the cumulative at each display age.
6. **The origin axis and vectors are strict.** A write at a coarse origin display is refused: `You cannot enter data unless the display origin length matches the data storage origin length (1).` A write at a coarse vector period is refused: `You cannot enter data unless the displayed period length corresponds with the data storage period length (1).` `StoredPeriodLength` cannot be set on an origin vector (`You may not set the stored period length on this vector.`); it follows `PeriodLength` while the vector is empty, exactly as rule 3. An origin vector shown at 12 over a monthly store reads the sum of the 12 months.
7. **`SetValues(originDate, ageMonths, value)` is display-level too.** At an annual display, age 10 maps to the column spanning ages 6–17 and writes the 17m stored cell; age 17 written first was overwritten by it. Writing a particular stored cell requires showing the triangle at the stored length. The same goes for reads: `Values(date, month)` returns the display column that contains the month.
8. **An annual view of a monthly-origin store is calendar-anchored.** With origin 1 / development 1 filled with `cum(origin month k, age d) = 1000k + d`, the annual row for year Y at age a equals the sum over its origin months m = 0…11 of `cum(month m, a − m)` for `a − m ≥ 1` — every origin month contributes the cell it holds at the column's calendar end (0 mismatches over 55 cells). Rows read 2017…2026 and columns 5m…113m; an origin-12 / development-1 view of the same store is allowed (row widths 113, 101, … 5). This is what [triangle_rollup.py](../../python-api/src/arcrho_api/triangle_rollup.py) already computes. The annual grid ResQ returned for that fill, used by step 6:

   | Origin | 5m | 17m | 29m | 41m | 53m | 65m | 77m | 89m | 101m | 113m |
   | :--- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
   | 2017 | 15015 | 78138 | 78282 | 78426 | 78570 | 78714 | 78858 | 79002 | 79146 | 79290 |
   | 2018 | 75015 | 222138 | 222282 | 222426 | 222570 | 222714 | 222858 | 223002 | 223146 | |
   | 2019 | 135015 | 366138 | 366282 | 366426 | 366570 | 366714 | 366858 | 367002 | | |
   | 2020 | 195015 | 510138 | 510282 | 510426 | 510570 | 510714 | 510858 | | | |
   | 2021 | 255015 | 654138 | 654282 | 654426 | 654570 | 654714 | | | | |
   | 2022 | 315015 | 798138 | 798282 | 798426 | 798570 | | | | | |
   | 2023 | 375015 | 942138 | 942282 | 942426 | | | | | | |
   | 2024 | 435015 | 1086138 | 1086282 | | | | | | | |
   | 2025 | 495015 | 1230138 | | | | | | | | |
   | 2026 | 555015 | | | | | | | | | |

### Answers to the questions the GUI left open

1. **Paste into a non-empty finer triangle:** cleared. The whole triangle is rebuilt from the display grid; stored cells between the display ages go to cumulative 0 and only the display-age cells survive (rule 5).
2. **Which stored cell receives the value:** the one whose age equals the display column's age. Under ResQ's end-anchored grouping that is always the newest stored cell of the column's block, so the two readings coincide by construction (5m, 17m, … in the May project; 8m, 20m, … in the August one).
3. **Incremental display:** the running sum is stored as the cumulative at each display age; the stored layout is the same as a cumulative paste (rule 5).
4. **What counts as empty:** no saved non-zero value; explicit zeros are empty; `ClearData` unlocks immediately; the refusal reads `The stored development length may not be set in this triangle.` (rule 4).
5. **Origin stored length:** no API setter, ever. It is fixed by setting `OriginLength` while the triangle is empty, which the GUI does implicitly, hence the dimmed spinner (rule 3).
6. **Display length after a stored-length change:** nothing is ever refused while the triangle is empty because a display put resyncs the store; once data is saved the display must be a multiple of the store (rules 3 and 4).
7. **Vectors:** strict on both counts, like the origin axis (rule 6).
8. **API paste path:** `SetValuesByIndex` and `SetValues` both accept a coarse development display and write what the GUI paste writes; the mapping lives in ResQ's store, not in the GUI (rules 5 and 7).

## The gap

| Behaviour | ResQ | ArcRho today |
| :--- | :--- | :--- |
| Choose the stored shape | The display control sets it while the triangle is empty; `Stored at` can then lower the development store to any factor | The first save's display shape becomes the store; no way to store finer than the display |
| Lower the stored development length below the display | Yes, while empty (rule 3) | No |
| Set the stored origin length | Only through the display control while empty | Same in effect |
| Write at a development display coarser than the store | Yes; the display cell's cumulative goes to the stored cell at its age and every other stored cell of the triangle becomes 0 (rule 5) | Refused, whole grid read-only |
| Write at an origin display coarser than the store, or a vector at a coarse period | Refused (rule 6) | Refused |
| Development labels | Stored grid forward from the origin start, grouped from the newest cell (rules 1–2) | Ages counted back from the Development End Date on every display |
| Roll-up values | Stored cell at the column's age; calendar-diagonal sum over origin months (rules 2, 8) | Same |

## Decisions

All taken on 2026-09-06; nothing is open.

1. **A coarse development paste zeroes the rest of the store, as ResQ does.** Each display cell's cumulative value is written to the stored cell at its age and every other stored cell of the triangle becomes cumulative 0 (rule 5). Keeping the in-between cells was considered and rejected: it would be a silent departure from ResQ, and the export rebuilds the ResQ triangle from ArcRho's store, so the two would no longer match.
2. **The export macro may set ResQ's stored development length on an empty target.** After `ClearData` the ResQ triangle is empty by rule 4, so the macro aligns `StoredDevelopmentLength` with ArcRho's before writing rather than skipping the dataset. The stored origin length can never be changed on an existing triangle, so a mismatch there is a reportable skip.
3. **ArcRho keeps its own development labels.** Every display is labelled by age counted back from the Development End Date. Mirroring ResQ's 12m…120m labels on a 12/12 store would reproduce the limitation case 2 exists to avoid; on a monthly store, which both production cases use, the two systems already agree.
4. **No separate origin `Stored at` control.** As in ResQ, the Origin Length control sets the origin store while the dataset is empty; the Data tab shows the origin `Stored at` value dimmed, never editable.

Open decisions: none.

## Plan

Steps 1→2 and 3→4 are ordered pairs. Steps 1 and 3 both edit the dataset save and must not run at the same time; steps 5 and 6 are independent of everything else and of each other; step 7 is last.

### Step 1 — App server: an empty triangle can be stored finer than it shows

**Goal.** The sidecar save accepts a requested stored development length for a manual triangle that holds no value, so the store can be any factor of the display (ResQ rules 3–4). Nothing else about the save changes.

**Read first.** ResQ rules 3, 4 and 6 and Decision 4 above; [dataset_service.py:2025-2039](../../frontend/app_server/services/dataset_service.py#L2025-L2039) (what "holds a value" means), [dataset_service.py:2117-2215](../../frontend/app_server/services/dataset_service.py#L2117-L2215) (the empty-dataset relabel path that moves the stored pair today) and [dataset_service.py:2355-2375](../../frontend/app_server/services/dataset_service.py#L2355-L2375) (the pair reported back); the save request model and route at [dataset_router.py:282](../../frontend/app_server/api/dataset_router.py#L282); [sidecar_core_contract.py:30-45](../../python-api/src/arcrho_api/sidecar_core_contract.py#L30-L45) and [sidecar_core_contract.py:108-170](../../python-api/src/arcrho_api/sidecar_core_contract.py#L108-L170) (the stored fields already exist; no new field is added); [test_dataset_stored_shape_save.py](../../frontend/tests/test_dataset_stored_shape_save.py); the request model itself in [dataset.py:156-190](../../frontend/app_server/schemas/dataset.py#L156-L190), the CSV name builder [helpers.py:112-131](../../frontend/app_server/helpers.py#L112-L131) (the file is named for the shape it is written at), and the hand-entered save paragraph of [dataset.md](../../frontend/docs/app_server/domains/dataset.md) that the change updates. Memory notes: `python-test-runner`, `propagation-hold-and-test-isolation` (a service test that saves must use the propagation workspace stub). Skill: `arcrho-json-contract`.

**Do.**

- [x] Add an optional `stored_development_length` to the save request model and to `_save_dataset_sidecar_impl`; a request without it behaves exactly as today.
- [x] While the dataset's CSV holds no non-zero value, the stored pair becomes (`origin_length`, `stored_development_length`); a value that does not divide `development_length` is a 400 reading `The stored development length must be a factor of the development length.`
- [x] Once the CSV holds a value, a requested stored length that differs from the recorded one is a 400 reading `The stored development length cannot be changed while the dataset holds values.`
- [x] An empty CSV written by that save is at the stored shape.
- [x] A vector ignores the field (rule 6).

**Tests.** [test_dataset_stored_shape_save.py](../../frontend/tests/test_dataset_stored_shape_save.py) gains: an empty triangle saved at display 12/12 with stored development 1 records and reports back (12, 1); a non-factor is refused; a triangle holding a value refuses a change; a vector ignores the field.

**Done when.** A save of an empty manual triangle at display 12/12 with a stored development length of 1 returns the stored pair (12, 1), the sidecar on disk carries it, and the four new tests pass alongside the existing file.

### Step 2 — Data tab: `Stored at` beside each length

**Goal.** The Data tab shows ResQ's layout: a dimmed `Stored at` value beside Origin Length, and a `Stored at` control beside Development Length that is live only while the dataset is empty and offers the factors of the current display length. Case 2 becomes possible from the UI.

**Read first.** [FRONTEND_AGENT_GUIDELINES.md](../../frontend/FRONTEND_AGENT_GUIDELINES.md); the `arcrho-ui-design` skill; Decision 4 and ResQ rule 3; [dataset.md:106-109](../../frontend/docs/ui/dataset.md#L106-L109); [data_tab_request_controller.js:522-560](../../frontend/ui/shared/tabs/data/data_tab_request_controller.js#L522-L560) (the length ladder and its narrowing), [data_tab_persistence_controller.js:270-360](../../frontend/ui/shared/tabs/data/data_tab_persistence_controller.js#L270-L360) (the stored pair, the pending test, the hints) and the save payload built in `saveDatasetSidecarForCurrentContext` at [data_tab_persistence_controller.js:832](../../frontend/ui/shared/tabs/data/data_tab_persistence_controller.js#L832); the top-bar markup that holds `originLenSelect` in [dataset_viewer_view.js](../../frontend/ui/dataset_viewer/dataset_viewer_view.js) and [dfm.html](../../frontend/ui/method_pages/dfm/dfm.html); [dataset_length_lock.test.mjs:276-432](../../frontend/tests/dataset_length_lock.test.mjs#L276-L432). Also needed and added while working the step: the length-select change handlers and the deps list in [data_tab_controls.js](../../frontend/ui/shared/tabs/data/data_tab_controls.js) (where the new control's change is wired), the `LEN_DROPDOWN_CONFIG` ladder in [data_tab_inputs_controller.js:8-20](../../frontend/ui/shared/tabs/data/data_tab_inputs_controller.js#L8-L20), the top-bar control styling in [dataset_viewer.css](../../frontend/ui/dataset_viewer/dataset_viewer.css) and [dfm.css](../../frontend/ui/method_pages/dfm/dfm.css), the cache-version pins in [shared_tab_surfaces.test.mjs](../../frontend/tests/shared_tab_surfaces.test.mjs) and [color_theme.test.mjs](../../frontend/tests/color_theme.test.mjs), and [changes/README.md](../../frontend/changes/README.md) for the release fragment. Memory notes: `frontend-node-test-suite`, `arcrho-dev-ui-cache-restart`, `theme-css-version-pins`, `electron-ui-screenshot-check`.

**Do.**

- [x] Add a `Stored at` value beside each length control in both hosts' top bars. The origin one is always read-only. The development one is a select of the factors of the current display length, drawn from the same `LEN_CHOICES` ladder, enabled only while the dataset holds no value (the existing pending test) and dimmed otherwise with the tooltip `Stored at can be changed only while the dataset is empty.`
- [x] A display-length change on an empty dataset resets both `Stored at` values to the display (rule 3); lowering the development one leaves the display alone.
- [x] The save sends `stored_development_length`; the load and the save response fill both values; the length-list narrowing keeps reading the stored pair.
- [x] Retire the `This dataset is still empty: its first save stores it at …` hint in favour of the new values; update [dataset.md](../../frontend/docs/ui/dataset.md) to describe the control.

**Tests.** [dataset_length_lock.test.mjs](../../frontend/tests/dataset_length_lock.test.mjs) gains: the origin `Stored at` is read-only; the development one offers the factors and is live only while empty; a display change on an empty dataset resyncs both; the save payload carries the field. The test that reads the stored period off the list is adjusted to the new layout.

**Done when.** On an empty manual triangle a user sets 12/12, lowers the development `Stored at` to 1 and saves; the sidecar records (12, 1) and the grid still shows 10×10 with the project's valuation-anchored labels. The node suite passes with no new failures against the baseline.

### Step 3 — App server: values saved at a coarser development view

**Goal.** A save whose values are at a development display coarser than the store writes each display cell's cumulative value into the stored cell at its age and zeroes every other stored cell (Decision 1, rule 5). A coarser origin display is still refused (rule 6).

**Read first.** ResQ rules 2, 5 and 6 and Decision 1; [triangle_rollup.py:1-60](../../python-api/src/arcrho_api/triangle_rollup.py#L1-L60) and [triangle_rollup.py:137-200](../../python-api/src/arcrho_api/triangle_rollup.py#L137-L200) (the valuation arithmetic the write must invert); the values-to-CSV part of the save at [dataset_service.py:2117-2260](../../frontend/app_server/services/dataset_service.py#L2117-L2260) and the roll-up read at [dataset_service.py:1097-1105](../../frontend/app_server/services/dataset_service.py#L1097-L1105); [test_manual_dataset_rollup_view.py](../../frontend/tests/test_manual_dataset_rollup_view.py) (the fake-project and General Settings patch technique); [test_triangle_rollup.py](../../python-api/tests/test_triangle_rollup.py). Also needed and added while working the step: the hand-entered save paragraph of [dataset.md](../../frontend/docs/app_server/domains/dataset.md), which states the refusal this step relaxes, and [changes/README.md](../../frontend/changes/README.md) for the release fragment. Memory notes: `triangle-rollup-valuation-anchor`, `propagation-hold-and-test-isolation`, `python-test-runner`.

**Do.**

- [x] Add `scatter_triangle` beside `rollup_triangle` in [triangle_rollup.py](../../python-api/src/arcrho_api/triangle_rollup.py), sharing `_valued_at`: each coarse cell's cumulative value goes to the stored cell at its valuation month; every other stored cell is 0; an incremental display is turned into cumulative first (rule 5), and the result is written as the sidecar's own cumulative or incremental convention says.
- [x] In the sidecar save, when values arrive at a development display coarser than the stored one and an origin display equal to it, scatter them into the stored shape before the CSV is written; when the origin display is coarser, refuse with a 400 reading `Values can be entered only at the stored origin period.`
- [x] The stored pair reported back is unchanged by such a save.
- [x] Restate the hand-entered save rule in [dataset.md](../../frontend/docs/app_server/domains/dataset.md) and add the release fragment.

**Tests.** [test_triangle_rollup.py](../../python-api/tests/test_triangle_rollup.py) gains scatter cases pinned to the probe's numbers in a project valued 113 months after its origin start: the annual grid `1000·row + column` scattered into a 12/1 store lands at months 5, 17, … 113 with zeros elsewhere, the incremental case stores the running sums, and scatter followed by roll-up returns the input. [test_dataset_stored_shape_save.py](../../frontend/tests/test_dataset_stored_shape_save.py) gains: a save at 12/12 over a 12/1 store writes a 10×113 CSV with those values; an origin-coarse save is refused.

**Done when.** The annual grid saved at 12/12 into a 12/1 store reads back identical at 12/12, and row 1 of the CSV is non-zero only at months 5, 17, … 113.

### Step 4 — Data tab: editing at a coarser development view

**Goal.** The grid, paste and the Links tab accept values when only the development display is coarser than the store; the origin axis stays read-only with a message that names it.

**Read first.** ResQ rules 5–6 and Decision 1; [data_tab_persistence_controller.js:323-340](../../frontend/ui/shared/tabs/data/data_tab_persistence_controller.js#L323-L340); [data_tab_preferences_controller.js:367-387](../../frontend/ui/shared/tabs/data/data_tab_preferences_controller.js#L367-L387); the read-only check in [dataset_run_controller.js:553](../../frontend/ui/shared/dataset/dataset_run_controller.js#L553) and the paste guard in [dataset_grid_interactions.js](../../frontend/ui/shared/tabs/data/dataset_grid_interactions.js) (search for `isDatasetReadOnly`); [dataset.md:109](../../frontend/docs/ui/dataset.md#L109); [dataset_length_lock.test.mjs:403-432](../../frontend/tests/dataset_length_lock.test.mjs#L403-L432). Also needed and added while working the step: the run controller's post-load status line at [dataset_run_controller.js:530-537](../../frontend/ui/shared/dataset/dataset_run_controller.js#L530-L537) and its wiring in [data_tab_host_controller.js:890-894](../../frontend/ui/shared/tabs/data/data_tab_host_controller.js#L890-L894) (where the new note is passed in), the cache-version chain through [data_tab_controller.js](../../frontend/ui/shared/tabs/data/data_tab_controller.js) to both hosts and its pins in [shared_tab_surfaces.test.mjs](../../frontend/tests/shared_tab_surfaces.test.mjs) and [color_theme.test.mjs](../../frontend/tests/color_theme.test.mjs), and [changes/README.md](../../frontend/changes/README.md) for the release fragment. Memory notes: `frontend-node-test-suite`, `arcrho-dev-ui-cache-restart`, `theme-css-version-pins`.

**Do.**

- [x] Split the coarser-than-stored test by axis; `isDatasetReadOnly` uses the origin axis only, and the message becomes `Values can be entered only at the stored origin period (Origin 1). Set the origin length back to edit.` (a vector keeps its `Period` wording).
- [x] The save keeps sending the grid at the display shape; the server scatters it (step 3). The reload after save shows the roll-up of the new store, which already happens.
- [x] While a development display is coarser than the store, the status line says in one sentence that values are stored at their column ages and the months between are cleared.
- [x] Update [dataset.md](../../frontend/docs/ui/dataset.md).

**Tests.** [dataset_length_lock.test.mjs](../../frontend/tests/dataset_length_lock.test.mjs) gains: a development-coarse view is editable and shows the one-sentence note; an origin-coarse view is read-only with the origin message; a vector is unchanged.

**Done when.** With a 12/1 store shown at 12/12, typing and pasting work and Save persists the values at their ages; at a 24/12 display the grid refuses with the origin message. The node suite passes with no new failures against the baseline.

### Step 5 — Export macro: write at the stored shape

**Goal.** The export writes a hand-entered triangle to ResQ at ArcRho's stored shape, aligning ResQ's stored development length first (Decision 2), so the ResQ store ends identical to ArcRho's for every stored/display combination.

**Read first.** ResQ rules 3–7 and Decision 2; [export_reserving_class_to_resq.py:405-458](../../python-api/macros/export_reserving_class_to_resq.py#L405-L458) (the triangle path with its stale comment, and the vector path that already switches to the stored length); [extractors.py:249-282](../../python-api/migration/resq_migration/extractors.py#L249-L282) (the import's display switch this mirrors); [python-api/macros/README.md](../../python-api/macros/README.md) (version, backup, release-note rules); [test_export_reserving_class_macro.py:83-140](../../python-api/tests/test_export_reserving_class_macro.py#L83-L140); [resq_stored_length_probe.py](../../tools/resq_stored_length_probe.py) for a live check. Memory notes: `shared-macro-library-deploy`, `resq-stored-length-rules`, `macro-tests-poisoned-by-test-resq-dfm-v2`.

**Do.**

- [ ] `_write_triangle_values`: read the sidecar's stored pair. If ResQ's `StoredOriginLength` differs, raise `ExportSkipped("stored_origin_mismatch", …)` naming both values. Otherwise `ClearData`, set `DevelopmentLength` to the sidecar's display length and `StoredDevelopmentLength` to its stored one when they differ (rule 4 allows it after `ClearData`), show the triangle at the stored pair, write the CSV by index, put the display pair back, `Save`.
- [ ] Confirm the sequence once on the fake project with the probe (`--keep`, then inspect and delete), since rule 3's resync after `ClearData` on a saved triangle was not probed.
- [ ] Replace the stale "captured at the sidecar display lengths" comment; bump `# Version` and `MACRO_VERSION`; write the release note; archive the backup copy; publish to the shared library.

**Tests.** [test_export_reserving_class_macro.py](../../python-api/tests/test_export_reserving_class_macro.py) gains: a 12/1 sidecar writes 10×113 values at display 1 and restores 12; a matching store writes as before; a stored-origin mismatch is a skip with its message; the stored development length is set only when it differs.

**Done when.** Exporting a triangle stored 12/1 and shown 12/12 leaves the ResQ triangle with the same stored cells as ArcRho's CSV, the macro's tests pass, and the published library copy carries the new version.

### Step 6 — Roll-up regression fixture from the ResQ probe

**Goal.** Pin ArcRho's roll-up to ResQ's own numbers for a monthly-origin store (rule 8) so the origin diagonal cannot drift.

**Read first.** ResQ rule 8 and its grid above; the `t6_monthly_origin` section of [resq_stored_length_probe.py](../../tools/resq_stored_length_probe.py); [triangle_rollup.py](../../python-api/src/arcrho_api/triangle_rollup.py); [test_triangle_rollup.py](../../python-api/tests/test_triangle_rollup.py). Memory note: `python-test-runner`.

**Do.**

- [ ] Add a test that builds the 120×113 fill `cum(origin month k, age d) = 1000k + d` in code (no fixture file), rolls it up to 12/12 with 113 valuation months, and asserts the 55-cell grid from rule 8.
- [ ] Add the development-axis case from rule 2: a 12/1 store filled with `100000·row + age` reads `100005, 100017, …` at 12/12.

**Tests.** [test_triangle_rollup.py](../../python-api/tests/test_triangle_rollup.py) gains the two cases.

**Done when.** Both tests pass at HEAD, and the origin one fails when the diagonal read is replaced by a plain block sum.

### Step 7 — Deploy

**Goal.** The server components that bundle the app server carry steps 1 and 3.

**Read first.** [AGENT_GUIDELINES.md](../../AGENT_GUIDELINES.md) "Component Build and Deploy"; [component-deployment-authorization.md](../../agent-instructions/component-deployment-authorization.md). Memory notes: `remote-component-deploy`, `hosted-save-fix-needs-engine-deploy`, `deploy-staleness-is-mtime-based`, `bridge-restart-after-deploy`.

**Do.**

- [ ] Check the Build Listener's heartbeat names its own clone, then run `python server-components/deploy.py` with no arguments so the stale set (Engine, Gateway, and Bridge, which all bundle the app server) is derived rather than guessed.
- [ ] Verify the deployed canonical app-server copy holds the scatter write.
- [ ] Note in the Progress row that the frontend UI ships with the next app build.

**Done when.** `deploy.py` exits 0 for every stale component, and a save at a coarse development view from a Client PC persists through the hosted save.

## Rough size

Seven sessions: one each for steps 1–6 (step 6 is short) and a deploy session. Steps 5 and 6 can run beside the others; the rest are two ordered pairs followed by the deploy.
