# ResQ stored and display lengths: every interaction case

What ResQ does when a triangle or a vector is created, reshaped, typed into,
pasted into and read back, and what ArcRho does about each rule.

Established against the live ResQ COM API on the Server PC, in the fake project
`NJ_Annual_Prod_202605_Fake`, reserving class `HPPREF\HO+DF\NJ\Legacy\HOL`,
using throwaway datasets of the non-unique `Net Loss - ad hoc` and
`F 00 - Ultimate Net Loss ` types. Every rule below carries a case identifier
such as **B2**; those are the case ids in the `CASES` table of
[resq_stored_length_probe.py](../../tools/resq_stored_length_probe.py), so any
rule can be re-run and re-checked:

```
py -3.10 tools/resq_stored_length_probe.py                 # every case, then delete what it made
py -3.10 tools/resq_stored_length_probe.py --only B2 D1    # one or more cases or groups
py -3.10 tools/resq_stored_length_probe.py --keep          # leave the objects for a look in the ResQ GUI
py -3.10 tools/resq_stored_length_probe.py --json temp/probe.json
```

The probe creates, saves and deletes real datasets, so it must only ever be
pointed at the fake project. Anything below that carries no case identifier was
seen in the ResQ GUI rather than measured, and is marked as such. Last run
2026-09-07.

**The project this was probed in** (case **A0**). Origins are annual from
2017-01-01 to 2026-12-31 (10 of them) and the Development End Date is
2026-05-31, so the oldest origin's newest cell is **113 months** old. Every age
below is that project's; the arithmetic, not the numbers, is what carries to
another project.

---

## 1. The two shapes

Every ResQ triangle has two shapes, and so does every ArcRho dataset.

| | ResQ | ArcRho |
| :--- | :--- | :--- |
| **Stored** — the grid the figures are really held in | `StoredOriginLength`, `StoredDevelopmentLength` | `stored_origin_length`, `stored_development_length` in the sidecar; the CSV is written at this shape and named for it |
| **Displayed** — the grid on screen | `OriginLength`, `DevelopmentLength` | `origin_length`, `development_length` in the sidecar; the Data tab's two length controls |

A vector has one of each: `PeriodLength` / `StoredPeriodLength` against
`period_length` / `stored_period_length`.

Displaying a dataset more coarsely than it is stored is a **roll-up**, and it is
the whole reason the two shapes exist. The production case is a 10x10 annual
triangle kept over a monthly store, because only a store of 1 puts a column on
the valuation date (§2).

`Cumulative`, `Calendarised` and `Transposed` are not part of either shape
(**A8**, **D2**).

---

## 2. The label arithmetic

**The stored grid runs forward from each origin period's start**, in steps of the
stored length, so its columns are `k x store` and its last period may be short
(**A9**):

| Store | Columns | Development labels |
| :--- | ---: | :--- |
| 1 | 113 | 1m, 2m, 3m … 113m |
| 2 | 57 | 2m, 4m, 6m … 114m |
| 3 | 38 | 3m, 6m, 9m … 114m |
| 4 | 29 | 4m, 8m, 12m … 116m |
| 6 | 19 | 6m, 12m, 18m … 114m |
| 12 | 10 | 12m, 24m, 36m … 120m |

**Only a store of 1 lands its newest column on the valuation date.** A store of
12 labels its newest column 120m — the June 2025 – May 2026 period, cut off at
May. That is the ResQ limitation the monthly store exists to avoid.

**A coarser display groups the stored cells from the newest one backwards.**
ResQ's own help puts it as "ResQ always crunches in the development dimension
from the end", and *the end* is the store's own newest column, not the valuation
date (**D6**):

| Store | Shown at | Columns | Development labels |
| :--- | ---: | ---: | :--- |
| 2 | 4 | 29 | 2m, 6m, 10m … 114m |
| 2 | 6 | 19 | 6m, 12m, 18m … 114m |
| 2 | 12 | 10 | 6m, 18m, 30m … 114m |
| 3 | 6 | 19 | 6m, 12m, 18m … 114m |
| 3 | 12 | 10 | 6m, 18m, 30m … 114m |
| 6 | 12 | 10 | 6m, 18m, 30m … 114m |

Over a **store of 1**, whose newest column is the valuation date itself, that
comes out as the table the production case uses (**D1**):

| Display | Columns in row 1 | Development labels | Row widths, 2017 … 2026 |
| :--- | ---: | :--- | :--- |
| 1 | 113 | 1m, 2m, 3m … 113m | 113, 101, 89, 77, 65, 53, 41, 29, 17, 5 |
| 2 | 57 | 1m, 3m, 5m … 113m | 57, 51, 45, 39, 33, 27, 21, 15, 9, 3 |
| 3 | 38 | 2m, 5m, 8m … 113m | 38, 34, 30, 26, 22, 18, 14, 10, 6, 2 |
| 4 | 29 | 1m, 5m, 9m … 113m | 29, 26, 23, 20, 17, 14, 11, 8, 5, 2 |
| 6 | 19 | 5m, 11m, 17m … 113m | 19, 17, 15, 13, 11, 9, 7, 5, 3, 1 |
| 12 | 10 | 5m, 17m, 29m … 113m | 10, 9, 8, 7, 6, 5, 4, 3, 2, 1 |

So over a store of 1 a row's column count is
`floor((newest_age - 1) / display) + 1` and its ages are
`newest_age - k x display` while that stays positive. Over a coarser store,
`newest_age` is the store's own last column, `ceil(113 / store) x store`.

`GetDevelopmentDate(row, column)` returns the calendar date the column is valued
at — 2017-05-31, 2018-05-31 … 2026-05-31 for the annual display over a monthly
store — which is what makes the arithmetic checkable rather than guessed
(**D1**).

**A cumulative display column reads the one stored cell at its own age**, and
nothing else. A monthly store filled with `100000 x row + age` reads 100005,
100017, 100029 … at the annual display (**D5**). Read **incrementally**, a
coarse column is the difference of the cumulative view, not a sum of the stored
increments in its block (**D7**).

**A coarser origin display reads the calendar diagonal**: each coarse row sums,
over the finer origin rows it covers, the cell each of those rows holds at the
coarse column's valuation date (**C2**, **D4**). A monthly-origin store filled
with `cum(origin month k, age d) = 1000k + d` rolls up to an annual grid with
zero mismatches against that rule over all 55 cells.

---

## 3. Group A — creating a dataset and choosing its shape

| Case | Rule | ResQ's own words when it refuses |
| :--- | :--- | :--- |
| **A1** | A newly added triangle starts at the project's own lengths, stored equal to displayed: 12/12 here. | |
| **A2** | A display-length put of the value already in place does nothing at all. A real origin change **keeps** the development length as long as it still divides the new origin length: display 6 under origin 12 survives a change to origin 24. | |
| **A3** | While the triangle holds no saved value, **changing** a display length moves the matching stored length to the same value, with no factor check: display 4 over a store of 3 simply makes the store 4. Putting the value that is already there does **not** move it. | |
| **A4** | `StoredDevelopmentLength` may then be set to any factor of the display length. | `The stored development length must be a factor of development length.` — and `Division by zero` for 0. |
| **A5** | `StoredOriginLength` has no setter in the type library, ever. The origin store is chosen only by setting `OriginLength` while the dataset is empty. | `Invalid number of parameters.` |
| **A6** | The display development length must be a factor of the display origin length. At origin 12: 1, 2, 3, 4, 6, 12 are allowed and 5, 7, 24 are not. At origin 6, 12 is not. | `The development length must be a factor of the origin length` |
| **A7** | Saving an empty triangle records the stored pair; it survives an `UnloadChildren` and a fresh read. | |
| **A8** | An origin change that leaves the development length no longer dividing it resets that length **and its store** to 1: origin 12 → 6 under a display of 12 lands on O6/D1 stored O6/D1. Toggling `Cumulative` changes nothing. | |
| **A9** | The stored grid at each stored length, the first table in section 2. | |

Because a display change moves the store with it (**A3**) and an origin change
can reset the development length (**A8**), the order that always lands where you
meant is `OriginLength`, then `DevelopmentLength`, then `StoredDevelopmentLength`,
then `Save`.

---

## 4. Group B — entering data

| Case | Rule | ResQ's own words when it refuses |
| :--- | :--- | :--- |
| **B1** | Writing at display == store is the plain case: 55 cells in, 55 cells out. | |
| **B2** | **A 10x10 annual paste into a triangle stored at development 1 is accepted, and the stored length does not move — before the save or after it.** The 55 entered figures land in 55 of the 590 stored cells, at ages 5, 17, 29 … 113, and each figure reads back at its own age at every legal display length. | |
| **B3** | Writing one cell at a coarse display **rebuilds the whole triangle from the display grid, at the moment of the write**. Over a filled monthly store, one annual cell written left 54 stored cells holding their value, 535 at cumulative 0, and 1 changed — in every row, not just the one written. Visible before `Save`. | |
| **B4** | A partial coarse write is the same rule with fewer cells: writing display columns 1 and 3 stores at 5m and 29m and leaves everything else at 0. | |
| **B5** | An incremental display stores the **running sum** as the cumulative at each display age. Entering 100, 10, 1 stores 100 at 5m, 110 at 17m and 111 at 29m and onward, because the unwritten columns have an increment of 0. | |
| **B6** | A write at an origin display coarser than the store is refused outright. | `You cannot enter data unless the display origin length matches the data storage origin length (1).` |
| **B7** | `SetValues(originDate, ageMonths, value)` is a **display**-level call: it writes the column whose block contains that age. Ages 1–5 hit the 5m column and ages 6–17 the 17m column, so age 10 written after age 17 overwrites it. `Values(date, age)` reads the same way. | |
| **B8** | "Holds data" means a **saved** non-zero value. An unsaved `SetValuesByIndex` does not lock the store, a triangle saved with explicit zeros everywhere counts as empty, and `ClearData` unlocks it at once without a save. | `The stored development length may not be set in this triangle.` |
| **B9** | `ClearData` frees the origin axis too: on a cleared triangle a real `OriginLength` change moves the stored origin length again. `StoredOriginLength` still has no setter. | `Invalid number of parameters.` |

---

## 5. Group C — what saved data locks

| Case | Rule | ResQ's own words when it refuses |
| :--- | :--- | :--- |
| **C1** | With saved values, `StoredDevelopmentLength` is refused — even a put of the value it already holds. | `The stored development length may not be set in this triangle.` |
| **C1** | The display development length still has to divide the display origin length, exactly as when empty. | `The development length must be a factor of the origin length` |
| **C1** | The display origin length must be a whole multiple of the stored one. Over a store of 12: 12, 24 and 36 are accepted; 1 and 6 are not. | `The stored origin length must be a factor of the origin length.` |
| **C2** | A coarser origin display reads the calendar diagonal (§2). Writing there is refused — that is **B6**. | |
| **C3** | The stored pair, the display pair and `Cumulative` all persist and come back unchanged after `UnloadChildren` and a fresh read. | |

---

## 6. Group D — reading and presentation

| Case | Rule |
| :--- | :--- |
| **D1** | The label and row-width table of §2 for a store of 1, plus `GetOriginDate` / `GetDevelopmentDate`. An origin display of 24 labels its rows `2017 - 2018`, `2019 - 2020`, … |
| **D2** | `Calendarised = True` — the GUI's **Calendar** radio — relabels the columns as calendar bands (`01/17 - 05/17`, `06/17 - 05/18`, …) and returned exactly the same values as the development view for this development-anchored triangle. |
| **D2** | `Transposed = True` is a GUI flag only: `OriginCount`, `DevelopmentCountByIndex`, `OriginLabel` and `DevelopmentLabel` all keep answering as if it were false, so automation never has to undo it. |
| **D3** | `LeadingDiagonalByIndex(row)` returns each row's newest value and is the same at all six legal display lengths. |
| **D4** | The monthly-origin roll-up fixture: a 120x113 store filled `1000k + d` read at 12/12. |
| **D5** | The monthly-development roll-up fixture: a 12/1 store filled `100000 x row + age` read at 12/12 as 100005, 100017, … |
| **D6** | A coarse display over a store of 2, 3 or 6, the second table in §2. |
| **D7** | The same store read cumulatively and incrementally, at the store and at a coarse display. |

**Seen in the GUI, not measured:** the Data tab's Decimal Places is a display
setting with no COM property, so the API always returns full precision. That is
why an incremental annual view of whole-looking figures can show numbers that do
not tie to the differences of the cumulative view on screen — both columns are
rounded independently.

---

## 7. Group E — vectors

| Case | Rule | ResQ's own words when it refuses |
| :--- | :--- | :--- |
| **E1** | `StoredPeriodLength` has no usable setter on an origin vector, even while it is empty. | `You may not set the stored period length on this vector.` |
| **E1** | The stored period follows `PeriodLength` while the vector is empty, exactly as rule **A3**. | |
| **E1** | A coarser period display **sums** the finer periods — 12 monthly values of 1001…1012 read 12078 at period 12 — rather than reading one of them, which is the opposite of the triangle development axis. | |
| **E1** | A write at a coarser period display is refused. | `You cannot enter data unless the displayed period length corresponds with the data storage period length (1).` |

---

## 8. Group F — the export write sequence

| Case | Rule |
| :--- | :--- |
| **F1** | **`ClearData` resyncs the store to the display.** A triangle saved at display 12 over a store of 1 reads stored 12/12 straight after `ClearData`, before any put. So a store the export wants is set **after** the display is in place, never before. A same-value display put after that moves nothing. |
| **F2** | The sequence ArcRho's export macro uses is sound on a triangle that already holds data: `ClearData`, put the display pair, put `StoredDevelopmentLength`, show the triangle at the stored pair, write every stored cell by index, put the display pair back, `Save`. Putting the display back **before** the save is safe — the store stays where it was put and every written cell survives. Writing ArcRho's own scattered CSV (values at 5m, 17m … 113m, zeros between) this way reproduces ArcRho's annual grid in ResQ exactly. |
| **F3** | **The same sequence moves the stored origin length too, in both directions.** Over a filled triangle ResQ stores at O1/D1, it lands on stored O12/D1 shown at O12/D12 with ArcRho's annual grid intact; over one stored at O12/D12 it lands on stored O1/D1. `ClearData` on its own is enough — the extra `Save` and reload the ResQ window asks a person for changed nothing in any of the three runs. |

The F2 fixture puts only the development length, because that is the production
case. **F3** is the case the export used to refuse: the macro's own helper puts
both lengths and orders them so a shorter development period goes first, which
rule **A6** requires.

---

## 9. How ArcRho mirrors each rule

| ResQ rule | ArcRho | Where it lives |
| :--- | :--- | :--- |
| Two shapes per dataset | `stored_*` beside the display lengths in every sidecar; the CSV is at the stored shape and named for it | [sidecar_core_contract.py](../../python-api/src/arcrho_api/sidecar_core_contract.py), [helpers.py](../../frontend/app_server/helpers.py) |
| Coarse display reads the stored cell at the column's age; coarse origin reads the calendar diagonal (§2) | `rollup_triangle`, anchored on the project's Development End Date | [triangle_rollup.py](../../python-api/src/arcrho_api/triangle_rollup.py) |
| A coarse development write scatters to the column ages and zeroes the rest (**B2**–**B5**) | `scatter_triangle`, called by the sidecar save | [triangle_rollup.py](../../python-api/src/arcrho_api/triangle_rollup.py), [dataset_service.py](../../frontend/app_server/services/dataset_service.py) |
| "Holds data" means a **saved** non-zero value; explicit zeros are empty (**B8**) | `_stored_csv_holds_values` on the server; `savedDatasetHoldsNoValue` on the client, which reads the grid only while nothing is marked as an edit | [dataset_service.py](../../frontend/app_server/services/dataset_service.py), [data_tab_persistence_controller.js](../../frontend/ui/shared/tabs/data/data_tab_persistence_controller.js) |
| The store may be lowered only while the dataset is empty (**A4**, **C1**) | The `Stored at` control beside Development Length, live only while the dataset's own file holds nothing | [data_tab_persistence_controller.js](../../frontend/ui/shared/tabs/data/data_tab_persistence_controller.js) |
| The origin store has no control of its own (**A5**) | The origin `Stored at` value is shown but never editable | same file |
| Changing a display length while empty moves the store (**A3**) | The chosen store is remembered against the display length it was chosen at, so a display change drops it | `getStoredDevelopmentLengthChoice` in the same file |
| A put of the value already in place changes nothing (**A2**) | The Data tab fires a length change only when the value actually moves | [data_tab_request_controller.js](../../frontend/ui/shared/tabs/data/data_tab_request_controller.js) |
| An origin change keeps a development length that still divides it (**A2**, **A8**) | `enforceDevLenRule` | same file |
| Entering values never moves the store (**B2**) | The save states the store it wants and the server keeps it | `storedLengthIsPending` / `storedDevelopmentLengthForSave`, and `_save_dataset_sidecar_impl` |
| A coarse origin write is refused (**B6**) | 400 `Values can be entered only at the stored origin period.`; the grid is read-only on that axis | [dataset_service.py](../../frontend/app_server/services/dataset_service.py), [data_tab_preferences_controller.js](../../frontend/ui/shared/tabs/data/data_tab_preferences_controller.js) |
| Vectors refuse a coarse write (**E1**) | A vector ignores `stored_development_length` and refuses values at another period | [dataset_service.py](../../frontend/app_server/services/dataset_service.py) |
| The Calendar view (**D2**) | ArcRho has its own calendar-aligned roll-up, which aggregates forward-running blocks with a short last period rather than reading the valuation diagonal, so the two need not agree away from a store of 1 | `_rollup_calendar` in [triangle_rollup.py](../../python-api/src/arcrho_api/triangle_rollup.py) |
| The export sequence (**F1**, **F2**) | `_write_triangle_values` | [export_reserving_class_to_resq.py](../../python-api/macros/export_reserving_class_to_resq.py) |
| Restating the origin store on export (**F3**) | `_empty_and_reopen`, called only when ResQ's `StoredOriginLength` has to move | same file |
| The import reads a dataset at its own stored shape | `_displayed_at` around every read | [extractors.py](../../python-api/migration/resq_migration/extractors.py) |

### Where the two systems differ

- **Development labels, for every store other than 1.** ResQ's stored grid runs
  forward from the origin start and its displays end on the store's newest column
  (**A9**, **D6**); ArcRho values every row's newest cell on the Development End
  Date and counts back. Over a store of 1 the two are identical, which is why
  both production cases use one. Over a store of 3 ResQ ends on 114m and ArcRho
  on 113m. This is deliberate — mirroring ResQ would reproduce the very
  limitation the monthly store exists to avoid — but it means **a dataset stored
  at anything but 1 does not export with matching ages** in a project whose
  valuation date is not on an origin-period boundary.
- **The origin store.** ResQ has no setter for it and moves it with the display
  while empty; ArcRho does the same. On export a stored-origin difference is
  restated rather than refused: the triangle is emptied, saved and read again,
  and the origin store then follows the display put (**F3**).
- **How the store is stated.** ResQ is driven by a sequence of property puts, so
  it can tell a display *change* from a display *state*. ArcRho's save is one
  request carrying the whole state, so the Data tab — the only part that knows a
  control moved — states the store it wants in `stored_development_length` on
  every save while the dataset's file is still empty. **A save that omits the
  field is read as "store follows the display"**, which is why the field must be
  sent rather than left out.
- **The period ladder.** ResQ accepts any factor of the display length, so at
  origin 12 a store or a display of 2 or 4 is legal (**A4**, **A6**). ArcRho's
  one ladder is 12, 6, 3 and 1, so those two periods are never offered.
- **Values entered but not yet saved hold the display still.** ResQ would let the
  display move and rebuild the grid, since the triangle holds no saved value.
  ArcRho refuses to show unsaved figures at a period below the one they were
  entered at until they are saved or set back to 0, and its length ladder is
  narrowed to match that refusal rather than offering a length it will bounce.

---

## 10. The bug this document was written for

A user opened an empty hand-entered triangle, set the display to 12/12, lowered
the development `Stored at` to 1, pasted a 10x10 annual triangle from Excel and
pressed Save. ArcRho recorded the stored development length as **12**, not 1.

ResQ keeps 1 (**B2**). ArcRho did not, because the Data tab decided whether the
store could still move by looking at the **grid on screen** rather than at the
dataset's **file**: the paste filled the grid, the tab concluded the store was
already fixed, and the save went out with no stored length in it — which the
server reads as "store follows the display".

The grid cannot be asked once it has been edited: an edit writes straight into
the model and the cell's old value is gone. So the tab now takes the answer while
the grid still matches the file — that is, while nothing is marked as an edit —
and keeps it until the next load or save replaces it.

The cost was not cosmetic. The export macro writes ArcRho's stored shape into
ResQ, and a ResQ triangle stored at anything but development 1 carries column
ages ArcRho does not agree with (see the divergence above), so the exported
triangle would have shown the wrong development ages.

Fixed 2026-09-07. The server side already did the right thing once the field
arrives. Two tests pin the result against the column counts, row widths and ages
ResQ itself reported for this triangle:

- `test_a_finer_store_and_the_values_can_arrive_in_one_save` in
  [test_dataset_stored_shape_save.py](../../frontend/tests/test_dataset_stored_shape_save.py)
  — one save carrying both the finer store and the pasted 10x10 records (12, 1)
  and writes the figures at months 5, 17 … 113.
- `test_every_display_length_reads_the_store_the_way_resq_does` in
  [test_triangle_rollup.py](../../python-api/tests/test_triangle_rollup.py) —
  moving the Development Length afterwards to 1, 2, 3, 4, 6 or 12 reproduces
  ResQ's column counts, row widths and ages exactly (**B2**, **D1**).

---

## 11. Related reading

- [manual_input_stored_length_resq_alignment.md](../plans/completed/manual_input_stored_length_resq_alignment.md)
  — the plan that built the `Stored at` control, the scatter write and the
  export sequence, and the decisions behind them.
- [manual_input_period_rollup.md](../plans/completed/manual_input_period_rollup.md)
  — the plan that gave every dataset a stored shape separate from its display.
- [resq-api-reference.md](../../agent-instructions/resq-api-reference.md) — where
  the ResQ manual, the toolbox and the GUI screenshots live, and how to reach a
  live ResQ instance.
