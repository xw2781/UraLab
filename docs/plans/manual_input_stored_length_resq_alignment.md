# Manual Input Triangles: Matching ResQ's Stored-Length Editing

Status: Investigation notes, 2026-09-06 — no code changed. The remaining ResQ rules are to be established against the ResQ COM API on the Server PC before any step is planned.
Last updated: 2026-09-06.

Follows on from [completed/manual_input_period_rollup.md](completed/manual_input_period_rollup.md), which gave every dataset a stored shape separate from its displayed shape and made a coarser view a read-only roll-up. Three things ResQ's GUI allows on top of that are not yet in ArcRho. This note records what has been observed so far, what ArcRho does today, the gap between them, and the questions still open, so the next agent can pick the work up without re-deriving it.

## What ArcRho does today

- **One control per axis.** The Data tab's Origin Length and Development Length are the displayed shape. The stored shape is shown as a muted caption (`stored 1`) and is never edited directly ([dataset.md](../../frontend/docs/ui/dataset.md), "Origin Length and Development Length are the shape…").
- **The first save fixes the stored shape.** While a manual dataset holds nothing but blanks and zeros, the caption reads `will be stored at 12 on first save` and follows the control; the save of a still-empty dataset is the one thing that moves the stored pair ([dataset_service.py:2355-2374](../../frontend/app_server/services/dataset_service.py#L2355-L2374)).
- **Each control offers only whole multiples of its own stored length**, origin and development independently, narrowed out of the one `LEN_CHOICES` ladder ([data_tab_request_controller.js:522-530](../../frontend/ui/shared/tabs/data/data_tab_request_controller.js#L522-L530)).
- **A coarser view is read-only.** When the display is coarser than the stored shape on either axis, the grid, the Links tab and paste all refuse, with the status line `Values can be entered only at the stored period (Origin 1, Development 1). Set the lengths back to edit.` ([data_tab_preferences_controller.js:369-393](../../frontend/ui/shared/tabs/data/data_tab_preferences_controller.js#L369-L393), [data_tab_persistence_controller.js:325-340](../../frontend/ui/shared/tabs/data/data_tab_persistence_controller.js#L325-L340)).
- **The roll-up is anchored on the Development End Date.** A coarser view counts development periods back from the valuation date, so a monthly triangle shown yearly in an August-valued project reads the stored columns at 8, 20, 32, … months of age ([triangle_rollup.py](../../python-api/src/arcrho_api/triangle_rollup.py), [dataset_service.py:1097-1105](../../frontend/app_server/services/dataset_service.py#L1097-L1105)). Landed 2026-09-06 (commit 46726a2a).
- **The CSV is cumulative or incremental as the sidecar says**, at the stored shape, and only that file is ever the dataset's data.

## What ResQ does (observed in the GUI, 2026-09-06)

Observed on `HPPREF\HO+DF\NJ\Legacy\HOL\Net Loss--Incurred Adjusted` and a `*5` copy of it, in a project valued at August 2026 with half-year and annual origins.

### 1. The stored development length can be lowered while the triangle is empty

- The Edit Triangle dialog has a `Stored at` spinner beside each length. On an all-zero half-year triangle displayed at Origin 6 / Development 3, the development `Stored at` was editable and was changed from 3 to 1 while the display stayed at 3. The origin `Stored at` (6) was dimmed on the same empty triangle.
- The display length must always be a whole multiple of the stored length.
- "Empty" evidently includes a triangle whose cells all read 0, not only one that has never been written: the triangle in the screenshot showed 0 in every cell and the spinner was still live.
- **Column ages follow the stored length, not the display length.** With stored 3 the columns read 3m, 6m, 9m, …; after lowering the stored length to 1, the same display of 3 read 2m, 5m, 8m, …. The newest cell of the newest row (2026 H2, starting July) is 2 months old at an August valuation; a quarterly store rounds that up to its 3-month boundary and labels it 3m. So the age of the latest diagonal is the months from the row's start to the valuation, rounded up to a multiple of the stored length, and every earlier column is one display length before it.

### 2. A coarse triangle can be pasted straight into a finer stored triangle (development axis)

- A 10×10 annual cumulative triangle was pasted into a triangle stored at Development 1 while it was displayed at Origin 12 / Development 12. ResQ accepted it.
- Shown afterwards at Development 4 (cumulative), the values sit only in the 8m, 20m, 32m, 44m, … columns — the stored columns the 12-month display columns stand for at an August valuation — and every other column reads 0.
- Shown incremental, each pasted value appears at 8m and its negative appears at 12m (for example +1,234 at 8m and −1,234 at 12m; 20m holds 1,481 and 24m holds −1,481). The untouched stored cells therefore kept a cumulative value of 0, and the pasted cumulative figure was written into the single stored cell its display column represents.
- In other words ResQ treats a paste at a coarse development display as "set the cumulative value of the stored cell at this age", leaving every other stored cell as it was. The two screenshots at Development 4 show this for a triangle that was empty before the paste.

### 3. The origin axis still requires display = stored to enter values

- Copying and pasting values requires the displayed origin length to equal the stored origin length. Only the development axis has the relaxation in point 2.
- This matches the rule the ResQ COM help states and our own export macro depends on: it sets `PeriodLength` back to `StoredPeriodLength` before `SetValuesByIndex` and restores the display afterwards ([export_reserving_class_to_resq.py:431-450](../../python-api/macros/export_reserving_class_to_resq.py#L431-L450)). Whether `SetValuesByIndex` itself accepts a coarse development display the way the GUI paste does is not yet known.

## The gap

| Behaviour | ResQ | ArcRho today |
| :--- | :--- | :--- |
| Set the stored development length below the display length on an empty triangle | Yes, a `Stored at` spinner | No; the stored shape is whatever the first save's display shape is |
| Set the stored origin length on an empty triangle | Spinner dimmed in the case observed | Same as development: fixed by the first save |
| Paste or type at a development display coarser than the stored length | Yes; writes the cumulative value into the stored cell at that age, other stored cells unchanged | Refused, whole grid read-only |
| Paste or type at an origin display coarser than the stored length | Refused | Refused |
| Column age labels | Multiples of the stored length counted back from the valuation, then stepped by the display length | To be checked against the 3m/2m observation above |

## Open questions for the ResQ API probe

Take these on the Server PC with the COM probe technique already recorded (`arcrho_bridge` venv, `gencache.EnsureDispatch`, read named getters only, never call `Set*`/`Select*`/`Refresh*` members blindly; the decompiled help is at `E:\XWSpace\ResQ API Doc`). Answer them in this section, one line each, before any implementation step is written.

1. **Paste into a non-empty finer triangle.** When the stored monthly triangle already holds values (a ResQ-imported one, say) and an annual cumulative triangle is pasted at the 12-month display, do the stored cells between the pasted ages keep their old cumulative values, so only the pasted age's incremental absorbs the difference? Or does ResQ clear the row block back to cumulative 0 first? (This was the question pending when the investigation paused.)
2. **Which stored cell receives the value.** The observed case put the value at the last stored age inside each display column (8m for the 8m column). Confirm that this is "the stored cell whose age equals the display column's age" rather than "the last stored cell of the block", using a project whose valuation falls on a period boundary (12m rather than 8m).
3. **Incremental display.** What does a paste at a coarse development display write when the triangle is shown incremental rather than cumulative?
4. **What counts as empty.** All-zero cells kept the development `Stored at` spinner live. Does a triangle with any non-zero value lock it, and does clearing the values back to zero unlock it again? Does the same rule hold for `StoredDevelopmentLength` set through the API, and what error does it raise when refused?
5. **Origin stored length.** The origin `Stored at` was dimmed on the empty triangle. Is it ever editable after the triangle exists, or only fixed at creation? Does the API allow `StoredOriginLength` to be set on an empty triangle?
6. **Display length after a stored-length change.** Lowering the stored length from 3 to 1 left the display at 3. Does raising it (1 back to 3 on an empty triangle) force the display up to a multiple, or refuse when the display is not one?
7. **Vectors.** Do `StoredPeriodLength` and paste follow the development-axis rule (relaxed) or the origin-axis rule (strict)?
8. **API paste path.** Does `SetValuesByIndex` (or whichever member the GUI paste uses) accept a development display coarser than the stored length and write the way the GUI does, or does the GUI do the mapping itself before calling the store?

## What ArcRho would need, once the questions are answered

Sketch only; not yet a plan.

- **A `Stored at` value per axis on the Data tab**, editable for a manual triangle only while it is empty and only on the axes ResQ allows (development for certain; origin depending on question 5). The display control keeps offering whole multiples of it.
- **Development-axis writes at a coarse display.** The read-only rule in `isDatasetReadOnly` splits by axis: a coarser origin display stays read-only; a coarser development display accepts typed values, paste and links, and the save maps each display cell to the one stored cell at its age and writes the cumulative value there, leaving every other stored cell as it was (or as question 1 dictates). The status wording in `getDatasetReadOnlyMessage` changes to name the origin axis only.
- **Column ages** labelled from the stored length as observed in point 1 of the ResQ section, if ArcRho's current labels differ.
- **The roll-up already reads the right stored cell** for the cumulative case (8, 20, 32, …), so a value written there shows again at the coarse display without further change. The incremental view of such a triangle will show the negative at the next stored age exactly as ResQ does, because it is derived from the same cumulative store.
- **Import/export symmetry.** The ResQ import already reads a hand-entered triangle at its stored shape; the export macro switches to the stored shape before writing. Neither needs to change for the storage itself, but the export could use the development-axis relaxation to write annual data into a monthly store without the switch, if question 8 says the API supports it.
