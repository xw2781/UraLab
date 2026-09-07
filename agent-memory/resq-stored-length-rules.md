---
name: resq-stored-length-rules
description: "ResQ stored vs display length rules, probed live; the store moves only on a display-length CHANGE or an explicit stored put, never on a data write, and docs/reference/resq_stored_and_display_lengths.md is the full case matrix"
metadata: 
  node_type: memory
  type: project
  originSessionId: a470642f-4b79-4631-a216-d7fd62626739
  modified: 2026-09-07T20:37:43.525Z
---

Probed 2026-09-06 and extended 2026-09-07 on the Server PC in `NJ_Annual_Prod_202605_Fake` (origins 2017–2026, Development End Date 2026-05-31, so a yearly view is at 5m, 17m … 113m). **The full case matrix, with ResQ's own refusal texts and the ArcRho mapping, is `docs/reference/resq_stored_and_display_lengths.md`** — read it before touching anything that decides a dataset's stored or displayed lengths. Re-runnable: `py -3.10 tools/resq_stored_length_probe.py [--only A B2] [--keep] [--json PATH]`, which creates and deletes `ArcRho probe …` datasets of the non-unique `Net Loss - ad hoc` type in `HPPREF\HO+DF\NJ\Legacy\HOL`; its case ids (A1…F2) are what the reference cites.

- **The store moves only on a display-length CHANGE while the dataset holds no saved value, on an explicit `StoredDevelopmentLength` put, or on `ClearData`.** A put of the value already in place moves nothing (A2, A3) and a data write never moves it (B2) — that is the rule ArcRho kept getting wrong. `ClearData` both unlocks the store **and resyncs it to the display** (F1), so the export must put the display it wants before the store it wants.
- **Stored grid runs forward from the origin start** in steps of the store, columns `k·store`, last period short: stored 12 labels 12m…120m, stored 3 labels 3m…114m, stored 1 labels 1m…113m. **Only a store of 1 lands a column on the valuation date.** A coarser display groups from the store's own newest column backwards, so over a store of 2 a display of 12 reads 6m, 18m…114m, not 5m, 17m…113m (D6). Over a store of 1 that is `floor((newest_age - 1) / display) + 1` columns at ages `113 - k·display`. A cumulative display column reads the one stored cell at its age; an incremental one is the difference of the cumulative view (D7).
- **Empty triangle:** `StoredDevelopmentLength` may be any factor of the display; `StoredOriginLength` has **no setter** (`Invalid number of parameters`); the display development length must divide the display origin length. An `OriginLength` change resets the development length and its store to 1 **only when the old one no longer divides the new origin length** (A2, A8) — the same rule ArcRho's `enforceDevLenRule` applies. Set `OriginLength`, then `DevelopmentLength`, then `StoredDevelopmentLength`, then `Save`. "Empty" = no saved non-zero value; explicit zeros are empty; an unsaved `SetValuesByIndex` does not lock.
- **Non-empty:** display must be a whole multiple of the store on each axis; `StoredDevelopmentLength` refused.
- **Coarse development write rebuilds the whole triangle from the display grid immediately**: each display cell's cumulative goes to the stored cell at its age, every other stored cell in every row becomes cumulative 0. An incremental display stores the running sum. `SetValues(date, age)` maps the age to its display column (ages 1–5 → the 5m column, 6–17 → the 17m column).
- **Origin axis and vectors are strict:** coarse writes refused; a coarse vector period *sums* its finer periods on read, unlike the triangle development axis, which reads one cell.
- **Roll-ups match ArcRho's `triangle_rollup`** on both axes (0 mismatches over the 55-cell origin fixture).

**Why:** ArcRho's manual-input design (the `Stored at` chooser, the coarse-display paste, export symmetry) must mirror these rules exactly, and a stored development length other than 1 makes ResQ label the exported triangle 12m, 24m … which is wrong in a project not valued on an origin-period boundary.

**How to apply:** ArcRho's save is one full-state request, so it cannot tell a display *change* from a display *state* — the Data tab must state the store it wants in `stored_development_length` on every save while the dataset's **file** is still empty, and "still empty" is asked of the file, never of the grid on screen (fixed 2026-09-07; see [[stored-at-must-ask-the-file]]). Related: [[resq-com-probe]], [[triangle-rollup-valuation-anchor]], [[origin-length-is-not-row-count]].
