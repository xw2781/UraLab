---
name: resq-stored-length-rules
description: "ResQ stored vs display length rules established by COM probe 2026-09-06 — display puts move the store while empty, StoredOriginLength has no setter, a coarse development write rebuilds the whole triangle (other stored cells go to 0), origin axis and vectors refuse coarse writes; export macro triangle path writes at display lengths (defect)"
metadata: 
  node_type: memory
  type: project
  originSessionId: a470642f-4b79-4631-a216-d7fd62626739
  modified: 2026-09-06T23:21:58.805Z
---

Probed 2026-09-06 on the Server PC in `NJ_Annual_Prod_202605_Fake` (origins 2017–2026, Development End Date 2026-05-31, so a yearly view is at 5m, 17m … 113m). Re-runnable: `py -3.10 tools/resq_stored_length_probe.py` (creates and deletes `ArcRho probe …` triangles of the non-unique `Net Loss - ad hoc` type in `HPPREF\HO+DF\NJ\Legacy\HOL`). Full write-up with error texts: `docs/plans/manual_input_stored_length_resq_alignment.md`.

- **Stored grid runs forward from the origin start** with a partial last period (stored 12 in a May project labels 12m…120m; stored 1 labels 1m…113m). A coarser display groups stored cells from the newest backwards (stored 1 at display 12 → 5m…113m) and a cumulative display column reads the one stored cell at its age.
- **Empty triangle:** any `OriginLength`/`DevelopmentLength` put moves the matching stored length to the same value (no multiple check). `StoredDevelopmentLength` can then be any factor of the display. `StoredOriginLength` has **no setter** (`Invalid number of parameters`). On a never-saved triangle an `OriginLength` put resets development to 1, so set `OriginLength`, then `DevelopmentLength`, then `StoredDevelopmentLength`, then `Save`. "Empty" = no saved non-zero value; explicit zeros are empty; `ClearData` unlocks at once.
- **Non-empty:** display must be a whole multiple of the store; `StoredDevelopmentLength` refused.
- **Coarse development write (`SetValuesByIndex`/`SetValues`) is accepted and rebuilds the whole triangle from the display grid immediately**: written display cells go to the stored cell at their age, every other stored cell in every row becomes cumulative 0, other display-age cells keep their values. `SetValues(date, age)` at a coarse display maps the age to its display column (age 10 → the 17m cell), so writing a stored cell needs the display at the stored length.
- **Origin axis and vectors are strict:** coarse writes refused (`You cannot enter data unless the display origin length matches the data storage origin length (1).`); `StoredPeriodLength` cannot be set on an origin vector.
- **Roll-ups match ArcRho's `triangle_rollup`** on both axes (development: stored cell at the column's age; origin: calendar-diagonal sum over origin months, 0 mismatches).

**Why:** ArcRho's manual-input design (stored-at chooser, coarse-display paste, export symmetry) must mirror these rules exactly for the export macro to leave ResQ identical.

**How to apply:** when writing a triangle to ResQ, show it at the sidecar's *stored* lengths first (the vector path already does; `_write_triangle_values` in `export_reserving_class_to_resq.py` still uses the display lengths — a recorded defect). Related: [[resq-com-probe]], [[triangle-rollup-valuation-anchor]], [[origin-length-is-not-row-count]].
