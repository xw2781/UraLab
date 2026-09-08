---
name: linked-origin-is-the-stored-origin
description: "A dataset cell link can only be entered at the stored origin period, so a sidecar's linked origin length is never independent of stored_origin_length; only the development width can differ"
metadata: 
  node_type: memory
  type: project
  originSessionId: 339cef0c-b486-464d-a125-5b303d82538b
  modified: 2026-09-08T00:24:57.896Z
---

A cell link (Excel, dataset, formula) can only be written where the grid can be
typed into, and on the origin axis that is always the period the file is held
at: a coarser origin display is read-only (`datasetOriginDisplayIsCoarserThanStored`
feeds `isDatasetReadOnly`), and a display finer than the store is refused while
the dataset holds values. A coarser *development* display does stay editable,
which is the one axis where the links' shape can differ from the store's.

**Why:** on 2026-09-07 a monthly triangle (`NJ_Annual_Prod_2026 Q3-Aug`, HOL,
`Net Loss--Incurred Adjusted***`, 6,786 linked cells) carried
`linked_origin_length: 12` over a 1/1 store, taken from the display the sidecar
happened to have been saved at. Switching the window to the yearly view then
looked like "at the linked shape" and every link reported "The linked dataset
cell is no longer part of this dataset." `linked_origin_length` /
`linked_period_length` are now retired fields; `linked_lengths` reads the origin
axis from `stored_lengths` and only `linked_development_length` is persisted.

**How to apply:** never infer a link's shape from a display, which moves; derive
the origin axis from the store. When a linked dataset misbehaves at one view,
compare the links' furthest `target_cells` row against the row count the
candidate shape would have — the targets name the grid they were written on.
See [[stored-at-must-ask-the-file]] and [[resq-stored-length-rules]].
