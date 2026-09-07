---
name: stored-at-must-ask-the-file
description: "A Data tab question of the form 'is this dataset still empty?' must be asked of the dataset's file, not the grid on screen, and an edit destroys the grid's answer — so take it while nothing is dirty and keep it"
metadata: 
  node_type: memory
  type: project
  originSessionId: cd2f5f48-538c-4bad-8ac5-862be4f5ff0f
  modified: 2026-09-07T20:22:13.546Z
---

Fixed 2026-09-07. A user opened an empty hand-entered triangle, set the display to 12/12, lowered the development `Stored at` to 1, pasted a 10×10 annual block from Excel and pressed Save once. ArcRho recorded the stored development length as **12**, not 1, and the pasted values went in at the yearly shape.

`storedLengthIsPending()` in `frontend/ui/shared/tabs/data/data_tab_persistence_controller.js` decides whether a save may state the store, and it called `datasetValuesAreAllZero()`, which reads `state.model.values` — the grid on screen. The paste filled the grid, the tab concluded the store was already fixed, `storedDevelopmentLengthForSave()` returned 0, and the payload carried `stored_development_length: null`, which the save service reads as "the store follows the display".

**The first fix was wrong in an instructive way.** It skipped the cells in `state.dirty` and called what was left "the file". But an edit writes straight into `state.model.values` (`dataset_grid_interactions.js`) and `state.dirty` stores the *new* value, so the file's old value is gone: skipping dirty cells answers "is every cell I did not touch zero", which is true of a **populated** dataset the user pastes over completely. That flipped `storedLengthIsPending()` true on datasets holding values, which unlocked the `Stored at` control, un-narrowed the length ladder, lifted the link-shape guard, and turned a working save into a 400 (`The stored development length cannot be changed while the dataset holds values`). An adversarial review caught it before it shipped.

The shipped fix is `savedDatasetHoldsNoValue()`: read the grid **only while nothing is dirty** — that is the moment it still matches the file — and keep that answer in a variable until the next load or save replaces it. `applyStoredLengthsFromResponse` re-takes it, because a save has just written the grid to the file.

**Why:** ResQ fixes a triangle's stored period on the first value it is *saved* with, not on the first value typed into it, and ArcRho mirrors ResQ here. The cost was not cosmetic: the export macro writes ArcRho's stored shape into ResQ, and a ResQ triangle stored at development 12 can only ever label its columns 12m, 24m … 120m, so the exported triangle carried the wrong development ages.

**How to apply:** in the Data tab, "does this dataset hold anything?" has two answers and they are not interchangeable. Anything about the stored shape, the length ladder, the link-shape guard or what a save may change asks the **file**; anything about refusing a reshape that would throw away unsaved work asks the **grid** (`validateManualDatasetLengthChange` correctly still does). When the two disagree they must at least not contradict each other on screen — `manualDatasetLadderFloor` exists so the ladder never offers a length the validator snaps straight back. And never try to reconstruct the file from the grid plus the dirty map: the map holds the new values, not the old ones. Related: [[resq-stored-length-rules]], [[arcrho-dev-ui-cache-restart]].
