---
name: dataset-window-two-open-paths
description: "A Dataset window has two open paths with different length sources — the Project Instance open treats the /dataset/cache/load response as the sidecar, the run path loads the real sidecar; check both when a length or Stored-at bug appears, and read file mtimes (not gateway.log) to know whether a dev-mode save on the Server PC ever happened"
metadata: 
  node_type: memory
  type: project
  originSessionId: 766675f8-ed39-4256-be1c-c2b45f52357b
  modified: 2026-09-07T12:38:06.940Z
---

Found 2026-09-07 while chasing "an imported 12/12 triangle stored at 12/1 opens as 12/1 and cannot be saved back to 12/12".

- **Two open paths.** Opening a dataset from Project Instance goes through `data_tab_host_controller.readProjectInstanceCachedDataset`, which posts `/dataset/cache/load` and hands that response to `syncSidecarForCurrentDataset({ sidecarData })` as if it were the sidecar load, so whatever `origin_length` / `development_length` / `stored_*` that response carries becomes the window's settings. The run path (`loadDataset` in `dataset_run_controller`) loads the real sidecar afterwards with `applyLengths: false`, which resets `lastSavedDatasetSettings` but not the controls — that is why the Save button went dead after the user set 12/12: current == last saved.
- **The cached load describes the file it read.** `load_cached_dataset_values` reports the CSV's own shape as `origin_length`/`development_length`; only with `at_display_shape` (sent by the PI open alone) does it roll a hand-entered dataset up to the sidecar's display pair and carry the stored pair. Method pages (BS/BF/CC/RS) call the same route for raw stored rows, so never make the roll-up unconditional.
- **On the Server PC a dev-mode app saves locally.** `is_network_path(E:\)` is false there, so hosted transport is skipped and `gateway.log` / `hosted_saves.log` show nothing for the save; the CSV and sidecar mtimes under the reserving class's `datasets/` and `sidecars/` folders are the evidence of whether a save landed (listing names and mtimes is not reading project metadata).

**How to apply:** when a Data tab shows lengths or a Stored-at value that disagree with the sidecar, compare the `/dataset/cache/load` response with `load_dataset_sidecar` before touching the controls code; when checking whether a save happened, look at file mtimes first. Related: [[triangle-rollup-valuation-anchor]], [[resq-stored-length-rules]].
