# One-way ResQ reserving-class export

The `Export Reserving Class to ResQ` macro
(`python-api/macros/export_reserving_class_to_resq.py`) pushes one ArcRho
reserving class into the identically scoped ResQ reserving class, one way and
in one piece. It is the push counterpart of the
[sync macro](resq_reserving_class_sync.md): the same Bridge queue, the same
canonical session, the same ResQ writer, the same shared baseline, and the
same results window — minus the per-row direction and the signatures. The
review before the write is the one the Import macro also opens; see
[the shared ResQ transfer review](resq_reserving_class_transfer_review.md).

The ArcRho project name is also used as the ResQ project name, and the
selected reserving-class path must exist in that project on both sides. The
UI does not offer a project or path mapping override.

## What is pushed

For the reserving-class path selected in the active Project Instance page,
every item below is written, each after the items it reads:

- **Input datasets** — every sidecar whose `method_type` is `None`, that is
  neither `calculated` nor an `engine` dataset, and that has a CSV cache on
  disk. Triangle and vector values are written cell by cell
  (`SetValuesByIndex`) and the sidecar Notes go into the ResQ `Notes`. ResQ
  takes values at the period lengths a dataset is shown at, so a triangle is
  emptied (`ClearData`), given the sidecar's `stored_development_length`, shown
  at the sidecar's stored pair for the write, and put back to its display pair
  before the save. When ResQ stores the triangle at a different origin length,
  the emptied triangle is also saved and read again (`Save`,
  `UnloadChildren`, re-find) before the shape is restated: `StoredOriginLength`
  has no setter, the origin store follows `OriginLength` while the triangle
  holds nothing, and that is the sequence the ResQ window itself asks for.
  A dataset that is `Calculated` in ResQ is skipped,
  because ResQ recomputes it, even when ArcRho's library treats the type as an
  editable input.
- **DFM methods** — ratio exclusions (`SetExcludedRatios`), User Entry factors
  (`SetUserRatios`, up to the last ratio column), each average row's `- Ult`
  tail factor (`CustomAverages(i).TailFactor`), the selected average per
  column including the tail column (`SetSelectedRatios`), the Curves tab
  (`FutureDevelopmentPeriods`, `FreeFitC`, `SetIncludedRatios`, the User
  Entry columns through `SetCurveColumnDescription` and `SetCurveValues`
  with `DevIndex = 0` for a column's tail, `SetSelectedEstimates` per period,
  `SelectedTailFactor` and `SelectedTailCurve`), and Notes, the writer the
  sync's apply phase uses too. `FittingMethod` is never written: ArcRho fits
  by log regression only, so a ResQ method fitted by least squares keeps that
  setting; a prior-analysis, pattern or benchmark user column keeps ResQ's
  own values. Before anything is
  written, the first column of every ResQ average formula is read; a DFM with
  a formula ResQ cannot evaluate is skipped with that formula named. The read
  covers the `RatioAverageCount` rows the DFM really has, never the phantom
  rows ResQ reports past them.
- **Result Selections** — the loaded source datasets, weights (`SetWeights`),
  selected-ultimate overrides (`ClearOverriddenUltimates` + `SetUltimates`),
  and Notes.
- **B&S Case Reserve Adequacy methods** — the `Avg. Selections` tab and
  Notes. For each development column of both grids the exporter writes the
  `User Value` row (`SetUserAvgInflation`, `SetUserAvgCaseReserves`) and then
  the estimator selected for the column (`SetSelectedAvgInflation`,
  `SetSelectedAvgCaseReserves`), using the ResQ ordinals the import's label
  maps in `resq_migration.extractors` translate from. The method JSON holds
  the `User Value` row as the numbers the page evaluated, with a formula's
  text kept beside them, so a formula-backed cell reaches ResQ as its plain
  value. The method is then saved.
- **Bornhuetter Ferguson, Cape Cod, and B&S Settlement Rate methods** — saved
  only. The exporter finds the ResQ method by its ArcRho output name and calls
  `Save()`, so ResQ recalculates it from the datasets and DFMs written before
  it and re-stamps it. No field is carried across: ArcRho's own settings for
  these methods are not pushed, and a method ResQ does not hold is reported
  as skipped rather than created.

Left out, and not shown in the results:

- **Calculated datasets** — propagation recomputes them from their formula
  inputs in ArcRho and in ResQ alike.
- **Engine-generated datasets** (`source_kind: engine`) — ArcRho rebuilds them
  through the Engine and ResQ through its own generator.
- **Bootstrap methods** — ResQ has no write path for them yet.
- **Method output datasets** — they are written through their method, never
  as datasets.

**The export never creates anything in ResQ.** A dataset or method that
exists in ArcRho but not in ResQ is shown as `Skipped` (a warning) and left
alone; a new object reaches ResQ through ResQ itself. Datasets that exist in
ResQ but not in ArcRho are never touched, because only ArcRho's inventory is
walked. A dataset without a CSV cache, or a method-owned output sidecar whose
method JSON is missing, is shown as `Skipped` with the reason as well.

## Write order

Items are written in ArcRho's dependency order, the same topological walk
the sync's apply phase uses (`resq_migration.sync_session`). The graph is the
sidecar `precedents`/`dependents` of the whole reserving class, calculated
datasets included, plus the links in each method row's tabs. A row is written
only after every row it reads, wherever that row sits in the inventory; a
calculated dataset is never written, so a row that reads one is written after
the rows that dataset derives from instead. Rows with no link between them
keep a kind order — datasets, then DFMs and Berquist Sherman adjustments,
then Bornhuetter Ferguson and Cape Cod, then Result Selections — and then the
inventory order.

The graph is genuinely needed. In the fake project, `C 92 - Current Qtr
Selected` (a Result Selection of claim counts) feeds the B&S Settlement Rate
adjustment of `Gross Loss--Paid`, whose adjusted triangle `D 18 - BS Paid
DFM` reads, which `D 92 - Current Qtr Selected` loads in turn; the walk writes
them in exactly that order, so each save in ResQ finds its inputs already
written.

Looking through calculated datasets is needed as well. `C 91 - Current Qtr
Indicated` loads `C 62 Reported *(CWOP/Reported) CDF`, a calculated vector
derived from the `C 52 - CWOP/Reported DFM` output. The B&S adjustment above
pulls `C 92` and `C 91` forward in the walk; without the calculated link,
`C 91` was saved before `C 52`, and ResQ marked it "Needs Review" the moment
`C 52` was saved a second later.

## The review table

Before anything is written, the macro runs the queue's `transfer_preview`
phase and opens the shared review table — the same window, columns, and tick
rules the Import macro opens. That window and the selection it remembers are
described once, in
[the shared ResQ transfer review](resq_reserving_class_transfer_review.md).

What is specific to the export:

- The table is opened with **Export Selected to ResQ** and **Cancel**. It is
  the only confirmation asked for; accepting it starts the write.
- The **This Run** column reads `Overwrites ResQ copy`, `Overwrites newer
  ResQ copy` (the warning, raised only when `ResQ` or `Both` changed since
  the saved pair), or `Not exported`.
- An ArcRho item ResQ does not hold is listed as `ArcRho only` and cannot be
  ticked, because the export creates nothing in ResQ.
- The ticked names go on the `export` request as `SelectedNames`, narrow the
  rows before the dependency walk orders them, and are saved as the default
  for the next export once the writes are done.
- A comparison that fails is not a gate. The failure is shown with
  `Export Anyway` and `Cancel`, and an `Export Anyway` sends no selection at
  all, so the whole class is pushed exactly as it was before selection
  existed. An unreachable Bridge is reported as such instead, since nothing
  could be published either way.

## No row-by-row review, but a baseline

Beyond that table the export compares nothing and verifies nothing before a
write. Every ticked item is written over the ResQ copy, however recently
ResQ changed it. It is the tool for the moment ArcRho is the source of truth
for a class and ResQ should simply follow; use `Sync Reserving Class with
ResQ` when the two sides need reconciling row by row.

What the export does record, once the writes are done, is the baseline: the
ArcRho and ResQ timestamps each written item ends up carrying. Without it,
ResQ's `Save()` re-stamps every written object and the very next review — this
macro's preview and the sync macro's alike — reports every exported item as
`ResQ changed` or `Both changed`, which is noise, not information.

- **Where** — the same shared document the sync macro keeps, one per project,
  reserving class and ResQ connection, under `projects/<project>/sync/resq/`
  on the ArcRho server (`sync.sync_state_path`). It is server-side and
  scoped to the reserving class, so every user reviews against the same pair;
  no copy of it lives on anyone's machine.
- **What** — one entry per logical item, holding both timestamps and when
  they were recorded (`sync.record_synced_items`). Only an item ResQ
  confirmed as `Exported` or `Saved` is baselined; a skipped or failed one
  keeps its old pair, so the next review still reports the difference.
- **The ArcRho side** is baselined at the values the export actually pushed,
  not at a fresh read, so an ArcRho edit made while the export ran stays
  pending rather than being recorded as delivered.
- **Ripple** — ResQ recalculates whatever reads a written item, which
  re-stamps rows the export never wrote. Those moves are the export's own
  doing, so they are absorbed into the baseline
  (`sync.absorb_propagated_changes`) instead of surfacing as ResQ edits.
- **Failure is never fatal** — the writes are already durable when the
  baseline is saved, so a baseline that cannot be read or written is reported
  in the results header and the export still reports its writes. The next
  review simply falls back to comparing timestamps.

Because the document is shared with `Sync Reserving Class with ResQ`, an
export also settles that macro's next preview for everything it wrote.

## Results window

The results open inside the active Project Instance page as the same nested,
read-only review window the sync macro uses (`ui.reviewTableOpen` with
`host: "projectInstance"` and `selectable: false`): one row per item in write
order, with the type, the logical name, an outcome of `Exported`, `Saved`,
`Skipped`, or `Failed`, and the Bridge's message for the item, under a header
naming the project, the reserving class, the ResQ connection, the counts, and
how many timestamp pairs were saved for the next review to compare against.
The window is non-modal, so it can be minimized to the toolbar while the
class is inspected; the macro keeps running until it is closed.

Before anything is published, the macro refuses while the active nested
window has unsaved changes, since an unsaved edit would not be part of the
export. The review table above is the only confirmation asked for.

## Runtime

ResQ automation exists only where ResQ itself is installed, which is usually
not the machine ArcRho runs on. The macro therefore owns no ResQ session and
reads no reserving-class file: it publishes an `export` request to the same
Bridge queue the sync macro uses (`SyncResQReservingClass`, contract version
4, under `requests\RPC bridge\resq_reserving_class_sync\`), and a
ResQ-connected ArcRho Bridge worker on the Server PC runs
`resq_migration.sync_session.export_reserving_class` on its behalf. The
worker takes the reserving-class job lease a ResQ import and a sync apply
take, so no two of them write one reserving class at the same time, and
connects to ResQ with the shared service account from the server
`config.json`.

The client side is shared with the sync macro: `arcrho_api.resq_sync_queue`
builds, publishes, and waits on the request and refuses before publishing
when no ResQ-connected worker heartbeat is live, and
`arcrho_api.ui.await_review_table` hosts the results window. Any Client PC
can therefore export, provided some machine is running ResQ with ArcRho open.
Inside the app the request is published through the
`resq_sync_request_publish` hosted mutation and its status is polled through
the hosted Bridge-liveness read, so the queue is never touched over the
share; see the transport notes in
[resq_reserving_class_sync.md](resq_reserving_class_sync.md).

The ResQ writer, `ResQReservingClassExporter`, lives in the macro file. The
Bridge freezes that file beside the canonical migration
(`arcrho_bridge/bundled_sources.py`) and loads it as its writer, so an edit to
the exporter or to the session has no effect on an export or a sync until the
Bridge is rebuilt and redeployed. The worker refuses a bundle whose
`SYNC_SESSION_API_VERSION` it was not built against rather than driving it.

For headless use, build a runtime with
`resq_migration.sync_session.build_runtime(migration, exporter_module)` and
call `export_reserving_class(runtime, project_name, rc_path, server_root=...)`
on a machine with ResQ.

## ResQ COM findings

Verified on 2026-08-11 against ResQ connection `JGO_CO1SQLWPV22`, project
`NJ_Annual_Prod_202605_Fake`, when the writer was first built, and still what
the writer relies on:

- 749 read-back comparisons across vector values, DFM
  exclusions/selections/User Entry factors, RS weights, and BF linkage came
  back identical to the ArcRho sources for
  `PRNJ - PA\PA\NY\Direct Group\BI Total`.
- Dataset creation was exercised end to end when the writer was built and
  removed on 2026-08-28: the export never creates a dataset, Dataset Type, or
  method in ResQ any more.
- The pywin32 write convention for parameterized VBA property puts is
  `Set<Property>(indices..., value)` (e.g. `SetValuesByIndex`,
  `SetSelectedRatios`); it appears nowhere in the ResQ documentation but is
  proven by `python-api/migration/references/ResQToolBox2.py`, the API example
  notebook, and the ArcRho Bridge `SyncDFM` implementation.
- **ResQ names carry stray whitespace.** Real objects exist with trailing or
  doubled spaces, while ArcRho normalized all names on import. A plain
  `collection.Item(name)` misses those objects, so every lookup falls back to
  a cached whitespace-normalized name map.
- **A DFM average formula ResQ cannot evaluate fails inside ResQ, not on the
  connection.** `D 14 - Paid DFM w/ External LDFs` in the fake project fails
  on every path that makes ResQ evaluate its averages: the import cannot read
  `AverageRatioValues` at formula 7 ("Vol + 0.9 - all"), the sync's read-back
  of the selected average fails at column 1, and the export's write surfaces
  as `Access violation at address ... in module 'ResQ3Automation.dll'` from
  `xDFMMethod`. The items written after it on the same connection succeed,
  so reconnecting does not help; the writer probes every formula first and
  skips the DFM (`resq_average_unreadable`) naming the formula to fix in ResQ.
- **`AverageFormula` never ends, so `RatioAverageCount` is the only end of the
  list.** Every DFM of the fake project has 13 average rows, `10: User Entry`
  through `12: User Entry` among them and the reserving class's own
  `13: Aug 2024` below them. Asked for row 14 or beyond, ResQ does not fail:
  it keeps answering `"14: User Entry"`, `"15: User Entry"` and so on out of
  unallocated memory, and evaluating one of those rows crashes in
  `ResQ3Automation.dll` — which skipped every DFM in the class as
  `resq_average_unreadable` until the walk was bounded by the count. The rows
  are read through `RatioAverageCount`, and the repeated `User Entry` rows are
  collapsed onto the first exactly as the import collapses them.
- **Template-implementation methods are locked.** Structurally changing a
  method that belongs to a ResQ reserving-class template fails with "it is
  part of the ... template implementation". Value and selection writes and a
  plain `Save()` still work.
- **The ArcRho `datasets/` CSV cache is lazy.** A dataset never opened in
  ArcRho has no CSV on disk and is skipped; open it once (or build the cache)
  before exporting.
- **COM collection state is cached per connection.** After `Delete()` the
  item still appears in the same session's collection; a fresh connection
  sees the truth. The exporter uses one connection and neither creates nor
  deletes.

## Unclear / undocumented areas

ResQ COM API:

1. Enum ordinals are undocumented; only `ResQMethodType` 0-4/8/9,
   `ResQDataFormat` 0/1, `RatioExclusionType` 0/1/2, and `PercDevelopedType`
   0-3 are empirically confirmed.
2. `AddMethod` accepts 1-4 in live code; whether it supports Berquist Sherman
   (8/9) or Bootstrap (6) is unknown, which is one reason those are saved
   rather than created.
3. The User Entry average row is documented as row 11; this database has three
   of them, rows 10-12, and ArcRho keeps only the first. The writer resolves
   the row dynamically via `AverageFormula` labels, never by fixed index, and
   writes User Entry factors to the first row only. Which of the three a ResQ
   user has actually filled in is not visible to the writer, so a selection
   ArcRho holds as User Entry always comes back as row 10.
4. No bulk/SafeArray write path exists; every cell is one COM round trip.
5. There is no explicit `Recalculate`; recalculation appears synchronous on
   property set and on `Save()`, but save-failure semantics (partial
   in-memory state) are undocumented.
6. `Save()` on `Selected` must be isolated per the docs; the writer never
   touches dataset `Selected` flags for this reason.

ArcRho → ResQ mapping gaps of the DFM and Result Selection writers:

7. DFM `excluded == 2` (no data) is never written; ResQ derives empty cells.
8. ArcRho User Entry *formula text* (`average formulas.inputs`) cannot be
   represented in ResQ; only the resolved numeric factor is written.
9. Custom average definitions are matched by label only. An ArcRho-authored
   average whose label does not exist in the ResQ method is skipped for that
   column; creating or reordering ResQ averages via `CustomAverages(i)` is
   untested.
10. Result Selection dataset ordering after `AddDataset` is assumed stable;
    weights are addressed through a name→index map rebuilt after adds, but
    ResQ's `CustomSortIndex` semantics are undocumented.
11. Ultimate overrides push ArcRho `ultimate_overrides` as ResQ overridden
    ultimates; non-overridden ultimates are not pushed.
12. On the Curves tab, ArcRho user columns map onto ResQ's by position
    (column 6 onward); `CurveUserValueColCount` is raised when ArcRho holds
    more, never lowered, and a column ResQ types as prior analysis, pattern
    or benchmark is left as ResQ has it. `SetSelectedEstimates(DevIndex)`
    changes the stored number without moving the tail's selected value, so
    the tail is always written through `SelectedTailFactor` (probed
    2026-09-03). A ResQ Curves tab fitted by least squares is read as
    `fitting_method = least_squares` and shown in ArcRho with its
    log-regression fits.

The Bornhuetter Ferguson and Cape Cod field writers remain in the exporter
for the sync macro's apply phase; their known gaps (one BF prior, no scaling
type, collapsed prior-ultimate modes) are the sync documentation's
`supported fields only` caveat and do not affect the export, which only saves
those methods.
