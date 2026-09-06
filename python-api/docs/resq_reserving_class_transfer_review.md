# The shared ResQ transfer review

`Import ResQ Reserving Class` and `Export Reserving Class to ResQ` move a
whole reserving class between ArcRho and ResQ in opposite directions. Before
either writes anything it opens the same window: one table of every dataset
and method output either system holds, with a tick box on each row. What was
ticked is remembered beside the reserving class on the ArcRho server, so the
next run in that direction opens with the same rows already ticked — for
whoever opens it next, not only for the person who saved it.

This page owns the description of that window and of the saved selection. The
[export](resq_reserving_class_export.md) and
[sync](resq_reserving_class_sync.md) pages describe what each macro does with
the answer, and
[import backups](resq_import_backups.md) describes the copy an import takes of
the existing reserving class before it writes.

## Where the rows come from

The rows are the Bridge's `transfer_preview` phase
(`resq_migration.sync_session.preview_transfer`), on the same queue and
through the same canonical session as the sync macro's own preview. Nothing
about a row is judged on the client: the session reads both inventories,
builds the synchronization plan for the items both sides hold, and adds the
items only one side holds, which the plan deliberately never contains.

An item only one side holds matters here in a way it does not for a sync. A
ResQ dataset ArcRho has never seen is precisely what an import brings across,
so it is a row, tickable, and shown as `ResQ only`. The same item can never be
exported, because the export writes ResQ objects and never creates them, so
in the export direction it is a row too — greyed out, with the reason.

`arcrho_api.resq_transfer_review` projects those rows into the review-table
contract and reads the ticked names back out. Both macros call it, so the two
cannot drift into describing one comparison two different ways.

## The columns

| Column | What it says |
| --- | --- |
| **Type** | `Dataset`, `DFM`, `Bornhuetter Ferguson`, `Cape Cod`, `Result Selection`, `B&S Settlement Rate`, `B&S Case Reserve Adequacy`. |
| **Dataset / Method Output** | The logical name both systems are paired on, and the name the request and the saved selection are written with. |
| **Held By** | `Both`, `ArcRho only`, or `ResQ only`. |
| **ArcRho Timestamp**, **ResQ Timestamp** | The pair the comparison read, as text, never re-parsed by the macro. `-` where that side has no copy. |
| **Newer** | `ArcRho`, `ResQ`, or `Same` (`sync.newer_side`). The plain fact about the two times, with no warning tone: after a transfer the target side always holds the newer stamp, so on its own it means very little. |
| **Changed Since Last Run** | `ArcRho`, `ResQ`, `Both`, or `None`, measured by `sync.changed_since_baseline` against the timestamp pair saved when the item was last exported or synchronized. A pair whose two timestamps match is `None` whatever the saved pair says: only a copy of one side over the other leaves both with the same stamp, and an import records no baseline. `No baseline yet` where no pair has been saved, in which case the row falls back to the raw timestamp comparison. |
| **This Run** | What the ticked run would do to the item, in the direction being reviewed — see below. |
| **Details** | The session's own sentence for the row, or the reason an untickable item cannot move. |

**This Run**, exporting:

- `Overwrites ResQ copy`
- `Overwrites newer ResQ copy` — the warning, raised only when `ResQ` or
  `Both` changed since the saved pair.
- `Not exported` — nothing would be written, so the row cannot be ticked.

**This Run**, importing. The Import macro always overwrites what is ticked, so
the column reads:

- `Added to ArcRho` — ResQ holds it and ArcRho does not.
- `Overwrites ArcRho copy`
- `Overwrites newer ArcRho copy` — when `ArcRho` or `Both` changed since the
  saved pair. Once the table is accepted, the ticked rows in this state are
  listed again (`edits_at_risk`) in a floating message box whose names open
  the item in the Project Instance page (`open_item_args`), and the import
  waits for a second Overwrite before it starts.
- `Keeps the newer ArcRho copy` — the same situation under a merge, which
  only the batch Import ResQ Reserving Classes macro still offers.
- `Not imported` — ArcRho cannot receive it (an unconfigured Dataset Type, a
  method object ResQ is missing), so the row cannot be ticked.

The header names the project, the reserving class and the ResQ connection, the
latest change on each side, how many items can be written, how many are
ticked, how many of those carry a change on the target side that the run would
overwrite, and where the ticked state came from.

## What may be ticked

A row is tickable when the direction can actually carry it
(`sync.transfer_support`). Everything else is shown disabled with its reason,
so an item that cannot move is visible rather than missing.

Two kinds are worth naming:

- **Berquist Sherman methods, exporting.** The sync's own write-back does not
  cover them, but the export saves each so ResQ recalculates it from the
  datasets and DFMs written before it, and writes a Case Reserve Adequacy
  method's `Avg. Selections` on the way. That is a real export, so they are
  tickable.
- **Calculated and engine-generated datasets.** Both systems recompute them
  from their inputs, so neither inventory lists them and neither direction
  offers them. An import still carries every one of them, ticked or not: the
  review is the only way to narrow an import, and a dataset it never shows
  cannot be left out by it (`catalog._is_unreviewed_dataset` is the one rule
  the review, the import, and the commit share).

Accepting with nothing ticked writes nothing and says so. Cancelling, or
closing the window, publishes nothing at all.

## The saved selection

The ticked names are saved once the run's writes are durable, in a document
beside the synchronization baseline:
`projects/<project>/sync/resq/<digest>.selection.json`, scoped by the same
project, reserving class, and ResQ connection
(`resq_migration.transfer_selection`). It is server-side and shared, so a
selection one person saves is the default every other person sees.

- **One document, two directions.** They are kept apart because they cannot
  hold the same answer: a ResQ-only item can be imported and can never be
  exported, so a single shared list would keep offering each direction rows
  the other one chose.
- **Names, not ids.** The document stores display names, de-duplicated by the
  same logical identity that pairs the two inventories, so it stays readable
  and survives a row id changing.
- **Nothing saved means everything ticked.** That is also what a run that
  carried no selection at all leaves behind, so a reserving class nobody has
  ever narrowed keeps behaving exactly as it did before.
- **Only the writing side saves.** A cancelled review, and a review that could
  not be produced, change nothing.
- **A default, never a decision.** Nothing is written from the document
  without the review being accepted, so an unreadable or foreign document
  falls back to everything ticked rather than stopping the run.
- **Failure is never fatal.** The run's writes are already durable when the
  document is written, so a document that cannot be saved is reported beside
  the results and the run still reports what it wrote.

## What a partial run does

**Exporting**, the ticked names narrow the rows before the dependency walk
orders them, so what is written is still written in ArcRho's dependency order
— just fewer rows. Only what ResQ confirmed as written is baselined, exactly
as before.

**Importing**, the ticked names narrow the ResQ inventory the staged import
walks. Every method is imported through its output dataset, so narrowing the
two dataset lists narrows the methods with them. Calculated and
engine-generated datasets stay in the inventory whatever was ticked; their
Dataset Type is read from ResQ only for a name the review did not tick.

The staged import then commits by overlaying the live reserving class onto the
stage. A live item the run never asked for was never offered to the stage, so
its absence there says nothing about ResQ: the merge always keeps it
(`merge.merge_preserved_arcrho_artifacts`, `requested_names`). Without that
rule a partial import would delete every dataset it was told to leave alone.
A calculated or engine-generated live item counts as requested, since the
stage always holds a fresh copy of it, and the ordinary merge or overwrite
rule decides between the two.

## Contract versions

Both queues refuse a request from a client they were not built against, rather
than answering with a shape the macro would misread:

- `SyncResQReservingClass` **contract version 4** — adds the
  `transfer_preview` phase with its `Direction`, and `SelectedNames` on the
  `export` phase.
- `ImportResQReservingClass` **contract version 2** — adds `SelectedNames`.
  A Bridge on version 1 would import everything regardless of what was
  ticked, which is why the version had to move.
- `resq_migration.sync_session` **API version 4** — `preview_transfer`, and
  `export_reserving_class` honouring a selection.

**Both macros and the Bridge must be redeployed together.** A Client PC on the
new macros with an old Bridge gets a clear refusal; the reverse is also safe.
