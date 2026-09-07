# ResQ API Reference Material

Read this before writing or debugging code that drives the ResQ COM API, migrates ResQ data, or reproduces a ResQ method in ArcRho.

## 1. Where The Reference Material Lives

Everything is consolidated under `E:\XWSpace\ResQ API Doc`, which is on the Server PC share:

| Path | What it is |
| --- | --- |
| `reference\resq_help_manual.chm` | The official ResQ scripting manual. |
| `reference\resq_help_manual_decompiled` | The same manual decompiled to searchable HTML. Grep this rather than opening the CHM. |
| `reference\ResQToolBox2.py` | The user's production Python wrapper around the ResQ COM layer. The richest source of real call patterns. |
| `reference\XLToolBox2.py` | Companion Excel automation module used alongside the toolbox. |
| `reference\ProjectDateTime.py` | Small helper used by the toolbox. |
| `reference\ResQ API Example.ipynb` | Worked examples of the COM surface. |
| `assets` | ResQ GUI screenshots, for matching ArcRho layouts to the ResQ originals. |

Two of these are also mirrored into the repository so migration code can be read offline: `python-api/migration/references/ResQToolBox2.py` and `python-api/migration/references/ResQ API Example.ipynb`.

[docs/reference/resq_stored_and_display_lengths.md](../docs/reference/resq_stored_and_display_lengths.md) is the repository's own reference for one corner the manual covers poorly: what ResQ does when a triangle or vector is created, reshaped, typed into, pasted into and read back at a period other than the one it is stored at. Every rule there was established against the live COM API and can be re-run with `tools/resq_stored_length_probe.py`. Read it before changing anything that decides a dataset's stored or displayed lengths.

Production reserve-review notebooks live under `E:\ResQ\Automations\Reserve Review\<quarter>` — for example the 2026Q1 `COL`, `HOL`, and `CMPxCAT` notebooks. They are the best evidence of which calls users actually depend on, and of the naming users already recognize.

## 2. How To Treat It

- **The manual is conceptual, not prescriptive.** Carry over the object hierarchy and the workflow ideas; do not carry over misleading COM names or argument conventions when ArcRho can offer a clearer Python name.
- **ResQ COM names are frequently unintuitive**, and the documented spelling sometimes disagrees with the working spelling — `GetCapeCodMethod` in the manual against `GetCapeCodeMethod` in the toolbox is the known case. Confirm against the decompiled HTML and the toolbox before trusting either, and ask the user when they still disagree.
- **`ResQToolBox2.py` is illustrative, not authoritative, and is incomplete in places.** Its `PercentageDevelopedType` map lists only codes 0-2 and omits `pdCumDevFactorsAdjusted=3`, which most production Cape Cod methods actually use. Verify enum coverage against the type library rather than against the toolbox.
- **The notebooks are a migration input, not a design constraint.** Use them to find the required helpers and the vocabulary; do not preserve a weaker legacy practice when a safer or more Pythonic API costs less to maintain.

## 3. Object Hierarchy To Preserve

```text
Application
  Projects
    Project
      ReservingClasses
        ReservingClass
          Triangles
          Vectors
          Methods / DFM methods
```

- A project is the container; a reserving-class path scopes datasets and methods.
- A method belongs to a reserving class and exposes input and output dataset links.
- ResQ's own help repeatedly recommends reaching a method through its reserving class or dataset rather than through a broad project-level collection. ArcRho follows that scoping model.
- DFM concepts map onto the ArcRho GUI tabs directly: Details (input triangle, output vector, name, lengths), Ratios (inclusions/exclusions, selected averages, average formulas), Results (ultimates and ultimate triangles), Notes (method notes and cell notes).

## 4. Running Against A Live ResQ Instance

ResQ is installed only on the Server PC, so in-process COM calls fail elsewhere with `Invalid class string`. Connection and sandbox rules — including the SSPI failure under sandboxed exec and the read-only sample project agents may use — are in [docs/plans/build_new_methods.md](../docs/plans/build_new_methods.md) under **ResQ Data Access**.
