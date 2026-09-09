# Method calculation precision parity with ResQ

## Status

Implemented on 2026-09-08. This document supersedes
`dfm_average_row_import_rounding.md`, which diagnosed one site of a problem that
turned out to run through every method contract in ArcRho.

## The rule

**A number ArcRho stores or calculates with is kept at the precision it was
observed with. A number ArcRho shows is rounded by the formatter that shows it.**

`arcrho_api.dfm_contract.canonical_input_number` is that rule for Python, and it
already owned the input triangle for the same reason. The DFM factor chain, the
Bornhuetter-Ferguson and Cape Cod contracts, and Result Selection now use it too,
and each JavaScript mirror was changed in the same commit so a page and the Engine
still agree byte for byte.

Two precisions remain deliberate, and neither is a projection of a computed number:

- **`ROUND(term, digits)` inside a User Entry formula.** The formula names its own
  rounding, so it is visible and auditable. `round_half_up` still rounds the
  decimal text, not the binary double.
- **Cape Cod's rate boxes**, at eight decimals. That is the precision the box
  itself offers, so a rate a user typed round-trips exactly.

## What was wrong

Three separate mechanisms lost precision that ResQ keeps.

### 1. A display precision used as a storage precision

The ResQ import rounded every average-row value and tail factor to the method's
Details-tab Decimal Places (4 in a reserve review). ArcRho never recomputes a
User Entry or benchmark row, so that rounded number *was* the factor its ultimate
chained: ResQ's `1.8189589399817268` was stored as `1.819`.

### 2. A six-decimal projection chained ten times

`canonical_number` rounds to six decimals so a persisted file stays readable. That
is right for a figure a reader checks by eye and wrong for a factor that gets
multiplied: each factor carried ~5e-7 of relative error, and a cumulative
development factor multiplies ten of them together. The same six-decimal quantum
was re-applied by every downstream contract, so a percentage developed — a
ratio-scale number where an absolute quantum of 0.000001 is coarse — reached
Bornhuetter-Ferguson and Cape Cod already rounded, twice.

### 3. `pandas.read_csv` without `float_precision="round_trip"`

The default C parser loses up to ~1e-12 relative precision. This was fixed in
`calculated_dataset_service.py` on 2026-09-07 for the F 31 issue; thirteen other
readers on method-value paths still had it, including every method's precedent
read and the Engine's own source-table load.

## What changed

| Area | Change |
| --- | --- |
| ResQ import | Average-row values and tail factors stored as ResQ gives them, not at the Details-tab precision |
| DFM contract | 16 chain sites moved from `canonical_number` to `canonical_input_number`: average rows, the selected chain, cumulative factors, % developed, the ultimate vector, and the ratio basis |
| DFM formula evaluation | An average row enters a User Entry formula whole; only an explicit `ROUND` rounds it |
| BF and Cape Cod | `_number` keeps observed precision, so a DFM's vectors are not re-quantized on the way in |
| Result Selection | All three producers (service, page, migration) carry a value whole; a weighted average of several ultimates now matches ResQ's |
| Bridge | `_snapshot_value` no longer rounds a ResQ factor to four decimals on the wire |
| CSV reads | `float_precision="round_trip"` added to 13 readers across DFM, BF, CC, B&S, Bootstrap, RS, dataset, roll-up cache and the Engine |

### The macros and the notes

The 2026-09-07 change (`9b371e92`) had made a User Entry formula read a referenced
row at the Details-tab Decimal Places implicitly, so the notes reconciled with the
Ratios tab without the formula saying so. That implicit rounding is what made
ArcRho's User Entry factor differ from ResQ's, so it is gone, and the rounding is
written down instead:

- **"Apply Growth and Cutoff Adjustments" (v1.4.0)** writes
  `= ROUND("Simple - 2", 4) * [Accounting Cutoff][-1] * [Growth Adjustment--Counts][-1]`
  again. The `ROUND` wraps only the DFM's own average row: the growth and cutoff
  vectors are already stored at four decimals and are multiplied in as they stand.
  Four decimals is the default the macro uses, and `BASE_FACTOR_DECIMALS` in
  `arcrho_api.combined_adjustment` is the one place it is written.
- **"Generate Notes for Combined Adjustment" (v1.3.0)** states the base factor at
  the precision the formula's own `ROUND` names, four by default, so the notes'
  arithmetic still reconciles with the adjusted value.
- **The DFM formula bar hides the `ROUND` in its rendered view.** The raw formula
  text keeps it, so clicking into the cell shows what is really stored.
  `stripRoundWrappers` in `frontend/ui/shared/components/formula_bar/formula_text.js`
  unwraps only a complete `ROUND(term, digits)`; a half-typed one is left alone.

Formulas written by either earlier version still read back: the macro's
`_GENERATED_FORMULA_RE` accepts both the bare label and the `ROUND` form.

## What a reader will notice

A DFM's persisted numbers are longer. An ultimate that read `1440.0` may now read
`1439.9999999999998`, because that is the double the chain produced and the number
ResQ chains too. The review scripts compare ultimates at two decimals and ratios at
1e-7, so this is well inside tolerance; it is the ~0.01 to ~0.2 differences at the
least-mature origins that the change removes.

## Verification

- `python-api/tests`: 746 passed, 21 failed — the same 21 that fail at `ce365350`,
  all reaching `E:\ArcRho Server` or naming a report directory. No new failures.
- `frontend/tests`: 1067 passed, 12 failed — identical to the set at `ce365350`.
- The ResQ import/sync macro suites, run separately because
  `test_resq_dfm_v2.py` slows them in the same process: 68 passed.

## Not done

Found and verified during the audit, left for a separate change because each is
its own contract:

- `roundRatio` in `dfm_ratio_calc.js` still rounds the binary double rather than
  the decimal text, so a value sitting exactly on a half can round the wrong way
  in the browser. It reaches storage only through the Excel paste paths.
- The Curves tab's user-value editor seeds from rendered text, so re-opening and
  committing a cell without changing it can save the displayed number.
- A custom average row's "- Ult" cell is replaced by `1.0` instead of keeping the
  row's own stored tail.
- Bootstrap's scale parameters are still quantized before the simulation reads
  them.
- The Result Selection import recomputes the ultimate rather than reading ResQ's
  `Ultimates(i)`.
