# DFM average row values are rounded to display precision at import

Status: Diagnosed 2026-09-08; fix not started.

## Symptom

`dfm_ratio_side_by_side_review.py` (extended 2026-09-08 to also compare each DFM's output/ultimate
vector, not just its ratio triangle) flags small but real differences between ArcRho's and ResQ's
ultimate values, even when every ratio triangle cell matches ResQ to float precision.

Example: `PRNJ - PA\PA\All States\Direct Group\COL`, DFM `C 12 - CWP DFM w/ Selected LDFs`.

| Origin | ArcRho ultimate | ResQ ultimate | Diff |
| --- | --- | --- | --- |
| 2017-2019 | matches exactly | matches exactly | 0.00 |
| 2020-2025 | e.g. 15,928.019344 | e.g. 15,928.032059 | ~-0.01 |
| 2026 | 10,693.875636 | 10,693.644937 | **-0.231** |

The pattern (zero, then small, then largest at the least-mature origin) is the fingerprint of one
rounded input compounding through a cumulative-development-factor chain, not of unrelated
calculation drift — the same shape as the already-fixed [F 31 diff](../issues/f31-claim-count-severity-diff.md),
but from a different cause.

## Root cause

The selected development factor for this DFM's first period is a "User Entry" average row. ResQ's
own live value for it is `1.8189589399817268`. ArcRho's persisted method JSON stores it as
`1.819` — rounded to the DFM's Details-tab Decimal Places setting (4, for this method; the Ratios
tab prints 4 decimals here even though `1.8190` collapses to `1.819`).

The rounding happens on import, unconditionally, for every average row cell — not only for rows a
formula references:

- [`python-api/migration/resq_migration/dfm.py:886-887`](../../python-api/migration/resq_migration/dfm.py#L886-L887)
  ```python
  v = dfm.AverageRatioValues(j, resq_formula_idx)
  row.append(round(float(v), decimal_places) if v is not None else None)
  ```
- the same rounding is applied to each row's tail factor at [dfm.py:883](../../python-api/migration/resq_migration/dfm.py#L883).

`decimal_places` here is the DFM's Details-tab display precision (4 in this case), not the 6-decimal
round-trip precision the ratio triangle itself is stored at (`DFM_VALUE_DECIMAL_PLACES` in
`dfm_contract.py`). So the value that feeds ArcRho's own ultimate recalculation
(`_stored_selected_ratios` -> `_cumulative_from_normalized` -> `_calculate_ultimate` in
[dfm_contract.py](../../python-api/src/arcrho_api/dfm_contract.py)) is the *displayed* factor, not
ResQ's internal one.

This is a separate mechanism from the 2026-09-07 change ("read an average row at the method's
decimal places", commit `9b371e92`), which rounds a row's value only at the moment a *different*
row's User Entry formula references it, so the reviewer's displayed-digit multiplication lands on
the shown factor exactly. That rounding is intentional and scoped to formula evaluation. The
import-time rounding at `dfm.py:886-887` is broader: it destroys precision in the stored value
itself, before any formula is ever evaluated, for every average row regardless of whether anything
references it.

## Why it is not acceptable

The lost precision does not stay contained to the one DFM. The DFM's output (ultimate) vector is a
published dataset that other methods and calculated datasets read as an input — a Bornhuetter-
Ferguson or Cape Cod expected-losses method built on this DFM's percentage developed, a calculated
dataset that sums or ratios this DFM's ultimate against another vector, a later reserve-review
adjustment layered on top. Each of those downstream consumers inherits the rounding error and can
amplify it further (the F 31 issue above shows how even a simple product doubles a small relative
error), so a change here changes the diagnosis needed for every one of those methods too.

## Proposed fix

Stop rounding the average row's stored value to the DFM's display decimal places at import time.
Store it at the same round-trip precision the ratio triangle itself already uses
(`DFM_VALUE_DECIMAL_PLACES = 6` in `dfm_contract.py`, or unrounded), and let the Ratios tab's own
display formatting handle presenting it at the Details-tab decimal places. This keeps:

- the existing 2026-09-07 formula-reference rounding behavior unchanged (it is a distinct,
  intentional step at formula-evaluation time), and
- the calculation path (`_stored_selected_ratios` / `_calculate_ultimate`) working from
  ResQ-equivalent precision instead of the rounded display value.

Candidate change: in `python-api/migration/resq_migration/dfm.py`, replace the `decimal_places`
argument to `round()` at lines 883 and 887 with the round-trip precision used elsewhere for average
row values, and re-run `dfm_ratio_side_by_side_review.py` against `NJ_Annual_Prod_2026 Q3-Aug` to
confirm the output-vector diffs disappear (or shrink to genuine floating-point noise, matching the
ratio triangle's own ~1e-7 tolerance).

## Status

Diagnosed only. Not yet fixed or scoped for a specific release.
