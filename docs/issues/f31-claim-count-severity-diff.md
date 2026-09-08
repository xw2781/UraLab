# F 31 - Claim Count * Severity: small diffs vs ResQ

**Reserving class:** `HPPREF\HO+DF\NJ\Legacy\HOL`
**Found:** 2026-09-07, via `combined_side_by_side_review.py --rc "HPPREF\HO+DF\NJ\Legacy\HOL"`

## Symptom

The F 31 vector (Claim Count * Severity) matches ResQ exactly for 2017-2020, then shows small
diffs from 2021 onward, growing much larger at the current diagonal (2026):

| Origin | ArcRho | ResQ | Diff |
| --- | --- | --- | --- |
| 2021 | 2,227.0004 | 2,227.0006 | -0.0002 |
| 2022 | 2,037.0494 | 2,037.0496 | -0.0001 |
| 2023 | 1,282.1005 | 1,282.1006 | -0.0000 |
| 2024 | 1,423.2058 | 1,423.2059 | -0.0001 |
| 2025 | 1,803.6270 | 1,803.6271 | -0.0001 |
| 2026 | 2,033.5156 | 2,033.5564 | -0.0408 |

## Root cause

F 31 is a literal cell-wise product of the Claim Count and Severity component matrices, evaluated
in `_eval_ast` (`frontend/app_server/services/calculated_dataset_service.py:461`).

Both components are loaded from CSV without round-trip float precision:

- `_load_component_matrix` — `frontend/app_server/services/calculated_dataset_service.py:921`
- `_read_numeric_csv` — `frontend/app_server/services/calculated_dataset_service.py:704`

Both call `pd.read_csv(path, header=None, dtype="float64", keep_default_na=True)` with no
`float_precision="round_trip"`. This is the same gap noted for method services generally (see
memory `pandas-read-csv-not-round-trip`): the default C float parser loses precision on values
read from source CSVs. Because F 31 multiplies two slightly-off floats together, the error doesn't
cancel — it shows up as a small, fairly uniform offset for every closed-out origin year.

The much larger 2026 diff is not a different bug: `_latest_diagonal_or_vector_values`
(`frontend/app_server/services/calculated_dataset_service.py:1252`) walks each origin row backward
and takes the last finite value. For a closed-out year that lands on a mature, stable development
column; for the still-developing 2026 origin it lands on an early, immature column. Immature-period
Claim Count/Severity values are more volatile, so the same relative rounding noise produces a much
larger absolute diff there.

## Fix

Applied 2026-09-07: added `float_precision="round_trip"` to the `pd.read_csv` calls in
`_load_component_matrix` and `_read_numeric_csv` in `calculated_dataset_service.py`.

## Status

Fixed.
