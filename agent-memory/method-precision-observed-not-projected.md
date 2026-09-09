---
name: method-precision-observed-not-projected
description: ArcRho keeps a stored or calculated number at the precision it was observed with; only a formatter or an explicit ROUND rounds one
metadata: 
  node_type: memory
  type: project
  originSessionId: 10cf579a-15e3-47d1-91f9-e76d69fe9a77
  modified: 2026-09-09T00:55:02.723Z
---

Since 2026-09-08 every ArcRho method contract follows one rule: **a number that is
stored or calculated with is kept at the precision it was observed with; a number
that is shown is rounded by the formatter that shows it.**

`arcrho_api.dfm_contract.canonical_input_number` is that rule in Python (it already
owned the input triangle for the same reason). `canonical_number`, the six-decimal
quantum, is now only for a value nothing chains. The DFM factor chain, BF, Cape Cod
and Result Selection all use the input form, and each JavaScript mirror
(`dfm_ratio_calc.js`, `*_json_contract.js`) was changed in the same commit — a page
and the Engine must produce byte-identical payloads.

Two roundings are deliberate and must stay:

- `ROUND(term, digits)` written inside a User Entry formula. "Apply Growth and
  Cutoff Adjustments" (v1.4.0) writes `ROUND("Simple - 2", 4)` around the DFM's own
  average row only — the growth and cutoff vectors already arrive at four decimals.
  `BASE_FACTOR_DECIMALS` in `arcrho_api.combined_adjustment` is the one place the 4
  is written. The DFM formula bar hides a complete `ROUND(...)` in its *rendered*
  view and keeps it in the raw editable text (`stripRoundWrappers` in
  `frontend/ui/shared/components/formula_bar/formula_text.js`).
- Cape Cod's rate boxes at eight decimals — the precision the box itself offers.

**Why:** three mechanisms were each losing precision ResQ keeps — a display
precision used as a storage precision at import, a six-decimal projection chained
ten times through the cumulative development factors, and `pandas.read_csv` without
`float_precision="round_trip"` (13 readers besides the one fixed for F 31). Together
they moved an ultimate by up to ~0.2 at the least-mature origin, and every
downstream BF, CC and RS inherited it.

**How to apply:** before adding any `round()`, `quantize`, `toFixed` or
`canonical_number` to a method value, ask whether anything reads it back to
calculate with. If so, keep it whole and round in the formatter instead. Expect
persisted numbers to look long — an ultimate of `1439.9999999999998` is the double
the chain produced and the one ResQ chains too. Full write-up:
`docs/plans/completed/method_calculation_precision_parity.md`. Related:
[[pandas-read-csv-not-round-trip]], [[origin-length-is-not-row-count]].
