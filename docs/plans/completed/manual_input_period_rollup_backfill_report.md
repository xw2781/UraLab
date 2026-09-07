# Stored period length backfill — what the run did

Step 4 of [manual_input_period_rollup.md](manual_input_period_rollup.md), run on
2026-09-05 from the Client PC with `tools/backfill_stored_period_lengths.py`
against `E:\ArcRho Server`. Kept here because a one-time pass over every
project's data is not repeatable evidence: the share now looks the way the run
left it, and this is the only record of what it looked like before.

## The run

| | |
| :--- | :--- |
| Dry run | 2026-09-05 21:05–21:25, nothing written |
| Real run | 2026-09-05 21:24–22:06, `--apply --rebuild-index`, 41 minutes |
| Projects | 38 |
| Reserving classes | 116 |
| Sidecars given a stored shape | 10,997 |
| Sidecars that already had one | 0 |
| Reserving-class indexes rebuilt | 116, none refused |

By what wrote the dataset:

| Source | Sidecars |
| :--- | ---: |
| Engine-generated | 6,078 |
| Development factor methods | 1,816 |
| Hand entered | 1,195 |
| Calculated | 738 |
| Result Selection | 676 |
| Bornhuetter-Ferguson | 356 |
| Cape Cod | 64 |
| Berquist Sherman | 72 |

## What it wrote

Each sidecar took the shape it already recorded: a triangle's
`origin_length` / `development_length` became `stored_origin_length` /
`stored_development_length`, and a vector's `period_length` became
`stored_period_length`. Nothing else in a file moved — the new fields go in
behind the display lengths, every other key keeps its place and its value, and
the text is what `arcrho_api.io` writes for every ArcRho JSON file. Checked
against the untouched originals on a 122-file class and on 25 files picked at
random from the imperfect ones: the only difference is the one or two added
lines.

## What the dry run confirmed, and what it did not

**The plan's assumption held.** Not one sidecar on the share recorded lengths
that disagreed with the `@origin@development@` in its own `csv_file` name, so
"copy the current lengths, because they are the CSV's" was safe everywhere.
Nothing had to be guessed at, and nothing was skipped for disagreeing with
itself.

**A generated dataset took the shape it was last built at.** 6,078 of them, and
that is a placeholder: their real granularity is the source data's, which
nothing records until Step 5 puts it in the project's field mapping. They are
counted separately in the report for that reason.

**The share has never been through the v4 conversion.** This the plan had not
accounted for. 4,034 of the sidecars fall short of the shared core for reasons
that predate this work — 2,927 of them carry no `json_format` stamp at all, and
others are missing `status`, `show_subtotal`, or the two dependency lists. The
backfill therefore checks only the period-length rule it owns, counts the rest,
and leaves them to `tools/migrate_persisted_json_v4.py` (step 6 of
[persisted_json_contract_v4.md](persisted_json_contract_v4.md), which by decision converted only `NJ_Annual_Prod_202605_Fake`),
which fills exactly those fields. The most common shortfalls:

| Count | Missing |
| ---: | :--- |
| 1,103 | `json_format`, `show_subtotal`, `precedents`, `dependents` |
| 912 | `status` |
| 670 | `json_format`, `status`, `show_subtotal`, `precedents`, `dependents`, `audit_log` |
| 582 | `json_format`, `show_subtotal`, `precedents`, `dependents`, `audit_log` |
| 348 | `json_format`, `method_type`, `status`, `show_subtotal`, `precedents`, `dependents`, `audit_log` |
| 419 | fifteen smaller combinations, including 16 method outputs not marked calculated |

## The two files that needed a second pass

Two sidecars inside never-merged ResQ import staging copies under
`NJ_Annual_Prod_2026 Q2-May` refused to write: their folder path is long enough
that the `.json.tmp` an atomic write puts beside them crosses Windows' 260
character limit. The script now asks for those paths in the extended form, the
same way `tools/migrate_persisted_json_v4.py` does, and a second run over that
one project wrote both. Nothing was lost in between — a refused atomic write
leaves the original file exactly as it was.

## Running it again

The script is repeatable: a sidecar that already carries the stored fields is
counted and left alone, so a second whole-share run writes nothing. That is
worth doing after the v4 conversion lands, and after any ResQ import that
brings in projects written by an older build.
