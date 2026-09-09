"""The combined growth and accounting cutoff adjustment on a DFM's User Entry row.

Three producers share one shape. The "Apply Growth and Cutoff Adjustments"
macro writes the User Entry formula, for example
``= ROUND("Simple - 2", 4) * [Accounting Cutoff][-1] * [Growth Adjustment--Counts][-1]``.
The "Generate Notes for Combined Adjustment" macro writes the method notes that
explain it::

    For development period 12-24:
      - Apply accounting cutoff of 1+1.17% = 1.0117;
      - Apply growth adjustment--counts of 1+4.26% = 1.0426;
      - Selected average factor: "Simple - 2" (2.8539)
      - Selected LDF after adjustments: 2.8539 * 1.0117 * 1.0426 = 3.0102

The ResQ import reads those notes back, because ResQ keeps only the number the
formula produced, and rebuilds the formula from them. Whatever one producer
writes the others must read, so the dataset names, the formula shape and the
note lines are defined here once.
"""

from __future__ import annotations

import re
from typing import Any

ACCOUNTING_CUTOFF_DATASET = "Accounting Cutoff"
# The reserving class carries one growth vector per adjustment basis.
GROWTH_ADJUSTMENT_DATASETS = {
    "counts": "Growth Adjustment--Counts",
    "incurred": "Growth Adjustment--Incurred",
    "paid": "Growth Adjustment--Paid",
}
ADJUSTMENT_DATASETS = (ACCOUNTING_CUTOFF_DATASET, *GROWTH_ADJUSTMENT_DATASETS.values())

# The decimals the notes show a factor at, and the precision the formula names
# in its own ROUND. An average row otherwise enters a formula with every digit
# it holds, so the rounding a reserve review wants is written down rather than
# assumed, and the notes' arithmetic reconciles with the adjusted value. Only
# the DFM's own average row is rounded: the growth and cutoff vectors are
# already stored at four decimals and are multiplied in as they stand.
BASE_FACTOR_DECIMALS = 4

NOTE_HEADER_PREFIX = "For development period "
APPLY_LINE_PREFIX = "Apply "
BASE_FACTOR_LINE_PREFIX = "Selected average factor: "
SELECTED_LDF_LINE_PREFIX = "Selected LDF after adjustments: "
NOTE_BULLET = "  - "
# Bullet markers a note line may open with; the notes were written with "-",
# the legacy macro used "◦", and pasted text can carry the others.
NOTE_BULLET_CHARS = "-–—*•◦·○▪●"


def clean_text(value: Any) -> str:
    return " ".join(str(value if value is not None else "").split()).strip()


def adjustment_description(dataset_name: str) -> str:
    """How a note names an adjustment dataset: ``Growth Adjustment--Counts``
    reads as ``growth adjustment--counts`` and ``C 01 - Foo`` as ``foo adjustment``."""
    text = clean_text(dataset_name)
    text = re.sub(r"^[A-Za-z]{1,3}\s*\d+\s*[-–]\s*", "", text)
    description = text.lower()
    if description and "adjustment" not in description and "cutoff" not in description:
        description += " adjustment"
    return description or "other adjustment"


def note_header(period: str) -> str:
    return f"{NOTE_HEADER_PREFIX}{period}:"


def apply_line(description: str, percent_text: str, multiplier_text: str) -> str:
    return f"{NOTE_BULLET}{APPLY_LINE_PREFIX}{description} of {percent_text} = {multiplier_text};"


def base_factor_line(label: str, value: float) -> str:
    return f'{NOTE_BULLET}{BASE_FACTOR_LINE_PREFIX}"{label}" ({value:.{BASE_FACTOR_DECIMALS}f})'


def selected_ldf_line(product_text: str, value: float) -> str:
    return f"{NOTE_BULLET}{SELECTED_LDF_LINE_PREFIX}{product_text} = {value:.{BASE_FACTOR_DECIMALS}f}"


def adjustment_formula(base_label: str, terms: list[tuple[str, str]], column: int) -> str:
    """The User Entry formula for development column ``column`` (0-based).

    ``terms`` is ``[(op, dataset_name)]``. Each vector holds one already
    compounded factor per origin period, so development period n reads row -n.
    """
    row_idx = f"-{column + 1}"
    parts = [f'= ROUND("{base_label}", {BASE_FACTOR_DECIMALS})']
    parts.extend(f"{op} [{dataset_name}][{row_idx}]" for op, dataset_name in terms)
    return " ".join(parts)


_HEADER_RE = re.compile(rf"^{re.escape(NOTE_HEADER_PREFIX)}(.+?)\s*:$")
_APPLY_RE = re.compile(
    rf"^{re.escape(APPLY_LINE_PREFIX)}(?P<description>.+?) of (?P<inverse>1/\()?1[+-][^ )]+%\)? = [\d.]+;?$"
)
_BASE_RE = re.compile(rf'{re.escape(BASE_FACTOR_LINE_PREFIX)}"([^"]+)"')
_SELECTED_LDF_RE = re.compile(rf"^{re.escape(SELECTED_LDF_LINE_PREFIX)}.*= (\d+(?:\.\d+)?)$")


def parse_adjustment_notes(notes: str) -> dict[str, dict[str, Any]]:
    """Read the adjustment blocks back out of a method's notes.

    Returns ``{period: {"base_label", "terms", "value"}}``. ``terms`` is the
    ``[(op, dataset_name)]`` list ``adjustment_formula`` takes, or ``None``
    when an adjustment line names something that is not one of the adjustment
    datasets -- a typed factor, or a note written by hand -- so the formula
    cannot be rebuilt faithfully. ``value`` is the selected LDF the block ends
    on, ``None`` when the block does not state one.
    """
    by_description = {adjustment_description(name): name for name in ADJUSTMENT_DATASETS}
    blocks: dict[str, dict[str, Any]] = {}
    block: dict[str, Any] | None = None
    for raw_line in str(notes or "").splitlines():
        text = clean_text(raw_line)
        header = _HEADER_RE.match(text)
        if header:
            block = {"base_label": None, "terms": [], "value": None}
            blocks.setdefault(clean_text(header.group(1)), block)
            continue
        if block is None:
            continue
        text = text.lstrip(NOTE_BULLET_CHARS).lstrip()
        base = _BASE_RE.search(text)
        if base:
            block["base_label"] = clean_text(base.group(1))
            continue
        applied = _APPLY_RE.match(text)
        if applied:
            dataset_name = by_description.get(applied.group("description"))
            if dataset_name is None:
                block["terms"] = None
            elif block["terms"] is not None:
                block["terms"].append(("/" if applied.group("inverse") else "*", dataset_name))
            continue
        if text.startswith(APPLY_LINE_PREFIX):
            block["terms"] = None
            continue
        selected = _SELECTED_LDF_RE.match(text)
        if selected:
            block["value"] = float(selected.group(1))
    return blocks
