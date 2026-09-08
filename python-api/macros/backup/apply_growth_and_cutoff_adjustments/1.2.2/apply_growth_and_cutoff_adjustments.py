# <arcrho-macro>
# Title: Apply Growth and Cutoff Adjustments
# Version: 1.2.2
# Release Note: The macro now names the Flight Deck icon a button made from it starts with, so everyone who loads it gets the same glyph; you can still change the icon on your own button.
# Description: Write the combined growth and accounting cutoff adjustment into the active
#   DFM's User Entry row as a live in-cell formula, for example
#   = ROUND("Simple - 2", 4) * [Accounting Cutoff][-1] * [Growth Adjustment--Counts][-1].
#   The adjustment basis comes from the method's own input triangle: claim counts, paid,
#   incurred, or a severity ratio of incurred over counts. Only the first three development
#   periods are considered, each reading one row further back in the vectors, and a period
#   whose factor is 1 is left alone. The average factor is rounded to the four decimals the
#   notes show it at; the vectors are multiplied in as stored. The average cell the
#   adjustment was built from is marked with a "Selected before adjustments." cell note
#   before it is overwritten. Run "Generate Notes for Combined Adjustment" afterwards to
#   write the matching method notes.
# Scope: DFM
# Icon: calculator
# </arcrho-macro>

from __future__ import annotations

import copy
import re
from typing import Any, Callable

try:
    from arcrho_api.exceptions import DfmDataError
except Exception:  # pragma: no cover - script can still show useful errors
    DfmDataError = ValueError

from arcrho_api.combined_adjustment import (
    ACCOUNTING_CUTOFF_DATASET,
    BASE_FACTOR_DECIMALS,
    GROWTH_ADJUSTMENT_DATASETS,
    adjustment_formula,
    parse_adjustment_notes,
)
from arcrho_api.dfm_contract import round_half_up

MACRO_TITLE = "Apply Growth and Cutoff Adjustments"

# An adjustment reaches at most three development periods; beyond that the
# vectors are 1 in every reserve review, annual and quarterly alike.
MAX_ADJUSTED_COLUMNS = 3

# A factor this close to 1 shows as 0.00% and carries no adjustment, so its
# term is left out of the formula.
UNITY_TOLERANCE = 0.00005

TARGET_ROW_LABEL = "User Entry"

# The note this macro leaves on the average cell an adjustment was built from,
# before that cell's selection moves to the User Entry row it just wrote.
PRE_ADJUSTMENT_CELL_NOTE = "Selected before adjustments."


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def _clean_text(value: Any) -> str:
    return " ".join(str(value if value is not None else "").split()).strip()


def _label_key(value: Any) -> str:
    label = _clean_text(value)
    if ":" in label:
        prefix, rest = label.split(":", 1)
        if prefix.strip().isdigit():
            label = rest.strip()
    return label.lower()


def _number(value: Any) -> float | None:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if number != number or number in (float("inf"), float("-inf")):
        return None
    return number


def _is_unity(factor: float) -> bool:
    return abs(float(factor) - 1.0) < UNITY_TOLERANCE


def _coerce_matrix(value: Any) -> list[list[Any]]:
    if not isinstance(value, list):
        return []
    return [row if isinstance(row, list) else [] for row in value]


def _cell(matrix: list[list[Any]], row: int, col: int) -> Any:
    if 0 <= row < len(matrix) and 0 <= col < len(matrix[row]):
        return matrix[row][col]
    return None


def _selected_row(selected: list[list[Any]], col: int) -> int | None:
    for row, row_values in enumerate(selected):
        if col < len(row_values) and bool(row_values[col]):
            return row
    return None


def _error_text(exc: Exception) -> str:
    detail = getattr(exc, "detail", None)
    return _clean_text(detail) or _clean_text(str(exc)) or exc.__class__.__name__


# ---------------------------------------------------------------------------
# Which adjustment basis the method sits on
# ---------------------------------------------------------------------------

def adjustment_basis(dfm: Any) -> dict[str, Any]:
    """Decide which growth vector, if any, this method's ratios move with.

    Returns {"kind", "growth" [(op, dataset_name)], "cutoff" bool, "reason"}.
    ``kind`` is "ratio" when the method develops one basis expressed as a
    percentage of another, where the growth cancels out and no adjustment
    belongs.
    """
    input_triangle = _clean_text(dfm.input_triangle)
    category = _clean_text((dfm.details or {}).get("output_category")).lower()
    method_name = _clean_text(dfm.name).lower()
    triangle = input_triangle.lower()

    if "as % of" in triangle or "as % of" in method_name:
        return {
            "kind": "ratio",
            "growth": [],
            "cutoff": False,
            "reason": (
                f"{input_triangle or dfm.name} develops one basis as a percentage of "
                "another, so the growth adjustment cancels out."
            ),
        }
    if triangle.startswith("severity") or category.startswith("h severity"):
        return {
            "kind": "severity",
            "growth": [
                ("*", GROWTH_ADJUSTMENT_DATASETS["incurred"]),
                ("/", GROWTH_ADJUSTMENT_DATASETS["counts"]),
            ],
            # A severity is incurred over counts and both diagonals stop at the
            # same accounting month, so the cutoff cancels.
            "cutoff": False,
            "reason": "",
        }
    if triangle.startswith("claim counts") or category.startswith("c claim count"):
        kind = "counts"
    elif "--incurred" in triangle:
        kind = "incurred"
    elif "--paid" in triangle or "--received" in triangle:
        kind = "paid"
    elif "incurred" in method_name:
        kind = "incurred"
    elif "paid" in method_name or "salv" in method_name or "subr" in method_name:
        kind = "paid"
    else:
        return {
            "kind": "",
            "growth": [],
            "cutoff": False,
            "reason": (
                f"Could not tell which growth adjustment {input_triangle or dfm.name!r} "
                "belongs to; expected a claim count, paid, incurred or severity basis."
            ),
        }
    return {
        "kind": kind,
        "growth": [("*", GROWTH_ADJUSTMENT_DATASETS[kind])],
        "cutoff": True,
        "reason": "",
    }


# ---------------------------------------------------------------------------
# Reading the base average row back off a formula this macro wrote before
# ---------------------------------------------------------------------------

# The opening term is ROUND("label", n) since 1.2.0 and a bare "label" before it.
_GENERATED_FORMULA_RE = re.compile(
    r'^=\s*(?:ROUND\(\s*"([^"]+)"\s*,\s*\d+\s*\)|"([^"]+)")\s*(.*)$', re.S | re.I
)
_ADJUSTMENT_TERM_RE = re.compile(r'^\s*([*/])\s*\[([^\]]+)\]\s*\[\s*-\d+\s*\]')


def base_labels_from_notes(notes: str) -> dict[str, str]:
    """Return {development period: average label} from the method notes.

    A method imported from ResQ arrives with the actuary's own note naming the
    row the adjusted factor was built on -- 'Selected average factor: "Simple -
    2" (2.8539)' -- which is the only record of that choice once the adjusted
    number sits in the User Entry row.
    """
    return {
        period: block["base_label"]
        for period, block in parse_adjustment_notes(notes).items()
        if block["base_label"]
    }


def base_label_from_generated_formula(formula: str) -> str | None:
    """Return the quoted average label when *formula* is one this macro wrote.

    Re-running the macro must not fold the User Entry row into its own formula,
    so a column already carrying a generated formula is rebuilt from the label
    that formula started with.
    """
    match = _GENERATED_FORMULA_RE.match(_clean_text(formula))
    if not match:
        return None
    label, rest = match.group(1) or match.group(2), match.group(3)
    known = {ACCOUNTING_CUTOFF_DATASET.lower()} | {
        name.lower() for name in GROWTH_ADJUSTMENT_DATASETS.values()
    }
    while rest.strip():
        term = _ADJUSTMENT_TERM_RE.match(rest)
        if not term or term.group(2).strip().lower() not in known:
            return None
        rest = rest[term.end():]
    return label


# ---------------------------------------------------------------------------
# Dataset reference resolution through the ArcRho app server
# ---------------------------------------------------------------------------

def resolve_references_via_app_server(
    project_name: str,
    reserving_class: str,
    references: list[dict[str, Any]],
) -> tuple[list[dict[str, Any] | None], list[str]]:
    """Resolve references with the same service the Ratios tab formula bar uses."""
    try:
        from app_server.services import dfm_service
    except Exception as exc:  # pragma: no cover - depends on app runtime
        raise DfmDataError(
            "Dataset references could not be resolved because the ArcRho app "
            f"server is not available in this Python session ({exc}). Run this "
            "macro from the ArcRho Macro window."
        ) from exc

    if not references:
        return [], []
    try:
        response = dfm_service.resolve_dfm_dataset_references(
            project_name, reserving_class, references
        )
        return list(response.get("results") or []), []
    except Exception:
        # One bad reference fails the whole batch, so fall back to naming the
        # ones that actually failed.
        results: list[dict[str, Any] | None] = []
        errors: list[str] = []
        for reference in references:
            try:
                response = dfm_service.resolve_dfm_dataset_references(
                    project_name, reserving_class, [reference]
                )
                results.append((response.get("results") or [None])[0])
            except Exception as exc:
                results.append(None)
                errors.append(f"[{reference['dataset_name']}][{reference['row_idx']}]: {_error_text(exc)}")
        return results, errors


# ---------------------------------------------------------------------------
# Building the adjustments
# ---------------------------------------------------------------------------

def _column_terms(basis: dict[str, Any], column: int) -> list[tuple[str, str, str]]:
    """Return [(op, dataset_name, row_idx)] for one development column."""
    row_idx = f"-{column + 1}"
    terms: list[tuple[str, str, str]] = []
    if basis["cutoff"]:
        terms.append(("*", ACCOUNTING_CUTOFF_DATASET, row_idx))
    for op, dataset_name in basis["growth"]:
        terms.append((op, dataset_name, row_idx))
    return terms


def plan_adjustments(
    dfm: Any,
    basis: dict[str, Any],
    *,
    resolver: Callable[[str, str, list[dict[str, Any]]], tuple[list[dict[str, Any] | None], list[str]]],
    project_name: str,
    reserving_class: str,
) -> dict[str, Any]:
    """Work out the formula, value and base label for each adjustable column."""
    formulas = dfm.average_formulas
    labels = [_clean_text(label) for label in (formulas.get("label") or [])]
    selected = _coerce_matrix(formulas.get("selected"))
    values = _coerce_matrix(formulas.get("values"))
    inputs = _coerce_matrix(formulas.get("inputs"))
    target_key = _label_key(TARGET_ROW_LABEL)

    note_bases = base_labels_from_notes(dfm.notes)

    columns = min(MAX_ADJUSTED_COLUMNS, dfm._average_col_count())
    candidates: list[dict[str, Any]] = []
    references: list[dict[str, Any]] = []
    skipped: list[str] = []

    for col in range(columns):
        period = dfm.dev_period(col + 1)
        row = _selected_row(selected, col)
        if row is None or row >= len(labels):
            skipped.append(f"{period}: no average row is selected.")
            continue
        base_label = labels[row]
        if _label_key(base_label) == target_key:
            # The User Entry row cannot be its own base, so recover the average
            # row the adjusted factor was built on.
            recovered = base_label_from_generated_formula(_cell(inputs, row, col)) or note_bases.get(period)
            if recovered is None:
                skipped.append(
                    f"{period}: the User Entry row is selected and neither its formula nor the "
                    "method notes name the average row it was built on. Select that row on the "
                    "Ratios tab and run the macro again."
                )
                continue
            base_row = next(
                (index for index, label in enumerate(labels) if _label_key(label) == _label_key(recovered)),
                None,
            )
            if base_row is None:
                skipped.append(f'{period}: this method has no "{recovered}" average row to adjust.')
                continue
            base_label, row = labels[base_row], base_row
        base_value = _number(_cell(values, row, col))
        if base_value is None:
            skipped.append(f'{period}: "{base_label}" has no value to adjust.')
            continue
        terms = _column_terms(basis, col)
        candidates.append({
            "col": col,
            "base_label": base_label,
            "base_value": base_value,
            "terms": terms,
            "first_reference": len(references),
        })
        references.extend(
            {"dataset_name": dataset_name, "row_idx": row_idx}
            for _op, dataset_name, row_idx in terms
        )

    resolved, errors = resolver(project_name, reserving_class, references) if references else ([], [])

    # A quarterly reserving class holds quarterly adjustment vectors and can
    # still carry annual-origin methods. Reading one row back per development
    # period would then step a quarter where the method steps a year, so the
    # method is left alone rather than adjusted onto the wrong grid.
    origin_keys = {_label_key(label) for label in (dfm.data_tab.get("origin_labels") or [])}
    for payload in resolved:
        row_label = _clean_text((payload or {}).get("row_label"))
        if row_label and origin_keys and _label_key(row_label) not in origin_keys:
            return {
                "plans": [],
                "skipped": [],
                "errors": errors,
                "grid_mismatch": (
                    f"The adjustment vectors run on {row_label} periods and this method runs on "
                    f"{_clean_text((dfm.data_tab.get('origin_labels') or [''])[-1])} periods, so "
                    "there is no matching row to read."
                ),
            }

    plans: list[dict[str, Any]] = []
    for candidate in candidates:
        terms: list[tuple[str, str]] = []
        display_parts: list[str] = []
        factor = 1.0
        unresolved = False
        for offset, (op, dataset_name, _row_idx) in enumerate(candidate["terms"]):
            payload = resolved[candidate["first_reference"] + offset] if resolved else None
            value = _number((payload or {}).get("value"))
            if value is None:
                unresolved = True
                break
            effective = (1.0 / value) if op == "/" and value else value
            if _is_unity(effective):
                continue
            terms.append((op, dataset_name))
            # The formula bar spells the row out by its own label rather than
            # its position, the way a hand-typed reference reads back.
            display_parts.append(f"{op} [{dataset_name}][{_clean_text(payload.get('row_label'))}]")
            factor *= effective
        if unresolved:
            skipped.append(
                f"{dfm.dev_period(candidate['col'] + 1)}: an adjustment vector could not be read."
            )
            continue
        if not terms:
            continue
        opening = adjustment_formula(candidate["base_label"], [], candidate["col"])
        base_value = round_half_up(candidate["base_value"], BASE_FACTOR_DECIMALS)
        plans.append({
            "col": candidate["col"],
            "base_label": candidate["base_label"],
            "formula": adjustment_formula(candidate["base_label"], terms, candidate["col"]),
            "display_formula": " ".join([opening] + display_parts),
            "value": round(base_value * factor, 6),
            "base_value": base_value,
        })
    return {"plans": plans, "skipped": skipped, "errors": errors, "grid_mismatch": ""}


def _mark_selected_before_adjustment(dfm: Any, col: int, base_label: str) -> None:
    """Note which average cell fed the adjustment before it is overwritten.

    Re-running the macro can move a column's adjustment onto a different base
    row (or drop it once a vector goes back to 1), so any note this process
    left on that development column is cleared first rather than left behind
    on a row the selection has moved past.
    """
    dfm.clear_cell_notes_for_development(col + 1)
    dfm.set_cell_note(base_label, col + 1, PRE_ADJUSTMENT_CELL_NOTE)


def apply_adjustments(dfm: Any, plans: list[dict[str, Any]]) -> None:
    for plan in plans:
        _mark_selected_before_adjustment(dfm, plan["col"], plan["base_label"])
        dfm.set_user_formula(plan["formula"], plan["value"], plan["col"] + 1)
    if not plans:
        return
    formulas = dfm.average_formulas
    row = dfm._ensure_average_label(TARGET_ROW_LABEL)
    columns = dfm._average_col_count()
    display = formulas.get("display_inputs")
    if not isinstance(display, list):
        display = []
        formulas["display_inputs"] = display
    while len(display) <= row:
        display.append([])
    for index, existing in enumerate(display):
        display[index] = (existing if isinstance(existing, list) else [])[:columns]
        display[index] += [""] * (columns - len(display[index]))
    for plan in plans:
        display[row][plan["col"]] = plan["display_formula"]


# ---------------------------------------------------------------------------
# Macro entry point
# ---------------------------------------------------------------------------

def run_macro(active_dfm, active_context=None):
    if active_dfm is None:
        return {"success": False, "message": "Open a DFM method before running this macro."}

    fields = (active_context or {}).get("fields") if isinstance(active_context, dict) else {}
    fields = fields if isinstance(fields, dict) else {}
    project_name = _clean_text(fields.get("project")) or _clean_text(getattr(active_dfm, "project_name", ""))
    reserving_class = _clean_text(fields.get("reservingClass")) or _clean_text(
        getattr(active_dfm, "reserving_class", "")
    )
    if not project_name or not reserving_class:
        raise DfmDataError(
            "The active DFM does not carry its project and reserving class, "
            "which are required to read the adjustment vectors."
        )

    basis = adjustment_basis(active_dfm)
    if not basis["growth"]:
        return {"success": True, "payload": copy.deepcopy(active_dfm.to_dict()), "message": basis["reason"]}

    result = plan_adjustments(
        active_dfm,
        basis,
        resolver=resolve_references_via_app_server,
        project_name=project_name,
        reserving_class=reserving_class,
    )
    if result["grid_mismatch"]:
        return {
            "success": True,
            "payload": copy.deepcopy(active_dfm.to_dict()),
            "message": result["grid_mismatch"],
        }
    apply_adjustments(active_dfm, result["plans"])

    if result["plans"]:
        applied = ", ".join(
            f"{active_dfm.dev_period(plan['col'] + 1)} {plan['base_value']:.4f} -> {plan['value']:.4f}"
            for plan in result["plans"]
        )
        message_parts = [f"Adjusted {len(result['plans'])} development period(s): {applied}."]
    else:
        message_parts = ["No development period needed a growth or accounting cutoff adjustment."]
    if result["skipped"]:
        message_parts.append("Left alone: " + " | ".join(result["skipped"]))
    if result["errors"]:
        message_parts.append("Could not read: " + " | ".join(result["errors"]))

    return {
        "success": True,
        "payload": copy.deepcopy(active_dfm.to_dict()),
        "message": " ".join(message_parts),
    }
