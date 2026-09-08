# <arcrho-macro>
# Title: Generate Notes for Combined Adjustment
# Version: 1.1.2
# Release Note: The macro now names the Flight Deck icon a button made from it starts with, so everyone who loads it gets the same glyph; you can still change the icon on your own button.
# Description: Read the selected User Entry formulas on the DFM Ratios tab that pull
#   adjustment factors from other ArcRho datasets (for example
#   = ROUND("Simple - 2", 4) * [Accounting Cutoff][-1] * [C 01 - Growth Adjustment][-1]),
#   resolve each referenced cell, and generate method notes in the legacy
#   "Apply Growth Adjustments" style. A term wrapped in ROUND is shown at that
#   precision. Adjustment factors equal to 1 are left out of the notes. Complex
#   formulas fall back to a resolved-formula note.
# Scope: DFM
# Icon: document
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
    APPLY_LINE_PREFIX,
    BASE_FACTOR_LINE_PREFIX,
    NOTE_BULLET_CHARS,
    NOTE_HEADER_PREFIX,
    SELECTED_LDF_LINE_PREFIX,
    adjustment_description,
    apply_line,
    base_factor_line,
    note_header,
    selected_ldf_line,
)
from arcrho_api.dfm_contract import round_half_up

MACRO_TITLE = "Generate Notes for Combined Adjustment"
NO_ADJUSTMENT_NOTE = "No combined adjustments were needed for this method."

# A displayed percent below this threshold rounds to 0.00%, so the factor is
# treated as 1 and its adjustment line is omitted from the notes.
UNITY_TOLERANCE = 0.00005


# ---------------------------------------------------------------------------
# Dataset reference parsing (mirrors frontend dfm_dataset_reference.js)
# ---------------------------------------------------------------------------

class DatasetReferenceSyntaxError(ValueError):
    pass


def _split_coordinates(raw: str) -> list[str]:
    parts: list[str] = []
    current = ""
    quote = ""
    for character in str(raw or ""):
        if quote:
            current += character
            if character == quote:
                quote = ""
            continue
        if character in ('"', "'"):
            quote = character
            current += character
            continue
        if character == ",":
            parts.append(current.strip())
            current = ""
            continue
        current += character
    if quote:
        raise DatasetReferenceSyntaxError("Dataset reference contains an unclosed quote.")
    parts.append(current.strip())
    return parts


def find_dataset_references(raw_formula: str) -> list[dict[str, Any]]:
    """Return [{match, start, end, dataset_name, row_idx, col_idx}] for a formula."""
    text = str(raw_formula or "")
    references: list[dict[str, Any]] = []
    start = 0
    while start < len(text):
        if text[start] != "[":
            start += 1
            continue
        dataset_end = text.find("]", start + 1)
        if dataset_end < 0:
            break
        coordinate_start = dataset_end + 1
        while coordinate_start < len(text) and text[coordinate_start].isspace():
            coordinate_start += 1
        if coordinate_start >= len(text) or text[coordinate_start] != "[":
            start += 1
            continue
        quote = ""
        coordinate_end = -1
        for index in range(coordinate_start + 1, len(text)):
            character = text[index]
            if quote:
                if character == quote:
                    quote = ""
                continue
            if character in ('"', "'"):
                quote = character
                continue
            if character == "]":
                coordinate_end = index
                break
        if coordinate_end < 0:
            raise DatasetReferenceSyntaxError("Dataset reference is missing its closing bracket.")
        dataset_name = text[start + 1:dataset_end].strip()
        coordinates = _split_coordinates(text[coordinate_start + 1:coordinate_end])
        if not dataset_name:
            raise DatasetReferenceSyntaxError("Dataset reference name cannot be blank.")
        if not coordinates[0]:
            raise DatasetReferenceSyntaxError("Dataset reference row index is required.")
        if len(coordinates) > 2 or (len(coordinates) == 2 and not coordinates[1]):
            raise DatasetReferenceSyntaxError(
                "Use [Dataset][row] for a vector or [Dataset][row, col] for a triangle."
            )
        references.append({
            "match": text[start:coordinate_end + 1],
            "start": start,
            "end": coordinate_end + 1,
            "dataset_name": dataset_name,
            "row_idx": coordinates[0],
            "col_idx": coordinates[1] if len(coordinates) == 2 else None,
        })
        start = coordinate_end + 1
    return references


# ---------------------------------------------------------------------------
# Formula tokenizing: split a User Entry formula into top-level product terms
# ---------------------------------------------------------------------------

def split_product_terms(expression: str) -> list[tuple[str, str]] | None:
    """Split "a * b / c" into [("*", "a"), ("*", "b"), ("/", "c")].

    Returns None when the expression is not a plain product/quotient chain
    (top-level + or -, unbalanced quotes/brackets, or an empty term), in which
    case the caller falls back to a resolved-formula note.
    """
    terms: list[tuple[str, str]] = []
    current: list[str] = []
    op = "*"
    quote = ""
    bracket_depth = 0
    paren_depth = 0
    for character in str(expression or ""):
        if quote:
            current.append(character)
            if character == quote:
                quote = ""
            continue
        if character in ('"', "'"):
            quote = character
            current.append(character)
            continue
        if character == "[":
            bracket_depth += 1
            current.append(character)
            continue
        if character == "]":
            bracket_depth -= 1
            if bracket_depth < 0:
                return None
            current.append(character)
            continue
        if character == "(":
            paren_depth += 1
            current.append(character)
            continue
        if character == ")":
            paren_depth -= 1
            if paren_depth < 0:
                return None
            current.append(character)
            continue
        if bracket_depth == 0 and paren_depth == 0:
            if character in "*/":
                text = "".join(current).strip()
                if not text:
                    return None
                terms.append((op, text))
                current = []
                op = character
                continue
            if character == "+":
                return None
            if character == "-" and "".join(current).strip():
                return None
        current.append(character)
    if quote or bracket_depth or paren_depth:
        return None
    text = "".join(current).strip()
    if not text:
        return None
    terms.append((op, text))
    return terms


def _strip_outer_parens(text: str) -> str:
    out = str(text or "").strip()
    while len(out) >= 2 and out[0] == "(" and out[-1] == ")":
        depth = 0
        balanced = True
        for index, character in enumerate(out):
            if character == "(":
                depth += 1
            elif character == ")":
                depth -= 1
                if depth == 0 and index < len(out) - 1:
                    balanced = False
                    break
        if not balanced:
            break
        out = out[1:-1].strip()
    return out


_ROUND_TERM_RE = re.compile(r"^round\s*\(\s*(.+?)\s*(?:,\s*(\d+)\s*)?\)$", re.I | re.S)


def classify_term(term: str) -> dict[str, Any] | None:
    """Classify a product term as an average label, dataset reference, or number.

    A term wrapped in ROUND(term, digits) classifies as the term inside and
    carries ``round_digits`` so the note shows it at that precision.
    """
    text = _strip_outer_parens(term)
    if not text:
        return None
    rounded = _ROUND_TERM_RE.match(text)
    if rounded:
        inner = classify_term(rounded.group(1))
        if inner is None:
            return None
        return {**inner, "round_digits": int(rounded.group(2) or 0)}
    if len(text) >= 2 and text[0] in ('"', "'") and text[-1] == text[0]:
        label = text[1:-1]
        if label and text[0] not in label:
            return {"kind": "label", "label": label}
        return None
    try:
        references = find_dataset_references(text)
    except DatasetReferenceSyntaxError:
        return None
    if len(references) == 1 and references[0]["match"] == text:
        return {"kind": "reference", "reference": references[0]}
    if references:
        return None
    try:
        return {"kind": "number", "value": float(text)}
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Text helpers (note style shared with the legacy Apply Growth Adjustments macro)
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


def _display_average_label(label: str) -> str:
    text = _clean_text(label)
    if ":" in text:
        _prefix, rest = text.split(":", 1)
        if rest.strip():
            return rest.strip()
    return text


def _number(value: Any) -> float | None:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if number != number or number in (float("inf"), float("-inf")):
        return None
    return number


def _format_note_multiplier(value: float) -> str:
    return f"{float(value):.4f}".rstrip("0").rstrip(".")


def _factor_percent_text(factor: float) -> str:
    delta = float(factor) - 1.0
    percent = f"{round(abs(delta) * 100, 2):g}%"
    return f"1+{percent}" if delta >= 0 else f"1-{percent}"


def _is_unity(factor: float) -> bool:
    return abs(float(factor) - 1.0) < UNITY_TOLERANCE


def _reference_location(resolved: dict[str, Any], reference: dict[str, Any]) -> str:
    name = _clean_text(resolved.get("dataset_name")) or _clean_text(reference.get("dataset_name"))
    row_label = _clean_text(resolved.get("row_label"))
    col_label = _clean_text(resolved.get("col_label"))
    data_format = _clean_text(resolved.get("data_format")).lower()
    if col_label and data_format != "vector":
        return f"[{name}] @ {row_label}, {col_label}"
    return f"[{name}] @ {row_label}"


# ---------------------------------------------------------------------------
# Dataset reference resolution through the ArcRho app server
# ---------------------------------------------------------------------------

def _error_text(exc: Exception) -> str:
    detail = getattr(exc, "detail", None)
    return _clean_text(detail) or _clean_text(str(exc)) or exc.__class__.__name__


def resolve_references_via_app_server(
    project_name: str,
    reserving_class: str,
    references: list[dict[str, Any]],
) -> tuple[list[dict[str, Any] | None], list[str]]:
    """Resolve references with the same service the Ratios tab formula bar uses.

    Returns (results, errors) where results[i] is the resolved payload for
    references[i] or None when that reference could not be resolved.
    """
    try:
        from app_server.services import dfm_service
    except Exception as exc:  # pragma: no cover - depends on app runtime
        raise DfmDataError(
            "Dataset references could not be resolved because the ArcRho app "
            f"server is not available in this Python session ({exc}). Run this "
            "macro from the ArcRho Macro window."
        ) from exc

    payload = [
        {
            "dataset_name": reference["dataset_name"],
            "row_idx": reference["row_idx"],
            **({"col_idx": reference["col_idx"]} if reference.get("col_idx") else {}),
        }
        for reference in references
    ]
    if not payload:
        return [], []
    try:
        response = dfm_service.resolve_dfm_dataset_references(project_name, reserving_class, payload)
        return list(response.get("results") or []), []
    except Exception:
        results: list[dict[str, Any] | None] = []
        errors: list[str] = []
        for reference in payload:
            try:
                response = dfm_service.resolve_dfm_dataset_references(
                    project_name, reserving_class, [reference]
                )
                results.append((response.get("results") or [None])[0])
            except Exception as exc:
                results.append(None)
                errors.append(f"[{reference['dataset_name']}]: {_error_text(exc)}")
        return results, errors


# ---------------------------------------------------------------------------
# Note generation
# ---------------------------------------------------------------------------

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


def _collect_column_formulas(dfm: Any) -> list[dict[str, Any]]:
    """Return one entry per development column whose selected row has an input formula."""
    formulas = dfm.average_formulas
    labels = [_clean_text(label) for label in (formulas.get("label") or [])]
    selected = _coerce_matrix(formulas.get("selected"))
    values = _coerce_matrix(formulas.get("values"))
    inputs = _coerce_matrix(formulas.get("inputs"))
    columns: list[dict[str, Any]] = []
    for col in range(dfm._average_col_count()):
        row = _selected_row(selected, col)
        if row is None:
            continue
        formula = _clean_text(_cell(inputs, row, col))
        if not formula:
            continue
        columns.append({
            "col": col,
            "row": row,
            "label": labels[row] if row < len(labels) else "",
            "formula": formula,
            "value": _number(_cell(values, row, col)),
            "labels": labels,
            "values": values,
        })
    return columns


def _parse_column_formula(entry: dict[str, Any]) -> dict[str, Any]:
    """Parse one column's formula into a base label plus multiplicative factors."""
    expression = entry["formula"]
    if expression.startswith("="):
        expression = expression[1:].strip()
    references = []
    try:
        references = find_dataset_references(expression)
    except DatasetReferenceSyntaxError:
        return {"ok": False, "references": []}
    terms = split_product_terms(expression)
    if terms is None:
        return {"ok": False, "references": references}
    base_label = None
    base_round_digits = None
    factors: list[dict[str, Any]] = []
    for op, term in terms:
        classified = classify_term(term)
        if classified is None:
            return {"ok": False, "references": references}
        if classified["kind"] == "label":
            if base_label is not None or op != "*":
                return {"ok": False, "references": references}
            base_label = classified["label"]
            base_round_digits = classified.get("round_digits")
            continue
        factors.append({"op": op, **classified})
    # Re-link reference factors to the whole-formula reference objects so the
    # resolved values attached to those objects are visible on each factor.
    reference_factors = [factor for factor in factors if factor["kind"] == "reference"]
    if len(reference_factors) != len(references):
        return {"ok": False, "references": references}
    for factor, reference in zip(reference_factors, references):
        factor["reference"] = reference
    return {
        "ok": True,
        "base_label": base_label,
        "base_round_digits": base_round_digits,
        "factors": factors,
        "references": references,
    }


def _base_value(entry: dict[str, Any], base_label: str) -> float | None:
    wanted = _label_key(base_label)
    for row, label in enumerate(entry["labels"]):
        if _label_key(label) == wanted:
            return _number(_cell(entry["values"], row, entry["col"]))
    return None


def _factor_lines(factors: list[dict[str, Any]]) -> tuple[list[str], list[str], bool]:
    """Build the "Apply ..." note lines and multiplier texts for non-unity factors."""
    lines: list[str] = []
    multipliers: list[str] = []
    for factor in factors:
        value = factor["resolved_value"]
        effective = 1.0 / value if factor["op"] == "/" else value
        if _is_unity(effective):
            continue
        multiplier = _format_note_multiplier(effective)
        if factor["kind"] == "reference":
            description = adjustment_description(factor["reference"]["dataset_name"])
        else:
            description = "other adjustment"
        if factor["op"] == "/":
            percent = f"1/({_factor_percent_text(value)})"
        else:
            percent = _factor_percent_text(value)
        lines.append(apply_line(description, percent, multiplier))
        multipliers.append(multiplier)
    return lines, multipliers, bool(lines)


def _column_note(dfm: Any, entry: dict[str, Any], parsed: dict[str, Any]) -> str | None:
    factors = parsed["factors"]
    has_reference_factor = any(factor["kind"] == "reference" for factor in factors)
    if not has_reference_factor and parsed["base_label"] is None:
        # A typed value or numeric expression is not a combined adjustment.
        return None
    for factor in factors:
        if factor.get("resolved_value") is None:
            return _fallback_note(dfm, entry, parsed)
        if factor["op"] == "/" and not factor["resolved_value"]:
            return _fallback_note(dfm, entry, parsed)

    lines, multipliers, meaningful = _factor_lines(factors)
    if not meaningful:
        return None

    base_label = parsed["base_label"]
    base_value = _base_value(entry, base_label) if base_label else None
    final_value = entry["value"]
    product = 1.0
    for factor in factors:
        value = factor["resolved_value"]
        product *= (1.0 / value) if factor["op"] == "/" else value
    if base_label is not None and base_value is None and final_value is not None and product:
        base_value = final_value / product
    if base_value is not None and parsed.get("base_round_digits") is not None:
        base_value = round_half_up(base_value, parsed["base_round_digits"])
    if final_value is None:
        if base_label is not None and base_value is None:
            return _fallback_note(dfm, entry, parsed)
        final_value = (base_value if base_value is not None else 1.0) * product

    note_lines = [note_header(dfm.dev_period(entry["col"] + 1))]
    note_lines.extend(lines)
    product_parts = list(multipliers)
    if base_label is not None:
        note_lines.append(base_factor_line(_display_average_label(base_label), base_value))
        product_parts.insert(0, f"{base_value:.4f}")
    note_lines.append(selected_ldf_line(" * ".join(product_parts), final_value))
    return "\n".join(note_lines)


def _fallback_note(dfm: Any, entry: dict[str, Any], parsed: dict[str, Any]) -> str | None:
    """Resolved-formula note for formulas outside the plain product shape."""
    references = parsed.get("references") or []
    resolved_parts = []
    for reference in references:
        resolved = reference.get("resolved")
        if resolved:
            location = _reference_location(resolved, reference)
            value = _number(resolved.get("value"))
            if value is not None:
                resolved_parts.append(f"{location} = {_format_note_multiplier(value)}")
    if not references:
        # Without dataset references there is no combined adjustment to explain.
        return None
    note_lines = [note_header(dfm.dev_period(entry["col"] + 1))]
    note_lines.append(f"  - User Entry formula: {entry['formula']};")
    if resolved_parts:
        note_lines.append(f"  - Resolved references: {'; '.join(resolved_parts)};")
    if entry["value"] is not None:
        note_lines.append(f"  - Selected LDF after adjustments: {entry['value']:.4f}")
    return "\n".join(note_lines)


# Lines that stand on their own (no bullet marker required).
GENERATED_HEADER_PREFIXES = (
    NOTE_HEADER_PREFIX,
    NO_ADJUSTMENT_NOTE,
    "No growth/accounting cutoff adjustments were needed",
)
# Lines that must carry a bullet marker to count as generated adjustment notes.
GENERATED_BULLET_PREFIXES = (
    APPLY_LINE_PREFIX,
    BASE_FACTOR_LINE_PREFIX,
    SELECTED_LDF_LINE_PREFIX,
    "User Entry formula: ",
    "Resolved references: ",
)


def _is_generated_note_line(line: str) -> bool:
    text = line.strip()
    if any(text.startswith(prefix) for prefix in GENERATED_HEADER_PREFIXES):
        return True
    if text[:1] not in NOTE_BULLET_CHARS:
        return False
    text = text[1:].lstrip()
    return any(text.startswith(prefix) for prefix in GENERATED_BULLET_PREFIXES)


def _clear_generated_notes(dfm: Any) -> None:
    lines = [line for line in dfm.notes.splitlines() if not _is_generated_note_line(line)]
    text = "\n".join(lines)
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    dfm.update_notes(text)


def generate_combined_adjustment_notes(
    dfm: Any,
    *,
    resolver: Callable[[str, str, list[dict[str, Any]]], tuple[list[dict[str, Any] | None], list[str]]] | None = None,
    project_name: str = "",
    reserving_class: str = "",
) -> dict[str, Any]:
    """Generate combined-adjustment notes for every selected User Entry formula.

    Returns {"note_blocks", "errors", "columns_seen"} without touching dfm.notes;
    the caller decides how to merge the blocks into the method notes.
    """
    resolve = resolver or resolve_references_via_app_server
    project = _clean_text(project_name) or _clean_text(getattr(dfm, "project_name", ""))
    rc = _clean_text(reserving_class) or _clean_text(getattr(dfm, "reserving_class", ""))

    columns = _collect_column_formulas(dfm)
    parsed_by_col: list[tuple[dict[str, Any], dict[str, Any]]] = []
    flat_references: list[dict[str, Any]] = []
    for entry in columns:
        parsed = _parse_column_formula(entry)
        parsed_by_col.append((entry, parsed))
        flat_references.extend(parsed["references"])

    errors: list[str] = []
    if flat_references:
        if not project or not rc:
            raise DfmDataError(
                "The active DFM does not carry its project and reserving class, "
                "which are required to resolve dataset references."
            )
        results, errors = resolve(project, rc, flat_references)
        for reference, resolved in zip(flat_references, results):
            reference["resolved"] = resolved

    note_blocks: list[str] = []
    for entry, parsed in parsed_by_col:
        if parsed["ok"]:
            for factor in parsed["factors"]:
                if factor["kind"] == "reference":
                    resolved = factor["reference"].get("resolved")
                    factor["resolved_value"] = _number((resolved or {}).get("value"))
                else:
                    factor["resolved_value"] = factor["value"]
                if factor.get("round_digits") is not None and factor["resolved_value"] is not None:
                    factor["resolved_value"] = round_half_up(factor["resolved_value"], factor["round_digits"])
            note = _column_note(dfm, entry, parsed)
        else:
            note = _fallback_note(dfm, entry, parsed)
        if note:
            note_blocks.append(note)

    return {"note_blocks": note_blocks, "errors": errors, "columns_seen": len(columns)}


def apply_notes(dfm: Any, note_blocks: list[str]) -> None:
    _clear_generated_notes(dfm)
    dfm.add_notes("\n\n".join(note_blocks) if note_blocks else NO_ADJUSTMENT_NOTE)


def run_macro(active_dfm, active_context=None):
    if active_dfm is None:
        return {
            "success": False,
            "message": "Open a DFM method before running this macro.",
        }
    fields = (active_context or {}).get("fields") if isinstance(active_context, dict) else {}
    fields = fields if isinstance(fields, dict) else {}
    original_notes = str(active_dfm.notes or "")
    result = generate_combined_adjustment_notes(
        active_dfm,
        project_name=str(fields.get("project") or ""),
        reserving_class=str(fields.get("reservingClass") or ""),
    )
    apply_notes(active_dfm, result["note_blocks"])
    suggested_notes = str(active_dfm.notes or "")

    block_count = len(result["note_blocks"])
    message_parts = [
        f"Generated combined-adjustment notes for {block_count} development period(s)."
        if block_count
        else "No combined adjustments were found in the selected User Entry formulas."
    ]
    if result["errors"]:
        message_parts.append(
            "Some dataset references could not be resolved: " + " | ".join(result["errors"])
        )
    preview = {
        "type": "notes_diff",
        "title": MACRO_TITLE,
        "summary": " ".join(message_parts) + " Review the suggested Notes text before applying.",
        "original_notes": original_notes,
        "suggested_notes": suggested_notes,
        "has_changes": suggested_notes != original_notes,
        "changes": [f"{block_count} note block(s) generated."],
    }
    return {
        "success": True,
        "payload": copy.deepcopy(active_dfm.to_dict()),
        "preview": preview,
        "message": " ".join(message_parts),
    }
