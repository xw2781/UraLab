from __future__ import annotations

import re
from collections import Counter
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path

from arcrho_api.combined_adjustment import (
    BASE_FACTOR_DECIMALS,
    adjustment_formula,
    parse_adjustment_notes,
)
from arcrho_api.dfm_contract import (
    DFM_JSON_FORMAT,
    apply_owned_patch,
    canonical_input_number,
    canonical_number,
    persisted_projection,
    recalculate_dfm_method,
)

from .catalog import _is_known_dataset_type, _unknown_dataset_type_skip_detail
from .core import (
    _clean_name,
    _encode_name_part,
    _iso_or_text,
    _normalize_import_name,
    persisted_json_text,
    _safe_attr,
    _safe_read_json,
)
from .extractors import (
    build_dfm_ultimate_publication,
    export_dfm_ultimate_vector,
    export_triangle,
    publish_dfm_artifacts,
)
from .number_formats import dataset_type_decimal_places, dataset_type_number_format


MAX_AVERAGE_FORMULA_PROBE = 30


def configure_dfm(*, dfm_json_format: str) -> None:
    if str(dfm_json_format) != DFM_JSON_FORMAT:
        raise ValueError(
            f"DFM JSON format is owned by arcrho_api.dfm_contract and must be {DFM_JSON_FORMAT!r}."
        )


def _dict_child(parent: dict, key: str) -> dict:
    value = parent.get(key)
    if isinstance(value, dict):
        return value
    value = {}
    parent[key] = value
    return value


def _dict_path(payload: dict, keys: tuple[str, ...]) -> dict:
    current = payload
    for key in keys:
        if not isinstance(current, dict):
            return {}
        current = current.get(key)
    return current if isinstance(current, dict) else {}


def _merge_cell_note_dicts(remote_notes: dict, local_notes: dict) -> dict:
    merged = deepcopy(remote_notes) if isinstance(remote_notes, dict) else {}
    if not isinstance(local_notes, dict):
        return merged

    for table_name, local_rows in local_notes.items():
        if not isinstance(local_rows, dict):
            merged[table_name] = deepcopy(local_rows)
            continue
        merged_rows = merged.setdefault(table_name, {})
        if not isinstance(merged_rows, dict):
            merged_rows = {}
            merged[table_name] = merged_rows
        for row_label, local_cols in local_rows.items():
            if not isinstance(local_cols, dict):
                merged_rows[row_label] = deepcopy(local_cols)
                continue
            merged_cols = merged_rows.setdefault(row_label, {})
            if not isinstance(merged_cols, dict):
                merged_cols = {}
                merged_rows[row_label] = merged_cols
            for col_label, note_text in local_cols.items():
                merged_cols[col_label] = deepcopy(note_text)
    return merged


def _references_another_average_row(row_inputs) -> bool:
    """True when a row's cells name other summary rows, as a translated
    ResQ User Calculation row does and a hand-typed User Entry row does not."""

    if not isinstance(row_inputs, list):
        return False
    return any('"' in str(cell or "") for cell in row_inputs)


def _average_formula_user_entry_index(average_formulas: dict) -> int | None:
    """The row that holds ResQ's own User Entry values.

    Since ResQ User Calculation rows import as User Entry rows too, being the
    first row of that type no longer identifies it. The row ResQ calls "User
    Entry" is preferred, then any User Entry row that is not driven by a
    formula over the other rows, and only then the first one of the type.
    """

    labels = average_formulas.get("label")
    if isinstance(labels, list):
        for index, label in enumerate(labels):
            normalized = _clean_name(label).lower()
            if normalized == "user entry" or normalized.startswith("user entry "):
                return index

    settings = average_formulas.get("custom_average_formula_settings")
    average_types = settings.get("average_type") if isinstance(settings, dict) else None
    if not isinstance(average_types, list):
        return None
    inputs = average_formulas.get("inputs")
    if not isinstance(inputs, list):
        inputs = average_formulas.get("formulas")
    user_entry_rows = [
        index
        for index, average_type in enumerate(average_types)
        if _clean_name(average_type).lower() == "user_entry"
    ]
    for index in user_entry_rows:
        row_inputs = inputs[index] if isinstance(inputs, list) and index < len(inputs) else None
        if not _references_another_average_row(row_inputs):
            return index
    return user_entry_rows[0] if user_entry_rows else None


def _dfm_ratio_development_labels(payload: dict) -> list[str]:
    ratio_triangle = _dict_path(payload, ("ratios_tab", "ratio_triangle"))
    labels = ratio_triangle.get("development_labels")
    return [_clean_name(label) for label in labels] if isinstance(labels, list) else []


def _ensure_matrix_row(matrix: list, row_index: int) -> list:
    while len(matrix) <= row_index:
        matrix.append([])
    if not isinstance(matrix[row_index], list):
        matrix[row_index] = []
    return matrix[row_index]


def _copy_local_user_entry_inputs(remote_payload: dict, local_payload: dict) -> bool:
    remote_avg = _dict_path(remote_payload, ("ratios_tab", "average_formulas"))
    local_avg = _dict_path(local_payload, ("ratios_tab", "average_formulas"))
    remote_user_row = _average_formula_user_entry_index(remote_avg)
    local_user_row = _average_formula_user_entry_index(local_avg)
    if remote_user_row is None or local_user_row is None:
        return False

    local_inputs = local_avg.get("inputs")
    if not isinstance(local_inputs, list):
        local_inputs = local_avg.get("formulas")
    local_input_row = (
        local_inputs[local_user_row]
        if isinstance(local_inputs, list)
        and local_user_row < len(local_inputs)
        and isinstance(local_inputs[local_user_row], list)
        else []
    )
    local_values = local_avg.get("values")
    local_value_row = (
        local_values[local_user_row]
        if isinstance(local_values, list)
        and local_user_row < len(local_values)
        and isinstance(local_values[local_user_row], list)
        else []
    )
    if not local_input_row and not local_value_row:
        return False

    remote_inputs = remote_avg.get("inputs")
    if not isinstance(remote_inputs, list):
        remote_inputs = []
        remote_avg["inputs"] = remote_inputs
    remote_row = _ensure_matrix_row(remote_inputs, remote_user_row)
    remote_values = remote_avg.get("values")
    if not isinstance(remote_values, list):
        remote_values = []
        remote_avg["values"] = remote_values
    remote_value_row = _ensure_matrix_row(remote_values, remote_user_row)
    local_display_inputs = local_avg.get("display_inputs")
    local_display_input_row = (
        local_display_inputs[local_user_row]
        if isinstance(local_display_inputs, list)
        and local_user_row < len(local_display_inputs)
        and isinstance(local_display_inputs[local_user_row], list)
        else []
    )
    remote_display_inputs = remote_avg.get("display_inputs")
    if not isinstance(remote_display_inputs, list):
        remote_display_inputs = []
        remote_avg["display_inputs"] = remote_display_inputs
    remote_display_input_row = _ensure_matrix_row(remote_display_inputs, remote_user_row)

    remote_dev_labels = _dfm_ratio_development_labels(remote_payload)
    local_dev_labels = _dfm_ratio_development_labels(local_payload)
    remote_label_to_col = {
        label: index
        for index, label in enumerate(remote_dev_labels)
        if label
    }

    copied = False
    for local_col in range(max(len(local_input_row), len(local_value_row))):
        formula_text = _clean_name(local_input_row[local_col] if local_col < len(local_input_row) else "")
        local_value = local_value_row[local_col] if local_col < len(local_value_row) else None
        remote_col = local_col
        if local_col < len(local_dev_labels):
            remote_col = remote_label_to_col.get(local_dev_labels[local_col], local_col)
        while len(remote_row) <= remote_col:
            remote_row.append("")
        while len(remote_value_row) <= remote_col:
            remote_value_row.append(None)
        while len(remote_display_input_row) <= remote_col:
            remote_display_input_row.append("")
        remote_row[remote_col] = formula_text
        remote_value_row[remote_col] = canonical_input_number(local_value)
        remote_display_input_row[remote_col] = _clean_name(
            local_display_input_row[local_col] if local_col < len(local_display_input_row) else ""
        )
        copied = copied or bool(formula_text) or local_value is not None
    return copied


def _preserve_local_dfm_data(remote_payload: dict, local_payload: dict) -> tuple[dict, set[str]]:
    """Rebase every canonical ArcRho-owned setting onto fresh ResQ snapshots."""
    preserved: set[str] = set()
    if not isinstance(local_payload, dict):
        return remote_payload, preserved

    if local_payload.get("json_format") == DFM_JSON_FORMAT:
        base = deepcopy(remote_payload)
        remote_details = _dict_child(base, "details_tab")
        local_details = _dict_path(local_payload, ("details_tab",))
        if _clean_name(local_details.get("input_triangle")) != _clean_name(remote_details.get("input_triangle")):
            base["data_tab"] = deepcopy(_dict_path(local_payload, ("data_tab",)))
            preserved.add("input selection and snapshot")
        remote_results = _dict_child(base, "results_tab")
        local_results = _dict_path(local_payload, ("results_tab",))
        if _clean_name(local_results.get("ratio_basis_dataset")) != _clean_name(
            remote_results.get("ratio_basis_dataset")
        ):
            for key in (
                "ratio_basis_origin_labels",
                "ratio_basis_values",
                "ratio_basis_data_format",
                "ratio_basis_number_format",
                "ratio_basis_decimal_places",
                "ratio_basis_source_revision",
            ):
                remote_results[key] = deepcopy(local_results.get(key))
            preserved.add("ratio basis selection and snapshot")
        refreshed_at = _dict_path(remote_payload, ("method_metadata",)).get("data_refreshed")
        rebased = apply_owned_patch(base, local_payload, timestamp=refreshed_at)
        local_last_modified = _dict_path(local_payload, ("method_metadata",)).get("last_modified")
        if local_last_modified:
            _dict_child(rebased, "method_metadata")["last_modified"] = local_last_modified
        preserved.update({
            "exclusions",
            "formula definitions and selections",
            "stored user values",
            "cell_notes",
        })
        return rebased, preserved

    # Legacy files have no complete owned-state contract. Preserve the two
    # historically local fields until their transactional v2 upgrade succeeds.
    remote_ratios = _dict_child(remote_payload, "ratios_tab")
    remote_notes = remote_ratios.get("cell_notes")
    local_notes = _dict_path(local_payload, ("ratios_tab", "cell_notes"))
    if isinstance(local_notes, dict) and local_notes:
        remote_ratios["cell_notes"] = _merge_cell_note_dicts(
            remote_notes if isinstance(remote_notes, dict) else {},
            local_notes,
        )
        preserved.add("cell_notes")

    if _copy_local_user_entry_inputs(remote_payload, local_payload):
        preserved.add("user entry formulas")
    return remote_payload, preserved


def _strip_formula_index_prefix(raw: str) -> str:
    """Remove leading '0: ' or '13: ' index that ResQ prepends to formula names."""
    raw = raw.strip()
    m = re.match(r"^\d+:\s*", raw)
    return raw[m.end():].strip() if m else raw


# ResQ's AverageType enumeration, in the order the automation help lists it:
# atCustom, atMedian, atGeoMean, atMin, atMax, atUserEntry, atCalculated,
# atPriorAnalysis, atPattern, atBenchmark. Only the two ArcRho reads are named.
RESQ_AVERAGE_TYPE_USER_ENTRY = 5
RESQ_AVERAGE_TYPE_CALCULATED = 6

# "User Calculation" in the ResQ dialog. The formula may hold scalars, the four
# arithmetic operators and Average(<row>) references to other average rows; the
# row number is the position shown in the ratios grid, counting from one.
_RESQ_AVERAGE_REFERENCE_RE = re.compile(r"average\s*\(\s*(\d+)\s*\)", re.I)
_ARCRHO_FORMULA_RESIDUE_RE = re.compile(r"^[\d\s.+\-*/()]*$")


def _read_resq_average_definition(dfm, row_index: int) -> dict:
    """Read one ResQ custom-average row: its type and, if any, its formula.

    ResQ keeps a row's formula even after its type is changed away from User
    Calculation, so a non-empty ``Formula`` proves nothing on its own -- the
    automation help is explicit that the formula "only has any effect if the
    average type is set to atCalculated". Every caller therefore has to gate on
    the type, never on the formula text.
    """

    definition = {"average_type": None, "formula": "", "tail_factor": 1.0}
    try:
        average = dfm.CustomAverages(row_index)
    except Exception:
        return definition
    if average is None:
        return definition
    try:
        definition["average_type"] = int(average.AverageType)
    except Exception:
        definition["average_type"] = None
    try:
        definition["formula"] = str(average.Formula or "").strip()
    except Exception:
        definition["formula"] = ""
    # The row's "- Ult" value on the Ratios tab. ``AverageRatioValues`` at the
    # tail column answers with unallocated memory for most rows, so the tail is
    # read from the row's own TailFactor, which is where ResQ keeps it.
    try:
        tail = float(average.TailFactor)
        definition["tail_factor"] = tail if tail > 0 else 1.0
    except Exception:
        definition["tail_factor"] = 1.0
    return definition


# ResQ's DFMCurveType columns: cvValue (1) is the Initial Selection, cvExpDecay
# .. cvWeibull (2-5) the fitted curves, and cvUser1 .. cvUser10 (6-15) the user
# value columns. Ordinals confirmed against the type library on 2026-09-03.
RESQ_CURVE_FIXED_COLUMNS = 5
# DFMCurveColumnType ordinals for a user value column.
RESQ_CURVE_COLUMN_TYPES = {3: "user_entry", 4: "prior_analysis", 5: "pattern", 6: "benchmark"}
# DFMFittingMethod ordinals.
RESQ_FITTING_METHODS = {0: "log_regression", 1: "least_squares"}


def _read_resq_curves_tab(dfm, period_count: int, *, strict: bool) -> dict:
    """Read the ResQ Curves tab settings and selections into ArcRho's ``curves_tab``.

    Only the choices come across: the fitted curves are recomputed by ArcRho
    (``arcrho_api.dfm_curves``), which reproduces ResQ's log-regression fits.
    A ResQ method fitted by least squares keeps that setting in the payload so
    the Curves tab can say the curves shown are ArcRho's log-regression fits.
    """

    def read(what, default, getter):
        try:
            return getter()
        except Exception as exc:
            _strict_dfm_failure(strict, f"Could not read the ResQ DFM Curves tab {what}.", exc)
            return default

    periods = range(1, period_count + 1)
    user_count = read("user column count", 0, lambda: max(int(dfm.CurveUserValueColCount), 0))
    user_columns = []
    for offset in range(user_count):
        column = RESQ_CURVE_FIXED_COLUMNS + 1 + offset
        column_type = RESQ_CURVE_COLUMN_TYPES.get(
            read(f"type of column {column}", 3, lambda column=column: int(dfm.CurveColumnType(column))),
            "user_entry",
        )
        user_columns.append(
            {
                "label": read(
                    f"description of column {column}",
                    "",
                    lambda column=column: str(dfm.CurveColumnDescription(column) or "").strip(),
                )
                or "User Entry",
                "column_type": column_type,
                "values": [
                    read(
                        f"value of column {column} at period {j}",
                        1.0,
                        lambda column=column, j=j: float(dfm.CurveValues(column, j)),
                    )
                    for j in periods
                ],
                "tail": read(f"tail of column {column}", 1.0, lambda column=column: float(dfm.CurveValues(column, 0))),
            }
        )
    return {
        "fitting_method": RESQ_FITTING_METHODS.get(
            read("fitting method", 0, lambda: int(dfm.FittingMethod)), "log_regression"
        ),
        "future_development_periods": read("future development periods", 1, lambda: int(dfm.FutureDevelopmentPeriods)),
        "free_fit_c": read("Free Fit C flag", False, lambda: bool(dfm.FreeFitC)),
        "included": [
            1 if read(f"inclusion at period {j}", True, lambda j=j: bool(dfm.IncludedRatios(j))) else 0
            for j in periods
        ],
        "user_columns": user_columns,
        "selected_estimates": [
            read(f"selected estimate at period {j}", 1, lambda j=j: int(dfm.SelectedEstimates(j))) for j in periods
        ],
        "selected_tail_factor": read("selected tail factor column", 1, lambda: int(dfm.SelectedTailFactor)),
        "selected_tail_curve": read("selected tail pattern column", 1, lambda: int(dfm.SelectedTailCurve)),
    }


def _translate_resq_average_formula(
    formula: str,
    resq_idx_map: list[int],
    formula_labels: list[str],
    own_row: int,
) -> str | None:
    """Rewrite a ResQ ``Average(n)`` formula as an ArcRho in-cell formula.

    ArcRho names another summary row by quoting its label, so
    ``(Average(5)+Average(6)+Average(7))/3`` becomes
    ``="Simple - 5"+"Simple - 3"+"Simple - 5 Ex hi/lo"`` over three, and the
    ratios tab then recalculates the row like any other User Entry cell.

    Returns ``None`` whenever the formula cannot be carried across faithfully --
    a row ArcRho did not import, a label shared by two rows, a self-reference,
    or anything left over that is not plain arithmetic. The caller then keeps
    the values ResQ computed rather than showing a formula that means something
    different here.
    """

    text = " ".join(str(formula or "").split())
    if not text:
        return None

    # ArcRho resolves a reference by label, so a label two rows share cannot be
    # named unambiguously and the whole formula has to decline.
    label_counts = Counter(_clean_name(label).casefold() for label in formula_labels)

    failed = False

    def replace(match: re.Match) -> str:
        nonlocal failed
        raw_index = int(match.group(1)) - 1
        row = next(
            (index for index, mapped in enumerate(resq_idx_map) if mapped == raw_index),
            None,
        )
        if row is None or row == own_row:
            failed = True
            return ""
        label = _clean_name(formula_labels[row])
        if not label or label_counts[label.casefold()] > 1:
            failed = True
            return ""
        return f'"{label}"'

    translated = _RESQ_AVERAGE_REFERENCE_RE.sub(replace, text)
    if failed or translated == text:
        # Nothing was substituted: either a reference could not be mapped, or
        # the formula never referenced another row and is a bare constant that
        # the stored values already carry.
        return None
    if not _ARCRHO_FORMULA_RESIDUE_RE.match(_RESQ_AVERAGE_REFERENCE_RE.sub("", text)):
        # A function call, a cell reference or anything else ArcRho's arithmetic
        # evaluator would not understand.
        return None
    return f"={translated}"


def _infer_avg_settings(label: str) -> dict:
    norm = " ".join(label.split()).strip()
    lower = norm.lower()
    if lower.startswith("user"):
        return {"average_type": "user_entry", "base": "simple", "periods": "all", "exclude": 0}
    if "benchmark" in lower:
        return {"average_type": "custom", "base": "benchmark", "periods": "all", "exclude": 0}
    m = re.match(
        r"^(volume|simple)\s*-\s*(all|[1-9]\d*)(\s+ex\s+hi/lo(?:\s*x\s*([1-9]\d*))?)?$",
        norm, re.I,
    )
    if m:
        base = m.group(1).lower()
        p = m.group(2).lower()
        periods: str | int = p if p == "all" else int(p)
        ex = int(m.group(4) or 0)
        if m.group(3) and ex == 0:
            ex = 1
        return {"average_type": "custom", "base": base, "periods": periods, "exclude": ex}
    return {"average_type": "custom", "base": "simple", "periods": "all", "exclude": 0}

def _recreate_adjustment_formulas(
    notes: str,
    ratio_dev_labels: list[str],
    formula_labels: list[str],
    selected: list[list[int]],
    values: list[list],
    avg_inputs: list[list[str]],
    decimal_places: int,
) -> int:
    """Give a User Entry value back the combined-adjustment formula it came from.

    ResQ keeps only the number, but the notes the "Generate Notes for Combined
    Adjustment" macro wrote beside it name the average row and the adjustment
    vectors it was built from. A column whose selected User Entry value still
    matches the selected LDF its note block states gets that formula back; the
    value stays ResQ's until the next recalculation. A value that no longer
    matches was changed by hand after the notes were written and stays a number.
    """

    user_row = next((row for row, label in enumerate(formula_labels) if label == "User Entry"), None)
    if user_row is None:
        return 0
    labels = {_clean_name(label).casefold() for label in formula_labels}
    columns = {_clean_name(label).casefold(): col for col, label in enumerate(ratio_dev_labels)}
    precision = min(int(decimal_places), BASE_FACTOR_DECIMALS)
    recreated = 0
    for period, block in parse_adjustment_notes(notes).items():
        col = columns.get(_clean_name(period).casefold())
        base_label = block["base_label"]
        if col is None or not base_label or not block["terms"] or block["value"] is None:
            continue
        if _clean_name(base_label).casefold() not in labels or not selected[user_row][col]:
            continue
        value = values[user_row][col] if col < len(values[user_row]) else None
        if value is None or abs(round(float(value), precision) - round(block["value"], precision)) > 1e-9:
            continue
        avg_inputs[user_row][col] = adjustment_formula(base_label, block["terms"], col)
        recreated += 1
    return recreated


def _ratio_label_endpoints(label: str) -> tuple[int | None, int | None]:
    text = re.sub(r"^\(?\s*\d+\s*\)?\s*", "", label).strip()
    m = re.match(r"^(\d+)\s*[-–]\s*(\d+)$", text)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None

def _build_data_dev_labels(ratio_dev_labels: list[str]) -> list[str]:
    """
    Derive cumulative-age labels ('2m', '5m', …) for the data tab from
    the period-to-period ratio labels ('(1) 2-5', '(2) 5-8', …, '119 - Ult').
    """
    ages: list[int] = []
    for label in ratio_dev_labels:
        if "ult" in label.lower():
            break
        start, end = _ratio_label_endpoints(label)
        if start is not None and not ages:
            ages.append(start)
        if end is not None:
            ages.append(end)
    return [f"{a}m" for a in ages]

def _parse_cell_notes(raw: str, origin_labels: list[str], avg_labels: list[str]) -> dict:
    """
    Parse ResQ CellNotes text into the JSON cell-notes dict.

    ResQ format per line:
        "Ratios.Ratios & Average Selection", Cell[<dev>, <row>], "<note>", User: ..., Date: ...

    If <row> matches an origin label → ratio main table.
    If <row> matches an average formula label → ratio summary table.
    """
    result: dict = {"ratio_main_table": {}, "ratio_summary_table": {}}
    if not raw:
        return result

    origin_lower = {l.strip().lower() for l in origin_labels}
    avg_lower = {l.strip().lower() for l in avg_labels}

    pattern = re.compile(r'"[^"]+",\s*Cell\[([^\]]+)\],\s*"([^"]*)"')
    for m in pattern.finditer(raw):
        cell_ref = m.group(1)
        note_text = m.group(2)
        parts = [p.strip() for p in cell_ref.split(",", 1)]
        if len(parts) != 2:
            continue
        dev_part, row_part = parts
        row_lower = row_part.lower()
        if row_lower in origin_lower or any(row_lower in ol for ol in origin_lower):
            table = "ratio_main_table"
        elif row_lower in avg_lower or any(row_lower in al for al in avg_lower):
            table = "ratio_summary_table"
        else:
            table = "ratio_summary_table"  # default
        result[table].setdefault(dev_part, {})[row_part] = note_text

    return result

RESQ_RATIO_INCLUDED = 0
RESQ_RATIO_EXCLUDED = 1
RESQ_RATIO_EMPTY_CELL = 2


def _excluded_ratio_flag(value: object) -> int:
    """Map a ResQ ``ExcludedRatios`` code to ArcRho's 0/1 exclusion flag.

    ResQ reports 0=included, 1=excluded, 2=empty cell. Only 1 records an
    actuary's exclusion; an empty cell carries no judgement and must not import
    as excluded.
    """
    return 1 if int(value) == RESQ_RATIO_EXCLUDED else 0


def _get_ratio_value(dfm, i: int, j: int) -> float | None:
    try:
        v = dfm.Ratios(OriginIndex=i, DevIndex=j)
        return float(v) if v is not None else None
    except Exception:
        return None


def _strict_dfm_failure(strict: bool, message: str, exc: Exception | None = None) -> None:
    if not strict:
        return
    error = RuntimeError(message)
    if exc is None:
        raise error
    raise error from exc


def _dfm_attr(source, name: str, default, *, strict: bool, context: str):
    if not strict:
        return _safe_attr(source, name, default)
    try:
        return getattr(source, name)
    except Exception as exc:
        _strict_dfm_failure(strict, f"Could not read ResQ DFM {context}.{name}.", exc)


def _ratio_basis_snapshot(
    dfm,
    name: str,
    origin_labels: list[str],
    rc_path: str,
    *,
    strict: bool = False,
) -> dict:
    if not name:
        return {}
    basis = _safe_attr(dfm, "SummaryRatioBasis", None)
    if basis is None:
        raise ValueError(f"Unable to read DFM Ratio Basis dataset {name!r} from ResQ.")

    try:
        basis_values = [
            canonical_input_number(basis.ValuesByIndex(index))
            for index in range(1, int(dfm.OriginCount) + 1)
        ]
    except Exception as exc:
        raise ValueError(f"Unable to read DFM Ratio Basis dataset {name!r} from ResQ: {exc}") from exc

    dataset_type_obj = _dfm_attr(
        basis, "DatasetType", None, strict=strict, context="ratio-basis dataset"
    )
    dataset_type = _normalize_import_name(
        _dfm_attr(
            dataset_type_obj,
            "Name",
            "",
            strict=strict,
            context="ratio-basis Dataset Type",
        )
    ) or name
    revision = _iso_or_text(
        _dfm_attr(basis, "Modified", "", strict=strict, context="ratio-basis dataset")
    )
    return {
        "name": name,
        "origin_labels": origin_labels,
        "values": basis_values,
        "data_format": "Vector",
        "number_format": dataset_type_number_format(rc_path, dataset_type),
        "decimal_places": dataset_type_decimal_places(rc_path, dataset_type),
        "revision": revision,
    }

def export_dfm(
    dfm,
    rc_path: str,
    project_data_dir: Path,
    *,
    max_average_formula_probe: int = MAX_AVERAGE_FORMULA_PROBE,
    ratio_basis_snapshot: dict | None = None,
    strict: bool = False,
) -> dict:
    """Extract all DFM data from a ResQ DFM COM object and return a JSON-ready dict."""
    del project_data_dir
    name = _normalize_import_name(_dfm_attr(dfm, "Name", "", strict=strict, context="method"))
    input_triangle = _dfm_attr(dfm, "InputTriangle", None, strict=strict, context=name or "method")
    output_vector = _dfm_attr(dfm, "OutputVector", None, strict=strict, context=name or "method")
    input_tri_name = _normalize_import_name(
        _dfm_attr(input_triangle, "Name", "", strict=strict, context="input_triangle")
    )
    output_vec_name = _normalize_import_name(
        _dfm_attr(output_vector, "Name", "", strict=strict, context="output vector")
    )
    output_dataset_type_obj = _dfm_attr(
        output_vector, "DatasetType", None, strict=strict, context="output vector"
    )
    output_dataset_type = _normalize_import_name(
        _dfm_attr(output_dataset_type_obj, "Name", "", strict=strict, context="output Dataset Type")
    ) or output_vec_name
    output_category_obj = _dfm_attr(
        output_dataset_type_obj, "Category", None, strict=strict, context="output Dataset Type"
    )
    output_category = (
        _normalize_import_name(
            _dfm_attr(output_category_obj, "Name", "", strict=strict, context="output_category")
        )
        if output_category_obj is not None
        else ""
    )
    origin_length = int(_dfm_attr(dfm, "OriginLength", 0, strict=strict, context=name or "method"))
    dev_length = int(_dfm_attr(dfm, "DevelopmentLength", 0, strict=strict, context=name or "method"))
    decimal_places = int(_dfm_attr(dfm, "RatioDecimalPlaces", 0, strict=strict, context=name or "method"))

    try:
        ultimate_dp: int = dfm.SummaryRatioDecimalPlaces
    except Exception as exc:
        _strict_dfm_failure(strict, "Could not read the ResQ DFM SummaryRatioDecimalPlaces.", exc)
        ultimate_dp = 2

    try:
        ratio_basis_object = dfm.SummaryRatioBasis
    except Exception as exc:
        _strict_dfm_failure(strict, "Could not read the ResQ DFM SummaryRatioBasis.", exc)
        ratio_basis_object = None
    try:
        ratio_basis = _normalize_import_name(ratio_basis_object.Name) if ratio_basis_object is not None else ""
    except Exception as exc:
        _strict_dfm_failure(strict, "Could not read the ResQ DFM SummaryRatioBasis name.", exc)
        ratio_basis = ""

    try:
        modified = dfm.OutputVector.Modified
        last_modified = _iso_or_text(modified)
    except Exception as exc:
        _strict_dfm_failure(strict, "Could not read the ResQ DFM output Modified timestamp.", exc)
        last_modified = datetime.now(timezone.utc).astimezone().isoformat()

    input_payload = export_triangle(input_triangle, strict=strict)
    origin_labels = [str(item) for item in input_payload.get("origin_labels", [])]
    data_dev_labels = [str(item) for item in input_payload.get("development_labels", [])]
    input_values = input_payload.get("values") if isinstance(input_payload.get("values"), list) else []
    input_mask = [
        [canonical_number(value) is not None for value in row] if isinstance(row, list) else []
        for row in input_values
    ]
    input_dataset_type = _normalize_import_name(input_payload.get("dataset_type")) or input_tri_name

    origin_count: int = len(origin_labels)
    dev_count: int = len(data_dev_labels)
    org_rng = range(1, origin_count + 1)
    dev_rng = range(1, dev_count + 1)
    ratio_dev_labels = [_normalize_import_name(dfm.DevelopmentLabel(j)) for j in dev_rng]

    # Ratio triangle values (staircase shape)
    ratio_values: list[list] = []
    excluded: list[list] = []
    for i in org_rng:
        row_dev = dfm.DevelopmentCount(i)
        rv_row: list = []
        ex_row: list = []
        for j in range(1, row_dev + 1):
            val = _get_ratio_value(dfm, i, j)
            if strict and val is None:
                raise RuntimeError(f"Could not read ResQ DFM ratio at ({i}, {j}).")
            rv_row.append(round(val, decimal_places) if val is not None else 0)
            try:
                ex_row.append(_excluded_ratio_flag(dfm.ExcludedRatios(i, j)))
            except Exception as exc:
                _strict_dfm_failure(strict, f"Could not read ResQ DFM exclusion at ({i}, {j}).", exc)
                ex_row.append(0)
        ratio_values.append(rv_row)
        excluded.append(ex_row)

    # Enumerate average formula names from ResQ (1-based, strip index prefix)
    raw_names: list[str] = []
    for idx in range(1, max_average_formula_probe + 1):
        try:
            f = dfm.AverageFormula(idx)
            if f is None:
                break
            raw_names.append(f)
            if strict and _strip_formula_index_prefix(f).lower().startswith("user"):
                break
        except Exception as exc:
            if strict and not any(
                _strip_formula_index_prefix(value).lower().startswith("user")
                for value in raw_names
            ):
                _strict_dfm_failure(
                    strict,
                    f"Could not finish enumerating ResQ DFM average formulas at index {idx}.",
                    exc,
                )
            break

    # Deduplicate: keep only the first User Entry; record its ResQ index
    formula_labels: list[str] = []
    resq_idx_map: list[int] = []   # formula_labels[k] came from ResQ formula index resq_idx_map[k]+1
    user_entry_seen = False
    for list_idx, raw in enumerate(raw_names):
        cleaned = _strip_formula_index_prefix(raw)
        is_user = cleaned.lower().startswith("user")
        if is_user:
            if not user_entry_seen:
                user_entry_seen = True
                formula_labels.append("User Entry")
                resq_idx_map.append(list_idx)
        else:
            formula_labels.append(cleaned)
            resq_idx_map.append(list_idx)

    n_formulas = len(formula_labels)

    # ResQ's "User Calculation" rows: an average defined as arithmetic over the
    # other average rows rather than over the ratio triangle. ArcRho has no such
    # row type, but its User Entry row accepts exactly the same kind of in-cell
    # formula, so each one is imported as a User Entry row under its ResQ name
    # with the formula rewritten into ArcRho's own reference syntax.
    calculated_formulas: dict[int, str] = {}
    definitions = [_read_resq_average_definition(dfm, raw_idx_0 + 1) for raw_idx_0 in resq_idx_map]
    for row, definition in enumerate(definitions):
        if definition["average_type"] != RESQ_AVERAGE_TYPE_CALCULATED:
            continue
        translated = _translate_resq_average_formula(
            definition["formula"], resq_idx_map, formula_labels, row
        )
        if translated:
            calculated_formulas[row] = translated
        else:
            # Nothing portable to carry across, so the row keeps the numbers
            # ResQ computed and stops moving with the triangle, the same
            # treatment a loaded benchmark row gets.
            calculated_formulas[row] = ""

    # selected[formula_row][dev_col] = 1 when that formula is selected
    selected = [[0] * dev_count for _ in range(n_formulas)]
    for j in dev_rng:
        try:
            sel = int(dfm.SelectedRatios(DevIndex=j))  # 1-based ResQ formula index
        except Exception as exc:
            _strict_dfm_failure(strict, f"Could not read the selected ResQ DFM average at column {j}.", exc)
            continue
        # sel is 1-based index into raw_names; find in resq_idx_map
        raw_idx_0 = sel - 1  # 0-based into raw_names
        matched = False
        for k, mapped_raw_idx in enumerate(resq_idx_map):
            if mapped_raw_idx == raw_idx_0:
                selected[k][j - 1] = 1
                matched = True
                break
        if strict and not matched:
            raise RuntimeError(
                f"ResQ DFM selected average index {sel} at column {j} was not present "
                "in the enumerated formula list."
            )

    # values[formula_row][dev_col] = computed average LDF; the last column is
    # the row's "- Ult" tail factor.
    #
    # Both are kept exactly as ResQ holds them. A User Entry row and a benchmark
    # row are never re-derived from the triangle, so whatever is stored here is
    # the factor ArcRho's own ultimate chains; rounding it to the Details tab's
    # Decimal Places would make the chain reproduce the printed number instead
    # of ResQ's. The Ratios tab still prints it at the display precision.
    values: list[list] = []
    for k, raw_idx_0 in enumerate(resq_idx_map):
        resq_formula_idx = raw_idx_0 + 1  # back to 1-based
        row: list = []
        for j in dev_rng:
            if j == dev_count:
                row.append(float(definitions[k]["tail_factor"]))
                continue
            try:
                v = dfm.AverageRatioValues(j, resq_formula_idx)
                row.append(float(v) if v is not None else None)
            except Exception as exc:
                _strict_dfm_failure(
                    strict,
                    f"Could not read ResQ DFM average value at formula {resq_formula_idx}, column {j}.",
                    exc,
                )
                row.append(None)
        # trim trailing None
        while row and row[-1] is None:
            row.pop()
        values.append(row)

    # Custom average formula settings
    avg_settings: dict = {"average_type": [], "base": [], "periods": [], "exclude": []}
    for row, label in enumerate(formula_labels):
        if row in calculated_formulas:
            s = (
                {"average_type": "user_entry", "base": "simple", "periods": "all", "exclude": 0}
                if calculated_formulas[row]
                else {"average_type": "custom", "base": "benchmark", "periods": "all", "exclude": 0}
            )
        else:
            s = _infer_avg_settings(label)
        avg_settings["average_type"].append(s["average_type"])
        avg_settings["base"].append(s["base"])
        avg_settings["periods"].append(s["periods"])
        avg_settings["exclude"].append(s["exclude"])

    # A translated User Calculation row carries its formula in every ratio
    # column, the way ResQ applies one definition across the whole row; the
    # tail column is the row's own entered tail factor, never the formula.
    avg_inputs = [[""] * dev_count for _ in range(n_formulas)]
    for row, formula in calculated_formulas.items():
        if formula:
            avg_inputs[row] = [formula] * max(dev_count - 1, 0) + [""] * min(dev_count, 1)
    _recreate_adjustment_formulas(
        str(_safe_attr(dfm, "Notes", "") or ""),
        ratio_dev_labels,
        formula_labels,
        selected,
        values,
        avg_inputs,
        decimal_places,
    )

    curves_tab =_read_resq_curves_tab(dfm, max(dev_count - 1, 0), strict=strict)

    # Notes
    # Cell notes
    try:
        cell_notes_raw: str = dfm.CellNotes or ""
    except Exception as exc:
        _strict_dfm_failure(strict, "Could not read ResQ DFM cell notes.", exc)
        cell_notes_raw = ""
    cell_notes = _parse_cell_notes(cell_notes_raw, origin_labels, formula_labels)

    base_payload = {
        "json_format": DFM_JSON_FORMAT,
        "details_tab": {
            "name": name,
            "output_type": output_dataset_type,
            "output_dataset": output_vec_name,
            "output_category": output_category,
            "input_triangle": input_tri_name,
            "origin_length": origin_length,
            "development_length": dev_length,
            "decimal_places": decimal_places,
        },
        "data_tab": {
            "origin_labels": origin_labels,
            "development_labels": data_dev_labels,
            "input_data_triangle_values": input_values,
            "input_data_triangle_mask": input_mask,
            "data_format": "Triangle",
            "number_format": dataset_type_number_format(rc_path, input_dataset_type),
            "decimal_places": dataset_type_decimal_places(rc_path, input_dataset_type),
            "source_revision": _iso_or_text(input_payload.get("modified")),
        },
        "ratios_tab": {
            "ratio_triangle": {
                "origin_labels": origin_labels,
                "development_labels": ratio_dev_labels,
                "ratio_values": ratio_values,
                "excluded": excluded,
            },
            "average_formulas": {
                "label": formula_labels,
                "custom_average_formula_settings": avg_settings,
                "selected": selected,
                "values": values,
                "inputs": avg_inputs,
                "display_inputs": [[""] * dev_count for _ in range(n_formulas)],
            },
            "cell_notes": cell_notes,
        },
        "curves_tab": curves_tab,
        "results_tab": {
            "ratio_basis_dataset": ratio_basis,
            "ratio_basis_origin_labels": [],
            "ratio_basis_values": [],
            "ratio_basis_data_format": "Vector",
            "ratio_basis_number_format": "#,##0",
            "ratio_basis_decimal_places": 0,
            "ratio_basis_source_revision": "",
            "ultimate_ratio_decimal_places": ultimate_dp,
            "ultimate_vector": [],
        },
        "method_metadata": {
            "last_modified": last_modified,
            "data_refreshed": last_modified,
        },
    }
    input_snapshot = {
        "name": input_tri_name,
        "origin_labels": origin_labels,
        "development_labels": data_dev_labels,
        "values": input_values,
        "mask": input_mask,
        "data_format": "Triangle",
        "number_format": dataset_type_number_format(rc_path, input_dataset_type),
        "decimal_places": dataset_type_decimal_places(rc_path, input_dataset_type),
        "revision": _iso_or_text(input_payload.get("modified")),
    }
    return recalculate_dfm_method(
        base_payload,
        input_snapshot=input_snapshot,
        ratio_basis_snapshot=(
            ratio_basis_snapshot
            if ratio_basis_snapshot is not None
            else _ratio_basis_snapshot(
                dfm,
                ratio_basis,
                origin_labels,
                rc_path,
                strict=strict,
            )
            if ratio_basis
            else None
        ),
        timestamp=last_modified,
    )


def dfm_methods_by_output_name(reserving_class, dfm_names: list[str] | None = None) -> dict[str, tuple[str, object]]:
    try:
        dfm_collection = reserving_class.DFMMethods()
    except Exception:
        return {}
    requested = {
        _clean_name(name).casefold()
        for name in (dfm_names or [])
        if _clean_name(name)
    } if dfm_names is not None else None
    out: dict[str, tuple[str, object]] = {}
    for dfm in dfm_collection:
        clean_name = _clean_name(_safe_attr(dfm, "Name", ""))
        if not clean_name:
            continue
        if requested is not None and clean_name.casefold() not in requested:
            continue
        output_vector = _safe_attr(dfm, "OutputVector", None)
        output_name = _normalize_import_name(_safe_attr(output_vector, "Name", ""))
        key = output_name.lower()
        if key and key not in out:
            out[key] = (clean_name, dfm)
    return out


def export_dfm_output_dataset(
    dfm,
    rc_path: str,
    rc_dir: Path,
    *,
    project_name: str,
    project_data_dir: Path,
    method_data_dir: str,
    debug_log,
    log,
    known_dataset_type_keys: set[str] | None = None,
    max_average_formula_probe: int = MAX_AVERAGE_FORMULA_PROBE,
    ratio_basis_snapshot: dict | None = None,
    preserve_local_owned_state: bool = True,
    strict: bool = False,
    verbose: bool = True,
) -> tuple[str, str, bool]:
    """Publish one ResQ DFM output and its canonical method artifacts.

    The default keeps the existing migration behavior: ArcRho-owned settings
    are rebased onto the fresh ResQ snapshot.  Callers performing an explicitly
    ResQ-authoritative synchronization may disable that rebase so the ResQ
    method payload, including its ``last modified`` timestamp, is retained.
    """
    dfm_name = _normalize_import_name(
        _dfm_attr(dfm, "Name", "", strict=strict, context="method")
    )
    output_vector = _dfm_attr(
        dfm, "OutputVector", None, strict=strict, context=dfm_name or "method"
    )
    output_dataset_name = _normalize_import_name(
        _dfm_attr(output_vector, "Name", "", strict=strict, context="output vector")
    ) or dfm_name
    dataset_type_obj = _dfm_attr(
        output_vector, "DatasetType", None, strict=strict, context="output vector"
    )
    output_dataset_type = _normalize_import_name(
        _dfm_attr(dataset_type_obj, "Name", "", strict=strict, context="output Dataset Type")
    ) or output_dataset_name
    if not _is_known_dataset_type(output_dataset_type, known_dataset_type_keys):
        detail = _unknown_dataset_type_skip_detail("DFM", output_dataset_name, output_dataset_type)
        log(verbose, detail)
        return output_dataset_name, detail, True
    file_name = f"DFM@{_encode_name_part(dfm_name)}.json"
    out_path = rc_dir / method_data_dir / file_name
    export_kwargs = {
        "max_average_formula_probe": max_average_formula_probe,
        "strict": strict,
    }
    if ratio_basis_snapshot is not None:
        export_kwargs["ratio_basis_snapshot"] = ratio_basis_snapshot
    payload = export_dfm(dfm, rc_path, project_data_dir, **export_kwargs)
    details_tab = payload.get("details_tab") if isinstance(payload.get("details_tab"), dict) else {}
    output_dataset_name = _normalize_import_name(details_tab.get("output_dataset")) or output_dataset_name
    debug_log(
        "dfm_export_payload",
        project_name=project_name,
        reserving_class=rc_path,
        method_name=payload.get("details_tab", {}).get("name") if isinstance(payload.get("details_tab"), dict) else dfm_name,
        input_triangle=payload.get("details_tab", {}).get("input_triangle") if isinstance(payload.get("details_tab"), dict) else "",
        origin_length=payload.get("details_tab", {}).get("origin_length") if isinstance(payload.get("details_tab"), dict) else "",
        development_length=payload.get("details_tab", {}).get("development_length") if isinstance(payload.get("details_tab"), dict) else "",
        input_source_revision=payload.get("data_tab", {}).get("source_revision") if isinstance(payload.get("data_tab"), dict) else "",
    )
    existing_payload = _safe_read_json(out_path)
    preserved: set[str] = set()
    if preserve_local_owned_state:
        payload, preserved = _preserve_local_dfm_data(payload, existing_payload)
    payload = recalculate_dfm_method(
        payload,
        timestamp=payload.get("method_metadata", {}).get("data_refreshed")
        if isinstance(payload.get("method_metadata"), dict)
        else None,
        update_refresh_timestamp=False,
    )
    ultimate_payload = export_dfm_ultimate_vector(
        dfm,
        payload["data_tab"]["origin_labels"],
        payload["details_tab"]["origin_length"],
        payload["details_tab"]["development_length"],
    )
    if strict and any(
        not isinstance(row, list) or not row or row[0] is None
        for row in ultimate_payload.get("values", [])
    ):
        raise RuntimeError("Could not read every ResQ DFM ultimate value in strict mode.")
    if not _is_known_dataset_type(ultimate_payload.get("dataset_type"), known_dataset_type_keys):
        detail = _unknown_dataset_type_skip_detail("DFM", output_dataset_name, ultimate_payload.get("dataset_type"))
        log(verbose, detail)
        return output_dataset_name, detail, True
    ultimate_payload["origin_labels"] = list(payload["data_tab"]["origin_labels"])
    ultimate_payload["origin_count"] = len(payload["data_tab"]["origin_labels"])
    ultimate_payload["values"] = [[value] for value in payload["results_tab"]["ultimate_vector"]]
    ultimate_payload["method_name"] = payload["details_tab"]["name"]
    ultimate_payload["precedents"] = [
        value
        for value in (
            payload["details_tab"].get("input_triangle"),
            payload["results_tab"].get("ratio_basis_dataset"),
        )
        if _clean_name(value)
    ]
    ultimate_payload["publication_revision"] = payload["method_metadata"]["publication_revision"]
    ultimate_csv_path, publication_files, sidecar_path = build_dfm_ultimate_publication(
        ultimate_payload,
        payload,
        rc_path,
        rc_dir,
    )
    publication_files[out_path] = persisted_json_text(persisted_projection(payload)).encode("utf-8")
    publish_dfm_artifacts(publication_files, sidecar_path=sidecar_path)

    suffix = f" (preserved {', '.join(sorted(preserved))})" if preserved else ""
    log(verbose, f"    OK  {_clean_name(ultimate_csv_path.name)}")
    detail = f"    OK  {file_name}{suffix}"
    log(verbose, detail)
    return output_dataset_name, detail, False
