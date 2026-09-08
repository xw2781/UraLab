"""Canonical, location-independent contract and calculations for DFM methods.

The functions in this module are intentionally free of filesystem and web
framework dependencies.  Persisted-data producers supply source snapshots and
delegate normalization and calculation here so an equivalent logical DFM emits
an equivalent JSON payload regardless of the producer or machine path.
"""

from __future__ import annotations

import ast
import math
import re
from copy import deepcopy
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
from typing import Any, Iterable, Mapping

from .dataset_display_contract import normalize_show_subtotal
from .dfm_curves import curves_table, curves_tab_is_default, normalize_curves_tab, owned_curves_tab
from .revision_contract import fingerprint
from .sidecar_audit_contract import (
    AUDIT_ACTION_INSERT,
    AUDIT_ACTION_UPDATE,
    append_audit_entry,
    normalize_audit_log,
)
from .sidecar_core_contract import (
    DATASET_SIDECAR_JSON_FORMAT,
    dependency_entries,
    dependency_names,
    stored_length_fields,
    validate_sidecar_core,
)
from .timestamps import persisted_timestamp as _timestamp


DFM_JSON_FORMAT = "arcrho-dfm-v4"
DFM_VALUE_DECIMAL_PLACES = 6
# The Details tab's Decimal Places: how many decimals the Ratios tab prints,
# and the precision an average row is read at inside a User Entry formula.
DFM_DETAILS_DECIMAL_PLACES = 4
_QUANTUM = Decimal("0.000001")
_EXCEL_REFERENCE_RE = re.compile(
    r"(?:\[[^\]]+\]|(?:^|[=+\-*/,(])\s*'(?:[^']|'')+'!\s*\$?[A-Za-z]{1,3}\$?\d+|"
    r"(?:^|[=+\-*/,(])\s*[^\s+\-*/(),]+!\s*\$?[A-Za-z]{1,3}\$?\d+)",
    re.IGNORECASE,
)
_INTERNAL_LABEL_RE = re.compile(r'"([^"]+)"')


class DfmContractError(ValueError):
    """Raised when a DFM payload cannot satisfy the canonical v2 contract."""


def _clean(value: Any) -> str:
    return " ".join(str(value if value is not None else "").split()).strip()


def _integer(value: Any, default: int, *, minimum: int = 0, maximum: int | None = None) -> int:
    try:
        result = int(value)
    except (TypeError, ValueError):
        result = default
    result = max(minimum, result)
    return min(result, maximum) if maximum is not None else result


def _canonical(value: Any, quantum: Decimal) -> float | int | None:
    if value is None or isinstance(value, bool) or value == "":
        return None
    try:
        number = float(value)
        if not math.isfinite(number):
            return None
        with localcontext() as context:
            # A wide origin figure carried to several decimals needs more digits
            # than the default 28-digit context allows, and an overflow there
            # would quietly turn the figure into "no value" instead of a number.
            context.prec = 60
            rounded = Decimal(str(abs(number))).quantize(quantum, rounding=ROUND_HALF_UP)
    except (TypeError, ValueError, InvalidOperation):
        return None
    result = float(rounded)
    if number < 0:
        result = -result
    if result == 0:
        result = 0.0
    if isinstance(value, int) and not isinstance(value, bool):
        return int(result)
    return result


def canonical_number(value: Any) -> float | int | None:
    """Return one JSON number rounded half-away-from-zero to six decimals."""

    return _canonical(value, _QUANTUM)


def round_half_up(value: float, digits: int = 0) -> float:
    """Round half-away-from-zero to *digits* decimals, as a reader rounds by hand.

    This is the ROUND function of a User Entry formula. It rounds the decimal
    text of the number rather than its binary double, so 2.38625 to four places
    is 2.3863 here, in the browser's evaluator and in the notes alike.
    """

    quantum = Decimal(1).scaleb(-int(digits))
    return float(_canonical(float(value), quantum))


def average_row_reference_value(value: Any, decimal_places: Any) -> float | None:
    """Return the value an average row contributes to a User Entry formula.

    A row enters the formula at the precision the Ratios tab prints it at, the
    method's own Decimal Places, rather than at the six decimals it is stored
    with. A reviewer then multiplies the digits shown in front of them and
    lands on the User Entry factor exactly, with no rounding step hidden inside
    the formula text. ``dfm_ratio_calc.js`` mirrors this for the browser, so
    both evaluators reach the same number.
    """

    number = canonical_number(value)
    if number is None:
        return None
    return round_half_up(
        float(number), _integer(decimal_places, DFM_DETAILS_DECIMAL_PLACES, minimum=0, maximum=8)
    )


def canonical_input_number(value: Any) -> float | int | None:
    """Return one input-triangle number at the precision it was observed with.

    Six decimals is the right precision for a *derived* figure a reader checks
    by eye. Any fixed decimal place is the wrong precision for the observed
    triangle every ratio and every average divides, because how much of a
    number it keeps depends only on how large that number happens to be. Ten
    decimals was generous for a loss figure and still too coarse for a near-zero
    "% of" figure, where the trimmed tail moved a ratio in its fourth decimal.
    The observation is therefore kept exactly as read, so a ratio is divided
    from the number the source holds rather than from a copy of it.

    Nothing is given up in the file. A JSON number round-trips a double
    exactly, and the text written is the shortest that reads back as the same
    value, so an ordinary figure still reads as an ordinary figure.
    """

    if value is None or isinstance(value, bool) or value == "":
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(number):
        return None
    if number == 0:
        number = 0.0
    if isinstance(value, int):
        return int(number)
    return number


_MONTH_BY_NAME = {
    "jan": 1, "january": 1, "feb": 2, "february": 2,
    "mar": 3, "march": 3, "apr": 4, "april": 4,
    "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10, "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}


def _origin_start_month(label: Any, period_length: int) -> tuple[int, int] | None:
    text = str(label if label is not None else "").strip()
    if period_length == 12 and re.fullmatch(r"\d{4}", text):
        return int(text), 1
    if period_length in {3, 6}:
        token = "Q" if period_length == 3 else "H"
        count = 4 if period_length == 3 else 2
        for pattern, reversed_parts in (
            (rf"(\d{{4}})\s*{token}([1-{count}])", False),
            (rf"{token}([1-{count}])\s*(\d{{4}})", True),
        ):
            match = re.fullmatch(pattern, text, re.IGNORECASE)
            if match:
                period, year = (int(match.group(1)), int(match.group(2))) if reversed_parts else (
                    int(match.group(2)), int(match.group(1))
                )
                return year, (period - 1) * period_length + 1
    if period_length == 1:
        match = re.fullmatch(r"(\d{4})(\d{2})", text)
        if match and 1 <= int(match.group(2)) <= 12:
            return int(match.group(1)), int(match.group(2))
        match = re.fullmatch(r"([A-Za-z]{3,9})\s+(\d{4})", text)
        if match and match.group(1).casefold() in _MONTH_BY_NAME:
            return int(match.group(2)), _MONTH_BY_NAME[match.group(1).casefold()]
    return None


def aggregate_vector_values(
    values: Iterable[Any],
    origin_labels: Iterable[Any],
    base_length: int,
    target_length: int,
) -> list[float | int | None]:
    """Aggregate a vector with the shared exact chronological-bucket rule."""

    vector = [canonical_number(item[0] if isinstance(item, list) and item else item) for item in values]
    labels = [str(item if item is not None else "") for item in origin_labels]
    factor = target_length // base_length if base_length and target_length % base_length == 0 else 0
    if factor <= 1 or not vector:
        return []
    buckets: dict[tuple[int, int], list[Any]] = {}
    order: list[tuple[int, int]] = []
    if len(labels) == len(vector) and base_length in {1, 3, 6, 12}:
        for label, value in zip(labels, vector):
            parsed = _origin_start_month(label, base_length)
            if parsed is None:
                buckets = {}
                break
            year, month = parsed
            key = (year, ((month - 1) // target_length) * target_length + 1)
            if key not in buckets:
                buckets[key] = []
                order.append(key)
            buckets[key].append(value)
    groups = [buckets[key] for key in order] if buckets else [
        vector[index:index + factor] for index in range(0, len(vector), factor)
    ]
    return [
        canonical_number(sum(float(value) for value in group if value is not None))
        if any(value is not None for value in group)
        else None
        for group in groups
    ]


def dfm_output_variants(payload: Mapping[str, Any]) -> dict[int, list[float | int | None]]:
    """Return primary and supported 3/6/12-period ultimate vector variants."""

    details = _tab(payload, "details_tab")
    data = _tab(payload, "data_tab")
    results = _tab(payload, "results_tab")
    base_length = _integer(details.get("origin_length"), 12, minimum=1)
    values = _numbers(results.get("ultimate_vector"))
    variants = {base_length: values}
    for target_length in (3, 6, 12):
        if target_length <= base_length or target_length % base_length:
            continue
        aggregate = aggregate_vector_values(
            values,
            data.get("origin_labels") if isinstance(data.get("origin_labels"), list) else [],
            base_length,
            target_length,
        )
        if aggregate:
            variants[target_length] = aggregate
    return variants


def _labels(value: Any) -> list[str]:
    return [str(item if item is not None else "") for item in value] if isinstance(value, list) else []


def _duplicate_labels(labels: Iterable[Any]) -> list[str]:
    seen: set[str] = set()
    duplicates: list[str] = []
    for raw in labels:
        label = str(raw if raw is not None else "")
        key = label
        if key in seen and label not in duplicates:
            duplicates.append(label)
        seen.add(key)
    return duplicates


def _numbers(value: Any) -> list[float | int | None]:
    return [canonical_number(item) for item in value] if isinstance(value, list) else []


def _number_matrix(value: Any) -> list[list[float | int | None]]:
    if not isinstance(value, list):
        return []
    return [_numbers(row) if isinstance(row, list) else [] for row in value]


def _input_numbers(value: Any) -> list[float | int | None]:
    return [canonical_input_number(item) for item in value] if isinstance(value, list) else []


def _input_number_matrix(value: Any) -> list[list[float | int | None]]:
    if not isinstance(value, list):
        return []
    return [_input_numbers(row) if isinstance(row, list) else [] for row in value]


def _bool_matrix(value: Any) -> list[list[bool]]:
    if not isinstance(value, list):
        return []
    return [[bool(item) for item in row] if isinstance(row, list) else [] for row in value]


def _int_matrix(value: Any) -> list[list[int]]:
    if not isinstance(value, list):
        return []
    return [
        [_integer(item, 0, minimum=0, maximum=2) for item in row]
        if isinstance(row, list)
        else []
        for row in value
    ]


def _text_matrix(value: Any) -> list[list[str]]:
    if not isinstance(value, list):
        return []
    return [
        [str(item if item is not None else "").strip() for item in row]
        if isinstance(row, list)
        else []
        for row in value
    ]


def _fit_matrix(matrix: list[list[Any]], rows: int, cols: int, fill: Any) -> list[list[Any]]:
    out: list[list[Any]] = []
    for row_index in range(rows):
        row = list(matrix[row_index]) if row_index < len(matrix) else []
        row = row[:cols]
        row.extend(deepcopy(fill) for _ in range(max(0, cols - len(row))))
        out.append(row)
    return out


def _trim_trailing_nulls(row: list[Any]) -> list[Any]:
    out = list(row)
    while out and out[-1] is None:
        out.pop()
    return out


def _settings(raw: Any, row_count: int) -> dict[str, list[Any]]:
    source = raw if isinstance(raw, dict) else {}
    average_types = _labels(source.get("average_type"))
    bases = _labels(source.get("base"))
    periods = list(source.get("periods")) if isinstance(source.get("periods"), list) else []
    excludes = list(source.get("exclude")) if isinstance(source.get("exclude"), list) else []
    out: dict[str, list[Any]] = {
        "average_type": [],
        "base": [],
        "periods": [],
        "exclude": [],
    }
    for index in range(row_count):
        average_type = average_types[index].lower() if index < len(average_types) else "custom"
        out["average_type"].append("user_entry" if average_type == "user_entry" else "custom")
        base = bases[index].lower() if index < len(bases) else "simple"
        out["base"].append(base if base in {"simple", "volume", "benchmark"} else "simple")
        period = periods[index] if index < len(periods) else "all"
        if not (isinstance(period, str) and period.strip().lower() == "all"):
            period = _integer(period, 0, minimum=0) or "all"
        out["periods"].append(period)
        out["exclude"].append(_integer(excludes[index] if index < len(excludes) else 0, 0, minimum=0))
    return out


def default_average_formulas() -> dict[str, Any]:
    labels = ["Volume - all", "Simple - all", "User Entry"]
    return {
        "label": labels,
        "custom_average_formula_settings": {
            "average_type": ["custom", "custom", "user_entry"],
            "base": ["volume", "simple", "simple"],
            "periods": ["all", "all", "all"],
            "exclude": [0, 0, 0],
        },
        "selected": [[], [], []],
        "values": [[], [], []],
        "inputs": [[], [], []],
        "display_inputs": [[], [], []],
    }


def _notes(raw: Any) -> dict[str, Any]:
    source = raw if isinstance(raw, dict) else {}
    return {
        "ratio_main_table": deepcopy(source.get("ratio_main_table"))
        if isinstance(source.get("ratio_main_table"), dict)
        else {},
        "ratio_summary_table": deepcopy(source.get("ratio_summary_table"))
        if isinstance(source.get("ratio_summary_table"), dict)
        else {},
    }


def _tab(payload: Mapping[str, Any], name: str) -> dict[str, Any]:
    value = payload.get(name)
    return value if isinstance(value, dict) else {}


def source_snapshot_revision(snapshot: Mapping[str, Any]) -> str:
    """Fingerprint canonical source content, ignoring producer-local timestamps/revisions.

    The projection below is the hashed vocabulary. It is spelled independently
    of the snapshot's own keys, so a snapshot read from a spaced-key file and
    one read from a snake_case file fingerprint identically.
    """

    origins = _labels(snapshot.get("origin_labels"))
    developments = _labels(snapshot.get("development_labels"))
    raw_values = snapshot.get("values")
    is_matrix = isinstance(raw_values, list) and any(isinstance(row, list) for row in raw_values)
    values: Any = _number_matrix(raw_values) if is_matrix else _numbers(raw_values)
    raw_mask = snapshot.get("mask")
    mask: Any = _bool_matrix(raw_mask) if isinstance(raw_mask, list) else []
    projection = {
        "name": _clean(snapshot.get("name")),
        "origin_labels": origins,
        "development_labels": developments,
        "values": values,
        "mask": mask,
        "data_format": _clean(snapshot.get("data_format")),
        "number_format": _clean(snapshot.get("number_format")),
        "decimal_places": _integer(
            snapshot.get("decimal_places"),
            0,
            minimum=0,
            maximum=8,
        ),
    }
    return fingerprint(projection)


def _is_direct_literal(value: Any) -> bool:
    text = str(value if value is not None else "").strip()
    if not text:
        return True
    if text.startswith("="):
        text = text[1:].strip()
    try:
        number = float(text)
    except (TypeError, ValueError):
        return False
    return math.isfinite(number)


def normalize_dfm_method(
    payload: Mapping[str, Any],
    *,
    require_complete: bool = True,
    timestamp: Any = None,
) -> dict[str, Any]:
    """Return the exact canonical v2 payload for a grouped DFM method."""

    if not isinstance(payload, Mapping):
        raise DfmContractError("DFM method payload must be a JSON object.")
    json_format = str(payload.get("json_format") or "").strip()
    if json_format not in {"", DFM_JSON_FORMAT}:
        raise DfmContractError(f"Unsupported DFM JSON format: {json_format!r}.")

    details_source = _tab(payload, "details_tab")
    data_source = _tab(payload, "data_tab")
    ratios_source = _tab(payload, "ratios_tab")
    ratio_source = _tab(ratios_source, "ratio_triangle")
    formulas_source = _tab(ratios_source, "average_formulas")
    results_source = _tab(payload, "results_tab")
    metadata_source = _tab(payload, "method_metadata")
    provided_revisions = {
        key: str(metadata_source.get(key) or "").strip()
        for key in ("owned_revision", "derived_revision", "publication_revision")
    }

    name = _clean(details_source.get("name"))
    output_type = _clean(details_source.get("output_type"))
    output_dataset = _clean(details_source.get("output_dataset")) or name
    input_name = _clean(details_source.get("input_triangle"))
    origin_labels = _labels(data_source.get("origin_labels"))
    development_labels = _labels(data_source.get("development_labels"))
    row_count = len(origin_labels)
    dev_count = len(development_labels)
    input_values = _fit_matrix(
        _input_number_matrix(data_source.get("input_data_triangle_values")), row_count, dev_count, None
    )
    raw_mask = _bool_matrix(data_source.get("input_data_triangle_mask"))
    if not raw_mask:
        raw_mask = [[value is not None for value in row] for row in input_values]
    input_mask = _fit_matrix(raw_mask, row_count, dev_count, False)
    for row in range(row_count):
        for col in range(dev_count):
            if not input_mask[row][col]:
                input_values[row][col] = None
            elif input_values[row][col] is None:
                input_mask[row][col] = False

    ratio_origin_labels = _labels(ratio_source.get("origin_labels")) or list(origin_labels)
    ratio_dev_labels = _labels(ratio_source.get("development_labels"))
    ratio_values = [_trim_trailing_nulls(row) for row in _number_matrix(ratio_source.get("ratio_values"))]
    excluded = _int_matrix(ratio_source.get("excluded"))

    formula_labels = _labels(formulas_source.get("label"))
    if not formula_labels:
        defaults = default_average_formulas()
        formula_labels = defaults["label"]
        formulas_source = defaults
    formula_count = len(formula_labels)
    formula_cols = len(ratio_dev_labels) or dev_count or max(
        (
            len(row)
            for key in ("selected", "values", "inputs", "display_inputs")
            for row in (formulas_source.get(key) if isinstance(formulas_source.get(key), list) else [])
            if isinstance(row, list)
        ),
        default=0,
    )
    selected = _fit_matrix(_int_matrix(formulas_source.get("selected")), formula_count, formula_cols, 0)
    formula_values = _fit_matrix(_number_matrix(formulas_source.get("values")), formula_count, formula_cols, None)
    formula_inputs = _fit_matrix(_text_matrix(formulas_source.get("inputs")), formula_count, formula_cols, "")
    formula_display_inputs = _fit_matrix(
        _text_matrix(formulas_source.get("display_inputs")), formula_count, formula_cols, ""
    )

    basis_name = _clean(results_source.get("ratio_basis_dataset"))
    # The basis labels are a forced copy of the input origin labels (the
    # validation below rejects anything else), so they are not persisted; a
    # stored file supplies none and the copy is re-derived here.
    basis_origin_labels = _labels(results_source.get("ratio_basis_origin_labels")) or list(origin_labels)
    basis_values = _numbers(results_source.get("ratio_basis_values"))
    if not basis_name:
        basis_origin_labels = []
        basis_values = []

    input_data_format = _clean(data_source.get("data_format")) or "Triangle"
    input_number_format = _clean(data_source.get("number_format")) or "#,##0"
    input_decimal_places = _integer(data_source.get("decimal_places"), 0, minimum=0, maximum=8)
    input_source_revision = source_snapshot_revision({
        "name": input_name,
        "origin_labels": origin_labels,
        "development_labels": development_labels,
        "values": input_values,
        "mask": input_mask,
        "data_format": input_data_format,
        "number_format": input_number_format,
        "decimal_places": input_decimal_places,
    }) if origin_labels and development_labels else ""
    basis_data_format = _clean(results_source.get("ratio_basis_data_format")) or "Vector"
    basis_number_format = _clean(results_source.get("ratio_basis_number_format")) or "#,##0"
    basis_decimal_places = _integer(
        results_source.get("ratio_basis_decimal_places"), 0, minimum=0, maximum=8
    )
    basis_source_revision = source_snapshot_revision({
        "name": basis_name,
        "origin_labels": basis_origin_labels,
        "values": basis_values,
        "data_format": basis_data_format,
        "number_format": basis_number_format,
        "decimal_places": basis_decimal_places,
    }) if basis_name and basis_origin_labels else ""

    default_time = _timestamp(timestamp)
    normalized = {
        "json_format": DFM_JSON_FORMAT,
        "details_tab": {
            "name": name,
            "output_type": output_type,
            "output_dataset": output_dataset,
            "output_category": _clean(details_source.get("output_category")),
            "input_triangle": input_name,
            "origin_length": _integer(details_source.get("origin_length"), 12, minimum=1),
            "development_length": _integer(details_source.get("development_length"), 12, minimum=1),
            "decimal_places": _integer(
                details_source.get("decimal_places"), DFM_DETAILS_DECIMAL_PLACES, minimum=0, maximum=8
            ),
        },
        "data_tab": {
            "origin_labels": origin_labels,
            "development_labels": development_labels,
            "input_data_triangle_values": input_values,
            "input_data_triangle_mask": input_mask,
            "data_format": input_data_format,
            "number_format": input_number_format,
            "decimal_places": input_decimal_places,
            "source_revision": input_source_revision,
        },
        "ratios_tab": {
            "ratio_triangle": {
                "origin_labels": ratio_origin_labels,
                "development_labels": ratio_dev_labels,
                "ratio_values": ratio_values,
                "excluded": excluded,
            },
            "average_formulas": {
                "label": formula_labels,
                "custom_average_formula_settings": _settings(
                    formulas_source.get("custom_average_formula_settings"), formula_count
                ),
                "selected": selected,
                "values": formula_values,
                "inputs": formula_inputs,
                "display_inputs": formula_display_inputs,
            },
            "cell_notes": _notes(ratios_source.get("cell_notes")),
        },
        "results_tab": {
            "ratio_basis_dataset": basis_name,
            "ratio_basis_origin_labels": basis_origin_labels,
            "ratio_basis_values": basis_values,
            "ratio_basis_data_format": basis_data_format,
            "ratio_basis_number_format": basis_number_format,
            "ratio_basis_decimal_places": basis_decimal_places,
            "ratio_basis_source_revision": basis_source_revision,
            "ultimate_ratio_decimal_places": _integer(
                results_source.get("ultimate_ratio_decimal_places"), 2, minimum=0, maximum=8
            ),
            "ultimate_vector": _numbers(results_source.get("ultimate_vector")),
        },
        "method_metadata": {
            "last_modified": str(metadata_source.get("last_modified") or "").strip() or default_time,
            "data_refreshed": str(metadata_source.get("data_refreshed") or "").strip() or default_time,
            "owned_revision": "",
            "derived_revision": "",
            "publication_revision": "",
        },
    }
    stored_chain = _stored_selected_ratios(normalized)
    normalized["curves_tab"] = normalize_curves_tab(
        _tab(payload, "curves_tab"), max(0, len(stored_chain) - 1), stored_chain[:-1]
    )
    _set_revisions(normalized)
    if require_complete:
        _validate_complete(normalized)
        for key, expected in method_revisions(normalized).items():
            if not provided_revisions[key]:
                raise DfmContractError(f"DFM method metadata.{key} is required.")
            if provided_revisions[key] != expected:
                raise DfmContractError(f"DFM method metadata.{key} does not match the canonical payload.")
    return normalized


def _validate_complete(payload: Mapping[str, Any]) -> None:
    details = _tab(payload, "details_tab")
    data = _tab(payload, "data_tab")
    results = _tab(payload, "results_tab")
    for key in ("name", "output_type", "output_dataset", "input_triangle"):
        if not _clean(details.get(key)):
            raise DfmContractError(f"DFM details tab.{key} is required.")
    origins = data.get("origin_labels") if isinstance(data.get("origin_labels"), list) else []
    devs = data.get("development_labels") if isinstance(data.get("development_labels"), list) else []
    if not origins or not devs:
        raise DfmContractError("DFM input snapshot must contain origin and development labels.")
    duplicates = _duplicate_labels(origins)
    if duplicates:
        raise DfmContractError("DFM input origin labels must be unique: " + ", ".join(duplicates))
    values = data.get("input_data_triangle_values")
    mask = data.get("input_data_triangle_mask")
    if not isinstance(values, list) or len(values) != len(origins):
        raise DfmContractError("DFM input values must contain one row per origin label.")
    if not isinstance(mask, list) or len(mask) != len(origins):
        raise DfmContractError("DFM input mask must contain one row per origin label.")
    if any(len(row) != len(devs) for row in values) or any(len(row) != len(devs) for row in mask):
        raise DfmContractError("DFM input values and mask must match the development-label geometry.")
    if not str(data.get("source_revision") or "").strip():
        raise DfmContractError("DFM input snapshot must contain a source revision.")
    ratios = _tab(payload, "ratios_tab")
    ratio = _tab(ratios, "ratio_triangle")
    expected_ratio_labels = _ratio_development_labels(list(devs))
    if ratio.get("origin_labels") != origins:
        raise DfmContractError("DFM ratio origin labels must equal the input origin labels.")
    if ratio.get("development_labels") != expected_ratio_labels:
        raise DfmContractError("DFM ratio development labels do not match the input geometry.")
    ratio_values = ratio.get("ratio_values")
    excluded = ratio.get("excluded")
    if not isinstance(ratio_values, list) or len(ratio_values) != len(origins):
        raise DfmContractError("DFM ratio values must contain one row per origin label.")
    if not isinstance(excluded, list) or len(excluded) != len(origins):
        raise DfmContractError("DFM exclusions must contain one row per origin label.")
    if any(not isinstance(row, list) or len(row) > max(0, len(devs) - 1) for row in ratio_values):
        raise DfmContractError("DFM ratio rows exceed the input development geometry.")
    if any(len(excluded[index]) != len(ratio_values[index]) for index in range(len(origins))):
        raise DfmContractError("DFM exclusion rows must match the corresponding ratio-value rows.")
    formulas = _tab(ratios, "average_formulas")
    formula_labels = formulas.get("label") if isinstance(formulas.get("label"), list) else []
    formula_cols = len(expected_ratio_labels)
    for key in ("selected", "values", "inputs", "display_inputs"):
        matrix = formulas.get(key)
        if not isinstance(matrix, list) or len(matrix) != len(formula_labels):
            raise DfmContractError(f"DFM average formulas.{key} must align to formula labels.")
        if any(not isinstance(row, list) or len(row) != formula_cols for row in matrix):
            raise DfmContractError(f"DFM average formulas.{key} must align to ratio columns.")
    curves = _tab(payload, "curves_tab")
    period_count = max(0, formula_cols - 1)
    for key in ("included", "selected_estimates"):
        if len(curves.get(key) or []) != period_count:
            raise DfmContractError(f"DFM curves tab.{key} must hold one entry per development period.")
    if any(len(column.get("values") or []) != period_count for column in curves.get("user_columns") or []):
        raise DfmContractError("DFM curves tab user columns must hold one value per development period.")
    if len(curves.get("selected_values") or []) not in {0, period_count + 1}:
        raise DfmContractError("DFM curves tab.selected_values must cover every development period and the tail.")
    ultimate = results.get("ultimate_vector")
    if not isinstance(ultimate, list) or len(ultimate) != len(origins):
        raise DfmContractError("DFM ultimate vector must contain one value per origin label.")
    if _clean(results.get("ratio_basis_dataset")):
        if results.get("ratio_basis_origin_labels") != origins:
            raise DfmContractError("DFM Ratio Basis labels must align exactly to the DFM origins.")
        if len(results.get("ratio_basis_values") or []) != len(origins):
            raise DfmContractError("DFM Ratio Basis values must align exactly to the DFM origins.")
        if not str(results.get("ratio_basis_source_revision") or "").strip():
            raise DfmContractError("DFM Ratio Basis snapshot must contain a source revision.")


def _split_coordinates_top_level(coordinates: str) -> list[str]:
    """Split ``row, col`` coordinates on top-level commas, honoring quotes."""

    parts: list[str] = []
    current = ""
    quote = ""
    for character in coordinates:
        if quote:
            current += character
            if character == quote:
                quote = ""
            continue
        if character in {'"', "'"}:
            quote = character
            current += character
            continue
        if character == ",":
            parts.append(current.strip())
            current = ""
            continue
        current += character
    parts.append(current.strip())
    return parts


def dataset_reference_tokens(value: Any) -> list[dict[str, Any]]:
    """Return valid ``[Dataset][coordinate]`` tokens with spans in formula order.

    Each token carries the exact reference text (``match``), its ``start``/``end``
    span inside the formula, the ``dataset_name``, and the raw ``row_idx`` /
    ``col_idx`` coordinate text exactly as written (``col_idx`` is ``None`` for a
    Vector reference). The coordinate text keeps quotes so it matches what the
    frontend sends to the dataset-reference resolver.
    """

    text = str(value if value is not None else "")
    tokens: list[dict[str, Any]] = []
    cursor = 0
    while cursor < len(text):
        dataset_start = text.find("[", cursor)
        if dataset_start < 0:
            break
        dataset_end = text.find("]", dataset_start + 1)
        if dataset_end < 0:
            break
        coordinate_start = dataset_end + 1
        while coordinate_start < len(text) and text[coordinate_start].isspace():
            coordinate_start += 1
        if coordinate_start >= len(text) or text[coordinate_start] != "[":
            cursor = dataset_start + 1
            continue

        quote = ""
        coordinate_end = -1
        comma_count = 0
        for index in range(coordinate_start + 1, len(text)):
            character = text[index]
            if quote:
                if character == quote:
                    quote = ""
                continue
            if character in {'"', "'"}:
                quote = character
            elif character == ",":
                comma_count += 1
            elif character == "]":
                coordinate_end = index
                break
        if coordinate_end < 0:
            break

        name = _clean(text[dataset_start + 1:dataset_end])
        coordinates = text[coordinate_start + 1:coordinate_end].strip()
        coordinate_parts = _split_coordinates_top_level(coordinates)
        if (
            name
            and coordinates
            and comma_count <= 1
            and len(coordinate_parts) in (1, 2)
            and all(coordinate_parts)
        ):
            tokens.append({
                "match": text[dataset_start:coordinate_end + 1],
                "start": dataset_start,
                "end": coordinate_end + 1,
                "dataset_name": name,
                "row_idx": coordinate_parts[0],
                "col_idx": coordinate_parts[1] if len(coordinate_parts) == 2 else None,
            })
        cursor = coordinate_end + 1
    return tokens


def _dataset_reference_names(value: Any) -> list[str]:
    """Return valid ``[Dataset][coordinate]`` identities in formula order."""

    return [token["dataset_name"] for token in dataset_reference_tokens(value)]


def _contains_dataset_reference(value: Any) -> bool:
    return bool(_dataset_reference_names(value))


def dfm_dataset_reference_tokens(payload: Mapping[str, Any]) -> list[dict[str, Any]]:
    """Return unique dataset-reference tokens across all average-formula inputs."""

    formulas = _tab(_tab(payload, "ratios_tab"), "average_formulas")
    inputs = formulas.get("inputs") if isinstance(formulas.get("inputs"), list) else []
    tokens: list[dict[str, Any]] = []
    seen: set[str] = set()
    for row in inputs:
        if not isinstance(row, list):
            continue
        for formula in row:
            for token in dataset_reference_tokens(formula):
                if token["match"] in seen:
                    continue
                seen.add(token["match"])
                tokens.append(token)
    return tokens


def _substitute_dataset_references(
    formula: Any,
    tokens: list[dict[str, Any]],
    reference_values: Mapping[str, Any] | None,
) -> str | None:
    """Replace dataset references with resolved numeric values.

    Returns ``None`` when any reference has no finite resolved value, so the
    caller keeps the stored evaluation instead of recomputing from a partial
    substitution.
    """

    if not isinstance(reference_values, Mapping) or not reference_values:
        return None
    text = str(formula if formula is not None else "")
    for token in reversed(tokens):
        try:
            value = float(reference_values.get(token["match"]))
        except (TypeError, ValueError):
            return None
        if not math.isfinite(value):
            return None
        text = f"{text[:token['start']]}{value}{text[token['end']:]}"
    return text


def _owned_formula_values(payload: Mapping[str, Any]) -> list[dict[str, Any]]:
    formulas = _tab(_tab(payload, "ratios_tab"), "average_formulas")
    labels = formulas.get("label") if isinstance(formulas.get("label"), list) else []
    settings = _tab(formulas, "custom_average_formula_settings")
    types = settings.get("average_type") if isinstance(settings.get("average_type"), list) else []
    bases = settings.get("base") if isinstance(settings.get("base"), list) else []
    values = formulas.get("values") if isinstance(formulas.get("values"), list) else []
    inputs = formulas.get("inputs") if isinstance(formulas.get("inputs"), list) else []
    out: list[dict[str, Any]] = []
    for index, label in enumerate(labels):
        average_type = str(types[index] if index < len(types) else "").strip().lower()
        base = str(bases[index] if index < len(bases) else "").strip().lower()
        row_inputs = inputs[index] if index < len(inputs) and isinstance(inputs[index], list) else []
        row_values = values[index] if index < len(values) and isinstance(values[index], list) else []
        owned_values = [
            deepcopy(row_values[col]) if col < len(row_values) else None
            for col, formula in enumerate(row_inputs)
            if base == "benchmark" or (
                average_type == "user_entry" and (
                    _is_direct_literal(formula)
                    or _contains_excel_reference(formula)
                    or _contains_dataset_reference(formula)
                )
            )
        ]
        owned_columns = [
            col
            for col, formula in enumerate(row_inputs)
            if base == "benchmark" or (
                average_type == "user_entry" and (
                    _is_direct_literal(formula)
                    or _contains_excel_reference(formula)
                    or _contains_dataset_reference(formula)
                )
            )
        ]
        if not owned_columns:
            continue
        out.append({
            "label": label,
            "columns": owned_columns,
            "values": owned_values,
        })
    return out


# The three revision projections below define the hashed vocabulary of a DFM
# method. Every key they emit is fixed here and spelled independently of the
# persisted field it is read from, so renaming an on-disk field never moves a
# stored fingerprint. Keys that carry user data (ratio-cell note coordinates)
# are copied as they are; they are content, not vocabulary.


def owned_projection(payload: Mapping[str, Any]) -> dict[str, Any]:
    """The person's choices: details, exclusions, formula selections, cell notes."""

    details = _tab(payload, "details_tab")
    ratios = _tab(payload, "ratios_tab")
    ratio = _tab(ratios, "ratio_triangle")
    formulas = _tab(ratios, "average_formulas")
    settings = _tab(formulas, "custom_average_formula_settings")
    notes = _tab(ratios, "cell_notes")
    results = _tab(payload, "results_tab")
    ratio_origins = ratio.get("origin_labels") if isinstance(ratio.get("origin_labels"), list) else []
    ratio_devs = ratio.get("development_labels") if isinstance(ratio.get("development_labels"), list) else []
    excluded = ratio.get("excluded") if isinstance(ratio.get("excluded"), list) else []
    excluded_cells = [
        {
            "origin_label": ratio_origins[row] if row < len(ratio_origins) else str(row),
            "development_label": ratio_devs[col] if col < len(ratio_devs) else str(col),
        }
        for row, values in enumerate(excluded)
        if isinstance(values, list)
        for col, value in enumerate(values)
        if value == 1
    ]
    projection = {
        "details": {
            "name": details.get("name"),
            "output_type": details.get("output_type"),
            "output_dataset": details.get("output_dataset"),
            "output_category": details.get("output_category"),
            "input_triangle": details.get("input_triangle"),
            "origin_length": details.get("origin_length"),
            "development_length": details.get("development_length"),
            "decimal_places": details.get("decimal_places"),
        },
        "excluded_cells": excluded_cells,
        "average_formulas": {
            "label": deepcopy(formulas.get("label") or []),
            "settings": {
                "average_type": deepcopy(settings.get("average_type") or []),
                "base": deepcopy(settings.get("base") or []),
                "periods": deepcopy(settings.get("periods") or []),
                "exclude": deepcopy(settings.get("exclude") or []),
            },
            "selected": deepcopy(formulas.get("selected") or []),
            "inputs": deepcopy(formulas.get("inputs") or []),
            "owned_values": _owned_formula_values(payload),
        },
        "cell_notes": {
            "ratio_main_table": deepcopy(notes.get("ratio_main_table") or {}),
            "ratio_summary_table": deepcopy(notes.get("ratio_summary_table") or {}),
        },
        "ratio_basis_dataset": results.get("ratio_basis_dataset", ""),
        "ultimate_ratio_decimal_places": results.get("ultimate_ratio_decimal_places", 2),
    }
    if not _curves_tab_is_default(payload):
        projection["curves"] = owned_curves_tab(_tab(payload, "curves_tab"))
    return projection


def _curves_tab_is_default(payload: Mapping[str, Any]) -> bool:
    """Whether the Curves tab can be left out of the revision fingerprints.

    Every method file written before the Curves tab existed reads as the
    default tab, and its factors are the same with or without it, so the
    vocabulary of the fingerprints only grows once a person changes something
    on that tab. A stored revision therefore keeps matching its file.
    """

    formulas = _tab(_tab(payload, "ratios_tab"), "average_formulas")
    if not isinstance(formulas.get("values"), list):
        return True
    return curves_tab_is_default(_tab(payload, "curves_tab"), _stored_selected_ratios(payload)[:-1])


def derived_projection(payload: Mapping[str, Any]) -> dict[str, Any]:
    """The computed numbers and the snapshots they were computed from."""

    data = _tab(payload, "data_tab")
    ratios = _tab(payload, "ratios_tab")
    ratio = _tab(ratios, "ratio_triangle")
    formulas = _tab(ratios, "average_formulas")
    results = _tab(payload, "results_tab")
    projection = {
        "input": {
            "origin_labels": deepcopy(data.get("origin_labels")),
            "development_labels": deepcopy(data.get("development_labels")),
            "values": deepcopy(data.get("input_data_triangle_values")),
            "mask": deepcopy(data.get("input_data_triangle_mask")),
            "data_format": data.get("data_format"),
            "number_format": data.get("number_format"),
            "decimal_places": data.get("decimal_places"),
            "source_revision": data.get("source_revision"),
        },
        "ratio_triangle": {
            "origin_labels": deepcopy(ratio.get("origin_labels") or []),
            "development_labels": deepcopy(ratio.get("development_labels") or []),
            "values": deepcopy(ratio.get("ratio_values") or []),
        },
        "average_formula_values": deepcopy(formulas.get("values") or []),
        "ratio_basis": {
            "origin_labels": deepcopy(results.get("ratio_basis_origin_labels")),
            "values": deepcopy(results.get("ratio_basis_values")),
            "data_format": results.get("ratio_basis_data_format"),
            "number_format": results.get("ratio_basis_number_format"),
            "decimal_places": results.get("ratio_basis_decimal_places"),
            "source_revision": results.get("ratio_basis_source_revision"),
        },
        "ultimate_vector": deepcopy(results.get("ultimate_vector")),
    }
    if not _curves_tab_is_default(payload):
        projection["curves_selected_values"] = deepcopy(_tab(payload, "curves_tab").get("selected_values") or [])
    return projection


def publication_projection(payload: Mapping[str, Any]) -> dict[str, Any]:
    """Only what a downstream method can see of this one."""

    details = _tab(payload, "details_tab")
    data = _tab(payload, "data_tab")
    results = _tab(payload, "results_tab")
    return {
        "output_dataset": details.get("output_dataset", ""),
        "output_type": details.get("output_type", ""),
        "output_category": details.get("output_category", ""),
        "origin_length": details.get("origin_length", 12),
        "decimal_places": details.get("decimal_places", 0),
        "origin_labels": deepcopy(data.get("origin_labels") or []),
        "ultimate_vector": deepcopy(results.get("ultimate_vector") or []),
    }


def persisted_projection(payload: Mapping[str, Any]) -> dict[str, Any]:
    """Return the on-disk form of a canonical DFM method.

    The input triangle is stored in a reduced form because every part of it that
    is dropped is exactly recoverable when the file is read back:

    * ``input data triangle mask`` is omitted. A cell is inside the triangle if
      and only if it holds a value -- an invariant ``normalize_dfm_method``
      enforces -- so the mask can only ever restate the values beside it.
    * Trailing nulls are trimmed from each input row, matching how
      ``ratio values`` and ``excluded`` are already stored. A null *inside* a row
      still marks a missing value inside the triangle.

    ``normalize_dfm_method`` derives the mask and refits every row, so this
    projection loses nothing; it exists to keep the persisted file readable.
    Applying it to an already-persisted payload is a no-op.

    This is a serialization projection only. The canonical in-memory payload
    keeps its rectangular geometry and its mask, so revisions, validation, and
    every calculation are unaffected.
    """

    persisted = deepcopy(dict(payload))
    results = persisted.get("results_tab")
    if isinstance(results, dict):
        # A forced copy of data_tab.origin_labels; re-derived on read.
        results.pop("ratio_basis_origin_labels", None)
    data = persisted.get("data_tab")
    if not isinstance(data, dict):
        return persisted
    data.pop("input_data_triangle_mask", None)
    values = data.get("input_data_triangle_values")
    if isinstance(values, list):
        data["input_data_triangle_values"] = [
            _trim_trailing_nulls(row) if isinstance(row, list) else row
            for row in values
        ]
    return persisted


def method_revisions(payload: Mapping[str, Any]) -> dict[str, str]:
    return {
        "owned_revision": fingerprint(owned_projection(payload)),
        "derived_revision": fingerprint(derived_projection(payload)),
        "publication_revision": fingerprint(publication_projection(payload)),
    }


def dfm_precedent_names(payload: Mapping[str, Any]) -> list[str]:
    details = _tab(payload, "details_tab")
    results = _tab(payload, "results_tab")
    formulas = _tab(_tab(payload, "ratios_tab"), "average_formulas")
    inputs = formulas.get("inputs") if isinstance(formulas.get("inputs"), list) else []
    formula_datasets = [
        dataset_name
        for row in inputs
        if isinstance(row, list)
        for formula in row
        for dataset_name in _dataset_reference_names(formula)
    ]
    return dependency_names([
        details.get("input_triangle"),
        results.get("ratio_basis_dataset"),
        *formula_datasets,
    ])


def build_dfm_output_sidecar(
    payload: Mapping[str, Any],
    *,
    project_name: Any,
    reserving_class: Any,
    csv_file: Any,
    existing: Mapping[str, Any] | None = None,
    existing_record: bool | None = None,
    dependents: Any = None,
    notes: Any = None,
    timestamp: Any = None,
    user: Any = "",
    output_changed: bool = True,
    append_audit: bool = True,
    audit_action: Any = None,
    status: Any = 0,
) -> dict[str, Any]:
    """Build the sole canonical parsed payload for a DFM output sidecar."""

    method = normalize_dfm_method(payload, require_complete=True, timestamp=timestamp)
    prior = existing if isinstance(existing, Mapping) else {}
    record_exists = bool(prior) if existing_record is None else bool(existing_record)
    details = _tab(method, "details_tab")
    data = _tab(method, "data_tab")
    metadata = _tab(method, "method_metadata")
    method_name = _clean(details.get("name"))
    output_dataset = _clean(details.get("output_dataset")) or method_name
    published_at = _timestamp(timestamp)
    actor = _clean(user)
    if not output_changed and record_exists:
        published_at = str(prior.get("updated_at") or "").strip() or published_at
        actor = _clean(prior.get("modified_by")) or actor
    created = str(prior.get("created") or "").strip() or published_at
    sidecar_notes = str(prior.get("notes") or "") if notes is None else str(notes)
    if append_audit:
        audits = append_audit_entry(
            prior.get("audit_log"),
            event_date=published_at,
            action=_clean(audit_action) or (AUDIT_ACTION_UPDATE if record_exists else AUDIT_ACTION_INSERT),
            user=actor,
        )
    else:
        audits = normalize_audit_log(prior.get("audit_log"))
    output_period = _integer(details.get("origin_length"), 12, minimum=1)
    return validate_sidecar_core({
        "json_format": DATASET_SIDECAR_JSON_FORMAT,
        "dataset_name": output_dataset,
        "dataset_type": _clean(details.get("output_type")) or output_dataset,
        "dataset_category": _clean(details.get("output_category")),
        "reserving_class": _clean(reserving_class),
        "project_name": _clean(project_name),
        "source_kind": "dfm",
        "calculated": True,
        "method_name": method_name,
        "method_type": "DFM",
        "data_format": "Vector",
        "period_length": output_period,
        # A method output is produced at its own origin period, so the
        # vector it publishes is stored at that period too.
        **stored_length_fields("Vector", output_period),
        "transposed": False,
        "show_subtotal": normalize_show_subtotal(prior.get("show_subtotal")),
        "number_format": _clean(prior.get("number_format")) or "#,##0",
        "decimal_places": _integer(details.get("decimal_places"), 0, minimum=0, maximum=8),
        "csv_file": _clean(csv_file),
        "notes": sidecar_notes,
        "origin_labels": deepcopy(data.get("origin_labels") or []),
        "development_labels": ["Ultimate"],
        "precedents": dependency_entries(dfm_precedent_names(method)),
        "dependents": dependency_entries(prior.get("dependents") if dependents is None else dependents),
        "created": created,
        "updated_at": published_at,
        "modified_by": actor,
        "status": _integer(status, 0, minimum=0),
        "publication_revision": str(metadata.get("publication_revision") or "").strip(),
        "audit_log": audits,
    })


def _set_revisions(payload: dict[str, Any]) -> None:
    metadata = payload.setdefault("method_metadata", {})
    metadata.update(method_revisions(payload))


def _apply_input_snapshot(payload: dict[str, Any], snapshot: Mapping[str, Any]) -> None:
    details = payload["details_tab"]
    old_ratio = payload["ratios_tab"]["ratio_triangle"]
    old_data_devs = _labels(payload["data_tab"].get("development_labels"))
    old_origins = _labels(old_ratio.get("origin_labels"))
    old_devs = _labels(old_ratio.get("development_labels"))
    old_excluded = _int_matrix(old_ratio.get("excluded"))

    name = _clean(snapshot.get("name"))
    if name:
        details["input_triangle"] = name
    origins = _labels(snapshot.get("origin_labels"))
    devs = _labels(snapshot.get("development_labels"))
    duplicates = _duplicate_labels(origins)
    if duplicates:
        raise DfmContractError("DFM input snapshot has duplicate origin labels: " + ", ".join(duplicates))
    if old_data_devs and old_data_devs != devs:
        raise DfmContractError(
            "DFM input development-label geometry changed; preserve the last valid method and require review."
        )
    values = _input_number_matrix(snapshot.get("values"))
    mask = _bool_matrix(snapshot.get("mask"))
    if not mask:
        mask = [[item is not None for item in row] for row in values]
    values = _fit_matrix(values, len(origins), len(devs), None)
    mask = _fit_matrix(mask, len(origins), len(devs), False)
    for row in range(len(origins)):
        for col in range(len(devs)):
            if not mask[row][col] or values[row][col] is None:
                values[row][col] = None
                mask[row][col] = False
    data = payload["data_tab"]
    data_format = _clean(snapshot.get("data_format")) or "Triangle"
    number_format = _clean(snapshot.get("number_format")) or "#,##0"
    decimal_places = _integer(
        snapshot.get("decimal_places"), 0, minimum=0, maximum=8
    )
    canonical_snapshot = {
        "name": details.get("input_triangle"),
        "origin_labels": origins,
        "development_labels": devs,
        "values": values,
        "mask": mask,
        "data_format": data_format,
        "number_format": number_format,
        "decimal_places": decimal_places,
    }
    data.update({
        "origin_labels": origins,
        "development_labels": devs,
        "input_data_triangle_values": values,
        "input_data_triangle_mask": mask,
        "data_format": data_format,
        "number_format": number_format,
        "decimal_places": decimal_places,
        "source_revision": source_snapshot_revision(canonical_snapshot),
    })
    ratio_labels = _ratio_development_labels(devs)
    origin_lookup = {label: index for index, label in enumerate(old_origins)}
    dev_lookup = {label: index for index, label in enumerate(old_devs)}
    remapped: list[list[int]] = []
    for origin in origins:
        old_row = origin_lookup.get(origin)
        row: list[int] = []
        for dev_label in ratio_labels:
            old_col = dev_lookup.get(dev_label)
            value = 0
            if old_row is not None and old_col is not None and old_row < len(old_excluded):
                source_row = old_excluded[old_row]
                if old_col < len(source_row):
                    value = source_row[old_col]
            row.append(value)
        remapped.append(row)
    old_ratio["origin_labels"] = origins
    old_ratio["development_labels"] = ratio_labels
    old_ratio["excluded"] = remapped


def _apply_ratio_basis_snapshot(payload: dict[str, Any], snapshot: Mapping[str, Any]) -> None:
    results = payload["results_tab"]
    name = _clean(snapshot.get("name"))
    if name:
        results["ratio_basis_dataset"] = name
    if not _clean(results.get("ratio_basis_dataset")):
        results.update({
            "ratio_basis_origin_labels": [],
            "ratio_basis_values": [],
            "ratio_basis_source_revision": "",
        })
        return
    method_origins = payload["data_tab"]["origin_labels"]
    source_origins = _labels(snapshot.get("origin_labels"))
    duplicates = _duplicate_labels(source_origins)
    if duplicates:
        raise DfmContractError("DFM Ratio Basis has duplicate origin labels: " + ", ".join(duplicates))
    raw_values = snapshot.get("values")
    if isinstance(raw_values, list) and raw_values and any(isinstance(row, list) for row in raw_values):
        raw_values = [row[0] if isinstance(row, list) and row else None for row in raw_values]
    source_values = _numbers(raw_values)
    lookup = {
        label: source_values[index] if index < len(source_values) else None
        for index, label in enumerate(source_origins)
    }
    missing = [label for label in method_origins if label not in lookup]
    if missing:
        raise DfmContractError(
            "DFM Ratio Basis is missing exact origin labels: " + ", ".join(str(label) for label in missing)
        )
    aligned_values = [lookup.get(label) for label in method_origins]
    data_format = _clean(snapshot.get("data_format")) or "Vector"
    number_format = _clean(snapshot.get("number_format")) or "#,##0"
    decimal_places = _integer(
        snapshot.get("decimal_places"), 0, minimum=0, maximum=8
    )
    canonical_snapshot = {
        "name": results.get("ratio_basis_dataset"),
        "origin_labels": list(method_origins),
        "values": aligned_values,
        "data_format": data_format,
        "number_format": number_format,
        "decimal_places": decimal_places,
    }
    results.update({
        "ratio_basis_origin_labels": list(method_origins),
        "ratio_basis_values": aligned_values,
        "ratio_basis_data_format": data_format,
        "ratio_basis_number_format": number_format,
        "ratio_basis_decimal_places": decimal_places,
        "ratio_basis_source_revision": source_snapshot_revision(canonical_snapshot),
    })


def _ratio_development_labels(development_labels: list[str]) -> list[str]:
    if not development_labels:
        return []
    labels: list[str] = []
    for index in range(max(0, len(development_labels) - 1)):
        left = str(development_labels[index])
        right = str(development_labels[index + 1])
        labels.append(f"({index + 1}) {_age_text(left)}-{_age_text(right)}")
    labels.append(f"{_age_text(development_labels[-1])} - Ult")
    return labels


def _age_text(value: Any) -> str:
    text = str(value if value is not None else "").strip()
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return match.group(0) if match else text


def _calculate_ratio_triangle(values: list[list[Any]], mask: list[list[bool]], dev_count: int) -> list[list[Any]]:
    out: list[list[Any]] = []
    for row_index, row_values in enumerate(values):
        row: list[Any] = []
        row_mask = mask[row_index] if row_index < len(mask) else []
        for col in range(max(0, dev_count - 1)):
            if col + 1 >= len(row_mask) or not row_mask[col] or not row_mask[col + 1]:
                row.append(None)
                continue
            left = canonical_input_number(row_values[col] if col < len(row_values) else None)
            right = canonical_input_number(row_values[col + 1] if col + 1 < len(row_values) else None)
            # A later value of zero is not a ratio of zero: the origin has
            # nothing to develop from, so the cell holds no ratio at all rather
            # than a 0 that would drag the column's averages down with it.
            if left in (None, 0) or right in (None, 0):
                row.append(None)
            else:
                row.append(canonical_number(float(right) / float(left)))
        out.append(_trim_trailing_nulls(row))
    return out


def _selected_rows(
    values: list[list[Any]],
    mask: list[list[bool]],
    excluded: list[list[int]],
    col: int,
    periods: Any,
    extra_exclude: int,
) -> list[int]:
    candidates: list[tuple[int, float]] = []
    for row in range(len(values)):
        if row >= len(mask) or col + 1 >= len(mask[row]) or not mask[row][col] or not mask[row][col + 1]:
            continue
        left = canonical_input_number(values[row][col])
        right = canonical_input_number(values[row][col + 1])
        if left in (None, 0) or right in (None, 0):
            continue
        ratio = float(right) / float(left)
        if not math.isfinite(ratio):
            continue
        if row < len(excluded) and col < len(excluded[row]) and excluded[row][col] == 1:
            continue
        candidates.append((row, ratio))
    lookback = 0 if isinstance(periods, str) and periods.lower() == "all" else _integer(periods, 0, minimum=0)
    if lookback:
        candidates = sorted(candidates, key=lambda item: item[0], reverse=True)[:lookback]
    # ResQ drops pairs of highest and lowest ratios "for as long as the
    # remaining number of ratios is greater than two", so a column down to two
    # ratios keeps both and averages them rather than excluding itself empty.
    # Allowing (n - 1) // 2 pairs is that rule written as a count.
    trim = max(0, min(int(extra_exclude), (len(candidates) - 1) // 2))
    if trim:
        sorted_values = sorted(candidates, key=lambda item: item[1])
        removed = {row for pair in (sorted_values[:trim], sorted_values[-trim:]) for row, _ratio in pair}
        candidates = [item for item in candidates if item[0] not in removed]
    return [row for row, _ratio in candidates]


def _calculate_average(
    values: list[list[Any]],
    mask: list[list[bool]],
    excluded: list[list[int]],
    col: int,
    *,
    base: str,
    periods: Any,
    extra_exclude: int,
) -> float:
    rows = _selected_rows(values, mask, excluded, col, periods, extra_exclude)
    if not rows:
        return 1.0
    if base == "volume":
        denominator = sum(float(values[row][col]) for row in rows)
        numerator = sum(float(values[row][col + 1]) for row in rows)
        return numerator / denominator if denominator else 1.0
    ratios = [float(values[row][col + 1]) / float(values[row][col]) for row in rows]
    return sum(ratios) / len(ratios) if ratios else 1.0


def _contains_excel_reference(value: Any) -> bool:
    text = str(value if value is not None else "").strip()
    return bool(text and _EXCEL_REFERENCE_RE.search(text))


def contains_excel_reference(value: Any) -> bool:
    """Public predicate used by background freshness-check consumers."""

    return _contains_excel_reference(value)


def _safe_arithmetic(expression: str) -> float | None:
    # ast.parse in eval mode rejects leading whitespace (IndentationError), so
    # a formula like "= expr" must not reach it with the space that stripping
    # the "=" leaves behind.
    try:
        tree = ast.parse(str(expression or "").strip(), mode="eval")
    except (SyntaxError, ValueError):
        return None
    binary = {
        ast.Add: lambda left, right: left + right,
        ast.Sub: lambda left, right: left - right,
        ast.Mult: lambda left, right: left * right,
        ast.Div: lambda left, right: left / right,
        ast.Pow: lambda left, right: left**right,
        ast.Mod: lambda left, right: left % right,
    }
    unary = {ast.UAdd: lambda value: value, ast.USub: lambda value: -value}

    def evaluate(node: ast.AST) -> float:
        if isinstance(node, ast.Expression):
            return evaluate(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
            return float(node.value)
        if isinstance(node, ast.BinOp) and type(node.op) in binary:
            return binary[type(node.op)](evaluate(node.left), evaluate(node.right))
        if isinstance(node, ast.UnaryOp) and type(node.op) in unary:
            return unary[type(node.op)](evaluate(node.operand))
        # ROUND(x) or ROUND(x, digits): the one function the formula language
        # offers, so a formula can fix an operand at the precision the notes show.
        if (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id.upper() == "ROUND"
            and not node.keywords
            and 1 <= len(node.args) <= 2
        ):
            digits = evaluate(node.args[1]) if len(node.args) == 2 else 0.0
            return round_half_up(evaluate(node.args[0]), int(digits))
        raise ValueError("unsupported expression")

    try:
        result = evaluate(tree)
    except (ArithmeticError, OverflowError, ValueError):
        return None
    return result if math.isfinite(result) else None


def _evaluate_internal_formula(
    formula: str,
    labels: list[str],
    computed: list[list[Any]],
    col: int,
    decimal_places: Any,
    resolver: Any = None,
) -> float | None:
    text = str(formula or "").strip()
    if not text or _contains_excel_reference(text) or _contains_dataset_reference(text):
        return None
    if text.startswith("="):
        text = text[1:]
    lookup = {_clean(label).casefold(): index for index, label in enumerate(labels)}

    def replace(match: re.Match[str]) -> str:
        row = lookup.get(_clean(match.group(1)).casefold())
        if row is None or row >= len(computed) or col >= len(computed[row]):
            return "nan"
        value = average_row_reference_value(
            resolver(row, col) if callable(resolver) else computed[row][col], decimal_places
        )
        return str(value) if value is not None else "nan"

    text = _INTERNAL_LABEL_RE.sub(replace, text)
    return _safe_arithmetic(text)


def _calculate_formula_values(
    payload: dict[str, Any],
    dataset_reference_values: Mapping[str, Any] | None = None,
) -> list[list[Any]]:
    data = payload["data_tab"]
    ratio = payload["ratios_tab"]["ratio_triangle"]
    formulas = payload["ratios_tab"]["average_formulas"]
    labels = formulas["label"]
    settings = formulas["custom_average_formula_settings"]
    values = data["input_data_triangle_values"]
    mask = data["input_data_triangle_mask"]
    excluded = ratio["excluded"]
    decimal_places = payload["details_tab"]["decimal_places"]
    old_values = _fit_matrix(_number_matrix(formulas.get("values")), len(labels), len(ratio["development_labels"]), None)
    inputs = _fit_matrix(_text_matrix(formulas.get("inputs")), len(labels), len(ratio["development_labels"]), "")
    col_count = len(ratio["development_labels"])
    tail_col = len(data["development_labels"]) - 1
    computed: list[list[Any]] = [[None] * col_count for _ in labels]

    def stored_tail(row: int) -> float:
        # The "- Ult" column is the row's own tail factor, entered rather than
        # averaged: ResQ keeps it as each average row's TailFactor. A computed
        # average row has none and stays at 1.0.
        value = canonical_number(old_values[row][tail_col]) if 0 <= tail_col < len(old_values[row]) else None
        return float(value) if value is not None and value > 0 else 1.0

    for row, _label in enumerate(labels):
        average_type = settings["average_type"][row]
        if average_type == "user_entry":
            continue
        if settings["base"][row] == "benchmark":
            computed[row] = [
                canonical_number(value) if canonical_number(value) is not None else 1.0
                for value in old_values[row]
            ]
            continue
        for col in range(col_count):
            computed[row][col] = 1.0 if col >= tail_col else canonical_number(
                _calculate_average(
                    values,
                    mask,
                    excluded,
                    col,
                    base=settings["base"][row],
                    periods=settings["periods"][row],
                    extra_exclude=settings["exclude"][row],
                )
            )
    resolving: set[tuple[int, int]] = set()

    def resolve(row: int, col: int) -> Any:
        existing = canonical_number(computed[row][col])
        if existing is not None:
            return existing
        stored = canonical_number(old_values[row][col])
        key = (row, col)
        if key in resolving:
            return stored if stored is not None and stored > 0 else 1.0
        if col >= tail_col:
            computed[row][col] = stored_tail(row)
            return computed[row][col]
        resolving.add(key)
        try:
            formula = inputs[row][col]
            reference_tokens = dataset_reference_tokens(formula)
            # The Excel-reference pattern treats any bracket segment as external,
            # so strip the dataset references before deciding whether an Excel
            # reference remains that only the client can evaluate.
            formula_without_references = str(formula or "")
            for token in reversed(reference_tokens):
                formula_without_references = (
                    formula_without_references[:token["start"]]
                    + formula_without_references[token["end"]:]
                )
            if _contains_excel_reference(formula_without_references):
                chosen = stored
            elif reference_tokens:
                # Re-evaluate a dataset-referencing formula only when the caller
                # supplies a resolved value for every reference; otherwise keep
                # the stored evaluation.
                substituted = _substitute_dataset_references(
                    formula, reference_tokens, dataset_reference_values
                )
                chosen = (
                    _evaluate_internal_formula(
                        substituted,
                        labels,
                        computed,
                        col,
                        decimal_places,
                        resolver=resolve,
                    )
                    if substituted is not None
                    else None
                )
                if chosen is None:
                    chosen = stored
            elif formula:
                chosen = _evaluate_internal_formula(
                    formula,
                    labels,
                    computed,
                    col,
                    decimal_places,
                    resolver=resolve,
                )
                if chosen is None:
                    chosen = stored
            else:
                chosen = stored
            computed[row][col] = canonical_number(chosen) if chosen is not None and chosen > 0 else 1.0
            return computed[row][col]
        finally:
            resolving.remove(key)

    for row, _label in enumerate(labels):
        if settings["average_type"][row] != "user_entry":
            continue
        for col in range(col_count):
            resolve(row, col)
    return computed


def selected_ratio_values(payload: Mapping[str, Any]) -> list[float]:
    """Return the selected development ratio per column at full precision.

    ``average formulas.values`` is stored through :func:`canonical_number`, which
    rounds to six decimals so a DFM file stays reviewable. That is the right
    precision to *display* a ratio and the wrong precision to *chain* one: a
    consumer that multiplies ten stored ratios together, as the Bootstrap method
    does when it back-fits a triangle, amplifies the rounding into a visible
    error. Re-deriving the computed averages from the stored triangle costs one
    pass and keeps a chained result exact.

    A ratio the user owns -- User Entry, or a benchmark row -- is returned as
    stored, because there the six-decimal value *is* the authoritative input
    rather than a rounded projection of one.
    """

    method = normalize_dfm_method(payload, require_complete=False)
    data = method["data_tab"]
    ratio = method["ratios_tab"]["ratio_triangle"]
    formulas = method["ratios_tab"]["average_formulas"]
    settings = formulas["custom_average_formula_settings"]
    selected = formulas["selected"]
    stored = formulas["values"]
    col_count = len(ratio["development_labels"])
    last_computed_col = len(data["development_labels"]) - 1

    ratios: list[float] = []
    for col in range(col_count):
        row = next(
            (
                index
                for index, flags in enumerate(selected)
                if col < len(flags) and flags[col] == 1
            ),
            0,
        )
        average_type = settings["average_type"][row] if row < len(settings["average_type"]) else "custom"
        base = settings["base"][row] if row < len(settings["base"]) else "volume"
        if average_type == "user_entry" or base == "benchmark" or col >= last_computed_col:
            value = canonical_number(stored[row][col] if row < len(stored) and col < len(stored[row]) else None)
            ratios.append(float(value) if value is not None else 1.0)
            continue
        ratios.append(
            float(
                _calculate_average(
                    data["input_data_triangle_values"],
                    data["input_data_triangle_mask"],
                    ratio["excluded"],
                    col,
                    base=base,
                    periods=settings["periods"][row],
                    extra_exclude=settings["exclude"][row],
                )
            )
        )
    return ratios


def _stored_selected_ratios(method: Mapping[str, Any]) -> list[float]:
    """Return the selected ratio per column exactly as the method stores it.

    This is the six-decimal projection a reader sees on the Ratios tab, not the
    re-derived full precision of :func:`selected_ratio_values`. The ultimate and
    the percentage developed are two views of the same chain, so both read the
    stored value and a reader can reproduce either one from the file.
    """

    formulas = method["ratios_tab"]["average_formulas"]
    values = formulas["values"]
    selected = formulas["selected"]
    col_count = len(method["ratios_tab"]["ratio_triangle"]["development_labels"])
    ratios: list[float] = []
    for col in range(col_count):
        chosen = 0
        for row, selected_row in enumerate(selected):
            if col < len(selected_row) and selected_row[col] == 1:
                chosen = row
                break
        value = canonical_number(values[chosen][col] if chosen < len(values) and col < len(values[chosen]) else None)
        ratios.append(float(value) if value is not None else 1.0)
    return ratios


def _selected_development_chain(method: Mapping[str, Any]) -> list[float]:
    """The factor per column the ultimates chain: the Curves tab's selected values.

    ``curves_tab.selected_values`` is refreshed by :func:`recalculate_dfm_method`
    from the Ratios tab's selected factors and the Curves tab's choices. A
    payload that has not been recalculated since the tab existed carries none,
    and then the Ratios tab's selection is the chain, exactly as before the
    tab: the Curves tab starts by selecting the Initial Selection everywhere.
    """

    ratios = _stored_selected_ratios(method)
    curves = _tab(method, "curves_tab")
    selected = curves.get("selected_values") if isinstance(curves.get("selected_values"), list) else []
    if len(selected) == len(ratios) and all(canonical_number(value) is not None for value in selected):
        return [float(canonical_number(value)) for value in selected]
    return ratios


def selected_development_factors(payload: Mapping[str, Any]) -> list[float]:
    """Return the selected development factor per ratio column, the tail last.

    This is the chain the ultimates and the percentage developed use: the
    Curves tab's selected value per period and its selected tail factor.
    """

    return _selected_development_chain(normalize_dfm_method(payload, require_complete=False))


def _cumulative_from_normalized(method: Mapping[str, Any]) -> list[float | None]:
    ratios = _selected_development_chain(method)
    col_count = len(ratios)
    cumulative: list[float | None] = [None] * col_count
    running: float | None = None
    for col in range(col_count - 1, -1, -1):
        value = ratios[col]
        running = value if col == col_count - 1 else (value * running if running is not None else None)
        cumulative[col] = running
    return cumulative


def _latest_column(data: Mapping[str, Any], row: int) -> int | None:
    """Return the development column holding the row's latest observation.

    A zero is an observation, so the column is found by presence rather than by
    value; an origin whose newest figure happens to be zero still sits at a
    known development age.
    """

    row_values = data["input_data_triangle_values"][row]
    row_mask = data["input_data_triangle_mask"][row]
    return next(
        (
            col
            for col in range(min(len(row_values), len(row_mask), len(data["development_labels"])) - 1, -1, -1)
            if row_mask[col] and canonical_input_number(row_values[col]) is not None
        ),
        None,
    )


def selected_cumulative_factors(payload: Mapping[str, Any]) -> list[float | None]:
    """Return the cumulative development factor per column, latest column first."""

    return _cumulative_from_normalized(normalize_dfm_method(payload, require_complete=False))


def dfm_percent_developed_vector(payload: Mapping[str, Any]) -> list[float | int | None]:
    """Return the percentage developed for each origin, one entry per origin row.

    The figure is the reciprocal of the cumulative development factor at the
    origin's own development age -- the same ``% Developed`` the Ratios tab
    shows, read at the column that origin has reached. It is a property of the
    selected factors alone, so an origin whose latest observation is zero, or
    whose ultimate is therefore zero, still carries a meaningful percentage.
    """

    method = normalize_dfm_method(payload, require_complete=False)
    data = method["data_tab"]
    cumulative = _cumulative_from_normalized(method)
    out: list[float | int | None] = []
    for row in range(len(data["input_data_triangle_values"])):
        latest_col = _latest_column(data, row)
        factor = cumulative[latest_col] if latest_col is not None and latest_col < len(cumulative) else None
        out.append(canonical_number(1.0 / float(factor)) if factor else None)
    return out


def _calculate_ultimate(payload: dict[str, Any]) -> list[Any]:
    data = payload["data_tab"]
    cumulative = _cumulative_from_normalized(payload)
    out: list[Any] = []
    for row, row_values in enumerate(data["input_data_triangle_values"]):
        latest_col = _latest_column(data, row)
        if latest_col is None or latest_col >= len(cumulative) or cumulative[latest_col] is None:
            out.append(None)
            continue
        out.append(canonical_number(float(row_values[latest_col]) * float(cumulative[latest_col])))
    return out


def recalculate_dfm_method(
    payload: Mapping[str, Any],
    *,
    input_snapshot: Mapping[str, Any] | None = None,
    ratio_basis_snapshot: Mapping[str, Any] | None = None,
    changed_precedents: Iterable[str] = (),
    timestamp: Any = None,
    update_refresh_timestamp: bool | None = None,
    dataset_reference_values: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Refresh DFM-derived state while preserving the DFM-owned projection.

    ``dataset_reference_values`` maps each ``[Dataset][coordinate]`` reference
    text to its resolved numeric value. When provided, User Entry formulas that
    reference datasets are re-evaluated with those values; without it, their
    stored evaluations are preserved (an ordinary Save trusts the values the
    client already evaluated).
    """

    changed = tuple(str(item) for item in changed_precedents)
    if update_refresh_timestamp is None:
        update_refresh_timestamp = input_snapshot is not None or ratio_basis_snapshot is not None or bool(changed)
    refreshed_at = _timestamp(timestamp)
    method = normalize_dfm_method(payload, require_complete=False, timestamp=refreshed_at)
    if input_snapshot is not None:
        _apply_input_snapshot(method, input_snapshot)
    if ratio_basis_snapshot is not None:
        _apply_ratio_basis_snapshot(method, ratio_basis_snapshot)
    ratio = method["ratios_tab"]["ratio_triangle"]
    data = method["data_tab"]
    ratio["origin_labels"] = list(data["origin_labels"])
    ratio["development_labels"] = _ratio_development_labels(data["development_labels"])
    ratio["ratio_values"] = _calculate_ratio_triangle(
        data["input_data_triangle_values"], data["input_data_triangle_mask"], len(data["development_labels"])
    )
    prior_excluded = _int_matrix(ratio.get("excluded"))
    ratio["excluded"] = [
        (prior_excluded[row] if row < len(prior_excluded) else [])[: len(ratio_values)]
        + [0] * max(0, len(ratio_values) - len(prior_excluded[row] if row < len(prior_excluded) else []))
        for row, ratio_values in enumerate(ratio["ratio_values"])
    ]
    formula_count = len(method["ratios_tab"]["average_formulas"]["label"])
    ratio_col_count = len(ratio["development_labels"])
    formulas = method["ratios_tab"]["average_formulas"]
    formulas["selected"] = _fit_matrix(_int_matrix(formulas.get("selected")), formula_count, ratio_col_count, 0)
    formulas["inputs"] = _fit_matrix(_text_matrix(formulas.get("inputs")), formula_count, ratio_col_count, "")
    formulas["display_inputs"] = _fit_matrix(
        _text_matrix(formulas.get("display_inputs")), formula_count, ratio_col_count, ""
    )
    formulas["values"] = _calculate_formula_values(method, dataset_reference_values)
    stored_chain = _stored_selected_ratios(method)
    method["curves_tab"] = normalize_curves_tab(method.get("curves_tab"), len(stored_chain) - 1, stored_chain[:-1])
    curves = curves_table(stored_chain[:-1], stored_chain[-1] if stored_chain else 1.0, method["curves_tab"])
    method["curves_tab"]["selected_values"] = [
        canonical_number(value) for value in [*curves["selected_values"], curves["selected_tail"]]
    ]
    method["results_tab"]["ultimate_vector"] = _calculate_ultimate(method)
    if update_refresh_timestamp:
        method["method_metadata"]["data_refreshed"] = refreshed_at
    _set_revisions(method)
    _validate_complete(method)
    return method


def preview_dfm_method(
    payload: Mapping[str, Any],
    *,
    input_snapshot: Mapping[str, Any] | None = None,
    ratio_basis_snapshot: Mapping[str, Any] | None = None,
    timestamp: Any = None,
) -> dict[str, Any]:
    """Calculate a complete in-memory preview; no I/O or external refresh occurs."""

    return recalculate_dfm_method(
        payload,
        input_snapshot=input_snapshot,
        ratio_basis_snapshot=ratio_basis_snapshot,
        timestamp=timestamp,
        update_refresh_timestamp=False,
    )


_OWNED_PATHS = (
    ("details_tab", "name"),
    ("details_tab", "output_type"),
    ("details_tab", "output_dataset"),
    ("details_tab", "output_category"),
    ("details_tab", "input_triangle"),
    ("details_tab", "origin_length"),
    ("details_tab", "development_length"),
    ("details_tab", "decimal_places"),
    ("ratios_tab", "average_formulas", "label"),
    ("ratios_tab", "average_formulas", "custom_average_formula_settings"),
    ("ratios_tab", "average_formulas", "selected"),
    ("ratios_tab", "average_formulas", "inputs"),
    ("ratios_tab", "average_formulas", "display_inputs"),
    ("ratios_tab", "average_formulas", "values"),
    ("ratios_tab", "cell_notes"),
    ("results_tab", "ratio_basis_dataset"),
    ("results_tab", "ultimate_ratio_decimal_places"),
    ("curves_tab", "fitting_method"),
    ("curves_tab", "future_development_periods"),
    ("curves_tab", "free_fit_c"),
    ("curves_tab", "included"),
    ("curves_tab", "user_columns"),
    ("curves_tab", "selected_estimates"),
    ("curves_tab", "selected_tail_factor"),
    ("curves_tab", "selected_tail_curve"),
)


def _apply_owned_exclusion_patch(base: dict[str, Any], patch: Mapping[str, Any]) -> None:
    patch_ratio = _tab(_tab(patch, "ratios_tab"), "ratio_triangle")
    if "excluded" not in patch_ratio:
        return
    patch_excluded = _int_matrix(patch_ratio.get("excluded"))
    patch_origins = _labels(patch_ratio.get("origin_labels"))
    patch_devs = _labels(patch_ratio.get("development_labels"))
    if not patch_origins or not patch_devs:
        raise DfmContractError(
            "A DFM exclusion patch must include its exact ratio origin and development labels."
        )
    if _duplicate_labels(patch_origins) or _duplicate_labels(patch_devs):
        raise DfmContractError("A DFM exclusion patch cannot contain duplicate labels.")
    base_ratio = base["ratios_tab"]["ratio_triangle"]
    base_origins = _labels(base_ratio.get("origin_labels"))
    base_devs = _labels(base_ratio.get("development_labels"))
    base_excluded = _int_matrix(base_ratio.get("excluded"))
    origin_lookup = {label: index for index, label in enumerate(base_origins)}
    dev_lookup = {label: index for index, label in enumerate(base_devs)}
    for patch_row, origin in enumerate(patch_origins):
        base_row = origin_lookup.get(origin)
        if base_row is None or patch_row >= len(patch_excluded):
            continue
        while len(base_excluded) <= base_row:
            base_excluded.append([])
        for patch_col, value in enumerate(patch_excluded[patch_row]):
            if patch_col >= len(patch_devs):
                break
            base_col = dev_lookup.get(patch_devs[patch_col])
            if base_col is None:
                continue
            while len(base_excluded[base_row]) <= base_col:
                base_excluded[base_row].append(0)
            base_excluded[base_row][base_col] = value
    base_ratio["excluded"] = base_excluded


def _path_value(payload: Mapping[str, Any], path: tuple[str, ...]) -> tuple[bool, Any]:
    current: Any = payload
    for key in path:
        if not isinstance(current, Mapping) or key not in current:
            return False, None
        current = current[key]
    return True, current


def _set_path(payload: dict[str, Any], path: tuple[str, ...], value: Any) -> None:
    current = payload
    for key in path[:-1]:
        child = current.get(key)
        if not isinstance(child, dict):
            child = {}
            current[key] = child
        current = child
    current[path[-1]] = deepcopy(value)


def apply_owned_patch(
    base: Mapping[str, Any],
    patch: Mapping[str, Any],
    *,
    timestamp: Any = None,
) -> dict[str, Any]:
    """Rebase an owned-state patch onto the newest embedded derived snapshot."""

    method = normalize_dfm_method(base, require_complete=False, timestamp=timestamp)
    _apply_owned_exclusion_patch(method, patch)
    for path in _OWNED_PATHS:
        exists, value = _path_value(patch, path)
        if exists:
            _set_path(method, path, value)
    modified_at = _timestamp(timestamp)
    method["method_metadata"]["last_modified"] = modified_at
    return recalculate_dfm_method(
        method,
        timestamp=modified_at,
        update_refresh_timestamp=False,
    )


def stamp_last_modified(payload: Mapping[str, Any], modified_at: Any) -> dict[str, Any]:
    """Return *payload* with only its ``last modified`` value replaced.

    A DFM uploaded to the RPC server is saved there, and ResQ stamps that save
    with its own ``Modified``. Nothing about the ArcRho method changed, so the
    two copies are identical in content while their recorded times are not, and
    the next sync reports the remote as newer. Recording the time ResQ actually
    wrote makes the pair agree again.

    This is deliberately not a save: no recalculation, no refresh timestamp, and
    no other field is touched. ``last modified`` is outside all three revision
    projections, so a stamp cannot shift a revision, invalidate an open editor's
    optimistic-concurrency token, or make a dependent look stale.
    """

    method = deepcopy(dict(payload))
    metadata = method.get("method_metadata")
    if not isinstance(metadata, dict):
        metadata = {}
        method["method_metadata"] = metadata
    metadata["last_modified"] = _timestamp(modified_at)
    return method


build_dfm_method_v2 = normalize_dfm_method


__all__ = [
    "DFM_JSON_FORMAT",
    "DFM_VALUE_DECIMAL_PLACES",
    "DfmContractError",
    "aggregate_vector_values",
    "apply_owned_patch",
    "build_dfm_method_v2",
    "build_dfm_output_sidecar",
    "canonical_number",
    "contains_excel_reference",
    "derived_projection",
    "dfm_percent_developed_vector",
    "dfm_precedent_names",
    "dfm_output_variants",
    "method_revisions",
    "selected_cumulative_factors",
    "selected_ratio_values",
    "normalize_dfm_method",
    "default_average_formulas",
    "owned_projection",
    "persisted_projection",
    "preview_dfm_method",
    "publication_projection",
    "recalculate_dfm_method",
    "source_snapshot_revision",
    "stamp_last_modified",
]
