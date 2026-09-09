"""Canonical, filesystem-free contract for Cape Cod methods.

All persisted-data producers delegate Cape Cod normalization, calculations,
trend-rate fitting, revision hashes, output variants, the as-if ultimates
triangle, and output-sidecar projection to this module.  A current v1 method
contains every value needed to open the Method and Ratios tabs without reading
its precedent datasets.

Every formula here is the ResQ Generalised Cape Cod calculation, verified
cell-exact against the ResQ COM API (see docs/plans/cape_cod_method_plan.md).
"""

from __future__ import annotations

import math
from copy import deepcopy
from datetime import datetime, timezone
from typing import Any, Iterable, Mapping

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

from .dataset_display_contract import normalize_show_subtotal
from .dfm_contract import aggregate_vector_values, canonical_input_number
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
    stored_length_fields,
    validate_sidecar_core,
)
from .timestamps import persisted_timestamp as _timestamp


_RATE_QUANTUM = Decimal("0.00000001")

CC_JSON_FORMAT = "arcrho-cape-cod-v4"
CC_METHOD_TYPE = "Cape Cod"
CC_SOURCE_KIND = "cape_cod"
CC_METHOD_TYPE_CODE = 3

CC_PRIOR_ULTIMATE_MODES = ("latest_ultimates", "pattern")
CC_SCALING_TYPES = ("percentage", "unscaled", "auto_scaled")

# Derived Method-tab columns in canonical order, excluding the effective trend rate.
CC_DERIVED_COLUMNS = (
    "trend_factors",
    "trended_latest_values",
    "percentage_developed",
    "development_factors",
    "developed_exposure_values",
    "future_exposure_values",
    "trended_developed_ratios",
    "expected_ultimate_ratios",
    "detrended_expected_ratios",
    "future_latest_values",
    "cape_cod_ultimate",
    "cape_cod_ultimate_ratios",
)


class CapeCodContractError(ValueError):
    """Raised when a Cape Cod payload cannot satisfy the canonical v1 contract."""


# Compact alias for callers that prefer the abbreviation used by the UI.
CcContractError = CapeCodContractError


def _clean(value: Any) -> str:
    return " ".join(str(value if value is not None else "").split()).strip()


def _integer(value: Any, default: int, *, minimum: int = 0, maximum: int | None = None) -> int:
    try:
        result = int(value)
    except (TypeError, ValueError):
        result = default
    result = max(minimum, result)
    return min(result, maximum) if maximum is not None else result


def _tab(payload: Mapping[str, Any], name: str) -> dict[str, Any]:
    value = payload.get(name)
    return dict(value) if isinstance(value, Mapping) else {}


def _labels(value: Any) -> list[str]:
    return [str(item if item is not None else "").strip() for item in value] if isinstance(value, list) else []


def _number(value: Any) -> float | int | None:
    """Canonicalize JSON numbers without retaining producer-specific 1 vs 1.0 types.

    The number is kept at the precision it was observed with. A percentage
    developed, an a-priori ratio and an ultimate all come from a DFM that
    chains its factors in full double precision, so quantizing the copy here
    would reintroduce, one method further down, exactly the drift from ResQ the
    chain was fixed to remove.
    """

    result = canonical_input_number(value)
    if isinstance(result, float) and result.is_integer():
        return int(result)
    return result


def _numbers(value: Any) -> list[float | int | None]:
    if not isinstance(value, list):
        return []
    return [_number(item[0] if isinstance(item, list) and item else item) for item in value]


def _fit(values: list[Any], size: int, fill: Any) -> list[Any]:
    if size <= 0:
        return list(values)
    return list(values[:size]) + [deepcopy(fill) for _ in range(max(0, size - len(values)))]


def _duplicates(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    duplicates: list[str] = []
    for value in values:
        if value in seen and value not in duplicates:
            duplicates.append(value)
        seen.add(value)
    return duplicates


def _snapshot_revision(name: str, origin_labels: list[str], values: list[Any]) -> str:
    if not name or not origin_labels or len(origin_labels) != len(values):
        return ""
    return fingerprint({
        "name": name,
        "origin_labels": origin_labels,
        "values": [_number(value) for value in values],
    })


def _prior_ultimate_snapshot_revision(
    name: str, origin_labels: list[str], values: list[Any], pattern: list[Any]
) -> str:
    """Fingerprint both vectors the prior ultimate precedent supplies.

    The percentage developed is not implied by the ultimate -- an origin whose
    latest observation is zero has a zero ultimate under any pattern -- so a
    revision covering only the ultimate could miss a changed selection.
    """

    if not name or not origin_labels or len(origin_labels) != len(values)             or len(origin_labels) != len(pattern):
        return ""
    return fingerprint({
        "name": name,
        "origin_labels": origin_labels,
        "values": [_number(value) for value in values],
        "percentage_developed": [_number(value) for value in pattern],
    })


def _finite(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def _rate(value: Any) -> float | int:
    """Canonicalize a rate/factor parameter at eight decimals.

    Rates are displayed as percentages with six decimals (the ResQ Trend Rate
    precision), so they need eight decimal places as decimals. Eight is what
    the box itself offers rather than a projection of a computed number, so a
    rate keeps every digit a user typed or ResQ supplied.
    """

    number = _finite(value)
    if number is None:
        return 0
    try:
        rounded = float(Decimal(str(abs(number))).quantize(_RATE_QUANTUM, rounding=ROUND_HALF_UP))
    except (InvalidOperation, ValueError):
        return 0
    result = -rounded if number < 0 else rounded
    if result == 0:
        return 0
    return int(result) if float(result).is_integer() else result


def _prior_mode(value: Any) -> str:
    mode = _clean(value).casefold().replace(" ", "_").replace("/", "_")
    return mode if mode in CC_PRIOR_ULTIMATE_MODES else CC_PRIOR_ULTIMATE_MODES[0]


def _scaling(value: Any) -> str:
    scaling = _clean(value).casefold().replace(" ", "_").replace("-", "_")
    return scaling if scaling in CC_SCALING_TYPES else CC_SCALING_TYPES[0]


def fit_cape_cod_trend_rate(latest: list[Any], developed_exposure: list[Any]) -> float | int:
    """ResQ ``FitTrendRate``: weighted log regression of the untrended developed
    ratio against origin position, weighted by developed exposure."""

    points: list[tuple[float, float, float]] = []
    for index in range(min(len(latest), len(developed_exposure))):
        latest_value = _finite(latest[index])
        weight = _finite(developed_exposure[index])
        if latest_value is None or weight is None or weight <= 0:
            continue
        ratio = latest_value / weight
        if ratio <= 0:
            continue
        points.append((float(index), math.log(ratio), weight))
    if len(points) < 2:
        return 0
    total_weight = sum(weight for _, _, weight in points)
    xw = sum(x * weight for x, _, weight in points)
    yw = sum(y * weight for _, y, weight in points)
    xxw = sum(x * x * weight for x, _, weight in points)
    xyw = sum(x * y * weight for x, y, weight in points)
    sxx = xxw - xw * xw / total_weight
    sxy = xyw - xw * yw / total_weight
    if sxx == 0:
        return 0
    return _rate(math.exp(sxy / sxx) - 1)


def _percentage_developed(
    latest: list[Any], prior_values: list[Any], pattern: list[Any], mode: str, row_count: int
) -> list[Any]:
    """Return the percentage developed for each origin.

    In ``pattern`` mode the selected vector *is* the pattern. In
    ``latest_ultimates`` mode the figure belongs to the development factors
    behind the prior ultimate, so a prior ultimate a DFM published carries its
    own pattern and that pattern is used directly -- an origin whose latest
    observation is zero has a zero ultimate under any pattern, and dividing one
    by the other would describe it as undeveloped. Only a prior ultimate with no
    DFM behind it, which Cape Cod also accepts, falls back to the ratio.
    """

    percentages: list[Any] = []
    for index in range(row_count):
        prior = _finite(prior_values[index] if index < len(prior_values) else None)
        if mode == "pattern":
            percentages.append(_number(prior))
            continue
        developed = _finite(pattern[index] if index < len(pattern) else None)
        if developed is not None:
            percentages.append(_number(developed))
            continue
        latest_value = _finite(latest[index] if index < len(latest) else None)
        if latest_value is None or prior is None or prior == 0:
            percentages.append(None)
        else:
            percentages.append(_number(latest_value / prior))
    return percentages


def _expected_ultimate_ratios(
    developed_exposure: list[Any], trended_developed_ratios: list[Any], decay: float, row_count: int
) -> list[Any]:
    usable: list[tuple[int, float, float]] = []
    for index in range(row_count):
        weight = _finite(developed_exposure[index] if index < len(developed_exposure) else None)
        ratio = _finite(trended_developed_ratios[index] if index < len(trended_developed_ratios) else None)
        if weight is None or ratio is None:
            continue
        usable.append((index, weight, ratio))
    expected: list[Any] = []
    for index in range(row_count):
        numerator = 0.0
        denominator = 0.0
        for other, weight, ratio in usable:
            decayed = weight * decay ** abs(index - other)
            numerator += decayed * ratio
            denominator += decayed
        expected.append(_number(numerator / denominator) if denominator != 0 else None)
    return expected


def _calculate_columns(method_tab: Mapping[str, Any]) -> dict[str, Any]:
    """Return the effective trend rate plus every derived Method-tab column."""

    origins = _labels(method_tab.get("origin_labels"))
    row_count = len(origins)
    latest = _fit(_numbers(method_tab.get("latest_values")), row_count, None)
    exposure = _fit(_numbers(method_tab.get("exposure_values")), row_count, None)
    prior = _fit(_numbers(method_tab.get("prior_ultimate_values")), row_count, None)
    pattern = _fit(_numbers(method_tab.get("prior_ultimate_percentage_developed")), row_count, None)
    mode = _prior_mode(method_tab.get("prior_ultimate_mode"))
    decay_value = _finite(method_tab.get("decay_factor"))
    decay = decay_value if decay_value is not None else 0.0
    auto_fit = bool(method_tab.get("auto_trend_fit", False))
    overrides = _fit(_numbers(method_tab.get("trend_factor_overrides")), row_count, None)
    if auto_fit:
        overrides = [None] * row_count

    percentages = _percentage_developed(latest, prior, pattern, mode, row_count)
    developed_exposure: list[Any] = []
    for index in range(row_count):
        exposure_value = _finite(exposure[index])
        percentage = _finite(percentages[index])
        developed_exposure.append(
            _number(exposure_value * percentage)
            if exposure_value is not None and percentage is not None
            else None
        )

    if auto_fit:
        trend_rate = fit_cape_cod_trend_rate(latest, developed_exposure)
    else:
        trend_rate = _rate(method_tab.get("trend_rate"))

    trend_factors: list[Any] = []
    for index in range(row_count):
        override = _finite(overrides[index])
        if override is not None:
            trend_factors.append(_number(override))
        else:
            trend_factors.append(_number((1.0 + float(trend_rate)) ** (row_count - 1 - index)))

    trended_latest: list[Any] = []
    development_factors: list[Any] = []
    future_exposure: list[Any] = []
    trended_developed_ratios: list[Any] = []
    for index in range(row_count):
        latest_value = _finite(latest[index])
        factor = _finite(trend_factors[index])
        trended = latest_value * factor if latest_value is not None and factor is not None else None
        trended_latest.append(_number(trended))
        percentage = _finite(percentages[index])
        development_factors.append(
            _number(1.0 / percentage) if percentage is not None and percentage != 0 else None
        )
        exposure_value = _finite(exposure[index])
        developed = _finite(developed_exposure[index])
        future_exposure.append(
            _number(exposure_value - developed)
            if exposure_value is not None and developed is not None
            else None
        )
        trended_developed_ratios.append(
            _number(trended / developed)
            if trended is not None and developed is not None and developed != 0
            else None
        )

    expected = _expected_ultimate_ratios(
        developed_exposure, trended_developed_ratios, float(decay), row_count
    )

    detrended: list[Any] = []
    future_latest: list[Any] = []
    ultimates: list[Any] = []
    ratios: list[Any] = []
    alternative = bool(method_tab.get("alternative_ultimate_calculation", False))
    for index in range(row_count):
        expected_value = _finite(expected[index])
        factor = _finite(trend_factors[index])
        detrended_value = (
            expected_value / factor
            if expected_value is not None and factor is not None and factor != 0
            else None
        )
        detrended.append(_number(detrended_value))
        future_exposure_value = _finite(future_exposure[index])
        future_value = (
            future_exposure_value * detrended_value
            if future_exposure_value is not None and detrended_value is not None
            else None
        )
        future_latest.append(_number(future_value))
        latest_value = _finite(latest[index])
        percentage = _finite(percentages[index])
        exposure_value = _finite(exposure[index])
        if (
            alternative
            and latest_value is not None
            and latest_value != 0
            and percentage == 0
            and detrended_value is not None
            and exposure_value is not None
        ):
            ultimate = detrended_value * exposure_value
        elif latest_value is not None and future_value is not None:
            ultimate = latest_value + future_value
        else:
            ultimate = None
        ultimates.append(_number(ultimate))
        ratios.append(
            _number(ultimate / exposure_value)
            if ultimate is not None and exposure_value is not None and exposure_value != 0
            else None
        )

    return {
        "trend_rate": trend_rate,
        "trend_factor_overrides": overrides,
        "trend_factors": trend_factors,
        "trended_latest_values": trended_latest,
        "percentage_developed": percentages,
        "development_factors": development_factors,
        "developed_exposure_values": developed_exposure,
        "future_exposure_values": future_exposure,
        "trended_developed_ratios": trended_developed_ratios,
        "expected_ultimate_ratios": expected,
        "detrended_expected_ratios": detrended,
        "future_latest_values": future_latest,
        "cape_cod_ultimate": ultimates,
        "cape_cod_ultimate_ratios": ratios,
    }


def owned_projection(payload: Mapping[str, Any]) -> dict[str, Any]:
    details = _tab(payload, "details_tab")
    method = _tab(payload, "method_tab")
    auto_fit = bool(method.get("auto_trend_fit", False))
    owned: dict[str, Any] = {
        "details_tab": deepcopy(details),
        "method_tab": {
            "latest_dataset": method.get("latest_dataset", ""),
            "exposure_dataset": method.get("exposure_dataset", ""),
            "prior_ultimate_dataset": method.get("prior_ultimate_dataset", ""),
            "prior_ultimate_mode": _prior_mode(method.get("prior_ultimate_mode")),
            "auto_trend_fit": auto_fit,
            "decay_factor": _number(method.get("decay_factor")),
            "scaling_type": _scaling(method.get("scaling_type")),
            "alternative_ultimate_calculation": bool(
                method.get("alternative_ultimate_calculation", False)
            ),
            "trend_factor_overrides": deepcopy(method.get("trend_factor_overrides") or []),
        },
    }
    if not auto_fit:
        owned["method_tab"]["trend_rate"] = _number(method.get("trend_rate"))
    return owned


def derived_projection(payload: Mapping[str, Any]) -> dict[str, Any]:
    method = _tab(payload, "method_tab")
    derived: dict[str, Any] = {
        "origin_labels": deepcopy(method.get("origin_labels") or []),
        "latest_values": deepcopy(method.get("latest_values") or []),
        "latest_source_revision": method.get("latest_source_revision", ""),
        "exposure_values": deepcopy(method.get("exposure_values") or []),
        "exposure_source_revision": method.get("exposure_source_revision", ""),
        "prior_ultimate_values": deepcopy(method.get("prior_ultimate_values") or []),
        "prior_ultimate_percentage_developed": deepcopy(
            method.get("prior_ultimate_percentage_developed") or []
        ),
        "prior_ultimate_source_revision": method.get("prior_ultimate_source_revision", ""),
    }
    if bool(method.get("auto_trend_fit", False)):
        derived["trend_rate"] = _number(method.get("trend_rate"))
    for key in CC_DERIVED_COLUMNS:
        derived[key] = deepcopy(method.get(key) or [])
    return derived


def publication_projection(payload: Mapping[str, Any]) -> dict[str, Any]:
    details = _tab(payload, "details_tab")
    method = _tab(payload, "method_tab")
    return {
        "dataset_name": details.get("name", ""),
        "dataset_type": details.get("output_type", ""),
        "dataset_category": details.get("dataset_category", ""),
        "origin_length": details.get("origin_length", 12),
        "statistic_decimal_places": details.get("statistic_decimal_places", 2),
        "origin_labels": deepcopy(method.get("origin_labels") or []),
        "cape_cod_ultimate": deepcopy(method.get("cape_cod_ultimate") or []),
    }


def method_revisions(payload: Mapping[str, Any]) -> dict[str, str]:
    """Return deterministic revisions for owned, derived, and published state."""

    return {
        "owned_revision": fingerprint(owned_projection(payload)),
        "derived_revision": fingerprint(derived_projection(payload)),
        "publication_revision": fingerprint(publication_projection(payload)),
    }


def _set_revisions(payload: dict[str, Any]) -> None:
    payload.setdefault("method_metadata", {}).update(method_revisions(payload))


def _current_payload(payload: Mapping[str, Any]) -> dict[str, Any]:
    """Stamp an unversioned producer input without upgrading any legacy marker."""

    json_format = str(payload.get("json_format") or "").strip()
    if json_format not in {"", CC_JSON_FORMAT}:
        raise CapeCodContractError(f"Unsupported Cape Cod JSON format: {json_format!r}.")
    stamped = deepcopy(dict(payload))
    stamped["json_format"] = CC_JSON_FORMAT
    return stamped


def _flat_vector(value: Any, expected: int, label: str) -> None:
    if not isinstance(value, list) or len(value) != expected:
        raise CapeCodContractError(
            f"Cape Cod {label} must contain exactly one scalar per origin label."
        )
    if any(isinstance(item, (list, tuple, dict)) for item in value):
        raise CapeCodContractError(f"Cape Cod {label} must be a flat scalar vector.")


def _validate_raw_complete_shape(payload: Mapping[str, Any], row_count: int) -> None:
    method = _tab(payload, "method_tab")
    for key in ("latest_values", "exposure_values", "prior_ultimate_values") + CC_DERIVED_COLUMNS:
        _flat_vector(method.get(key), row_count, key)
    _flat_vector(method.get("trend_factor_overrides"), row_count, "trend_factor_overrides")


def normalize_cape_cod_method(
    payload: Mapping[str, Any],
    *,
    require_complete: bool = True,
    timestamp: Any = None,
) -> dict[str, Any]:
    """Return the exact canonical, self-contained Cape Cod v1 payload."""

    if not isinstance(payload, Mapping):
        raise CapeCodContractError("Cape Cod method payload must be a JSON object.")
    json_format = str(payload.get("json_format") or "").strip()
    if json_format != CC_JSON_FORMAT:
        raise CapeCodContractError(f"Unsupported Cape Cod JSON format: {json_format!r}.")

    details_source = _tab(payload, "details_tab")
    method_source = _tab(payload, "method_tab")
    metadata_source = _tab(payload, "method_metadata")
    origins = _labels(method_source.get("origin_labels"))
    row_count = len(origins)
    if require_complete:
        _validate_raw_complete_shape(payload, row_count)

    latest_name = _clean(method_source.get("latest_dataset"))
    exposure_name = _clean(method_source.get("exposure_dataset"))
    prior_name = _clean(method_source.get("prior_ultimate_dataset"))
    latest_values = _fit(_numbers(method_source.get("latest_values")), row_count, None)
    exposure_values = _fit(_numbers(method_source.get("exposure_values")), row_count, None)
    prior_values = _fit(_numbers(method_source.get("prior_ultimate_values")), row_count, None)
    prior_pattern = _fit(
        _numbers(method_source.get("prior_ultimate_percentage_developed")), row_count, None
    )
    auto_fit = bool(method_source.get("auto_trend_fit", False))
    overrides = _fit(_numbers(method_source.get("trend_factor_overrides")), row_count, None)
    if auto_fit:
        overrides = [None] * row_count

    default_time = _timestamp(timestamp)
    last_modified = str(metadata_source.get("last_modified") or "").strip() or default_time
    data_refreshed = str(metadata_source.get("data_refreshed") or "").strip() or last_modified
    normalized = {
        "json_format": CC_JSON_FORMAT,
        "details_tab": {
            "name": _clean(details_source.get("name")),
            "method_type": CC_METHOD_TYPE,
            "output_type": _clean(details_source.get("output_type")),
            "dataset_category": _clean(details_source.get("dataset_category")),
            "origin_length": _integer(details_source.get("origin_length"), 12, minimum=1),
            "statistic_decimal_places": _integer(
                details_source.get("statistic_decimal_places"), 2, minimum=0, maximum=8
            ),
        },
        "method_tab": {
            "latest_dataset": latest_name,
            "latest_values": latest_values,
            "latest_source_revision": _snapshot_revision(latest_name, origins, latest_values),
            "exposure_dataset": exposure_name,
            "exposure_values": exposure_values,
            "exposure_source_revision": _snapshot_revision(exposure_name, origins, exposure_values),
            "prior_ultimate_dataset": prior_name,
            "prior_ultimate_mode": _prior_mode(method_source.get("prior_ultimate_mode")),
            "prior_ultimate_values": prior_values,
            "prior_ultimate_percentage_developed": prior_pattern,
            "prior_ultimate_source_revision": _prior_ultimate_snapshot_revision(
                prior_name, origins, prior_values, prior_pattern
            ),
            "trend_rate": _rate(method_source.get("trend_rate")),
            "auto_trend_fit": auto_fit,
            "decay_factor": _rate(method_source.get("decay_factor")),
            "scaling_type": _scaling(method_source.get("scaling_type")),
            "alternative_ultimate_calculation": bool(
                method_source.get("alternative_ultimate_calculation", False)
            ),
            "trend_factor_overrides": overrides,
            "origin_labels": origins,
        },
        "method_metadata": {
            "method_type": CC_METHOD_TYPE,
            "source_kind": CC_SOURCE_KIND,
            "last_modified": last_modified,
            "data_refreshed": data_refreshed,
            "owned_revision": "",
            "derived_revision": "",
            "publication_revision": "",
        },
    }
    method_tab = normalized["method_tab"]
    for key in CC_DERIVED_COLUMNS:
        method_tab[key] = _fit(_numbers(method_source.get(key)), row_count, None)
    _set_revisions(normalized)
    if require_complete:
        _validate_complete(normalized)
    return normalized


def _validate_complete(payload: Mapping[str, Any]) -> None:
    details = _tab(payload, "details_tab")
    method = _tab(payload, "method_tab")
    for key in ("name", "output_type"):
        if not _clean(details.get(key)):
            raise CapeCodContractError(f"Cape Cod details_tab.{key} is required.")
    if _integer(details.get("origin_length"), 0) not in {1, 3, 6, 12}:
        raise CapeCodContractError("Cape Cod origin_length must be 1, 3, 6, or 12 months.")
    origins = _labels(method.get("origin_labels"))
    if not origins or any(not label for label in origins):
        raise CapeCodContractError("Cape Cod origin_labels must be non-empty.")
    duplicates = _duplicates(origins)
    if duplicates:
        raise CapeCodContractError("Cape Cod origin_labels must be unique: " + ", ".join(duplicates))
    for name_key, values_key, revision_key in (
        ("latest_dataset", "latest_values", "latest_source_revision"),
        ("exposure_dataset", "exposure_values", "exposure_source_revision"),
        ("prior_ultimate_dataset", "prior_ultimate_values", "prior_ultimate_source_revision"),
    ):
        if not _clean(method.get(name_key)):
            raise CapeCodContractError(f"Cape Cod method_tab.{name_key} is required.")
        if len(method.get(values_key) or []) != len(origins):
            raise CapeCodContractError(f"Cape Cod {values_key} must align to origin_labels.")
        if not _clean(method.get(revision_key)):
            raise CapeCodContractError(f"Cape Cod {revision_key} is required.")
    decay = _finite(method.get("decay_factor"))
    if decay is None or decay < 0 or decay > 1:
        raise CapeCodContractError("Cape Cod decay_factor must be between 0 and 1.")
    expected = _calculate_columns(method)
    if _rate(method.get("trend_rate")) != expected["trend_rate"]:
        raise CapeCodContractError(
            "Cape Cod trend_rate does not match the embedded source snapshots."
        )
    for key in CC_DERIVED_COLUMNS:
        if method.get(key) != expected[key]:
            raise CapeCodContractError(
                f"Cape Cod {key} does not match the embedded source snapshots."
            )


def _snapshot_vector(snapshot: Mapping[str, Any], *, latest: bool) -> tuple[list[str], list[Any]]:
    labels = _labels(snapshot.get("origin_labels"))
    raw_values = snapshot.get("values") if isinstance(snapshot.get("values"), list) else []
    raw_mask = snapshot.get("mask") if isinstance(snapshot.get("mask"), list) else []
    values: list[Any] = []
    for row_index, raw in enumerate(raw_values):
        if not isinstance(raw, list):
            values.append(_number(raw))
            continue
        if not latest:
            values.append(_number(raw[0] if raw else None))
            continue
        mask_row = raw_mask[row_index] if row_index < len(raw_mask) and isinstance(raw_mask[row_index], list) else []
        selected = None
        for column in range(len(raw) - 1, -1, -1):
            if mask_row and (column >= len(mask_row) or not bool(mask_row[column])):
                continue
            candidate = _number(raw[column])
            if candidate is not None:
                selected = candidate
                break
        values.append(selected)
    return labels, _fit(values, len(labels), None)


def _align_by_labels(labels: list[str], values: list[Any], origins: list[str]) -> list[Any]:
    duplicates = _duplicates(labels)
    if duplicates:
        raise CapeCodContractError(
            "Cape Cod source snapshot has duplicate origins: " + ", ".join(duplicates)
        )
    if set(labels) != set(origins) or len(labels) != len(origins):
        raise CapeCodContractError(
            "Cape Cod source snapshot origins must match the Latest origins exactly."
        )
    lookup = {label: values[index] for index, label in enumerate(labels)}
    return [lookup[label] for label in origins]


def _aligned_snapshot(
    snapshot: Mapping[str, Any],
    origins: list[str],
    *,
    latest: bool,
) -> list[Any]:
    labels, values = _snapshot_vector(snapshot, latest=latest)
    return _align_by_labels(labels, values, origins)


def _aligned_percentages(snapshot: Mapping[str, Any], origins: list[str]) -> list[Any]:
    """Align a precedent's percentage-developed pattern onto Cape Cod origins."""

    labels = _labels(snapshot.get("origin_labels"))
    values = _fit(_numbers(snapshot.get("percentage_developed")), len(labels), None)
    return _align_by_labels(labels, values, origins)


def _source_snapshot(source_snapshots: Mapping[str, Any], role: str, name: str) -> Mapping[str, Any] | None:
    direct = source_snapshots.get(role)
    if isinstance(direct, Mapping):
        return direct
    by_name = source_snapshots.get(name)
    return by_name if isinstance(by_name, Mapping) else None


def recalculate_cape_cod_method(
    payload: Mapping[str, Any],
    *,
    source_snapshots: Mapping[str, Any] | None = None,
    changed_precedents: Iterable[str] = (),
    timestamp: Any = None,
    update_refresh_timestamp: bool | None = None,
) -> dict[str, Any]:
    """Refresh Cape Cod derived state from optional aggregate source snapshots."""

    changed = tuple(str(item) for item in changed_precedents)
    refreshed_at = _timestamp(timestamp)
    if update_refresh_timestamp is None:
        update_refresh_timestamp = source_snapshots is not None or bool(changed)
    method = normalize_cape_cod_method(
        _current_payload(payload), require_complete=False, timestamp=refreshed_at
    )
    tab = method["method_tab"]
    snapshots = source_snapshots if isinstance(source_snapshots, Mapping) else {}

    latest_snapshot = _source_snapshot(snapshots, "latest", tab["latest_dataset"])
    if latest_snapshot is not None:
        snapshot_name = _clean(latest_snapshot.get("name"))
        if not tab["latest_dataset"] and snapshot_name:
            tab["latest_dataset"] = snapshot_name
        elif snapshot_name and snapshot_name.casefold() != tab["latest_dataset"].casefold():
            raise CapeCodContractError(
                "Cape Cod Latest snapshot identity does not match its configured source."
            )
        new_origins, _ = _snapshot_vector(latest_snapshot, latest=True)
        if not new_origins or any(not label for label in new_origins) or _duplicates(new_origins):
            raise CapeCodContractError("Cape Cod Latest snapshot requires unique non-empty origins.")
        old_origins = _labels(tab.get("origin_labels"))
        override_map = dict(zip(old_origins, tab.get("trend_factor_overrides") or []))
        tab["origin_labels"] = new_origins
        tab["trend_factor_overrides"] = [_number(override_map.get(label)) for label in new_origins]

    origins = _labels(tab.get("origin_labels"))
    if latest_snapshot is not None:
        tab["latest_values"] = _aligned_snapshot(latest_snapshot, origins, latest=True)
    for role, name_key, values_key in (
        ("exposure", "exposure_dataset", "exposure_values"),
        ("prior_ultimate", "prior_ultimate_dataset", "prior_ultimate_values"),
    ):
        snapshot = _source_snapshot(snapshots, role, tab[name_key])
        if snapshot is None:
            continue
        snapshot_name = _clean(snapshot.get("name"))
        if not tab[name_key] and snapshot_name:
            tab[name_key] = snapshot_name
        elif snapshot_name and snapshot_name.casefold() != tab[name_key].casefold():
            raise CapeCodContractError(
                f"Cape Cod {role} snapshot identity does not match its configured source."
            )
        tab[values_key] = _aligned_snapshot(snapshot, origins, latest=False)
        if role == "prior_ultimate":
            tab["prior_ultimate_percentage_developed"] = _aligned_percentages(snapshot, origins)

    row_count = len(origins)
    tab["latest_values"] = _fit(_numbers(tab.get("latest_values")), row_count, None)
    tab["exposure_values"] = _fit(_numbers(tab.get("exposure_values")), row_count, None)
    tab["prior_ultimate_values"] = _fit(_numbers(tab.get("prior_ultimate_values")), row_count, None)
    tab["prior_ultimate_percentage_developed"] = _fit(
        _numbers(tab.get("prior_ultimate_percentage_developed")), row_count, None
    )
    tab["trend_factor_overrides"] = _fit(_numbers(tab.get("trend_factor_overrides")), row_count, None)
    tab["latest_source_revision"] = _snapshot_revision(tab["latest_dataset"], origins, tab["latest_values"])
    tab["exposure_source_revision"] = _snapshot_revision(tab["exposure_dataset"], origins, tab["exposure_values"])
    tab["prior_ultimate_source_revision"] = _prior_ultimate_snapshot_revision(
        tab["prior_ultimate_dataset"],
        origins,
        tab["prior_ultimate_values"],
        tab["prior_ultimate_percentage_developed"],
    )
    columns = _calculate_columns(tab)
    tab["trend_rate"] = columns["trend_rate"]
    tab["trend_factor_overrides"] = columns["trend_factor_overrides"]
    for key in CC_DERIVED_COLUMNS:
        tab[key] = columns[key]
    if update_refresh_timestamp:
        method["method_metadata"]["data_refreshed"] = refreshed_at
    _set_revisions(method)
    _validate_complete(method)
    return method


def apply_owned_patch(
    base: Mapping[str, Any], patch: Mapping[str, Any], *, timestamp: Any = None
) -> dict[str, Any]:
    """Rebase Cape Cod-owned edits onto the newest embedded derived snapshots."""

    method = normalize_cape_cod_method(base, require_complete=False, timestamp=timestamp)
    incoming = normalize_cape_cod_method(patch, require_complete=False, timestamp=timestamp)
    old_tab = method["method_tab"]
    incoming_tab = incoming["method_tab"]
    method["details_tab"] = deepcopy(incoming["details_tab"])
    for name_key, values_keys, revision_key in (
        ("latest_dataset", ("latest_values",), "latest_source_revision"),
        ("exposure_dataset", ("exposure_values",), "exposure_source_revision"),
        (
            "prior_ultimate_dataset",
            ("prior_ultimate_values", "prior_ultimate_percentage_developed"),
            "prior_ultimate_source_revision",
        ),
    ):
        new_name = incoming_tab[name_key]
        if _clean(new_name).casefold() != _clean(old_tab.get(name_key)).casefold():
            for values_key in values_keys:
                old_tab[values_key] = []
            old_tab[revision_key] = ""
        old_tab[name_key] = new_name
    old_tab["prior_ultimate_mode"] = incoming_tab["prior_ultimate_mode"]
    old_tab["auto_trend_fit"] = incoming_tab["auto_trend_fit"]
    old_tab["decay_factor"] = incoming_tab["decay_factor"]
    old_tab["scaling_type"] = incoming_tab["scaling_type"]
    old_tab["alternative_ultimate_calculation"] = incoming_tab["alternative_ultimate_calculation"]
    if not incoming_tab["auto_trend_fit"]:
        old_tab["trend_rate"] = incoming_tab["trend_rate"]
    base_origins = _labels(old_tab.get("origin_labels"))
    patch_origins = _labels(incoming_tab.get("origin_labels"))
    submitted = dict(zip(patch_origins, incoming_tab.get("trend_factor_overrides") or []))
    current = dict(zip(base_origins, old_tab.get("trend_factor_overrides") or []))
    old_tab["trend_factor_overrides"] = [
        _number(submitted.get(label, current.get(label))) for label in base_origins
    ]
    row_count = len(base_origins)
    old_tab["latest_values"] = _fit(_numbers(old_tab.get("latest_values")), row_count, None)
    old_tab["exposure_values"] = _fit(_numbers(old_tab.get("exposure_values")), row_count, None)
    old_tab["prior_ultimate_values"] = _fit(_numbers(old_tab.get("prior_ultimate_values")), row_count, None)
    columns = _calculate_columns(old_tab)
    old_tab["trend_rate"] = columns["trend_rate"]
    old_tab["trend_factor_overrides"] = columns["trend_factor_overrides"]
    for key in CC_DERIVED_COLUMNS:
        old_tab[key] = columns[key]
    method["method_metadata"]["last_modified"] = _timestamp(timestamp)
    _set_revisions(method)
    return method


def cape_cod_precedent_names(payload: Mapping[str, Any]) -> list[str]:
    method = _tab(payload, "method_tab")
    raw = [
        method.get("latest_dataset"),
        method.get("exposure_dataset"),
        method.get("prior_ultimate_dataset"),
    ]
    names: list[str] = []
    seen: set[str] = set()
    for value in raw:
        name = _clean(value)
        key = name.casefold()
        if name and key not in seen:
            seen.add(key)
            names.append(name)
    return names


def cape_cod_output_variants(
    payload: Mapping[str, Any],
) -> dict[int, list[float | int | None]]:
    """Return the native and supported 3/6/12-period Cape Cod output variants."""

    method = normalize_cape_cod_method(payload, require_complete=True)
    details = method["details_tab"]
    tab = method["method_tab"]
    base_length = _integer(details.get("origin_length"), 12, minimum=1)
    values = _numbers(tab.get("cape_cod_ultimate"))
    variants = {base_length: values}
    for target_length in (3, 6, 12):
        if target_length <= base_length or target_length % base_length:
            continue
        aggregate = aggregate_vector_values(values, tab["origin_labels"], base_length, target_length)
        if aggregate:
            variants[target_length] = aggregate
    return variants


def cape_cod_ultimates_triangle(
    payload: Mapping[str, Any],
    latest_triangle_values: list[Any],
) -> list[list[float | int | None]]:
    """Return the as-if diagnostic ultimates triangle (ResQ Ultimates tab).

    ``latest_triangle_values`` holds one row per origin (oldest first) of the
    observed cumulative Latest values on a regular triangle: row ``i`` must
    contain exactly ``n - i`` cells so that cells with equal ``origin + column``
    share a calendar diagonal.  Every cell is re-estimated with the method's
    current exposure, percentage-developed, decay, and trend-rate settings,
    restricted to the origins that exist on that cell's diagonal — verified
    cell-exact against ResQ ``UltimateTriangleValues`` at every stored point.
    """

    method = normalize_cape_cod_method(payload, require_complete=True)
    tab = method["method_tab"]
    origins = _labels(tab.get("origin_labels"))
    row_count = len(origins)
    rows = latest_triangle_values if isinstance(latest_triangle_values, list) else []
    if len(rows) != row_count or any(
        not isinstance(row, list) or len(row) != row_count - index
        for index, row in enumerate(rows)
    ):
        raise CapeCodContractError(
            "Cape Cod ultimates triangle requires one Latest row per origin with "
            "n - origin_index cells (a regular triangle)."
        )
    exposure = _fit(_numbers(tab.get("exposure_values")), row_count, None)
    percentages = _fit(_numbers(tab.get("percentage_developed")), row_count, None)
    decay_value = _finite(tab.get("decay_factor"))
    decay = decay_value if decay_value is not None else 0.0
    trend_rate = _finite(tab.get("trend_rate")) or 0.0
    alternative = bool(tab.get("alternative_ultimate_calculation", False))

    result: list[list[float | int | None]] = [[None] * len(rows[index]) for index in range(row_count)]
    for diagonal in range(1, row_count + 1):
        cells: list[tuple[int, float | None, float | None]] = []
        for origin in range(diagonal):
            column = diagonal - origin  # 1-based development column
            latest_value = _finite(_number(rows[origin][column - 1]))
            # The pattern for development column k is the current Method-tab
            # percentage developed of the origin whose leading diagonal sits in
            # column k (both share the same development age on a regular grid).
            percentage = _finite(percentages[row_count - column])
            cells.append((origin, latest_value, percentage))
        newest = diagonal - 1
        usable: list[tuple[int, float, float, float]] = []
        for origin, latest_value, percentage in cells:
            exposure_value = _finite(exposure[origin])
            if latest_value is None or percentage is None or exposure_value is None:
                continue
            developed = exposure_value * percentage
            if developed == 0:
                continue
            factor = (1.0 + trend_rate) ** (newest - origin)
            usable.append((origin, developed, factor * latest_value / developed, factor))
        for origin, latest_value, percentage in cells:
            exposure_value = _finite(exposure[origin])
            if latest_value is None or percentage is None or exposure_value is None:
                continue
            numerator = 0.0
            denominator = 0.0
            for other, developed, ratio, _factor in usable:
                weight = developed * decay ** abs(origin - other)
                numerator += weight * ratio
                denominator += weight
            if denominator == 0:
                continue
            factor = (1.0 + trend_rate) ** (newest - origin)
            if factor == 0:
                continue
            detrended = (numerator / denominator) / factor
            developed_exposure = exposure_value * percentage
            if alternative and latest_value != 0 and percentage == 0:
                ultimate = detrended * exposure_value
            else:
                ultimate = latest_value + (exposure_value - developed_exposure) * detrended
            result[origin][diagonal - origin - 1] = _number(ultimate)
    return result


def build_cape_cod_output_sidecar(
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
    """Build the canonical parsed payload for a Cape Cod output sidecar."""

    method = normalize_cape_cod_method(payload, require_complete=True, timestamp=timestamp)
    prior = existing if isinstance(existing, Mapping) else {}
    record_exists = bool(prior) if existing_record is None else bool(existing_record)
    details = method["details_tab"]
    tab = method["method_tab"]
    metadata = method["method_metadata"]
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
    return validate_sidecar_core({
        "json_format": DATASET_SIDECAR_JSON_FORMAT,
        "dataset_name": details["name"],
        "dataset_type": details["output_type"] or details["name"],
        "dataset_category": details.get("dataset_category", ""),
        "reserving_class": _clean(reserving_class),
        "project_name": _clean(project_name),
        "source_kind": CC_SOURCE_KIND,
        "calculated": True,
        "method_name": details["name"],
        "method_type": CC_METHOD_TYPE,
        "data_format": "Vector",
        "period_length": details["origin_length"],
        # A method output is produced at its own origin period, so the
        # vector it publishes is stored at that period too.
        **stored_length_fields("Vector", details["origin_length"]),
        "transposed": False,
        "show_subtotal": normalize_show_subtotal(prior.get("show_subtotal")),
        "number_format": _clean(prior.get("number_format")) or "#,##0",
        "decimal_places": details["statistic_decimal_places"],
        "csv_file": _clean(csv_file),
        "notes": sidecar_notes,
        "origin_labels": deepcopy(tab["origin_labels"]),
        "development_labels": ["Ultimate"],
        "precedents": dependency_entries(cape_cod_precedent_names(method)),
        "dependents": dependency_entries(prior.get("dependents") if dependents is None else dependents),
        "created": created,
        "updated_at": published_at,
        "modified_by": actor,
        "status": _integer(status, 0, minimum=0),
        "publication_revision": metadata["publication_revision"],
        "audit_log": audits,
    })


__all__ = [
    "CC_DERIVED_COLUMNS",
    "CC_JSON_FORMAT",
    "CC_METHOD_TYPE",
    "CC_METHOD_TYPE_CODE",
    "CC_PRIOR_ULTIMATE_MODES",
    "CC_SCALING_TYPES",
    "CC_SOURCE_KIND",
    "CapeCodContractError",
    "CcContractError",
    "apply_owned_patch",
    "build_cape_cod_output_sidecar",
    "cape_cod_output_variants",
    "cape_cod_precedent_names",
    "cape_cod_ultimates_triangle",
    "derived_projection",
    "fit_cape_cod_trend_rate",
    "method_revisions",
    "normalize_cape_cod_method",
    "owned_projection",
    "publication_projection",
    "recalculate_cape_cod_method",
]
