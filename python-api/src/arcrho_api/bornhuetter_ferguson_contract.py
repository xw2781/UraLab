"""Canonical, filesystem-free contract for Bornhuetter Ferguson methods.

All persisted-data producers delegate BF normalization, calculations, revision
hashes, output variants, and output-sidecar projection to this module.  A
current v3 method therefore contains every value needed to open the Method and
Chart tabs without reading its precedent datasets.
"""

from __future__ import annotations

from copy import deepcopy
from datetime import datetime, timezone
from typing import Any, Iterable, Mapping

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


BF_JSON_FORMAT = "arcrho-bornhuetter-ferguson-v4"
BF_METHOD_TYPE = "Bornhuetter Ferguson"
BF_SOURCE_KIND = "bornhuetter_ferguson"
BF_METHOD_TYPE_CODE = 2


class BornhuetterFergusonContractError(ValueError):
    """Raised when a BF payload cannot satisfy the canonical v3 contract."""


# Compact alias for callers that prefer the abbreviation used by the UI.
BfContractError = BornhuetterFergusonContractError


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


def _dfm_snapshot_revision(
    name: str, origin_labels: list[str], ultimates: list[Any], percentages: list[Any]
) -> str:
    """Fingerprint both vectors the DFM precedent supplies.

    The percentage developed is not implied by the ultimate -- an origin whose
    latest observation is zero has a zero ultimate under any pattern -- so a
    revision covering only the ultimate could miss a changed selection.
    """

    if not name or not origin_labels or len(origin_labels) != len(ultimates) \
            or len(origin_labels) != len(percentages):
        return ""
    return fingerprint({
        "name": name,
        "origin_labels": origin_labels,
        "values": [_number(value) for value in ultimates],
        "percentage_developed": [_number(value) for value in percentages],
    })


def _calculate_vectors(method_tab: Mapping[str, Any]) -> tuple[list[Any], list[Any]]:
    origins = _labels(method_tab.get("origin_labels"))
    latest = _fit(_numbers(method_tab.get("latest_values")), len(origins), None)
    percentages = _fit(_numbers(method_tab.get("percentage_developed")), len(origins), None)
    priors = method_tab.get("prior_datasets") if isinstance(method_tab.get("prior_datasets"), list) else []
    selected: list[Any] = []
    ultimates: list[Any] = []
    for index in range(len(origins)):
        latest_value = latest[index]
        percentage_raw = percentages[index]
        numerator = 0.0
        denominator = 0.0
        for prior in priors:
            if not isinstance(prior, Mapping):
                continue
            values = prior.get("values") if isinstance(prior.get("values"), list) else []
            weights = prior.get("weights") if isinstance(prior.get("weights"), list) else []
            value = _number(values[index] if index < len(values) else None)
            weight = _number(weights[index] if index < len(weights) else 0)
            weight = max(0.0, float(weight)) if weight is not None else 0.0
            if value is None or weight <= 0:
                continue
            numerator += float(value) * weight
            denominator += weight
        selected_raw = numerator / denominator if denominator > 0 else None
        selected_value = _number(selected_raw)
        if latest_value is None:
            ultimate = None
        elif selected_value is None:
            ultimate = latest_value
        elif percentage_raw is None:
            ultimate = None
        else:
            # The ultimate keeps its fraction at the same six decimals as every
            # other BF vector. ResQ never rounds it, and a whole-number ultimate
            # drifts every dependent that divides by it -- a claim-count BF feeding
            # a Berquist-Sherman settlement-rate adjustment moved the adjusted
            # triangle's earliest ages by up to 2%.
            ultimate = _number(
                float(latest_value) + (1.0 - float(percentage_raw)) * float(selected_raw)
            )
        selected.append(selected_value)
        ultimates.append(ultimate)
    return selected, ultimates


def owned_projection(payload: Mapping[str, Any]) -> dict[str, Any]:
    details = _tab(payload, "details_tab")
    method = _tab(payload, "method_tab")
    priors = method.get("prior_datasets") if isinstance(method.get("prior_datasets"), list) else []
    return {
        "details_tab": deepcopy(details),
        "method_tab": {
            "latest_dataset": method.get("latest_dataset", ""),
            "dfm_dataset": method.get("dfm_dataset", ""),
            "show_weights": method.get("show_weights", True),
            "show_effective_weights": method.get("show_effective_weights", False),
            "prior_datasets": [
                {
                    "name": prior.get("name", ""),
                    "weights": deepcopy(prior.get("weights") or []),
                }
                for prior in priors
                if isinstance(prior, Mapping)
            ],
        },
    }


def derived_projection(payload: Mapping[str, Any]) -> dict[str, Any]:
    method = _tab(payload, "method_tab")
    priors = method.get("prior_datasets") if isinstance(method.get("prior_datasets"), list) else []
    return {
        "origin_labels": deepcopy(method.get("origin_labels") or []),
        "latest_values": deepcopy(method.get("latest_values") or []),
        "latest_source_revision": method.get("latest_source_revision", ""),
        "dfm_ultimate_values": deepcopy(method.get("dfm_ultimate_values") or []),
        "dfm_source_revision": method.get("dfm_source_revision", ""),
        "prior_datasets": [
            {
                "name": prior.get("name", ""),
                "values": deepcopy(prior.get("values") or []),
                "source_revision": prior.get("source_revision", ""),
            }
            for prior in priors
            if isinstance(prior, Mapping)
        ],
        "percentage_developed": deepcopy(method.get("percentage_developed") or []),
        "selected_prior_values": deepcopy(method.get("selected_prior_values") or []),
        "new_ultimate": deepcopy(method.get("new_ultimate") or []),
    }


def publication_projection(payload: Mapping[str, Any]) -> dict[str, Any]:
    details = _tab(payload, "details_tab")
    method = _tab(payload, "method_tab")
    return {
        "dataset_name": details.get("name", ""),
        "dataset_type": details.get("output_type", ""),
        "dataset_category": details.get("dataset_category", ""),
        "origin_length": details.get("origin_length", 12),
        "statistic_decimal_places": details.get("statistic_decimal_places", 1),
        "origin_labels": deepcopy(method.get("origin_labels") or []),
        "new_ultimate": deepcopy(method.get("new_ultimate") or []),
    }


def method_revisions(payload: Mapping[str, Any]) -> dict[str, str]:
    """Return deterministic revisions for BF-owned, derived, and published state."""

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
    if json_format not in {"", BF_JSON_FORMAT}:
        raise BornhuetterFergusonContractError(f"Unsupported BF JSON format: {json_format!r}.")
    stamped = deepcopy(dict(payload))
    stamped["json_format"] = BF_JSON_FORMAT
    return stamped


def _flat_vector(value: Any, expected: int, label: str) -> None:
    if not isinstance(value, list) or len(value) != expected:
        raise BornhuetterFergusonContractError(
            f"BF {label} must contain exactly one scalar per origin label."
        )
    if any(isinstance(item, (list, tuple, dict)) for item in value):
        raise BornhuetterFergusonContractError(f"BF {label} must be a flat scalar vector.")


def _validate_raw_complete_shape(payload: Mapping[str, Any], row_count: int) -> None:
    method = _tab(payload, "method_tab")
    for key in (
        "latest_values",
        "dfm_ultimate_values",
        "percentage_developed",
        "selected_prior_values",
        "new_ultimate",
    ):
        _flat_vector(method.get(key), row_count, key)
    priors = method.get("prior_datasets")
    if not isinstance(priors, list) or not priors:
        raise BornhuetterFergusonContractError("BF requires at least one Prior dataset.")
    for prior in priors:
        if not isinstance(prior, Mapping):
            raise BornhuetterFergusonContractError("Every BF Prior dataset must be an object.")
        _flat_vector(prior.get("values"), row_count, "Prior values")
        _flat_vector(prior.get("weights"), row_count, "Prior weights")


def normalize_bornhuetter_ferguson_method(
    payload: Mapping[str, Any],
    *,
    require_complete: bool = True,
    timestamp: Any = None,
) -> dict[str, Any]:
    """Return the exact canonical, self-contained BF v3 payload."""

    if not isinstance(payload, Mapping):
        raise BornhuetterFergusonContractError("BF method payload must be a JSON object.")
    json_format = str(payload.get("json_format") or "").strip()
    if json_format != BF_JSON_FORMAT:
        raise BornhuetterFergusonContractError(f"Unsupported BF JSON format: {json_format!r}.")

    details_source = _tab(payload, "details_tab")
    method_source = _tab(payload, "method_tab")
    metadata_source = _tab(payload, "method_metadata")
    origins = _labels(method_source.get("origin_labels"))
    row_count = len(origins)
    if require_complete:
        _validate_raw_complete_shape(payload, row_count)
    latest_name = _clean(method_source.get("latest_dataset"))
    dfm_name = _clean(method_source.get("dfm_dataset"))
    latest_values = _fit(_numbers(method_source.get("latest_values")), row_count, None)
    dfm_values = _fit(_numbers(method_source.get("dfm_ultimate_values")), row_count, None)
    dfm_percentages = _fit(_numbers(method_source.get("percentage_developed")), row_count, None)

    raw_priors = method_source.get("prior_datasets") if isinstance(method_source.get("prior_datasets"), list) else []
    if not raw_priors and _clean(method_source.get("prior_dataset")):
        raw_priors = [{
            "name": method_source.get("prior_dataset"),
            "values": method_source.get("prior_ultimate_values"),
            "weights": [],
        }]
    priors: list[dict[str, Any]] = []
    for raw in raw_priors:
        if not isinstance(raw, Mapping):
            continue
        name = _clean(raw.get("name"))
        values = _fit(_numbers(raw.get("values")), row_count, None)
        weights = _fit(_numbers(raw.get("weights")), row_count, 1.0)
        weights = [_number(max(0.0, float(value))) if value is not None else 1 for value in weights]
        priors.append({
            "name": name,
            "values": values,
            "weights": weights,
            "source_revision": _snapshot_revision(name, origins, values),
        })

    default_time = _timestamp(timestamp)
    last_modified = str(metadata_source.get("last_modified") or "").strip() or default_time
    data_refreshed = str(metadata_source.get("data_refreshed") or "").strip() or last_modified
    normalized = {
        "json_format": BF_JSON_FORMAT,
        "details_tab": {
            "name": _clean(details_source.get("name")),
            "method_type": BF_METHOD_TYPE,
            "output_type": _clean(details_source.get("output_type")),
            "dataset_category": _clean(details_source.get("dataset_category")),
            "origin_length": _integer(details_source.get("origin_length"), 12, minimum=1),
            "statistic_decimal_places": _integer(
                details_source.get("statistic_decimal_places"), 1, minimum=0, maximum=8
            ),
        },
        "method_tab": {
            "latest_dataset": latest_name,
            "latest_values": latest_values,
            "latest_source_revision": _snapshot_revision(latest_name, origins, latest_values),
            "dfm_dataset": dfm_name,
            "dfm_ultimate_values": dfm_values,
            "dfm_source_revision": _dfm_snapshot_revision(dfm_name, origins, dfm_values, dfm_percentages),
            "show_weights": method_source.get("show_weights") is not False,
            "show_effective_weights": bool(method_source.get("show_effective_weights", False)),
            "prior_datasets": priors,
            "origin_labels": origins,
            "percentage_developed": dfm_percentages,
            "selected_prior_values": _fit(_numbers(method_source.get("selected_prior_values")), row_count, None),
            "new_ultimate": _fit(_numbers(method_source.get("new_ultimate")), row_count, None),
        },
        "method_metadata": {
            "method_type": BF_METHOD_TYPE,
            "source_kind": BF_SOURCE_KIND,
            "last_modified": last_modified,
            "data_refreshed": data_refreshed,
            "owned_revision": "",
            "derived_revision": "",
            "publication_revision": "",
        },
    }
    _set_revisions(normalized)
    if require_complete:
        _validate_complete(normalized)
    return normalized


def _validate_complete(payload: Mapping[str, Any]) -> None:
    details = _tab(payload, "details_tab")
    method = _tab(payload, "method_tab")
    for key in ("name", "output_type"):
        if not _clean(details.get(key)):
            raise BornhuetterFergusonContractError(f"BF details_tab.{key} is required.")
    if _integer(details.get("origin_length"), 0) not in {1, 3, 6, 12}:
        raise BornhuetterFergusonContractError("BF origin_length must be 1, 3, 6, or 12 months.")
    origins = _labels(method.get("origin_labels"))
    if not origins or any(not label for label in origins):
        raise BornhuetterFergusonContractError("BF origin_labels must be non-empty.")
    duplicates = _duplicates(origins)
    if duplicates:
        raise BornhuetterFergusonContractError("BF origin_labels must be unique: " + ", ".join(duplicates))
    for name_key, values_key, revision_key in (
        ("latest_dataset", "latest_values", "latest_source_revision"),
        ("dfm_dataset", "dfm_ultimate_values", "dfm_source_revision"),
    ):
        if not _clean(method.get(name_key)):
            raise BornhuetterFergusonContractError(f"BF method_tab.{name_key} is required.")
        if len(method.get(values_key) or []) != len(origins):
            raise BornhuetterFergusonContractError(f"BF {values_key} must align to origin_labels.")
        if not _clean(method.get(revision_key)):
            raise BornhuetterFergusonContractError(f"BF {revision_key} is required.")
    priors = method.get("prior_datasets") if isinstance(method.get("prior_datasets"), list) else []
    if not priors:
        raise BornhuetterFergusonContractError("BF requires at least one Prior dataset.")
    prior_keys: set[str] = set()
    for prior in priors:
        if not isinstance(prior, Mapping) or not _clean(prior.get("name")):
            raise BornhuetterFergusonContractError("Every BF Prior dataset requires a name.")
        prior_key = _clean(prior.get("name")).casefold()
        if prior_key in prior_keys:
            raise BornhuetterFergusonContractError("BF Prior dataset names must be unique.")
        prior_keys.add(prior_key)
        if len(prior.get("values") or []) != len(origins) or len(prior.get("weights") or []) != len(origins):
            raise BornhuetterFergusonContractError("BF Prior values and weights must align to origin_labels.")
        if not _clean(prior.get("source_revision")):
            raise BornhuetterFergusonContractError("Every BF Prior dataset requires a source_revision.")
    if len(method.get("percentage_developed") or []) != len(origins):
        raise BornhuetterFergusonContractError("BF percentage_developed must align to origin_labels.")
    expected = _calculate_vectors(method)
    for key, values in zip(("selected_prior_values", "new_ultimate"), expected):
        if method.get(key) != values:
            raise BornhuetterFergusonContractError(f"BF {key} does not match the embedded source snapshots.")


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
        raise BornhuetterFergusonContractError("BF source snapshot has duplicate origins: " + ", ".join(duplicates))
    if set(labels) != set(origins) or len(labels) != len(origins):
        raise BornhuetterFergusonContractError("BF source snapshot origins must match the Latest origins exactly.")
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
    """Align the DFM precedent's percentage-developed pattern onto BF origins."""

    labels = _labels(snapshot.get("origin_labels"))
    values = _fit(_numbers(snapshot.get("percentage_developed")), len(labels), None)
    return _align_by_labels(labels, values, origins)


def _source_snapshot(
    source_snapshots: Mapping[str, Any], role: str, name: str, index: int = 0
) -> Mapping[str, Any] | None:
    direct = source_snapshots.get(role)
    if role == "priors":
        if isinstance(direct, list) and index < len(direct) and isinstance(direct[index], Mapping):
            return direct[index]
        if isinstance(direct, Mapping):
            value = direct.get(name)
            if isinstance(value, Mapping):
                return value
    elif isinstance(direct, Mapping):
        return direct
    by_name = source_snapshots.get(name)
    return by_name if isinstance(by_name, Mapping) else None


def recalculate_bornhuetter_ferguson_method(
    payload: Mapping[str, Any],
    *,
    source_snapshots: Mapping[str, Any] | None = None,
    changed_precedents: Iterable[str] = (),
    timestamp: Any = None,
    update_refresh_timestamp: bool | None = None,
) -> dict[str, Any]:
    """Refresh BF-derived state from optional aggregate source snapshots."""

    changed = tuple(str(item) for item in changed_precedents)
    refreshed_at = _timestamp(timestamp)
    if update_refresh_timestamp is None:
        update_refresh_timestamp = source_snapshots is not None or bool(changed)
    method = normalize_bornhuetter_ferguson_method(
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
            raise BornhuetterFergusonContractError("BF Latest snapshot identity does not match its configured source.")
        new_origins, _ = _snapshot_vector(latest_snapshot, latest=True)
        if not new_origins or any(not label for label in new_origins) or _duplicates(new_origins):
            raise BornhuetterFergusonContractError("BF Latest snapshot requires unique non-empty origins.")
        old_origins = _labels(tab.get("origin_labels"))
        old_weight_maps = [
            dict(zip(old_origins, prior.get("weights") or []))
            for prior in tab.get("prior_datasets", [])
            if isinstance(prior, Mapping)
        ]
        tab["origin_labels"] = new_origins
        for index, prior in enumerate(tab.get("prior_datasets", [])):
            weights = old_weight_maps[index] if index < len(old_weight_maps) else {}
            prior["weights"] = [_number(weights.get(label, 1)) for label in new_origins]

    origins = _labels(tab.get("origin_labels"))
    if latest_snapshot is not None:
        tab["latest_values"] = _aligned_snapshot(latest_snapshot, origins, latest=True)
    dfm_snapshot = _source_snapshot(snapshots, "dfm", tab["dfm_dataset"])
    if dfm_snapshot is not None:
        snapshot_name = _clean(dfm_snapshot.get("name"))
        if not tab["dfm_dataset"] and snapshot_name:
            tab["dfm_dataset"] = snapshot_name
        elif snapshot_name and snapshot_name.casefold() != tab["dfm_dataset"].casefold():
            raise BornhuetterFergusonContractError("BF DFM snapshot identity does not match its configured source.")
        tab["dfm_ultimate_values"] = _aligned_snapshot(dfm_snapshot, origins, latest=False)
        tab["percentage_developed"] = _aligned_percentages(dfm_snapshot, origins)
    for index, prior in enumerate(tab.get("prior_datasets", [])):
        snapshot = _source_snapshot(snapshots, "priors", _clean(prior.get("name")), index)
        if snapshot is not None:
            snapshot_name = _clean(snapshot.get("name"))
            if not prior["name"] and snapshot_name:
                prior["name"] = snapshot_name
            elif snapshot_name and snapshot_name.casefold() != prior["name"].casefold():
                raise BornhuetterFergusonContractError("BF Prior snapshot identity does not match its configured source.")
            prior["values"] = _aligned_snapshot(snapshot, origins, latest=False)

    tab["latest_values"] = _fit(_numbers(tab.get("latest_values")), len(origins), None)
    tab["dfm_ultimate_values"] = _fit(_numbers(tab.get("dfm_ultimate_values")), len(origins), None)
    tab["percentage_developed"] = _fit(_numbers(tab.get("percentage_developed")), len(origins), None)
    tab["latest_source_revision"] = _snapshot_revision(tab["latest_dataset"], origins, tab["latest_values"])
    tab["dfm_source_revision"] = _dfm_snapshot_revision(
        tab["dfm_dataset"], origins, tab["dfm_ultimate_values"], tab["percentage_developed"]
    )
    for prior in tab.get("prior_datasets", []):
        prior["values"] = _fit(_numbers(prior.get("values")), len(origins), None)
        prior["weights"] = [
            _number(max(0.0, float(value))) if value is not None else 1
            for value in _fit(_numbers(prior.get("weights")), len(origins), 1.0)
        ]
        prior["source_revision"] = _snapshot_revision(prior["name"], origins, prior["values"])
    selected, ultimates = _calculate_vectors(tab)
    tab["selected_prior_values"] = selected
    tab["new_ultimate"] = ultimates
    if update_refresh_timestamp:
        method["method_metadata"]["data_refreshed"] = refreshed_at
    _set_revisions(method)
    _validate_complete(method)
    return method


def apply_owned_patch(
    base: Mapping[str, Any], patch: Mapping[str, Any], *, timestamp: Any = None
) -> dict[str, Any]:
    """Rebase BF-owned edits onto the newest embedded derived snapshots."""

    method = normalize_bornhuetter_ferguson_method(base, require_complete=False, timestamp=timestamp)
    incoming = normalize_bornhuetter_ferguson_method(patch, require_complete=False, timestamp=timestamp)
    old_tab = method["method_tab"]
    incoming_tab = incoming["method_tab"]
    old_sources = {
        _clean(prior.get("name")).casefold(): prior
        for prior in old_tab.get("prior_datasets", [])
        if isinstance(prior, Mapping)
    }
    method["details_tab"] = deepcopy(incoming["details_tab"])
    old_tab["show_weights"] = incoming_tab["show_weights"]
    old_tab["show_effective_weights"] = incoming_tab["show_effective_weights"]
    for name_key, values_keys, revision_key in (
        ("latest_dataset", ("latest_values",), "latest_source_revision"),
        ("dfm_dataset", ("dfm_ultimate_values", "percentage_developed"), "dfm_source_revision"),
    ):
        new_name = incoming_tab[name_key]
        if _clean(new_name).casefold() != _clean(old_tab.get(name_key)).casefold():
            for values_key in values_keys:
                old_tab[values_key] = []
            old_tab[revision_key] = ""
        old_tab[name_key] = new_name
    rebased_priors: list[dict[str, Any]] = []
    base_origins = _labels(old_tab.get("origin_labels"))
    patch_origins = _labels(incoming_tab.get("origin_labels"))
    for prior in incoming_tab.get("prior_datasets", []):
        name = _clean(prior.get("name"))
        old = old_sources.get(name.casefold(), {})
        submitted_weights = dict(zip(patch_origins, prior.get("weights") or []))
        current_weights = dict(zip(base_origins, old.get("weights") or []))
        rebased_priors.append({
            "name": name,
            "values": deepcopy(old.get("values") or []),
            "weights": [
                deepcopy(submitted_weights.get(label, current_weights.get(label, 1.0)))
                for label in base_origins
            ],
            "source_revision": str(old.get("source_revision") or ""),
        })
    old_tab["prior_datasets"] = rebased_priors
    selected, ultimates = _calculate_vectors(old_tab)
    old_tab["selected_prior_values"] = selected
    old_tab["new_ultimate"] = ultimates
    method["method_metadata"]["last_modified"] = _timestamp(timestamp)
    _set_revisions(method)
    return method


def bornhuetter_ferguson_precedent_names(payload: Mapping[str, Any]) -> list[str]:
    method = _tab(payload, "method_tab")
    raw = [method.get("latest_dataset"), method.get("dfm_dataset")]
    raw.extend(
        prior.get("name")
        for prior in method.get("prior_datasets", [])
        if isinstance(prior, Mapping)
    )
    names: list[str] = []
    seen: set[str] = set()
    for value in raw:
        name = _clean(value)
        key = name.casefold()
        if name and key not in seen:
            seen.add(key)
            names.append(name)
    return names


def bornhuetter_ferguson_output_variants(
    payload: Mapping[str, Any],
) -> dict[int, list[float | int | None]]:
    """Return the native and supported 3/6/12-period BF output variants."""

    method = normalize_bornhuetter_ferguson_method(payload, require_complete=True)
    details = method["details_tab"]
    tab = method["method_tab"]
    base_length = _integer(details.get("origin_length"), 12, minimum=1)
    values = _numbers(tab.get("new_ultimate"))
    variants = {base_length: values}
    for target_length in (3, 6, 12):
        if target_length <= base_length or target_length % base_length:
            continue
        aggregate = aggregate_vector_values(values, tab["origin_labels"], base_length, target_length)
        if aggregate:
            variants[target_length] = aggregate
    return variants


def build_bornhuetter_ferguson_output_sidecar(
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
    """Build the canonical parsed payload for a BF output sidecar."""

    method = normalize_bornhuetter_ferguson_method(payload, require_complete=True, timestamp=timestamp)
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
        "source_kind": BF_SOURCE_KIND,
        "calculated": True,
        "method_name": details["name"],
        "method_type": BF_METHOD_TYPE,
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
        "precedents": dependency_entries(bornhuetter_ferguson_precedent_names(method)),
        "dependents": dependency_entries(prior.get("dependents") if dependents is None else dependents),
        "created": created,
        "updated_at": published_at,
        "modified_by": actor,
        "status": _integer(status, 0, minimum=0),
        "publication_revision": metadata["publication_revision"],
        "audit_log": audits,
    })


__all__ = [
    "BF_JSON_FORMAT",
    "BF_METHOD_TYPE",
    "BF_METHOD_TYPE_CODE",
    "BF_SOURCE_KIND",
    "BfContractError",
    "BornhuetterFergusonContractError",
    "apply_owned_patch",
    "bornhuetter_ferguson_output_variants",
    "bornhuetter_ferguson_precedent_names",
    "build_bornhuetter_ferguson_output_sidecar",
    "derived_projection",
    "method_revisions",
    "normalize_bornhuetter_ferguson_method",
    "owned_projection",
    "publication_projection",
    "recalculate_bornhuetter_ferguson_method",
]
