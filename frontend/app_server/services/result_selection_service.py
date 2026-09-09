"""Persist, load, and eagerly refresh Result Selection methods."""
from __future__ import annotations

import getpass
import hashlib
import json
import math
import os
import re
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Dict, Iterable, List, Tuple

import pandas as pd
from fastapi import HTTPException

from arcrho_api.io import persisted_json_text
from arcrho_api.timestamps import utc_now_text
from app_server import config
from app_server.helpers import sanitize_dataset_file_name
from app_server.services import (
    dataset_sidecar_status_service,
    dependent_propagation_service,
    precedent_cache_service,
    user_identity_service,
)


RESULT_SELECTION_JSON_FORMAT = "arcrho-result-selection-v4"
VALUE_DECIMAL_PLACES = 6
MAX_RATIO_BASIS_COUNT = 3
READ_MAX_WORKERS = 6
MAX_REFRESH_VISITS_PER_DATASET = 32
MAX_REFRESH_GRAPH_NODES = 5000
_QUANTUM = Decimal("0.000001")
_READ_EXECUTOR = ThreadPoolExecutor(
    max_workers=READ_MAX_WORKERS,
    thread_name_prefix="arcrho-rs-read",
)
_ORIGIN_MONTH_NAME = {
    "jan": 1, "january": 1, "feb": 2, "february": 2,
    "mar": 3, "march": 3, "apr": 4, "april": 4, "may": 5,
    "jun": 6, "june": 6, "jul": 7, "july": 7, "aug": 8,
    "august": 8, "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10, "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}


def _clean(value: Any) -> str:
    return str(value if value is not None else "").strip()


def _key(value: Any) -> str:
    return " ".join(_clean(value).lower().split())


def _now() -> str:
    return utc_now_text()


def _current_user_name() -> str:
    """Configured full name for the saving account, or the raw login."""
    return user_identity_service.get_current_display_name() or getpass.getuser()


def _lock(project_name: str, reserving_class: str) -> threading.RLock:
    return dataset_sidecar_status_service.reserving_class_io_lock(project_name, reserving_class)


def _method_path(project_name: str, reserving_class: str, method_name: str) -> str:
    return dataset_sidecar_status_service.method_json_path(
        project_name,
        reserving_class,
        dataset_sidecar_status_service.METHOD_TYPE_RESULT_SELECTION,
        method_name,
    )


def _sidecar_path(project_name: str, reserving_class: str, dataset_name: str) -> str:
    return dataset_sidecar_status_service.sidecar_path(project_name, reserving_class, dataset_name)


def _read_json(path: str) -> Dict[str, Any]:
    try:
        with open(path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except FileNotFoundError:
        return {}
    except PermissionError as exc:
        raise HTTPException(423, f"Result Selection file is locked or inaccessible: {os.path.basename(path)}") from exc
    except (OSError, json.JSONDecodeError) as exc:
        raise HTTPException(500, f"Invalid Result Selection JSON: {os.path.basename(path)}: {exc}") from exc
    return payload if isinstance(payload, dict) else {}


def _json_text(payload: Dict[str, Any]) -> str:
    return persisted_json_text(payload)


def _revision_projection(payload: Dict[str, Any]) -> Dict[str, Any]:
    """The content a revision covers: everything except when it was written.

    ``last_modified`` records when a person last saved the file and
    ``data_refreshed`` when a propagation refresh last recomputed it; neither
    says what the file holds, and every other method family already leaves
    both out -- DFM, BF, CC and bootstrap all revision a projection through
    ``dfm_contract.method_revisions`` that never reaches them. Keeping them in
    meant an RPC upload recording ResQ's own save time would move the token an
    open editor holds, and that editor's next save would be refused even
    though nothing it edited had changed.
    """

    projection = dict(payload)
    metadata = projection.get("method_metadata")
    if isinstance(metadata, dict):
        projection["method_metadata"] = {
            key: value
            for key, value in metadata.items()
            if key not in _REVISION_FREE_METADATA_KEYS
        }
    return projection


# The two write stamps: when a person last saved the method, and when a
# propagation refresh last recomputed it. Neither describes the content.
_REVISION_FREE_METADATA_KEYS = frozenset({"last_modified", "data_refreshed"})


def _normalized_metadata(payload: Dict[str, Any]) -> Dict[str, Any]:
    """Keep the user-save stamp, and the refresh stamp when the file has one."""

    source = payload.get("method_metadata") if isinstance(payload.get("method_metadata"), dict) else {}
    metadata = {"last_modified": _clean(source.get("last_modified")) or _now()}
    data_refreshed = _clean(source.get("data_refreshed"))
    if data_refreshed:
        metadata["data_refreshed"] = data_refreshed
    return metadata


def _method_revision(payload: Dict[str, Any]) -> str:
    digest = hashlib.sha256(_json_text(_revision_projection(payload)).encode("utf-8")).hexdigest()
    return f"sha256:{digest}"


def _read_bytes_if_file(path: str) -> bytes | None:
    if not os.path.isfile(path):
        return None
    with open(path, "rb") as handle:
        return handle.read()


def _commit_text_files(files: Dict[str, str], *, last_paths: Iterable[str] = ()) -> None:
    last_keys = {os.path.normcase(os.path.abspath(path)) for path in last_paths}
    ordered_paths = sorted(
        files,
        key=lambda path: (
            os.path.normcase(os.path.abspath(path)) in last_keys,
            os.path.normcase(path),
        ),
    )
    ordered = [(path, files[path]) for path in ordered_paths]
    staged: Dict[str, str] = {}
    backups: Dict[str, bytes | None] = {}
    replaced: List[str] = []
    try:
        for path, value in ordered:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            backups[path] = _read_bytes_if_file(path)
            temporary = f"{path}.{uuid.uuid4()}.tmp"
            with open(temporary, "w", encoding="utf-8", newline="\n") as handle:
                handle.write(value)
            staged[path] = temporary
        for path, _value in ordered:
            dataset_sidecar_status_service.replace_staged_file(staged.pop(path), path)
            replaced.append(path)
    except Exception as exc:
        rollback_errors: List[str] = []
        for path in reversed(replaced):
            original = backups.get(path)
            try:
                if original is None:
                    if os.path.exists(path):
                        os.remove(path)
                    continue
                temporary = f"{path}.{uuid.uuid4()}.rollback"
                with open(temporary, "wb") as handle:
                    handle.write(original)
                os.replace(temporary, path)
            except OSError as rollback_exc:
                rollback_errors.append(f"{os.path.basename(path)}: {rollback_exc}")
        if rollback_errors:
            raise RuntimeError(
                f"{exc}; file rollback failed: {'; '.join(rollback_errors)}"
            ) from exc
        raise
    finally:
        for temporary in staged.values():
            try:
                os.remove(temporary)
            except OSError:
                pass


def _round_number(value: Any) -> float | int | None:
    """Return one Result Selection number at the precision it was observed with.

    Every value here is a method ultimate produced somewhere else and copied in:
    a DFM chain, a Bornhuetter-Ferguson, a Cape Cod. Quantizing the copy made
    the weighted average of several ultimates disagree with the same average
    taken in ResQ, and the published selection carried the difference on to
    everything that reads it. The number is therefore carried whole, exactly as
    the input triangle and the DFM factor chain already are.
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
    if isinstance(value, int) and not isinstance(value, bool):
        return int(number)
    return number


def _round_vector(values: Any) -> List[float | int | None]:
    return [_round_number(value) for value in values] if isinstance(values, list) else []


def _fit_vector(values: Any, row_count: int, *, fill: Any = None) -> List[Any]:
    fitted = list(values)[:row_count] if isinstance(values, list) else []
    fitted.extend([fill] * max(0, row_count - len(fitted)))
    return fitted


def _unique_names(values: Any, *, limit: int | None = None) -> List[str]:
    out: List[str] = []
    seen = set()
    for raw in values if isinstance(values, list) else []:
        name = _clean(raw)
        normalized = _key(name)
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        out.append(name)
        if limit is not None and len(out) >= limit:
            break
    return out


def _legacy_ratio_basis_names(details: Dict[str, Any]) -> List[str]:
    names = _unique_names(details.get("ratio_basis_datasets"), limit=MAX_RATIO_BASIS_COUNT)
    if names:
        return names
    fallback = _clean(details.get("ratio_basis_dataset") or details.get("ratio_basis"))
    return [fallback] if fallback else []


def _normalize_ratio_basis_values(raw: Any, names: List[str]) -> List[Dict[str, Any]]:
    if isinstance(raw, list) and raw and all(not isinstance(item, dict) for item in raw):
        raw = [{"name": names[0], "values": raw}] if names else []
    by_name: Dict[str, Dict[str, Any]] = {}
    for item in raw if isinstance(raw, list) else []:
        if not isinstance(item, dict):
            continue
        name = _clean(item.get("name"))
        normalized = _key(name)
        if normalized and normalized not in by_name:
            by_name[normalized] = {"name": name, "values": _round_vector(item.get("values"))}
    return [
        {"name": name, "values": list(by_name.get(_key(name), {}).get("values") or [])}
        for name in names
    ]


def _normalize_sources(raw: Any) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for item in raw if isinstance(raw, list) else []:
        if not isinstance(item, dict) or not _clean(item.get("name")):
            continue
        try:
            origin_length = int(item.get("origin_length") or 0)
        except (TypeError, ValueError):
            origin_length = 0
        out.append({
            "name": _clean(item.get("name")),
            "dataset_type": _clean(item.get("dataset_type")),
            "data_format": _clean(item.get("data_format")),
            "method_type": _clean(item.get("method_type")),
            "category": _clean(item.get("category")),
            "source_kind": _clean(item.get("source_kind")),
            "origin_length": origin_length if origin_length > 0 else None,
            "values": _round_vector(item.get("values")),
            "weights": [max(0.0, float(value or 0)) for value in _round_vector(item.get("weights"))],
        })
    return out


def normalize_method_payload(payload: Dict[str, Any], *, require_complete_basis: bool = True) -> Dict[str, Any]:
    if not isinstance(payload, dict):
        raise HTTPException(422, "Result Selection method payload is required.")
    details = payload.get("details_tab") if isinstance(payload.get("details_tab"), dict) else {}
    method = payload.get("method_tab") if isinstance(payload.get("method_tab"), dict) else {}
    name = _clean(details.get("name"))
    output_type = _clean(details.get("output_type"))
    if not name or not output_type:
        raise HTTPException(422, "Result Selection name and output_type are required.")
    try:
        origin_length = int(details.get("origin_length") or 12)
    except (TypeError, ValueError) as exc:
        raise HTTPException(422, "Result Selection origin_length must be positive.") from exc
    if origin_length <= 0:
        raise HTTPException(422, "Result Selection origin_length must be positive.")

    basis_names = _legacy_ratio_basis_names(details)
    active = _clean(details.get("active_ratio_basis_dataset") or details.get("ratio_basis_dataset") or details.get("ratio_basis"))
    active = next((item for item in basis_names if _key(item) == _key(active)), "")
    if basis_names and not active:
        active = basis_names[0]
    origin_labels = [str(value if value is not None else "") for value in method.get("origin_labels", [])] \
        if isinstance(method.get("origin_labels"), list) else []
    ratio_basis_values = _normalize_ratio_basis_values(method.get("ratio_basis_values"), basis_names)
    if require_complete_basis:
        for item in ratio_basis_values:
            if len(item["values"]) != len(origin_labels):
                raise HTTPException(
                    422,
                    f"Ratio Basis '{item['name']}' must contain exactly {len(origin_labels)} origin values.",
                )

    raw_decimal_places = details.get("statistic_decimal_places")
    try:
        statistic_decimal_places = max(
            0,
            min(8, int(1 if raw_decimal_places is None or raw_decimal_places == "" else raw_decimal_places)),
        )
    except (TypeError, ValueError) as exc:
        raise HTTPException(422, "Result Selection statistic_decimal_places must be an integer.") from exc

    normalized = {
        "json_format": RESULT_SELECTION_JSON_FORMAT,
        "details_tab": {
            "name": name,
            "output_type": output_type,
            "origin_length": origin_length,
            "ratio_basis_datasets": basis_names,
            "active_ratio_basis_dataset": active,
            "show_ratios_as_percentages": details.get("show_ratios_as_percentages") is not False,
            "statistic_decimal_places": statistic_decimal_places,
        },
        "method_tab": {
            "origin_labels": origin_labels,
            "show_weights": method.get("show_weights") is not False,
            "loaded_datasets": _normalize_sources(method.get("loaded_datasets")),
            "ratio_basis_values": ratio_basis_values,
            "calculated_ultimate": _round_vector(method.get("calculated_ultimate")),
            "selected_ultimate": _round_vector(method.get("selected_ultimate")),
            "ultimate_overrides": _round_vector(method.get("ultimate_overrides")),
        },
        "method_metadata": _normalized_metadata(payload),
    }
    if require_complete_basis:
        row_count = len(origin_labels)
        for source in normalized["method_tab"]["loaded_datasets"]:
            for field in ("values", "weights"):
                if len(source[field]) != row_count:
                    raise HTTPException(
                        422,
                        f"Result Selection source '{source['name']}' {field} must contain exactly {row_count} origin values.",
                    )
        for field in ("calculated_ultimate", "selected_ultimate", "ultimate_overrides"):
            if len(normalized["method_tab"][field]) != row_count:
                raise HTTPException(
                    422,
                    f"Result Selection {field} must contain exactly {row_count} origin values.",
                )
    return normalized


def _precedent_names(payload: Dict[str, Any]) -> List[str]:
    method = payload["method_tab"]
    details = payload["details_tab"]
    names = [item.get("name") for item in method.get("loaded_datasets", [])]
    names.extend(details.get("ratio_basis_datasets", []))
    return _unique_names(names)


def _weighted_ultimates(sources: List[Dict[str, Any]], row_count: int) -> List[float | int | None]:
    out: List[float | int | None] = []
    for row in range(row_count):
        numerator = 0.0
        denominator = 0.0
        for source in sources:
            values = source.get("values") if isinstance(source.get("values"), list) else []
            weights = source.get("weights") if isinstance(source.get("weights"), list) else []
            try:
                value = float(values[row])
                weight = max(0.0, float(weights[row]))
            except (IndexError, TypeError, ValueError):
                continue
            if not math.isfinite(value) or not math.isfinite(weight) or weight <= 0:
                continue
            numerator += value * weight
            denominator += weight
        out.append(_round_number(numerator / denominator) if denominator > 0 else None)
    return out


def _recalculate_method(payload: Dict[str, Any]) -> None:
    method = payload["method_tab"]
    row_count = len(method.get("origin_labels", []))
    calculated = _weighted_ultimates(method.get("loaded_datasets", []), row_count)
    overrides = list(method.get("ultimate_overrides") or [])[:row_count]
    overrides.extend([None] * (row_count - len(overrides)))
    method["ultimate_overrides"] = overrides
    method["calculated_ultimate"] = calculated
    method["selected_ultimate"] = [
        overrides[index] if overrides[index] is not None else calculated[index]
        for index in range(row_count)
    ]


def _parse_origin_start_month(label: str, base_length: int) -> Tuple[int, int] | None:
    text = _clean(label)
    if base_length == 12 and re.fullmatch(r"\d{4}", text):
        return int(text), 1
    if base_length == 6:
        match = re.fullmatch(r"(\d{4})\s*H([12])", text, re.IGNORECASE)
        if match:
            return int(match.group(1)), (int(match.group(2)) - 1) * 6 + 1
    if base_length == 3:
        match = re.fullmatch(r"(\d{4})\s*Q([1-4])", text, re.IGNORECASE)
        if match:
            return int(match.group(1)), (int(match.group(2)) - 1) * 3 + 1
    if base_length == 1:
        match = re.fullmatch(r"(\d{4})(\d{2})", text)
        if match and 1 <= int(match.group(2)) <= 12:
            return int(match.group(1)), int(match.group(2))
        match = re.fullmatch(r"([A-Za-z]{3,9})\s+(\d{4})", text)
        if match and match.group(1).lower() in _ORIGIN_MONTH_NAME:
            return int(match.group(2)), _ORIGIN_MONTH_NAME[match.group(1).lower()]
    return None


def _aggregate_vector(values: List[Any], labels: List[str], base_length: int, target_length: int) -> List[Any]:
    factor = target_length // base_length if base_length and target_length % base_length == 0 else 0
    if factor <= 1:
        return []
    if len(labels) == len(values) and base_length in {1, 3, 6, 12}:
        buckets: Dict[Tuple[int, int], List[Any]] = {}
        order: List[Tuple[int, int]] = []
        for index, value in enumerate(values):
            parsed = _parse_origin_start_month(labels[index], base_length)
            if not parsed:
                buckets = {}
                break
            year, month = parsed
            bucket = (year, ((month - 1) // target_length) * target_length + 1)
            if bucket not in buckets:
                buckets[bucket] = []
                order.append(bucket)
            buckets[bucket].append(value)
        if buckets:
            return [_sum_nullable(buckets[item]) for item in order]
    return [_sum_nullable(values[index:index + factor]) for index in range(0, len(values), factor)]


def _sum_nullable(values: List[Any]) -> float | int | None:
    numbers = [float(value) for value in values if value is not None and math.isfinite(float(value))]
    return _round_number(sum(numbers)) if numbers else None


def _csv_text(values: List[Any]) -> str:
    rows = []
    for value in values:
        number = _round_number(value)
        rows.append("" if number is None else str(number))
    return "\n".join(rows) + "\n"


def _output_files(project_name: str, reserving_class: str, payload: Dict[str, Any]) -> Dict[str, str]:
    details = payload["details_tab"]
    method = payload["method_tab"]
    name = details["name"]
    base_length = int(details["origin_length"])
    values = list(method.get("selected_ultimate") or [])
    labels = list(method.get("origin_labels") or [])
    data_dir = config.get_project_dataset_cache_dir(project_name, reserving_class)
    files = {
        os.path.join(data_dir, f"{sanitize_dataset_file_name(name)}@{base_length}.csv"): _csv_text(values),
    }
    for target in (3, 6, 12):
        if target <= base_length or target % base_length:
            continue
        aggregate = _aggregate_vector(values, labels, base_length, target)
        if aggregate:
            files[os.path.join(data_dir, f"{sanitize_dataset_file_name(name)}@{target}.csv")] = _csv_text(aggregate)
    return files


def _sidecar_response(payload: Dict[str, Any], *, exists: bool) -> Dict[str, Any]:
    if not exists:
        return {"exists": False, "audit_log": [], "notes": "", "origin_labels": []}
    vector = _clean(payload.get("data_format")).lower() == "vector"
    return {
        "exists": True,
        "dataset_name": _clean(payload.get("dataset_name")),
        "dataset_type": _clean(payload.get("dataset_type")),
        "data_format": _clean(payload.get("data_format")),
        # Display, not stored: this is what the page shows beside the output,
        # not a shape anything reads a file at.
        "origin_length": payload.get("period_length") if vector else payload.get("origin_length"),
        "origin_labels": [str(item) for item in payload.get("origin_labels", [])]
        if isinstance(payload.get("origin_labels"), list) else [],
        "notes": str(payload.get("notes") if payload.get("notes") is not None else ""),
        "status": dataset_sidecar_status_service.normalize_status(payload.get("status")),
        "audit_log": payload.get("audit_log") if isinstance(payload.get("audit_log"), list) else [],
        "updated_at": _clean(payload.get("updated_at")),
    }


def _validate_method_sidecar_pair(method_name: str, method: Dict[str, Any], sidecar: Dict[str, Any]) -> None:
    details = method["details_tab"]
    method_tab = method["method_tab"]
    if _key(details.get("name")) != _key(method_name):
        raise HTTPException(409, "Result Selection method name does not match the requested output.")
    if _key(sidecar.get("dataset_name")) != _key(method_name):
        raise HTTPException(409, "Result Selection sidecar identity does not match the method JSON.")
    method_type = dataset_sidecar_status_service.normalize_method_type(
        sidecar.get("method_type"),
        sidecar.get("source_kind"),
    )
    if method_type != dataset_sidecar_status_service.METHOD_TYPE_RESULT_SELECTION:
        raise HTTPException(409, "Result Selection sidecar does not identify a Result Selection output.")
    if _clean(sidecar.get("data_format")).lower() != "vector":
        raise HTTPException(409, "Result Selection output sidecar must use Vector data format.")
    # Stored, not displayed: the check is that the CSV this method wrote
    # holds its own periods.
    sidecar_period = precedent_cache_service.source_period(sidecar)
    if sidecar_period != int(details["origin_length"]):
        raise HTTPException(409, "Result Selection method and sidecar origin lengths do not match.")
    sidecar_labels = sidecar.get("origin_labels")
    if isinstance(sidecar_labels, list) and sidecar_labels:
        if [str(item) for item in sidecar_labels] != method_tab["origin_labels"]:
            raise HTTPException(409, "Result Selection method and sidecar origin labels do not match.")
    method_precedents = {_key(item) for item in _precedent_names(method)}
    sidecar_precedents = {
        _key(item)
        for item in dataset_sidecar_status_service.entry_names(sidecar.get("precedents"))
    }
    if method_precedents != sidecar_precedents:
        raise HTTPException(409, "Result Selection method and sidecar precedents do not match.")


def load_result_selection(
    project_name: str,
    reserving_class: str,
    method_name: str,
    *,
    include_method: bool = True,
) -> Dict[str, Any]:
    project = _clean(project_name)
    reserving = _clean(reserving_class)
    name = _clean(method_name)
    if not project or not reserving or not name:
        raise HTTPException(400, "project_name, reserving_class, and method_name are required.")
    method_path = _method_path(project, reserving, name)
    sidecar_path = _sidecar_path(project, reserving, name)
    with dataset_sidecar_status_service.sidecar_write_lock(sidecar_path):
        method_future = _READ_EXECUTOR.submit(_read_json, method_path) if include_method else None
        sidecar_future = _READ_EXECUTOR.submit(_read_json, sidecar_path)
        method = method_future.result() if method_future else {}
        sidecar = sidecar_future.result()
        upgraded = False
        if include_method and bool(method) != bool(sidecar):
            raise HTTPException(409, "Result Selection requires both its method JSON and output sidecar.")
        if include_method and method:
            json_format = _clean(method.get("json_format"))
            if json_format != RESULT_SELECTION_JSON_FORMAT:
                raise HTTPException(422, f"Unsupported Result Selection JSON format: {json_format or '(missing)' }.")
            method = normalize_method_payload(method, require_complete_basis=True)
            _validate_method_sidecar_pair(name, method, sidecar)
    return {
        "ok": True,
        "project_name": project,
        "reserving_class": reserving,
        "method_name": name,
        "method_exists": bool(method) if include_method else os.path.isfile(method_path),
        "method": method if include_method and method else None,
        "method_revision": _method_revision(method) if include_method and method else "",
        "sidecar": _sidecar_response(sidecar, exists=bool(sidecar)),
        "upgraded": upgraded,
    }


def save_result_selection(
    project_name: str,
    reserving_class: str,
    method: Dict[str, Any],
    notes: str = "",
    expected_revision: str | None = None,
) -> Dict[str, Any]:
    project = _clean(project_name)
    reserving = _clean(reserving_class)
    if not project or not reserving:
        raise HTTPException(400, "project_name and reserving_class are required.")
    # Dependent propagation runs on ArcRho Engine; block the save before any
    # write when no live Engine can pick the job up or another walk is still
    # rewriting this reserving class.
    dependent_propagation_service.require_reserving_class_writable(project, reserving)
    payload = normalize_method_payload(method, require_complete_basis=True)
    _recalculate_method(payload)
    details = payload["details_tab"]
    method_tab = payload["method_tab"]
    name = details["name"]
    method_path = _method_path(project, reserving, name)
    primary_path = os.path.join(
        config.get_project_dataset_cache_dir(project, reserving),
        f"{sanitize_dataset_file_name(name)}@{details['origin_length']}.csv",
    )
    output_sidecar_path = _sidecar_path(project, reserving, name)
    new_precedents = _precedent_names(payload)
    old_precedents: List[str] = []
    with _lock(project, reserving), dataset_sidecar_status_service.sidecar_write_lock(output_sidecar_path):
        current_method = _read_json(method_path)
        if current_method:
            current_sidecar = _read_json(output_sidecar_path)
            if _clean(current_method.get("json_format")) != RESULT_SELECTION_JSON_FORMAT:
                raise HTTPException(409, "Result Selection changed on disk; reload it before saving.")
            current_method = normalize_method_payload(current_method, require_complete_basis=True)
            old_precedents = _precedent_names(current_method)
            current_revision = _method_revision(current_method)
            if expected_revision is not None and _clean(expected_revision) != current_revision:
                raise HTTPException(409, "Result Selection changed on disk; reload the latest values before saving.")
            if dataset_sidecar_status_service.normalize_status(current_sidecar.get("status")) \
                    == dataset_sidecar_status_service.STATUS_REVIEW_NEEDED:
                _refresh_review_save_payload(project, reserving, payload)
        elif expected_revision is not None and _clean(expected_revision):
            raise HTTPException(409, "Result Selection was removed on disk; reload before saving.")
        _assert_new_precedents_do_not_cycle(project, reserving, name, new_precedents)
        output_files = _output_files(project, reserving, payload)
        previous_files = {
            path: _read_bytes_if_file(path)
            for path in [method_path, *output_files, output_sidecar_path]
        }
        _commit_text_files({method_path: _json_text(payload), **output_files})
        try:
            from app_server.services import dataset_service

            sidecar = dataset_service.save_dataset_sidecar(
                project,
                reserving,
                name,
                dataset_type=details["output_type"],
                instance_name=name,
                source_kind="result_selection",
                data_format="Vector",
                origin_length=details["origin_length"],
                development_length=details["origin_length"],
                cumulative=True,
                transposed=False,
                calendar=False,
                origin_labels=method_tab["origin_labels"],
                csv_file=os.path.basename(primary_path),
                method_type=dataset_sidecar_status_service.METHOD_TYPE_RESULT_SELECTION,
                status=dataset_sidecar_status_service.STATUS_CURRENT,
                notes=notes,
                precedents=new_precedents,
            )
        except Exception as exc:
            rollback_errors: List[str] = []
            try:
                dataset_sidecar_status_service.update_precedent_dependents(
                    project,
                    reserving,
                    name,
                    new_precedents,
                    old_precedents,
                    require_new_precedents=bool(old_precedents),
                )
            except Exception as rollback_exc:
                rollback_errors.append(f"dependency graph: {rollback_exc}")
            restore: Dict[str, str] = {}
            for path, original in previous_files.items():
                if original is None:
                    try:
                        os.remove(path)
                    except FileNotFoundError:
                        pass
                    except OSError as rollback_exc:
                        rollback_errors.append(f"{os.path.basename(path)}: {rollback_exc}")
                else:
                    restore[path] = original.decode("utf-8")
            if restore:
                try:
                    _commit_text_files(restore)
                except Exception as rollback_exc:
                    rollback_errors.append(f"persisted files: {rollback_exc}")
            if rollback_errors:
                raise RuntimeError(
                    f"{exc}; Result Selection rollback failed: {'; '.join(rollback_errors)}"
                ) from exc
            raise
    aggregate_paths = [path for path in output_files if os.path.normcase(path) != os.path.normcase(primary_path)]
    unreviewed_precedents = dataset_sidecar_status_service.review_needed_precedent_names(
        project,
        reserving,
        new_precedents,
    )
    return {
        "ok": True,
        "method": payload,
        "method_revision": _method_revision(payload),
        "method_path": method_path,
        "csv_path": primary_path,
        "aggregated_csv_paths": sorted(aggregate_paths, key=os.path.normcase),
        "sidecar": sidecar,
        "propagation_ok": bool(sidecar.get("propagation_ok", True)),
        "propagation": sidecar.get("calculated_updates"),
        "index_ok": bool(sidecar.get("index_ok", True)),
        "index_error": _clean(sidecar.get("index_error")),
        "unreviewed_precedents": unreviewed_precedents,
        "unreviewed_precedent_count": len(unreviewed_precedents),
    }


def save_propagation_roots(
    project_name: str,
    reserving_class: str,
    method: Dict[str, Any],
    notes: str = "",
    expected_revision: str | None = None,
    **_ignored: Any,
) -> List[Tuple[str, str]]:
    """Return the roots ``save_result_selection`` would propagate from.

    The two-step save plans the dependent closure before anything is written.
    Result Selection propagates through the output sidecar it writes, so the
    roots are the same ``(name, output_type)`` pair the save hands to
    ``dataset_service.save_dataset_sidecar``.
    """

    payload = normalize_method_payload(method, require_complete_basis=True)
    details = payload["details_tab"]
    name = _clean(details.get("name"))
    return [(name, _clean(details.get("output_type")) or name)]


def _dependency_values(
    project_name: str,
    reserving_class: str,
    dataset_name: str,
    sidecar: Dict[str, Any],
    origin_length: int,
    *,
    exact: bool,
) -> List[float | int | None]:
    stored_period = precedent_cache_service.source_period(sidecar)
    # A hand-entered precedent stored at a finer period is aggregated to this
    # method's own length in memory, from the CSV the sidecar names, so a
    # coarser copy left on disk by an earlier release is never read.
    rollup = bool(stored_period and stored_period != origin_length) \
        and not precedent_cache_service.rollup_reason(sidecar, origin_length)
    if rollup:
        path = precedent_cache_service.sidecar_csv_path(project_name, reserving_class, sidecar)
        if not path:
            raise RuntimeError(f"Cached dataset CSV is missing for '{dataset_name}'.")
    else:
        path = precedent_cache_service.precedent_csv_path(
            project_name,
            reserving_class,
            dataset_name,
            sidecar,
            origin_length,
            exact=exact,
        )
    try:
        frame = pd.read_csv(
            path, header=None, dtype="float64", keep_default_na=True, float_precision="round_trip"
        )
    except Exception as exc:
        raise RuntimeError(f"Unable to read '{dataset_name}': {exc}") from exc
    rows = frame.to_numpy().tolist()
    if rollup:
        try:
            rows = precedent_cache_service.rollup_rows(project_name, sidecar, rows, origin_length)
        except ValueError as exc:
            raise RuntimeError(
                f"Unable to roll '{dataset_name}' up to {origin_length} months: {exc}"
            ) from exc
    if _clean(sidecar.get("data_format")).lower() == "triangle":
        values: List[Any] = []
        for row in rows:
            value = next((item for item in reversed(row) if item is not None and not pd.isna(item)), None)
            values.append(value)
    else:
        values = [row[0] if row else None for row in rows]
    return _round_vector(values)


def _read_sidecars(project_name: str, reserving_class: str, names: Iterable[str]) -> Dict[str, Dict[str, Any]]:
    ordered = _unique_names(list(names))
    futures = {
        name: _READ_EXECUTOR.submit(_read_json, _sidecar_path(project_name, reserving_class, name))
        for name in ordered
    }
    return {name: futures[name].result() for name in ordered}


def _refresh_review_save_payload(
    project_name: str,
    reserving_class: str,
    payload: Dict[str, Any],
) -> None:
    """Rebase a review-needed save on exactly the revised incoming precedent set."""
    method = payload["method_tab"]
    details = payload["details_tab"]
    origin_length = int(details["origin_length"])
    row_count = len(method.get("origin_labels", []))
    sources = method.get("loaded_datasets", [])
    source_names = [source["name"] for source in sources]
    basis_names = list(details.get("ratio_basis_datasets", []))
    dependency_names = _unique_names([*source_names, *basis_names])
    sidecars = _read_sidecars(project_name, reserving_class, dependency_names)
    missing = [name for name in dependency_names if not sidecars.get(name)]
    if missing:
        raise HTTPException(
            409,
            "Result Selection cannot save because datasets in the revised precedent list "
            "are missing: " + ", ".join(missing)
            + ". Restore or remove those datasets, then save again.",
        )

    tasks: Dict[Tuple[str, bool], Any] = {}
    for name in source_names:
        tasks[(name, False)] = _READ_EXECUTOR.submit(
            _dependency_values,
            project_name,
            reserving_class,
            name,
            sidecars[name],
            origin_length,
            exact=False,
        )
    for name in basis_names:
        tasks[(name, True)] = _READ_EXECUTOR.submit(
            _dependency_values,
            project_name,
            reserving_class,
            name,
            sidecars[name],
            origin_length,
            exact=True,
        )

    for source in sources:
        name = source["name"]
        values = tasks[(name, False)].result()
        if len(values) != row_count:
            raise HTTPException(
                422,
                f"Result Selection source '{name}' returned {len(values)} values; expected {row_count}.",
            )
        sidecar = sidecars[name]
        source.update({
            "dataset_type": _clean(sidecar.get("dataset_type")) or source.get("dataset_type"),
            "data_format": _clean(sidecar.get("data_format")) or source.get("data_format"),
            "method_type": dataset_sidecar_status_service.normalize_method_type(
                sidecar.get("method_type"),
                sidecar.get("source_kind"),
            ),
            "category": _clean(sidecar.get("dataset_category") or sidecar.get("category"))
            or source.get("category"),
            "source_kind": _clean(sidecar.get("source_kind")) or source.get("source_kind"),
            "origin_length": precedent_cache_service.source_period(sidecar) or source.get("origin_length") or origin_length,
            "values": values,
        })
        source["weights"] = _fit_vector(source.get("weights"), row_count, fill=0.0)

    method["ratio_basis_values"] = [
        {"name": name, "values": tasks[(name, True)].result()}
        for name in basis_names
    ]
    for basis in method["ratio_basis_values"]:
        if len(basis["values"]) != row_count:
            raise HTTPException(
                422,
                f"Result Selection Ratio Basis '{basis['name']}' returned "
                f"{len(basis['values'])} values; expected {row_count}.",
            )
    _recalculate_method(payload)


def _read_sidecars_cached(
    project_name: str,
    reserving_class: str,
    names: Iterable[str],
    sidecar_snapshot: Dict[str, Dict[str, Any]],
) -> Dict[str, Dict[str, Any]]:
    ordered = _unique_names(list(names))
    missing = [name for name in ordered if _key(name) not in sidecar_snapshot]
    if missing:
        loaded = _read_sidecars(project_name, reserving_class, missing)
        for name in missing:
            sidecar_snapshot[_key(name)] = loaded.get(name) or {}
    return {
        name: sidecar_snapshot.get(_key(name), {})
        for name in ordered
    }


def _dependency_subgraph(
    project_name: str,
    reserving_class: str,
    roots: Iterable[str],
    *,
    sidecar_snapshot: Dict[str, Dict[str, Any]] | None = None,
) -> Tuple[Dict[str, str], Dict[str, List[str]]]:
    snapshot = sidecar_snapshot if sidecar_snapshot is not None else {}
    names_by_key: Dict[str, str] = {}
    adjacency: Dict[str, List[str]] = {}
    queue = _unique_names(list(roots))
    root_keys = {_key(name) for name in queue}
    while queue:
        frontier = [name for name in _unique_names(queue) if _key(name) not in adjacency]
        queue = []
        if not frontier:
            continue
        if len(adjacency) + len(frontier) > MAX_REFRESH_GRAPH_NODES:
            raise RuntimeError("Result Selection dependency graph exceeds the safe refresh limit.")
        sidecars = _read_sidecars_cached(
            project_name,
            reserving_class,
            frontier,
            snapshot,
        )
        for name in frontier:
            key = _key(name)
            names_by_key.setdefault(key, name)
            sidecar = sidecars.get(name) or {}
            if not sidecar:
                if key in root_keys:
                    adjacency[key] = []
                    continue
                raise RuntimeError(f"Dependency graph sidecar is missing for '{name}'.")
            dependents = dataset_sidecar_status_service.entry_names(sidecar.get("dependents"))
            adjacency[key] = []
            for dependent_name in dependents:
                dependent_key = _key(dependent_name)
                if not dependent_key:
                    continue
                names_by_key.setdefault(dependent_key, dependent_name)
                adjacency[key].append(dependent_key)
                if dependent_key not in adjacency:
                    queue.append(dependent_name)
    return names_by_key, adjacency


def _assert_acyclic_dependency_subgraph(
    project_name: str,
    reserving_class: str,
    roots: Iterable[str],
    *,
    sidecar_snapshot: Dict[str, Dict[str, Any]] | None = None,
) -> Tuple[Dict[str, str], Dict[str, List[str]]]:
    names_by_key, adjacency = _dependency_subgraph(
        project_name,
        reserving_class,
        roots,
        sidecar_snapshot=sidecar_snapshot,
    )
    visiting: List[str] = []
    visiting_set = set()
    visited = set()

    def visit(key: str) -> None:
        if key in visiting_set:
            start = visiting.index(key)
            cycle = visiting[start:] + [key]
            labels = [names_by_key.get(item, item) for item in cycle]
            raise RuntimeError(f"Result Selection dependency cycle detected: {' -> '.join(labels)}")
        if key in visited:
            return
        visiting.append(key)
        visiting_set.add(key)
        for dependent_key in adjacency.get(key, []):
            visit(dependent_key)
        visiting.pop()
        visiting_set.remove(key)
        visited.add(key)

    for root in _unique_names(list(roots)):
        visit(_key(root))
    return names_by_key, adjacency


def _assert_new_precedents_do_not_cycle(
    project_name: str,
    reserving_class: str,
    output_name: str,
    precedents: Iterable[str],
) -> None:
    output_key = _key(output_name)
    precedent_keys = {_key(name) for name in precedents if _key(name)}
    if output_key in precedent_keys:
        raise HTTPException(422, "Result Selection cannot use its own output as a precedent.")
    try:
        names, adjacency = _assert_acyclic_dependency_subgraph(
            project_name,
            reserving_class,
            [output_name],
        )
    except RuntimeError as exc:
        raise HTTPException(422, str(exc)) from exc
    reachable = set()
    queue = list(adjacency.get(output_key, []))
    while queue:
        key = queue.pop(0)
        if key in reachable:
            continue
        reachable.add(key)
        queue.extend(adjacency.get(key, []))
    creates_cycle = sorted(reachable & precedent_keys)
    if creates_cycle:
        raise HTTPException(
            422,
            "Result Selection precedents would create a dependency cycle: "
            + ", ".join(names.get(key, key) for key in creates_cycle),
        )


def _persist_refreshed_method(
    project_name: str,
    reserving_class: str,
    payload: Dict[str, Any],
    sidecar: Dict[str, Any],
    *,
    allow_status_current: bool,
    precedent_names: List[str] | None = None,
) -> Dict[str, Any]:
    name = payload["details_tab"]["name"]
    method_path = _method_path(project_name, reserving_class, name)
    sidecar_path = _sidecar_path(project_name, reserving_class, name)
    timestamp = _now()
    # A refresh recomputes the method from its inputs; nobody edited it here.
    # The user-save stamp stays where the last Save put it, as it does for
    # DFM, BF and Cape Cod, and the refresh records itself separately, so the
    # ResQ sync review never reads a propagation as an edit to push.
    payload["method_metadata"]["data_refreshed"] = timestamp
    with dataset_sidecar_status_service.sidecar_write_lock(sidecar_path):
        latest_sidecar = _read_json(sidecar_path) or sidecar
        updated_sidecar = dict(latest_sidecar)
        if precedent_names is not None:
            updated_sidecar["precedents"] = dataset_sidecar_status_service.name_entries(precedent_names)
        updated_sidecar["updated_at"] = timestamp
        user_name = _current_user_name()
        updated_sidecar["modified_by"] = user_name
        updated_sidecar["status"] = (
            dataset_sidecar_status_service.compute_status(
                project_name,
                reserving_class,
                name,
                updated_sidecar,
                sidecar_path,
            )
            if allow_status_current
            else dataset_sidecar_status_service.STATUS_REVIEW_NEEDED
        )
        from app_server.services.dataset_service import _append_dataset_audit_entry

        _append_dataset_audit_entry(
            updated_sidecar,
            "Update",
            event_date=timestamp,
            user_name=user_name,
        )
        files = {
            method_path: _json_text(payload),
            **_output_files(project_name, reserving_class, payload),
            sidecar_path: _json_text(updated_sidecar),
        }
        _commit_text_files(files, last_paths=[sidecar_path])
    return updated_sidecar


def _mark_refreshed_sidecar_current(
    project_name: str,
    reserving_class: str,
    output_name: str,
    sidecar: Dict[str, Any],
    *,
    allow_status_current: bool,
) -> Tuple[Dict[str, Any], bool]:
    if not allow_status_current:
        return sidecar, False
    sidecar_path = _sidecar_path(project_name, reserving_class, output_name)
    with dataset_sidecar_status_service.sidecar_write_lock(sidecar_path):
        latest_sidecar = _read_json(sidecar_path) or sidecar
        updated_sidecar = dict(latest_sidecar)
        timestamp = _now()
        user_name = _current_user_name()
        updated_sidecar["updated_at"] = timestamp
        updated_sidecar["modified_by"] = user_name
        updated_sidecar["status"] = dataset_sidecar_status_service.compute_status(
            project_name,
            reserving_class,
            output_name,
            updated_sidecar,
            sidecar_path,
        )
        before = dataset_sidecar_status_service.normalize_status(latest_sidecar.get("status"))
        after = dataset_sidecar_status_service.normalize_status(updated_sidecar.get("status"))
        if before == after:
            return latest_sidecar, False
        _commit_text_files({sidecar_path: _json_text(updated_sidecar)})
    return updated_sidecar, after == dataset_sidecar_status_service.STATUS_CURRENT


def _mark_output_review_needed(
    project_name: str,
    reserving_class: str,
    output_name: str,
) -> None:
    sidecar_path = _sidecar_path(project_name, reserving_class, output_name)
    with dataset_sidecar_status_service.sidecar_write_lock(sidecar_path):
        sidecar = _read_json(sidecar_path)
        if not sidecar:
            return
        if dataset_sidecar_status_service.normalize_status(sidecar.get("status")) \
                == dataset_sidecar_status_service.STATUS_REVIEW_NEEDED:
            return
        sidecar["method_type"] = dataset_sidecar_status_service.METHOD_TYPE_RESULT_SELECTION
        sidecar["status"] = dataset_sidecar_status_service.STATUS_REVIEW_NEEDED
        _commit_text_files({sidecar_path: _json_text(sidecar)})


def _refresh_one_method(
    project_name: str,
    reserving_class: str,
    output_name: str,
    output_sidecar: Dict[str, Any],
    changed_sidecars: Dict[str, Dict[str, Any]],
    dependency_value_cache: Dict[Tuple[str, int, bool], List[Any]],
    *,
    allow_status_current: bool,
    blocked_precedent_keys: set[str],
    sidecar_snapshot: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:
    payload = _read_json(_method_path(project_name, reserving_class, output_name))
    if not payload:
        return {"ok": False, "dataset_name": output_name, "reason": "method_json_missing"}
    json_format = _clean(payload.get("json_format"))
    if json_format != RESULT_SELECTION_JSON_FORMAT:
        raise RuntimeError(f"Unsupported Result Selection JSON format: {json_format or '(missing)' }.")
    payload = normalize_method_payload(payload, require_complete_basis=True)
    method = payload["method_tab"]
    origin_length = int(payload["details_tab"]["origin_length"])
    row_count = len(method.get("origin_labels", []))
    precedent_names = _precedent_names(payload)
    blocked_precedents = [
        name for name in precedent_names
        if _key(name) in blocked_precedent_keys
    ]
    if blocked_precedents:
        return {
            "ok": False,
            "dataset_name": output_name,
            "reason": "Precedent refresh failed: " + ", ".join(blocked_precedents),
        }
    changed_by_key = {_key(name): (name, sidecar) for name, sidecar in changed_sidecars.items() if sidecar}
    if dataset_sidecar_status_service.normalize_status(output_sidecar.get("status")) \
            == dataset_sidecar_status_service.STATUS_REVIEW_NEEDED:
        missing_precedents = [
            name for name in precedent_names
            if _key(name) not in changed_by_key
        ]
        reloaded_sidecars = _read_sidecars_cached(
            project_name,
            reserving_class,
            missing_precedents,
            sidecar_snapshot,
        )
        unavailable_precedents = [
            name for name in missing_precedents
            if not reloaded_sidecars.get(name)
        ]
        if unavailable_precedents:
            return {
                "ok": False,
                "dataset_name": output_name,
                "reason": "Required precedent sidecar is missing: " + ", ".join(unavailable_precedents),
            }
        for name, sidecar in reloaded_sidecars.items():
            changed_by_key[_key(name)] = (name, sidecar)
        # A precedent flagged review-needed no longer blocks this recompute:
        # its numbers are current — the flag only records that no person has
        # confirmed them — and this method comes out flagged itself, so the
        # human review demand survives while the walk stays healthy. Only a
        # precedent whose refresh actually failed (``blocked_precedents``
        # above) still stops the recompute.
    changed = False
    matched_input = False

    def values_for(name: str, exact: bool) -> List[Any]:
        normalized = _key(name)
        cache_key = (normalized, origin_length, exact)
        if cache_key not in dependency_value_cache:
            source_name, source_sidecar = changed_by_key[normalized]
            dependency_value_cache[cache_key] = _dependency_values(
                project_name,
                reserving_class,
                source_name,
                source_sidecar,
                origin_length,
                exact=exact,
            )
        return dependency_value_cache[cache_key]

    for source in method.get("loaded_datasets", []):
        source_key = _key(source.get("name"))
        if source_key not in changed_by_key:
            continue
        matched_input = True
        _source_name, source_sidecar = changed_by_key[source_key]
        metadata = {
            "dataset_type": _clean(source_sidecar.get("dataset_type")) or source.get("dataset_type"),
            "data_format": _clean(source_sidecar.get("data_format")) or source.get("data_format"),
            "method_type": dataset_sidecar_status_service.normalize_method_type(
                source_sidecar.get("method_type"),
                source_sidecar.get("source_kind"),
            ),
            "category": _clean(source_sidecar.get("dataset_category") or source_sidecar.get("category"))
            or source.get("category"),
            "source_kind": _clean(source_sidecar.get("source_kind")) or source.get("source_kind"),
            "origin_length": precedent_cache_service.source_period(source_sidecar) or source.get("origin_length"),
        }
        for field, value in metadata.items():
            if value != source.get(field):
                source[field] = value
                changed = True
        values = values_for(source["name"], False)
        if len(values) != row_count:
            raise RuntimeError(
                f"Source '{source['name']}' returned {len(values)} values; expected {row_count}."
            )
        if values != source.get("values"):
            source["values"] = values
            changed = True
    for basis in method.get("ratio_basis_values", []):
        if _key(basis.get("name")) not in changed_by_key:
            continue
        matched_input = True
        values = values_for(basis["name"], True)
        if len(values) != row_count:
            raise RuntimeError(
                f"Ratio Basis '{basis['name']}' returned {len(values)} values; expected {row_count}."
            )
        if values != basis.get("values"):
            basis["values"] = values
            changed = True
    if not changed:
        if not matched_input:
            return {"ok": False, "dataset_name": output_name, "reason": "stale_reverse_dependency_edge"}
        updated_sidecar, status_refreshed = _mark_refreshed_sidecar_current(
            project_name,
            reserving_class,
            output_name,
            output_sidecar,
            allow_status_current=allow_status_current,
        )
        return {
            "ok": True,
            "dataset_name": output_name,
            "skipped": True,
            "reason": "input_values_unchanged",
            "status_refreshed": status_refreshed,
            "sidecar": updated_sidecar,
        }
    previous_output = list(method.get("selected_ultimate") or [])
    _recalculate_method(payload)
    updated_sidecar = _persist_refreshed_method(
        project_name,
        reserving_class,
        payload,
        output_sidecar,
        allow_status_current=allow_status_current,
    )
    return {
        "ok": True,
        "dataset_name": output_name,
        "updated": True,
        "output_changed": previous_output != method.get("selected_ultimate"),
        "sidecar": updated_sidecar,
    }


def refresh_dependents(
    project_name: str,
    reserving_class: str,
    changed_dataset_names: Iterable[Any],
    *,
    rebuild_index: bool = True,
    allow_status_current: bool = True,
    blocked_precedent_names: Iterable[Any] = (),
    unchanged_precedent_names: Iterable[Any] = (),
    finalize_method_review_status: bool = True,
) -> Dict[str, Any]:
    project = _clean(project_name)
    reserving = _clean(reserving_class)
    changed_names = [_clean(item) for item in changed_dataset_names]
    fresh_precedent_keys = {_key(item) for item in changed_names if _key(item)}
    blocked_precedent_keys = {
        _key(item) for item in blocked_precedent_names
        if _key(item)
    }
    queue = _unique_names(changed_names)
    visit_counts: Dict[str, int] = {}
    updated = []
    status_refreshed = []
    skipped = []
    errors = []
    downstream_fresh_names: List[str] = []
    downstream_blocked_names: List[str] = []
    # Outputs whose values did not move: a status-only refresh, a recompute
    # that came out the same, or a DFM the wave before this one already
    # visited to the same publication (``unchanged_precedent_names``). Such
    # an output, or a DFM or calculated dependent reached only through them,
    # is still current as it stands, so it is skipped without being marked a
    # failed precedent.
    unchanged_source_keys: set[str] = {
        _key(item) for item in unchanged_precedent_names
        if _key(item)
    }
    index_error = ""
    dependency_value_cache: Dict[Tuple[str, int, bool], List[Any]] = {}
    sidecar_snapshot: Dict[str, Dict[str, Any]] = {}
    with _lock(project, reserving):
        try:
            _assert_acyclic_dependency_subgraph(
                project,
                reserving,
                changed_names,
                sidecar_snapshot=sidecar_snapshot,
            )
        except Exception as exc:
            return {
                "ok": False,
                "project_name": project,
                "reserving_class": reserving,
                "changed_dataset_names": _unique_names(changed_names),
                "updated": [],
                "status_refreshed": [],
                "skipped": [],
                "errors": [{"reason": str(exc)}],
                "downstream_fresh_names": [],
                "downstream_blocked_names": [],
                "review_status_updates": [],
                "index_error": "",
            }
        while queue:
            frontier = _unique_names(queue)
            queue = []
            if not frontier:
                break
            allowed_frontier = []
            for name in frontier:
                name_key = _key(name)
                visit_counts[name_key] = visit_counts.get(name_key, 0) + 1
                if visit_counts[name_key] > MAX_REFRESH_VISITS_PER_DATASET:
                    errors.append({
                        "dataset_name": name,
                        "reason": "Result Selection dependency cycle did not converge.",
                    })
                    continue
                allowed_frontier.append(name)
            frontier = allowed_frontier
            if not frontier:
                continue
            source_sidecars = _read_sidecars_cached(
                project,
                reserving,
                frontier,
                sidecar_snapshot,
            )
            dependent_sources: Dict[str, Dict[str, Dict[str, Any]]] = {}
            for source_name in frontier:
                source_sidecar = source_sidecars.get(source_name) or {}
                for dependent_name in dataset_sidecar_status_service.entry_names(source_sidecar.get("dependents")):
                    dependent_sources.setdefault(dependent_name, {})[source_name] = source_sidecar
            if not dependent_sources:
                continue
            dependent_sidecars = _read_sidecars_cached(
                project,
                reserving,
                dependent_sources,
                sidecar_snapshot,
            )
            for dependent_name in sorted(dependent_sources, key=lambda item: (_key(item), item)):
                sidecar = dependent_sidecars.get(dependent_name) or {}
                if not sidecar:
                    dependent_key = _key(dependent_name)
                    fresh_precedent_keys.discard(dependent_key)
                    blocked_precedent_keys.add(dependent_key)
                    errors.append({
                        "dataset_name": dependent_name,
                        "reason": "dependency_sidecar_missing",
                    })
                    continue
                method_type = dataset_sidecar_status_service.normalize_method_type(
                    sidecar.get("method_type"),
                    sidecar.get("source_kind"),
                )
                if method_type != dataset_sidecar_status_service.METHOD_TYPE_RESULT_SELECTION:
                    dependent_key = _key(dependent_name)
                    if dependent_key not in fresh_precedent_keys and (
                        dependent_key in unchanged_source_keys
                        or all(
                            _key(source_name) in unchanged_source_keys
                            for source_name in dependent_sources[dependent_name]
                        )
                    ):
                        # This output kept its values -- the DFM wave already
                        # recomputed it from the same roots to the same
                        # publication -- or every source that led here did, so
                        # this DFM or calculated output needs no recompute and
                        # is not blocked: blocking it would refuse every
                        # Result Selection further down that also loads it
                        # with "Precedent refresh failed" over a refresh
                        # nothing needed.
                        skipped.append({
                            "dataset_name": dependent_name,
                            "reason": "non_result_selection_dependent_inputs_unchanged",
                        })
                        continue
                    if dependent_key not in fresh_precedent_keys:
                        blocked_precedent_keys.add(dependent_key)
                    skipped.append({
                        "dataset_name": dependent_name,
                        "reason": "non_result_selection_dependent_requires_explicit_refresh",
                    })
                    continue
                try:
                    result = _refresh_one_method(
                        project,
                        reserving,
                        dependent_name,
                        sidecar,
                        dependent_sources[dependent_name],
                        dependency_value_cache,
                        allow_status_current=allow_status_current,
                        blocked_precedent_keys=blocked_precedent_keys,
                        sidecar_snapshot=sidecar_snapshot,
                    )
                except Exception as exc:
                    dependent_key = _key(dependent_name)
                    fresh_precedent_keys.discard(dependent_key)
                    blocked_precedent_keys.add(dependent_key)
                    _mark_output_review_needed(project, reserving, dependent_name)
                    sidecar_snapshot.pop(dependent_key, None)
                    errors.append({"dataset_name": dependent_name, "reason": str(exc)})
                    continue
                if result.get("ok") is False:
                    dependent_key = _key(dependent_name)
                    fresh_precedent_keys.discard(dependent_key)
                    blocked_precedent_keys.add(dependent_key)
                    _mark_output_review_needed(project, reserving, dependent_name)
                    sidecar_snapshot.pop(dependent_key, None)
                    errors.append({
                        "dataset_name": dependent_name,
                        "reason": result.get("reason") or "result_selection_refresh_failed",
                    })
                    continue
                refreshed_sidecar = result.get("sidecar") or sidecar
                dependent_key = _key(dependent_name)
                sidecar_snapshot[dependent_key] = refreshed_sidecar
                refreshed_status = dataset_sidecar_status_service.normalize_status(
                    refreshed_sidecar.get("status")
                )
                if refreshed_status not in (
                    dataset_sidecar_status_service.STATUS_CURRENT,
                    # Review-needed means freshly recomputed but awaiting a
                    # person's sign-off; deeper dependents keep cascading from
                    # those current numbers instead of declining, so one save
                    # no longer fails the whole walk over its own flags.
                    dataset_sidecar_status_service.STATUS_REVIEW_NEEDED,
                ):
                    fresh_precedent_keys.discard(dependent_key)
                    blocked_precedent_keys.add(dependent_key)
                else:
                    blocked_precedent_keys.discard(dependent_key)
                    fresh_precedent_keys.add(dependent_key)
                if result.get("updated"):
                    updated.append({"dataset_name": dependent_name})
                    for cache_key in list(dependency_value_cache):
                        if cache_key[0] == dependent_key:
                            dependency_value_cache.pop(cache_key, None)
                    output_is_current = dataset_sidecar_status_service.normalize_status(
                        (result.get("sidecar") or {}).get("status")
                    ) == dataset_sidecar_status_service.STATUS_CURRENT
                    if not output_is_current:
                        continue
                    if not result.get("output_changed"):
                        unchanged_source_keys.add(dependent_key)
                        queue.append(dependent_name)
                        continue
                    unchanged_source_keys.discard(dependent_key)
                    status_updates = dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
                        project, reserving, [dependent_name]
                    )
                    for status_update in status_updates:
                        sidecar_snapshot.pop(_key(status_update.get("dataset_name")), None)
                    queue.append(dependent_name)
                    try:
                        from app_server.services import calculated_dataset_service

                        calculated = calculated_dataset_service.recalculate_dependents(
                            project,
                            reserving,
                            dependent_name,
                            _clean(sidecar.get("dataset_type")) or dependent_name,
                            include_result_selection=False,
                            include_bornhuetter_ferguson=False,
                            include_cape_cod=False,
                            include_bootstrap=False,
                            finalize_method_review_status=False,
                            rebuild_index=False,
                        )
                        calculated_names = [
                            _clean(item.get("dataset_type_name"))
                            for item in calculated.get("updated", [])
                            if _clean(item.get("dataset_type_name"))
                        ]
                        failed_calculated_names = [
                            _clean(item.get("dataset_type_name"))
                            for item in calculated.get("skipped", [])
                            if _clean(item.get("dataset_type_name"))
                        ]
                        nested_dfm = calculated.get("dfm_updates") \
                            if isinstance(calculated.get("dfm_updates"), dict) else {}
                        nested_dfm_names = [
                            _clean(item.get("dataset_name") or item.get("dataset_type"))
                            for field in ("updated", "status_refreshed")
                            for item in nested_dfm.get(field, [])
                            if isinstance(item, dict)
                            and _clean(item.get("dataset_name") or item.get("dataset_type"))
                        ]
                        failed_nested_dfm_names = [
                            _clean(item.get("dataset_name") or item.get("dataset_type"))
                            for item in nested_dfm.get("errors", [])
                            if isinstance(item, dict)
                            and _clean(item.get("dataset_name") or item.get("dataset_type"))
                        ]
                        nested_fresh_names = _unique_names([
                            *calculated_names,
                            *nested_dfm_names,
                        ])
                        nested_failed_names = _unique_names([
                            *failed_calculated_names,
                            *failed_nested_dfm_names,
                        ])
                        downstream_fresh_names.extend(nested_fresh_names)
                        downstream_blocked_names.extend(nested_failed_names)
                        successful_calculated_keys = {_key(name) for name in nested_fresh_names}
                        failed_calculated_keys = {_key(name) for name in nested_failed_names}
                        blocked_precedent_keys.difference_update(successful_calculated_keys)
                        fresh_precedent_keys.update(successful_calculated_keys)
                        fresh_precedent_keys.difference_update(failed_calculated_keys)
                        blocked_precedent_keys.update(failed_calculated_keys)
                        for calculated_name in [*nested_fresh_names, *nested_failed_names]:
                            sidecar_snapshot.pop(_key(calculated_name), None)
                        for calculated_name in nested_fresh_names:
                            calculated_key = _key(calculated_name)
                            for cache_key in list(dependency_value_cache):
                                if cache_key[0] == calculated_key:
                                    dependency_value_cache.pop(cache_key, None)
                        # A nested DFM output that was only status-refreshed
                        # kept its values; the rest of the fresh names moved.
                        nested_changed_keys = {
                            _key(item.get("dataset_name") or item.get("dataset_type"))
                            for item in nested_dfm.get("updated", [])
                            if isinstance(item, dict) and item.get("output_changed", True)
                        }
                        nested_changed_keys.update(_key(name) for name in calculated_names)
                        for nested_name in nested_fresh_names:
                            if _key(nested_name) in nested_changed_keys:
                                unchanged_source_keys.discard(_key(nested_name))
                            else:
                                unchanged_source_keys.add(_key(nested_name))
                        queue.extend(nested_fresh_names)
                        if not calculated.get("ok", True):
                            failed_steps = calculated.get("skipped") or []
                            errors.append({
                                "dataset_name": dependent_name,
                                "reason": "Calculated dependent refresh failed: "
                                + "; ".join(
                                    _clean(item.get("reason")) or _clean(item.get("dataset_type_name"))
                                    for item in failed_steps
                                ),
                            })
                    except Exception as exc:
                        sidecar_snapshot.clear()
                        dependency_value_cache.clear()
                        errors.append({
                            "dataset_name": dependent_name,
                            "reason": f"Calculated dependent refresh failed: {exc}",
                        })
                else:
                    if result.get("status_refreshed"):
                        status_refreshed.append({"dataset_name": dependent_name})
                        unchanged_source_keys.add(dependent_key)
                        queue.append(dependent_name)
                    skipped.append({
                        "dataset_name": dependent_name,
                        "reason": result.get("reason") or "not_updated",
                    })
        review_status_updates = (
            dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
                project,
                reserving,
                changed_names,
            )
            if finalize_method_review_status
            else []
        )
        if (updated or status_refreshed or review_status_updates) and rebuild_index:
            try:
                from app_server.services import dataset_instance_index_service

                dataset_instance_index_service.rebuild_index(project, reserving)
            except Exception as exc:
                index_error = str(exc)
    def unique_updates(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        names = _unique_names([item.get("dataset_name") for item in items])
        return [{"dataset_name": name} for name in names]

    return {
        "ok": not errors,
        "project_name": project,
        "reserving_class": reserving,
        "changed_dataset_names": _unique_names(changed_names),
        "updated": unique_updates(updated),
        "status_refreshed": unique_updates(status_refreshed),
        "skipped": skipped,
        "errors": errors,
        "downstream_fresh_names": _unique_names(downstream_fresh_names),
        "downstream_blocked_names": _unique_names(downstream_blocked_names),
        "review_status_updates": review_status_updates,
        "index_error": index_error,
    }
