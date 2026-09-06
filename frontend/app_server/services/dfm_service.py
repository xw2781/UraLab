"""Persist, load, preview, and eagerly refresh self-contained DFM methods."""
from __future__ import annotations

import getpass
import json
import math
import os
import re
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Dict, Iterable, List, Mapping, Tuple

import pandas as pd
from fastapi import HTTPException

from arcrho_api.dfm_contract import (
    DFM_JSON_FORMAT,
    DfmContractError,
    apply_owned_patch,
    build_dfm_output_sidecar,
    dfm_dataset_reference_tokens,
    dfm_precedent_names,
    dfm_output_variants,
    method_revisions,
    normalize_dfm_method,
    persisted_projection,
    preview_dfm_method as canonical_preview_dfm_method,
    recalculate_dfm_method,
    stamp_last_modified,
)
from arcrho_api.io import persisted_json_text
from arcrho_api.sidecar_audit_contract import AUDIT_ACTION_AUTO_REFRESH
from arcrho_api.sidecar_core_contract import stored_lengths
from arcrho_api.timestamps import persisted_timestamp, utc_now_text
from app_server import config
from app_server.helpers import sanitize_dataset_file_name
from app_server.services import (
    dataset_sidecar_status_service,
    dependent_propagation_service,
    precedent_cache_service,
    user_identity_service,
)


READ_MAX_WORKERS = 4
_READ_EXECUTOR = ThreadPoolExecutor(
    max_workers=READ_MAX_WORKERS,
    thread_name_prefix="arcrho-dfm-read",
)
SnapshotCacheKey = Tuple[str, bool, Tuple[str, ...], Tuple[str, ...], int, int]


def _clean(value: Any) -> str:
    return str(value if value is not None else "").strip()


def _key(value: Any) -> str:
    return " ".join(_clean(value).lower().split())


def _axis_labels(values: Any) -> List[str]:
    if values is None or isinstance(values, (str, bytes, Mapping)):
        return []
    try:
        return [str(item if item is not None else "") for item in values]
    except TypeError:
        return []


def _positive_int(value: Any) -> int:
    try:
        return max(0, int(value or 0))
    except (TypeError, ValueError):
        return 0


def _now() -> str:
    return utc_now_text()


def _lock(project_name: str, reserving_class: str) -> threading.RLock:
    return dataset_sidecar_status_service.reserving_class_io_lock(project_name, reserving_class)


def _method_path(project_name: str, reserving_class: str, method_name: str) -> str:
    return dataset_sidecar_status_service.method_json_path(
        project_name, reserving_class, dataset_sidecar_status_service.METHOD_TYPE_DFM, method_name
    )


def _sidecar_path(project_name: str, reserving_class: str, output_dataset: str) -> str:
    return dataset_sidecar_status_service.sidecar_path(project_name, reserving_class, output_dataset)


def _read_json(path: str) -> Dict[str, Any]:
    try:
        with open(path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except FileNotFoundError:
        return {}
    except PermissionError as exc:
        raise HTTPException(423, f"DFM file is locked or inaccessible: {os.path.basename(path)}") from exc
    except (OSError, json.JSONDecodeError) as exc:
        raise HTTPException(500, f"Invalid DFM JSON: {os.path.basename(path)}: {exc}") from exc
    return payload if isinstance(payload, dict) else {}


def _json_text(payload: Mapping[str, Any]) -> str:
    return persisted_json_text(payload)


def _method_json_text(payload: Mapping[str, Any]) -> str:
    """Serialize a DFM method through the canonical on-disk projection.

    Every method-file write and every unchanged-file comparison must go through
    here, so a file is only rewritten when its persisted content really differs.
    """
    return _json_text(persisted_projection(payload))


def _read_bytes_if_file(path: str) -> bytes | None:
    if not os.path.isfile(path):
        return None
    with open(path, "rb") as handle:
        return handle.read()


def _commit_text_files(files: Mapping[str, str], *, last_paths: Iterable[str] = ()) -> List[str]:
    """Atomically replace changed files, rolling back all replacements on failure."""

    last_keys = {os.path.normcase(os.path.abspath(path)) for path in last_paths}
    changed = {
        path: value
        for path, value in files.items()
        if _read_bytes_if_file(path) != value.encode("utf-8")
    }
    ordered_paths = sorted(
        changed,
        key=lambda path: (
            os.path.normcase(os.path.abspath(path)) in last_keys,
            os.path.normcase(path),
        ),
    )
    staged: Dict[str, str] = {}
    backups: Dict[str, bytes | None] = {}
    replaced: List[str] = []
    try:
        for path in ordered_paths:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            backups[path] = _read_bytes_if_file(path)
            temporary = f"{path}.{uuid.uuid4()}.tmp"
            with open(temporary, "w", encoding="utf-8", newline="\n") as handle:
                handle.write(changed[path])
            staged[path] = temporary
        for path in ordered_paths:
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
            raise RuntimeError(f"{exc}; DFM rollback failed: {'; '.join(rollback_errors)}") from exc
        raise
    finally:
        for temporary in staged.values():
            try:
                os.remove(temporary)
            except OSError:
                pass
    return replaced


def _contract_call(func, *args: Any, **kwargs: Any) -> Dict[str, Any]:
    try:
        result = func(*args, **kwargs)
    except DfmContractError as exc:
        raise HTTPException(422, str(exc)) from exc
    if not isinstance(result, dict):
        raise HTTPException(500, "Canonical DFM calculation returned an invalid payload.")
    return result


def _details(payload: Mapping[str, Any]) -> Dict[str, Any]:
    value = payload.get("details_tab") if isinstance(payload, Mapping) else None
    return value if isinstance(value, dict) else {}


def _data_tab(payload: Mapping[str, Any]) -> Dict[str, Any]:
    value = payload.get("data_tab") if isinstance(payload, Mapping) else None
    return value if isinstance(value, dict) else {}


def _results_tab(payload: Mapping[str, Any]) -> Dict[str, Any]:
    value = payload.get("results_tab") if isinstance(payload, Mapping) else None
    return value if isinstance(value, dict) else {}


def _identity(payload: Mapping[str, Any]) -> Tuple[str, str]:
    details = _details(payload)
    method_name = _clean(details.get("name"))
    output_dataset = _clean(details.get("output_dataset")) or method_name
    if not method_name or not output_dataset:
        raise HTTPException(422, "DFM name and output dataset are required.")
    return method_name, output_dataset


def _precedent_names(payload: Mapping[str, Any]) -> List[str]:
    return dfm_precedent_names(payload)


def _revision_response(payload: Mapping[str, Any]) -> Dict[str, str]:
    revisions = method_revisions(payload)
    owned = _clean(revisions.get("owned_revision"))
    derived = _clean(revisions.get("derived_revision"))
    publication = _clean(revisions.get("publication_revision"))
    return {
        "owned_revision": owned,
        "derived_revision": derived,
        "publication_revision": publication,
        "method_revision": publication,
    }


def _sidecar_response(payload: Mapping[str, Any], *, exists: bool) -> Dict[str, Any]:
    if not exists:
        return {"exists": False, "notes": "", "audit_log": []}
    return {**dict(payload), "exists": True}


def _load_source_snapshot(
    project_name: str,
    reserving_class: str,
    dataset_name: str,
    *,
    vector: bool,
    allow_review_needed: bool = False,
    canonical_origin_labels: Iterable[Any] = (),
    canonical_development_labels: Iterable[Any] = (),
    expected_origin_length: int = 0,
    expected_development_length: int = 0,
) -> Dict[str, Any]:
    sidecar_path = _sidecar_path(project_name, reserving_class, dataset_name)
    sidecar = _read_json(sidecar_path)
    if not sidecar:
        raise HTTPException(404, f"DFM precedent sidecar is missing: {dataset_name}")
    source_method_type = dataset_sidecar_status_service.normalize_method_type(
        sidecar.get("method_type"), sidecar.get("source_kind")
    )
    source_status = dataset_sidecar_status_service.normalize_status(sidecar.get("status"))
    if not allow_review_needed \
            and source_method_type != dataset_sidecar_status_service.METHOD_TYPE_NONE \
            and source_status == dataset_sidecar_status_service.STATUS_REVIEW_NEEDED:
        raise HTTPException(409, f"DFM precedent requires review: {dataset_name}")
    data_format = _clean(sidecar.get("data_format")) or "Triangle"
    is_vector = data_format.lower() == "vector"
    if not vector and is_vector:
        raise HTTPException(422, f"DFM input '{dataset_name}' must be a Triangle dataset.")
    # Stored, not displayed: the method reads this precedent's own CSV, so the
    # shape that must match is the one that file holds.
    stored_origin_length, stored_development_length = stored_lengths(sidecar)
    source_origin_length = _positive_int(stored_origin_length)
    required_origin_length = _positive_int(expected_origin_length)
    origin_mismatch = bool(
        source_origin_length and required_origin_length
        and source_origin_length != required_origin_length
    )
    source_development_length = 0 if is_vector else _positive_int(stored_development_length)
    required_development_length = _positive_int(expected_development_length)
    development_mismatch = bool(
        not vector and source_development_length and required_development_length
        and source_development_length != required_development_length
    )
    rollup_target_origin = required_origin_length or source_origin_length
    rollup_target_development = None if vector else required_development_length
    needs_rollup = False
    if (origin_mismatch or development_mismatch) \
            and _clean(sidecar.get("source_kind")).lower() == "engine":
        # The Engine builds its own datasets at any period from the source
        # table, so a generated precedent cached at another period is rebuilt
        # at the method's lengths instead of refused.
        try:
            csv_path = precedent_cache_service.materialize_engine_source(
                project_name,
                reserving_class,
                dataset_name,
                sidecar,
                rollup_target_origin,
                development_length=rollup_target_development,
            )
        except RuntimeError as exc:
            raise HTTPException(
                422,
                f"DFM precedent '{dataset_name}' could not be generated at the method's period: {exc}",
            ) from exc
    elif (origin_mismatch or development_mismatch) and not precedent_cache_service.rollup_reason(
        sidecar, rollup_target_origin, rollup_target_development
    ):
        # A hand-entered precedent stored at a finer period is aggregated to the
        # method's own lengths in memory, from the CSV the sidecar names. No
        # coarser copy is written, so an Excel refresh of the stored figures is
        # picked up on the next load.
        needs_rollup = True
        csv_path = precedent_cache_service.sidecar_csv_path(project_name, reserving_class, sidecar)
        if not csv_path:
            raise HTTPException(422, f"DFM precedent '{dataset_name}' does not identify its cache CSV.")
    else:
        if origin_mismatch:
            raise HTTPException(
                422,
                f"DFM precedent '{dataset_name}' has incompatible origin period length "
                f"({source_origin_length}; expected {required_origin_length}).",
            )
        if development_mismatch:
            raise HTTPException(
                422,
                f"DFM input '{dataset_name}' has incompatible development period length "
                f"({source_development_length}; expected {required_development_length}).",
            )
        csv_path = precedent_cache_service.sidecar_csv_path(project_name, reserving_class, sidecar)
        if not csv_path:
            raise HTTPException(422, f"DFM precedent '{dataset_name}' does not identify its cache CSV.")
    try:
        frame = pd.read_csv(csv_path, header=None).astype(object)
    except FileNotFoundError as exc:
        raise HTTPException(404, f"DFM precedent CSV is missing: {dataset_name}") from exc
    except PermissionError as exc:
        raise HTTPException(423, f"DFM precedent CSV is locked: {dataset_name}") from exc
    except Exception as exc:
        raise HTTPException(422, f"DFM precedent CSV is invalid: {dataset_name}: {exc}") from exc
    frame = frame.where(pd.notnull(frame), None)
    raw_values = frame.values.tolist()
    if needs_rollup:
        try:
            raw_values = precedent_cache_service.rollup_rows(
                project_name, sidecar, raw_values, rollup_target_origin, rollup_target_development
            )
        except ValueError as exc:
            raise HTTPException(
                422,
                f"DFM precedent '{dataset_name}' could not be rolled up to the method's period: {exc}",
            ) from exc
    method_origin_labels = _axis_labels(canonical_origin_labels)
    origin_labels = method_origin_labels or _axis_labels(sidecar.get("origin_labels"))
    if len(origin_labels) != len(raw_values):
        raise HTTPException(
            422,
            f"DFM precedent '{dataset_name}' has {len(raw_values)} rows; "
            f"expected {len(origin_labels)}.",
        )
    try:
        decimal_places = int(sidecar.get("decimal_places") or 0)
    except (TypeError, ValueError):
        decimal_places = 0
    if vector:
        if is_vector:
            values = [row[0] if isinstance(row, list) and row else None for row in raw_values]
        else:
            # Triangle Ratio Basis follows the UI rule: latest available diagonal
            # value for each exact origin label.
            values = []
            for row in raw_values:
                latest = None
                for value in reversed(row if isinstance(row, list) else []):
                    if value is not None:
                        latest = value
                        break
                values.append(latest)
        snapshot: Dict[str, Any] = {
            "name": _clean(sidecar.get("dataset_name")) or dataset_name,
            "data_format": data_format,
            "origin_labels": origin_labels,
            "values": values,
            "number_format": _clean(sidecar.get("number_format")) or "#,##0",
            "decimal_places": decimal_places,
        }
    else:
        column_count = max((len(row) for row in raw_values), default=0)
        method_development_labels = _axis_labels(canonical_development_labels)
        if method_development_labels:
            if len(method_development_labels) != column_count:
                raise HTTPException(
                    422,
                    f"DFM input '{dataset_name}' has incompatible development geometry.",
                )
            development_labels = method_development_labels
        else:
            development_labels = _axis_labels(sidecar.get("development_labels"))
        if not method_development_labels and len(development_labels) != column_count:
            # These labels describe the CSV just opened: the method's own
            # lengths when an Engine precedent was rebuilt at them, and the
            # sidecar's stored shape otherwise.
            first_development = max(1, required_origin_length or source_origin_length or 12)
            development_step = max(1, required_development_length or source_development_length or 12)
            development_labels = [
                str(first_development + development_step * index)
                for index in range(column_count)
            ]
        snapshot = {
            "name": _clean(sidecar.get("dataset_name")) or dataset_name,
            "data_format": data_format,
            "origin_labels": origin_labels,
            "development_labels": development_labels,
            "values": raw_values,
            "mask": [[value is not None for value in row] for row in raw_values],
            "number_format": _clean(sidecar.get("number_format")) or "#,##0",
            "decimal_places": decimal_places,
        }
    snapshot["_method_type"] = source_method_type
    snapshot["_status"] = source_status
    return snapshot


def _dataset_reference_axis_index(
    raw_index: Any,
    labels: Iterable[Any],
    *,
    axis_name: str,
    dataset_name: str,
    negative_index_length: int | None = None,
) -> Tuple[int, str]:
    token = _clean(raw_index)
    if not token:
        raise HTTPException(422, f"{axis_name.capitalize()} index is required for '{dataset_name}'.")
    axis_labels = _axis_labels(labels)
    quoted = len(token) >= 2 and token[0] in {'"', "'"} and token[-1] == token[0]
    if quoted:
        requested_label = token[1:-1]
        matches = [index for index, label in enumerate(axis_labels) if label == requested_label]
        if len(matches) != 1:
            detail = "not found" if not matches else "ambiguous"
            raise HTTPException(
                422,
                f"{axis_name.capitalize()} label '{requested_label}' is {detail} in '{dataset_name}'.",
            )
        return matches[0], axis_labels[matches[0]]
    negative_match = re.fullmatch(r"-([1-9]\d*)", token)
    if negative_match:
        valid_length = (
            len(axis_labels)
            if negative_index_length is None
            else max(0, min(len(axis_labels), int(negative_index_length)))
        )
        from_end = int(negative_match.group(1))
        resolved_index = valid_length - from_end
        if resolved_index < 0:
            raise HTTPException(
                422,
                f"{axis_name.capitalize()} index -{from_end} is outside the valid range "
                f"of '{dataset_name}' ({valid_length} positions).",
            )
        return resolved_index, axis_labels[resolved_index]
    if token.isdigit():
        one_based = int(token)
        if 1 <= one_based <= len(axis_labels):
            return one_based - 1, axis_labels[one_based - 1]
        # Large numeric axis labels such as origin year 2024 are labels rather
        # than plausible positions. Quoting remains available to disambiguate a
        # numeric label that is also a valid position.
        matches = [index for index, label in enumerate(axis_labels) if label == token]
        if len(matches) == 1:
            return matches[0], axis_labels[matches[0]]
        raise HTTPException(
            422,
            f"{axis_name.capitalize()} index {one_based} is outside '{dataset_name}' "
            f"(1-{len(axis_labels)}).",
        )
    matches = [index for index, label in enumerate(axis_labels) if label == token]
    if len(matches) != 1:
        detail = "not found" if not matches else "ambiguous"
        raise HTTPException(
            422,
            f"{axis_name.capitalize()} label '{token}' is {detail} in '{dataset_name}'.",
        )
    return matches[0], axis_labels[matches[0]]


def _dataset_reference_has_value(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, str) and not value.strip():
        return False
    try:
        return not bool(pd.isna(value))
    except (TypeError, ValueError):
        return True


def _dataset_reference_valid_boundary(
    values: Iterable[Any],
    *,
    vector: bool,
    valuation_row_count: int | None = None,
) -> int:
    """Return the last valid vector position or triangle calendar diagonal.

    Dataset caches retain their full configured geometry. Sub-annual projects
    therefore have a trailing suffix (vectors) or empty calendar diagonals
    (triangles) beyond the valuation period. The last non-empty cell
    establishes that boundary; blanks inside it remain valid positions. A
    vector's rows after the Development End Date may also hold values, so its
    boundary never passes ``valuation_row_count`` when the project settings
    provide one: ``[-1]`` is the valuation period whether or not later rows
    are filled.
    """
    boundary = -1
    for row_index, raw_row in enumerate(values):
        row = raw_row if isinstance(raw_row, list) else [raw_row]
        for col_index, value in enumerate(row):
            if _dataset_reference_has_value(value):
                boundary = max(boundary, row_index if vector else row_index + col_index)
    if vector and valuation_row_count is not None:
        boundary = min(boundary, int(valuation_row_count) - 1)
    return boundary


def with_valuation_row_counts(
    project_name: str,
    datasets: Mapping[str, Dict[str, Any]],
) -> Mapping[str, Dict[str, Any]]:
    """Stamp each loaded vector with the project's valuation row count.

    Shared by the DFM reference resolver and the dataset internal-link
    resolver so a negative vector index resolves the same way in both. One
    General Settings read per distinct origin period length.
    """
    from app_server.services import dataset_service

    counts_by_length: Dict[int, int | None] = {}
    for dataset in datasets.values():
        if _clean(dataset.get("data_format")).casefold() != "vector":
            continue
        length = max(1, int(dataset.get("origin_length") or 1))
        if length not in counts_by_length:
            counts_by_length[length] = dataset_service.valuation_origin_row_count(project_name, length)
        dataset["valuation_row_count"] = counts_by_length[length]
    return datasets


def _resolved_dataset_reference(
    reference: Mapping[str, Any],
    dataset: Mapping[str, Any],
) -> Dict[str, Any]:
    requested_name = _clean(reference.get("dataset_name"))
    dataset_name = _clean(dataset.get("dataset_name")) or requested_name
    values = dataset.get("values") if isinstance(dataset.get("values"), list) else []
    origin_labels = dataset.get("origin_labels") if isinstance(dataset.get("origin_labels"), list) else []
    data_format = _clean(dataset.get("data_format")) or "Triangle"
    is_vector = data_format.casefold() == "vector"
    valid_boundary = _dataset_reference_valid_boundary(
        values,
        vector=is_vector,
        valuation_row_count=dataset.get("valuation_row_count"),
    )
    row_index, row_label = _dataset_reference_axis_index(
        reference.get("row_idx"),
        origin_labels,
        axis_name="row",
        dataset_name=dataset_name,
        negative_index_length=valid_boundary + 1,
    )
    raw_col = _clean(reference.get("col_idx"))
    if not is_vector and not raw_col:
        raise HTTPException(422, f"Column index is required for Triangle dataset '{dataset_name}'.")
    development_labels = (
        dataset.get("dev_labels")
        if isinstance(dataset.get("dev_labels"), list)
        else []
    )
    if is_vector and not development_labels:
        development_labels = ["Ultimate"]
    col_index, col_label = _dataset_reference_axis_index(
        raw_col or "1",
        development_labels,
        axis_name="column",
        dataset_name=dataset_name,
        negative_index_length=(
            len(development_labels)
            if is_vector
            else max(0, valid_boundary - row_index + 1)
        ),
    )
    row = values[row_index] if row_index < len(values) and isinstance(values[row_index], list) else []
    value = row[col_index] if col_index < len(row) else None
    try:
        numeric_value = float(value)
    except (TypeError, ValueError) as exc:
        raise HTTPException(
            422,
            f"Referenced cell [{dataset_name}][{row_label}, {col_label}] is blank or non-numeric.",
        ) from exc
    if not math.isfinite(numeric_value):
        raise HTTPException(
            422,
            f"Referenced cell [{dataset_name}][{row_label}, {col_label}] is blank or non-numeric.",
        )
    return {
        "dataset_name": dataset_name,
        "data_format": data_format,
        "row_label": row_label,
        "col_label": col_label,
        "value": numeric_value,
    }


def resolve_dfm_dataset_references(
    project_name: str,
    reserving_class: str,
    references: Iterable[Mapping[str, Any]],
) -> Dict[str, Any]:
    """Resolve DFM User Entry dataset references with one read per dataset."""
    from app_server.services import dataset_service

    project = _clean(project_name)
    rc = _clean(reserving_class)
    requested = [dict(reference) for reference in references]
    if not project or not rc:
        raise HTTPException(422, "Project and reserving class are required.")
    if not requested:
        raise HTTPException(422, "At least one dataset reference is required.")

    names_by_key: Dict[str, str] = {}
    for reference in requested:
        name = _clean(reference.get("dataset_name"))
        if not name:
            raise HTTPException(422, "Dataset name is required in every reference.")
        names_by_key.setdefault(_key(name), name)
    futures = {
        key: _READ_EXECUTOR.submit(
            dataset_service.load_cached_dataset_values,
            project,
            rc,
            name,
        )
        for key, name in names_by_key.items()
    }
    datasets = with_valuation_row_counts(project, {key: future.result() for key, future in futures.items()})
    return {
        "ok": True,
        "results": [
            _resolved_dataset_reference(reference, datasets[_key(reference.get("dataset_name"))])
            for reference in requested
        ],
    }


def _source_snapshots(
    project_name: str,
    reserving_class: str,
    payload: Mapping[str, Any],
    *,
    load_input: bool,
    load_basis: bool,
    allow_review_needed: bool = False,
    snapshot_cache: Dict[SnapshotCacheKey, Dict[str, Any]] | None = None,
) -> Tuple[Dict[str, Any] | None, Dict[str, Any] | None]:
    details = _details(payload)
    data = _data_tab(payload)
    results = _results_tab(payload)
    input_name = _clean(details.get("input_triangle"))
    basis_name = _clean(results.get("ratio_basis_dataset"))
    method_origin_labels = tuple(_axis_labels(data.get("origin_labels")))
    method_development_labels = tuple(_axis_labels(data.get("development_labels")))
    origin_length = _positive_int(details.get("origin_length"))
    development_length = _positive_int(details.get("development_length"))
    cache = snapshot_cache if snapshot_cache is not None else {}
    futures = {}
    input_cache_key: SnapshotCacheKey = (
        _key(input_name),
        False,
        method_origin_labels,
        method_development_labels,
        origin_length,
        development_length,
    )
    basis_cache_key: SnapshotCacheKey = (
        _key(basis_name),
        True,
        method_origin_labels,
        (),
        origin_length,
        0,
    )
    if load_input:
        if not input_name:
            raise HTTPException(422, "DFM input triangle is required.")
        if input_cache_key not in cache:
            futures["input"] = _READ_EXECUTOR.submit(
                _load_source_snapshot,
                project_name,
                reserving_class,
                input_name,
                vector=False,
                allow_review_needed=allow_review_needed,
                canonical_origin_labels=method_origin_labels,
                canonical_development_labels=method_development_labels,
                expected_origin_length=origin_length,
                expected_development_length=development_length,
            )
    if load_basis and basis_name:
        if basis_cache_key not in cache:
            futures["basis"] = _READ_EXECUTOR.submit(
                _load_source_snapshot,
                project_name,
                reserving_class,
                basis_name,
                vector=True,
                allow_review_needed=allow_review_needed,
                canonical_origin_labels=method_origin_labels,
                expected_origin_length=origin_length,
            )
    if "input" in futures:
        cache[input_cache_key] = futures["input"].result()
    if "basis" in futures:
        cache[basis_cache_key] = futures["basis"].result()
    input_snapshot = cache.get(input_cache_key) if load_input else None
    basis_snapshot = cache.get(basis_cache_key) if load_basis and basis_name else None
    return input_snapshot, basis_snapshot


def _recalculate_with_sources(
    project_name: str,
    reserving_class: str,
    payload: Mapping[str, Any],
    *,
    load_input: bool,
    load_basis: bool,
    allow_review_needed: bool = False,
    changed_precedents: Iterable[str] = (),
    snapshot_cache: Dict[SnapshotCacheKey, Dict[str, Any]] | None = None,
    dataset_reference_values: Mapping[str, Any] | None = None,
) -> Dict[str, Any]:
    input_snapshot, basis_snapshot = _source_snapshots(
        project_name,
        reserving_class,
        payload,
        load_input=load_input,
        load_basis=load_basis,
        allow_review_needed=allow_review_needed,
        snapshot_cache=snapshot_cache,
    )
    return _contract_call(
        recalculate_dfm_method,
        dict(payload),
        input_snapshot=input_snapshot,
        ratio_basis_snapshot=basis_snapshot,
        changed_precedents=tuple(changed_precedents),
        timestamp=_now(),
        dataset_reference_values=dataset_reference_values,
    )


def _resolved_reference_token_values(
    project_name: str,
    reserving_class: str,
    tokens: List[Dict[str, Any]],
) -> Dict[str, float]:
    """Resolve dataset-reference tokens to values keyed by their reference text."""

    response = resolve_dfm_dataset_references(
        project_name,
        reserving_class,
        [
            {
                "dataset_name": token["dataset_name"],
                "row_idx": token["row_idx"],
                **({"col_idx": token["col_idx"]} if token["col_idx"] else {}),
            }
            for token in tokens
        ],
    )
    results = response.get("results") or []
    return {
        token["match"]: result["value"]
        for token, result in zip(tokens, results)
    }


def _assert_refreshable_precedents(
    project_name: str,
    reserving_class: str,
    payload: Mapping[str, Any],
    snapshot_cache: Mapping[SnapshotCacheKey, Mapping[str, Any]],
    precedent_names: Iterable[str] | None = None,
) -> None:
    missing = []
    futures = {}
    names = list(precedent_names) if precedent_names is not None else _precedent_names(payload)
    for name in names:
        normalized = _key(name)
        cached = next(
            (
                snapshot
                for cache_key, snapshot in snapshot_cache.items()
                if cache_key[0] == normalized
            ),
            None,
        )
        if cached is not None:
            continue
        futures[name] = _READ_EXECUTOR.submit(
            _read_json,
            _sidecar_path(project_name, reserving_class, name),
        )
    for name, future in futures.items():
        sidecar = future.result()
        if not sidecar:
            missing.append(name)
    if missing:
        raise RuntimeError("DFM precedent sidecar is missing: " + ", ".join(missing))


def _csv_text(values: Iterable[Any]) -> str:
    rows = []
    for value in values:
        if value is None:
            rows.append("")
            continue
        try:
            number = float(value)
        except (TypeError, ValueError):
            rows.append("")
            continue
        rows.append(str(value) if math.isfinite(number) else "")
    return "\n".join(rows) + "\n"


def _output_files(project_name: str, reserving_class: str, payload: Mapping[str, Any]) -> Dict[str, str]:
    details = _details(payload)
    output_dataset = _clean(details.get("output_dataset")) or _clean(details.get("name"))
    data_dir = config.get_project_dataset_cache_dir(project_name, reserving_class)
    safe_name = sanitize_dataset_file_name(output_dataset)
    return {
        os.path.join(data_dir, f"{safe_name}@{period_length}.csv"): _csv_text(values)
        for period_length, values in dfm_output_variants(payload).items()
    }


def _build_sidecar(
    project_name: str,
    reserving_class: str,
    payload: Mapping[str, Any],
    existing: Mapping[str, Any],
    *,
    notes: str | None,
    changed: bool,
    automatic: bool,
) -> Dict[str, Any]:
    from app_server.services import calculated_dataset_service

    details = _details(payload)
    _method_name, output_dataset = _identity(payload)
    origin_length = int(details.get("origin_length") or 12)
    now = _now()
    user_name = user_identity_service.get_current_display_name() or getpass.getuser()
    output_files = _output_files(project_name, reserving_class, payload)
    primary = min(
        output_files,
        key=lambda path: 0 if path.endswith(f"@{origin_length}.csv") else 1,
    )
    canonical_existing: Dict[str, Any] = dict(existing)
    if not existing:
        graph_seed = {
            "dataset_name": output_dataset,
            "dataset_type": _clean(details.get("output_type")) or output_dataset,
            "project_name": project_name,
            "reserving_class": reserving_class,
            "source_kind": "dfm",
            "method_type": dataset_sidecar_status_service.METHOD_TYPE_DFM,
            "precedents": dataset_sidecar_status_service.name_entries(_precedent_names(payload)),
            "dependents": [],
        }
        calculated_dataset_service.apply_sidecar_graph_fields(
            graph_seed,
            project_name,
            graph_seed["dataset_type"],
        )
        canonical_existing = graph_seed
    return _contract_call(
        build_dfm_output_sidecar,
        payload,
        project_name=project_name,
        reserving_class=reserving_class,
        csv_file=os.path.basename(primary),
        existing=canonical_existing,
        existing_record=bool(existing),
        dependents=canonical_existing.get("dependents"),
        notes=notes,
        timestamp=now,
        user=user_name,
        output_changed=bool(changed or not automatic),
        append_audit=bool(not automatic or changed),
        audit_action=(
            AUDIT_ACTION_AUTO_REFRESH if automatic and changed
            else ("Update" if existing else "Insert")
        ),
        status=dataset_sidecar_status_service.STATUS_CURRENT,
    )


def _publish(
    project_name: str,
    reserving_class: str,
    payload: Mapping[str, Any],
    existing_sidecar: Mapping[str, Any],
    *,
    notes: str | None,
    changed: bool,
    automatic: bool,
    write_outputs: bool,
) -> Tuple[Dict[str, Any], List[str]]:
    method_name, output_dataset = _identity(payload)
    method_path = _method_path(project_name, reserving_class, method_name)
    sidecar_path = _sidecar_path(project_name, reserving_class, output_dataset)
    sidecar = _build_sidecar(
        project_name,
        reserving_class,
        payload,
        existing_sidecar,
        notes=notes,
        changed=changed,
        automatic=automatic,
    )
    old_precedents = dataset_sidecar_status_service.entry_names(existing_sidecar.get("precedents"))
    new_precedents = _precedent_names(payload)
    graph_updated = False
    graph_changed = {_key(item) for item in old_precedents} != {_key(item) for item in new_precedents}
    try:
        if graph_changed:
            from app_server.services import result_selection_service

            try:
                result_selection_service._assert_new_precedents_do_not_cycle(
                    project_name,
                    reserving_class,
                    output_dataset,
                    new_precedents,
                )
            except HTTPException as exc:
                raise HTTPException(
                    exc.status_code,
                    str(exc.detail).replace("Result Selection", "DFM"),
                ) from exc
            dataset_sidecar_status_service.update_precedent_dependents(
                project_name,
                reserving_class,
                output_dataset,
                old_precedents,
                new_precedents,
                require_new_precedents=True,
            )
            graph_updated = True
        files = {method_path: _method_json_text(payload)}
        if write_outputs:
            files.update(_output_files(project_name, reserving_class, payload))
        files[sidecar_path] = _json_text(sidecar)
        changed_paths = _commit_text_files(files, last_paths=[sidecar_path])
    except Exception:
        if graph_updated:
            dataset_sidecar_status_service.update_precedent_dependents(
                project_name,
                reserving_class,
                output_dataset,
                new_precedents,
                old_precedents,
                require_new_precedents=False,
            )
        raise
    return sidecar, changed_paths


def _validate_pair(
    requested_method_name: str,
    requested_output_dataset: str,
    method: Mapping[str, Any],
    sidecar: Mapping[str, Any],
) -> None:
    method_name, output_dataset = _identity(method)
    if _key(method_name) != _key(requested_method_name):
        raise HTTPException(409, "DFM method identity does not match the requested method.")
    if _key(output_dataset) != _key(requested_output_dataset):
        raise HTTPException(409, "DFM output identity does not match the requested sidecar.")
    if _key(sidecar.get("dataset_name")) != _key(output_dataset):
        raise HTTPException(409, "DFM sidecar identity does not match the method JSON.")
    sidecar_method = _clean(sidecar.get("method_name")) or output_dataset
    if _key(sidecar_method) != _key(method_name):
        raise HTTPException(409, "DFM sidecar is owned by a different method.")
    if dataset_sidecar_status_service.normalize_method_type(
        sidecar.get("method_type"), sidecar.get("source_kind")
    ) != dataset_sidecar_status_service.METHOD_TYPE_DFM:
        raise HTTPException(409, "DFM output sidecar does not identify a DFM output.")
    method_precedents = {_key(item) for item in _precedent_names(method)}
    sidecar_precedents = {
        _key(item) for item in dataset_sidecar_status_service.entry_names(sidecar.get("precedents"))
    }
    if method_precedents != sidecar_precedents:
        raise HTTPException(409, "DFM method and output sidecar precedents do not match.")
    publication_revision = _revision_response(method)["publication_revision"]
    if _clean(sidecar.get("publication_revision")) != publication_revision:
        raise HTTPException(409, "DFM method and output sidecar publication revisions do not match.")


def _method_response(
    project_name: str,
    reserving_class: str,
    method: Mapping[str, Any],
    sidecar: Mapping[str, Any],
    *,
    upgraded: bool = False,
    changed_paths: Iterable[str] = (),
) -> Dict[str, Any]:
    method_name, output_dataset = _identity(method)
    return {
        "ok": True,
        "project_name": project_name,
        "reserving_class": reserving_class,
        "method_name": method_name,
        "output_dataset": output_dataset,
        "method": dict(method),
        **_revision_response(method),
        "sidecar": _sidecar_response(sidecar, exists=bool(sidecar)),
        "upgraded": upgraded,
        "changed_paths": sorted(changed_paths, key=os.path.normcase),
    }


def load_dfm_method(
    project_name: str,
    reserving_class: str,
    method_name: str,
    *,
    output_dataset: str | None = None,
) -> Dict[str, Any]:
    project = _clean(project_name)
    reserving = _clean(reserving_class)
    name = _clean(method_name)
    if not project or not reserving or not name:
        raise HTTPException(400, "project_name, reserving_class, and method_name are required.")
    method_path = _method_path(project, reserving, name)
    requested_output = _clean(output_dataset)
    if requested_output:
        sidecar_path = _sidecar_path(project, reserving, requested_output)
        with dataset_sidecar_status_service.sidecar_write_lock(sidecar_path):
            method_future = _READ_EXECUTOR.submit(_read_json, method_path)
            sidecar_future = _READ_EXECUTOR.submit(_read_json, sidecar_path)
            method = method_future.result()
            sidecar = sidecar_future.result()
    else:
        with _lock(project, reserving):
            method = _read_json(method_path)
            if not method:
                raise HTTPException(404, f"DFM method not found: {name}")
            _loaded_name, requested_output = _identity(method)
            sidecar_path = _sidecar_path(project, reserving, requested_output)
            with dataset_sidecar_status_service.sidecar_write_lock(sidecar_path):
                sidecar = _read_json(sidecar_path)
    if not method or not sidecar:
        raise HTTPException(409, "DFM requires both its method JSON and output sidecar.")
    json_format = _clean(method.get("json_format"))
    if json_format != DFM_JSON_FORMAT:
        raise HTTPException(422, f"Unsupported DFM JSON format: {json_format or '(missing)'}.")
    normalized = _contract_call(normalize_dfm_method, method, require_complete=True)
    _validate_pair(name, requested_output, normalized, sidecar)
    return _method_response(project, reserving, normalized, sidecar)


def preview_dfm_method(method: Dict[str, Any]) -> Dict[str, Any]:
    payload = _contract_call(canonical_preview_dfm_method, method, timestamp=_now())
    return {"ok": True, "method": payload, **_revision_response(payload)}


def save_dfm_method(
    project_name: str,
    reserving_class: str,
    method: Dict[str, Any],
    *,
    notes: str | None = None,
    expected_owned_revision: str | None = None,
    expected_derived_revision: str | None = None,
) -> Dict[str, Any]:
    project = _clean(project_name)
    reserving = _clean(reserving_class)
    if not project or not reserving:
        raise HTTPException(400, "project_name and reserving_class are required.")
    # Dependent propagation runs on ArcRho Engine; block the save before any
    # write when no live Engine can pick the job up or another walk is still
    # rewriting this reserving class.
    dependent_propagation_service.require_reserving_class_writable(project, reserving)
    incoming = _contract_call(normalize_dfm_method, method, require_complete=False)
    method_name, output_dataset = _identity(incoming)
    method_path = _method_path(project, reserving, method_name)
    sidecar_path = _sidecar_path(project, reserving, output_dataset)
    with _lock(project, reserving), dataset_sidecar_status_service.sidecar_write_lock(sidecar_path):
        current = _read_json(method_path)
        existing_sidecar = _read_json(sidecar_path)
        if current:
            if _clean(current.get("json_format")) != DFM_JSON_FORMAT:
                raise HTTPException(409, "DFM changed on disk; reload it before saving.")
            current = _contract_call(normalize_dfm_method, current, require_complete=True)
            current_method_name, current_output = _identity(current)
            if _key(current_method_name) != _key(method_name) or _key(current_output) != _key(output_dataset):
                raise HTTPException(409, "An existing DFM cannot change its method or output identity during Save.")
            current_revisions = _revision_response(current)
            if expected_owned_revision is not None and _clean(expected_owned_revision) != current_revisions["owned_revision"]:
                raise HTTPException(409, "DFM owned settings changed on disk; reload before saving.")
            merged = _contract_call(apply_owned_patch, current, method, timestamp=_now())
        else:
            if expected_owned_revision is not None and _clean(expected_owned_revision):
                raise HTTPException(409, "DFM was removed on disk; reload before saving.")
            merged = incoming
        if existing_sidecar:
            owner = _clean(existing_sidecar.get("method_name")) or _clean(existing_sidecar.get("dataset_name"))
            if _key(owner) != _key(method_name):
                raise HTTPException(
                    409,
                    f"Output dataset '{output_dataset}' is already owned by DFM '{owner}'. Choose a unique output dataset.",
                )
        current_input = _clean(_details(current).get("input_triangle")) if current else ""
        current_basis = _clean(_results_tab(current).get("ratio_basis_dataset")) if current else ""
        next_input = _clean(_details(merged).get("input_triangle"))
        next_basis = _clean(_results_tab(merged).get("ratio_basis_dataset"))
        load_input = not current or _key(current_input) != _key(next_input)
        load_basis = bool(next_basis) and (not current or _key(current_basis) != _key(next_basis))
        if load_input and next_basis:
            load_basis = True
        if load_input or load_basis:
            refreshed = _recalculate_with_sources(
                project,
                reserving,
                merged,
                load_input=load_input,
                load_basis=load_basis,
                allow_review_needed=True,
                changed_precedents=[
                    name for name, changed_identity in (
                        (next_input, load_input),
                        (next_basis, load_basis),
                    ) if name and changed_identity
                ],
            )
        else:
            # The owned patch was already recalculated by the canonical merger
            # against the latest embedded snapshots. Do not leak out-of-band CSV
            # edits into an ordinary save.
            refreshed = merged
        previous_publication = (
            _revision_response(current)["publication_revision"] if current else ""
        )
        next_publication = _revision_response(refreshed)["publication_revision"]
        publication_changed = not current or previous_publication != next_publication
        sidecar, changed_paths = _publish(
            project,
            reserving,
            refreshed,
            existing_sidecar,
            notes=notes,
            changed=publication_changed,
            automatic=False,
            write_outputs=True,
        )
    response = _method_response(
        project,
        reserving,
        refreshed,
        sidecar,
        changed_paths=changed_paths,
    )
    response["derived_rebased"] = bool(
        current
        and expected_derived_revision is not None
        and _clean(expected_derived_revision) != _revision_response(current)["derived_revision"]
    )
    response["unreviewed_precedents"] = dataset_sidecar_status_service.review_needed_precedent_names(
        project,
        reserving,
        _precedent_names(refreshed),
    )
    response["unreviewed_precedent_count"] = len(response["unreviewed_precedents"])
    output_type = _clean(_details(refreshed).get("output_type")) or output_dataset
    if publication_changed:
        response["propagation"] = _enqueue_propagation_job(
            project, reserving, output_dataset, output_type
        )
    else:
        # A save whose publication revision is unchanged cannot alter any
        # dependent, so no Engine job is submitted.
        response["propagation"] = dependent_propagation_service.unchanged_propagation()
    response["propagation_ok"] = bool(response["propagation"].get("ok"))
    response["calculated_updates"] = response["propagation"]
    return response


def _enqueue_propagation_job(
    project: str,
    reserving: str,
    output_dataset: str,
    output_type: str,
) -> Dict[str, Any]:
    return dependent_propagation_service.enqueue_marked_save_propagation(
        project, reserving, output_dataset, output_type
    )


def save_propagation_roots(
    project_name: str,
    reserving_class: str,
    method: Dict[str, Any],
    **_ignored: Any,
) -> List[Tuple[str, str]]:
    """Return the changed roots ``save_dfm_method`` would propagate from.

    The two-step save plans the dependent closure before anything is written,
    and the roots must be derived exactly the way the save derives them. Save
    refuses to change a DFM's output identity (409), so the incoming payload's
    identity is the identity the save will publish.
    """

    incoming = _contract_call(normalize_dfm_method, method, require_complete=False)
    _method_name, output_dataset = _identity(incoming)
    output_type = _clean(_details(incoming).get("output_type")) or output_dataset
    return [(output_dataset, output_type)]


def _mark_review_needed(project_name: str, reserving_class: str, output_dataset: str) -> None:
    sidecar_path = _sidecar_path(project_name, reserving_class, output_dataset)
    with dataset_sidecar_status_service.sidecar_write_lock(sidecar_path):
        sidecar = _read_json(sidecar_path)
        if not sidecar:
            return
        if dataset_sidecar_status_service.normalize_status(sidecar.get("status")) \
                == dataset_sidecar_status_service.STATUS_REVIEW_NEEDED:
            return
        sidecar["status"] = dataset_sidecar_status_service.STATUS_REVIEW_NEEDED
        _commit_text_files({sidecar_path: _json_text(sidecar)})


def _refresh_one(
    project_name: str,
    reserving_class: str,
    output_dataset: str,
    sidecar: Mapping[str, Any],
    changed_names: Iterable[str],
    snapshot_cache: Dict[SnapshotCacheKey, Dict[str, Any]],
    method_payload: Mapping[str, Any] | None = None,
) -> Dict[str, Any]:
    previous_status = dataset_sidecar_status_service.normalize_status(sidecar.get("status"))
    method_name = _clean(sidecar.get("method_name")) or output_dataset
    method = dict(method_payload) if isinstance(method_payload, Mapping) else _read_json(
        _method_path(project_name, reserving_class, method_name)
    )
    if not method:
        raise RuntimeError("DFM method JSON is missing.")
    if _clean(method.get("json_format")) != DFM_JSON_FORMAT:
        raise RuntimeError("DFM must be upgraded to v2 before automatic refresh.")
    method = _contract_call(normalize_dfm_method, method, require_complete=True)
    input_name = _clean(_details(method).get("input_triangle"))
    basis_name = _clean(_results_tab(method).get("ratio_basis_dataset"))
    changed_keys = {_key(item) for item in changed_names if _key(item)}
    load_input = _key(input_name) in changed_keys
    load_basis = bool(basis_name and _key(basis_name) in changed_keys)
    reference_tokens = dfm_dataset_reference_tokens(method)
    load_references = any(
        _key(token["dataset_name"]) in changed_keys for token in reference_tokens
    )
    if not load_input and not load_basis and not load_references:
        return {"ok": True, "dataset_name": output_dataset, "skipped": True, "reason": "stale_reverse_dependency_edge"}
    # Re-resolve every dataset reference whenever this method recomputes, so
    # User Entry formulas that mix dataset references with average-formula row
    # references are refreshed as one consistent evaluation. A blank or
    # non-numeric referenced cell aborts the refresh; the caller marks the
    # output Review Needed and preserves the last valid publication.
    dataset_reference_values = (
        _resolved_reference_token_values(project_name, reserving_class, reference_tokens)
        if reference_tokens
        else None
    )
    refreshed = _recalculate_with_sources(
        project_name,
        reserving_class,
        method,
        load_input=load_input,
        load_basis=load_basis,
        allow_review_needed=True,
        changed_precedents=changed_names,
        snapshot_cache=snapshot_cache,
        dataset_reference_values=dataset_reference_values,
    )
    _assert_refreshable_precedents(
        project_name,
        reserving_class,
        refreshed,
        snapshot_cache,
        [
            name
            for name, was_loaded in ((input_name, load_input), (basis_name, load_basis))
            if name and was_loaded
        ],
    )
    before_revisions = _revision_response(method)
    after_revisions = _revision_response(refreshed)
    if before_revisions["derived_revision"] == after_revisions["derived_revision"]:
        # No persisted derived value changed. Preserve the prior refresh stamp
        # so the method file remains byte-identical and only a Review Needed
        # sidecar status needs restoration.
        refreshed["method_metadata"]["data_refreshed"] = method["method_metadata"]["data_refreshed"]
    output_changed = (
        before_revisions["publication_revision"]
        != after_revisions["publication_revision"]
    )
    before_text = _method_json_text(method)
    after_text = _method_json_text(refreshed)
    sidecar, changed_paths = _publish(
        project_name,
        reserving_class,
        refreshed,
        sidecar,
        notes=None,
        # A rewritten method is a modification of its output dataset even when
        # the published values held: the sidecar's Last Modified and Audit Log
        # move with the file.
        changed=output_changed or before_text != after_text,
        automatic=True,
        write_outputs=output_changed,
    )
    return {
        "ok": True,
        "dataset_name": output_dataset,
        "updated": before_text != after_text,
        "output_changed": output_changed,
        "status_refreshed": (
            previous_status == dataset_sidecar_status_service.STATUS_REVIEW_NEEDED
            and dataset_sidecar_status_service.normalize_status(sidecar.get("status"))
            == dataset_sidecar_status_service.STATUS_CURRENT
        ),
        "method": refreshed,
        "sidecar": sidecar,
        "changed_paths": changed_paths,
    }


def refresh_dependents(
    project_name: str,
    reserving_class: str,
    changed_dataset_names: Iterable[Any],
    *,
    blocked_precedent_names: Iterable[Any] = (),
    finalize_method_review_status: bool = True,
) -> Dict[str, Any]:
    """Refresh affected DFM methods transitively, without cascading other domains."""

    project = _clean(project_name)
    reserving = _clean(reserving_class)
    changed = []
    seen = set()
    for item in changed_dataset_names:
        name = _clean(item)
        normalized = _key(name)
        if normalized and normalized not in seen:
            seen.add(normalized)
            changed.append(name)
    blocked_keys = {_key(item) for item in blocked_precedent_names if _key(item)}
    updated: List[Dict[str, Any]] = []
    status_refreshed: List[Dict[str, Any]] = []
    skipped: List[Dict[str, Any]] = []
    errors: List[Dict[str, Any]] = []
    snapshot_cache: Dict[SnapshotCacheKey, Dict[str, Any]] = {}
    sidecar_cache: Dict[str, Dict[str, Any]] = {}
    processed_source_keys = set()
    queue = list(changed)
    with _lock(project, reserving):
        while queue:
            frontier = []
            for raw_name in queue:
                name = _clean(raw_name)
                normalized = _key(name)
                if not normalized or normalized in processed_source_keys:
                    continue
                processed_source_keys.add(normalized)
                frontier.append(name)
            queue = []
            if not frontier:
                break
            source_futures = {
                name: _READ_EXECUTOR.submit(_read_json, _sidecar_path(project, reserving, name))
                for name in frontier
                if _key(name) not in sidecar_cache
            }
            for name, future in source_futures.items():
                sidecar_cache[_key(name)] = future.result()
            dependent_sources: Dict[str, List[str]] = {}
            for source_name in frontier:
                source_sidecar = sidecar_cache.get(_key(source_name)) or {}
                for dependent in dataset_sidecar_status_service.entry_names(source_sidecar.get("dependents")):
                    dependent_sources.setdefault(dependent, []).append(source_name)
            dependent_futures = {
                name: _READ_EXECUTOR.submit(_read_json, _sidecar_path(project, reserving, name))
                for name in dependent_sources
                if _key(name) not in sidecar_cache
            }
            for name, future in dependent_futures.items():
                sidecar_cache[_key(name)] = future.result()
            method_futures = {}
            for output_dataset in dependent_sources:
                sidecar = sidecar_cache.get(_key(output_dataset)) or {}
                if dataset_sidecar_status_service.normalize_method_type(
                    sidecar.get("method_type"), sidecar.get("source_kind")
                ) != dataset_sidecar_status_service.METHOD_TYPE_DFM:
                    continue
                method_name = _clean(sidecar.get("method_name")) or output_dataset
                method_futures[output_dataset] = _READ_EXECUTOR.submit(
                    _read_json,
                    _method_path(project, reserving, method_name),
                )
            prefetched_methods = {
                name: future.result() for name, future in method_futures.items()
            }
            for output_dataset in sorted(dependent_sources, key=lambda value: (_key(value), value)):
                sidecar = sidecar_cache.get(_key(output_dataset)) or {}
                if dataset_sidecar_status_service.normalize_method_type(
                    sidecar.get("method_type"), sidecar.get("source_kind")
                ) != dataset_sidecar_status_service.METHOD_TYPE_DFM:
                    continue
                changed_sources = dependent_sources[output_dataset]
                blocked = [name for name in changed_sources if _key(name) in blocked_keys]
                dataset_type = _clean(sidecar.get("dataset_type")) or output_dataset
                if blocked:
                    _mark_review_needed(project, reserving, output_dataset)
                    blocked_keys.update({_key(output_dataset), _key(dataset_type)})
                    sidecar_cache.pop(_key(output_dataset), None)
                    queue.append(output_dataset)
                    errors.append({
                        "dataset_name": output_dataset,
                        "dataset_type": dataset_type,
                        "reason": "Precedent refresh failed: " + ", ".join(blocked),
                    })
                    continue
                try:
                    output_sidecar_path = _sidecar_path(project, reserving, output_dataset)
                    with dataset_sidecar_status_service.sidecar_write_lock(output_sidecar_path):
                        latest_sidecar = _read_json(output_sidecar_path) or sidecar
                        result = _refresh_one(
                            project,
                            reserving,
                            output_dataset,
                            latest_sidecar,
                            changed_sources,
                            snapshot_cache,
                            method_payload=prefetched_methods.get(output_dataset),
                        )
                except Exception as exc:
                    _mark_review_needed(project, reserving, output_dataset)
                    blocked_keys.update({_key(output_dataset), _key(dataset_type)})
                    sidecar_cache.pop(_key(output_dataset), None)
                    queue.append(output_dataset)
                    errors.append({
                        "dataset_name": output_dataset,
                        "dataset_type": dataset_type,
                        "reason": str(exc),
                    })
                    continue
                refreshed_sidecar = result.get("sidecar") or sidecar
                sidecar_cache[_key(output_dataset)] = refreshed_sidecar
                if result.get("updated"):
                    updated.append({
                        "dataset_name": output_dataset,
                        "dataset_type": _clean(refreshed_sidecar.get("dataset_type")) or output_dataset,
                        "output_changed": bool(result.get("output_changed")),
                    })
                    if result.get("output_changed") or result.get("status_refreshed"):
                        queue.append(output_dataset)
                elif result.get("status_refreshed"):
                    status_refreshed.append({"dataset_name": output_dataset})
                    queue.append(output_dataset)
                else:
                    skipped.append({
                        "dataset_name": output_dataset,
                        "reason": result.get("reason") or "not_updated",
                    })
        review_status_updates = (
            dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
                project,
                reserving,
                changed,
            )
            if finalize_method_review_status
            else []
        )
    return {
        "ok": not errors,
        "project_name": project,
        "reserving_class": reserving,
        "changed_dataset_names": changed,
        "updated": updated,
        "status_refreshed": status_refreshed,
        "skipped": skipped,
        "errors": errors,
        "review_status_updates": review_status_updates,
    }


def record_rpc_sync_last_modified(
    project_name: str,
    reserving_class: str,
    method_name: str,
    last_modified: str,
) -> Dict[str, Any]:
    """Record the time ResQ stamped on a method this workspace just uploaded.

    An upload writes ArcRho's settings into the RPC server and ResQ saves them
    under its own ``Modified``. The two copies then hold identical content and
    different times, so the next sync review calls the remote newer and invites
    the user to pull back the values they just pushed. Writing ResQ's own value
    here -- not this machine's clock, which would only move the disagreement to
    whichever side has the faster clock -- makes the pair compare equal.

    Only ``method metadata.last modified`` changes. It is outside every revision
    projection, so an editor open on this method keeps its optimistic-concurrency
    token and no dependent is disturbed.
    """

    project = _clean(project_name)
    reserving = _clean(reserving_class)
    name = _clean(method_name)
    stamped = _clean(last_modified)
    if not project or not reserving or not name:
        raise HTTPException(400, "project_name, reserving_class and method_name are required.")
    if not stamped:
        raise HTTPException(400, "last_modified is required.")
    # The bridge reports the instant ResQ stamped; persist it in the one
    # timestamp form so the file compares equal to what ResQ reports next time.
    stamped = persisted_timestamp(stamped)
    # A propagation walk rewrites whole method files in this class from another
    # process, and this is a read-modify-write of one of them. Stand aside while
    # it owns the class rather than risk reverting what it wrote; the caller
    # reports that beside a sync that has already succeeded. The read-only probe
    # is used rather than the save preflight because recording a timestamp needs
    # no Engine and must not fail merely because none is running.
    if dependent_propagation_service.get_reserving_class_busy(project, reserving)["busy"]:
        return {"ok": False, "status": "class_busy", "last_modified": ""}
    method_path = _method_path(project, reserving, name)
    with _lock(project, reserving):
        current = _read_json(method_path)
        if not current:
            return {"ok": False, "status": "missing", "last_modified": ""}
        if _clean(current.get("json_format")) != DFM_JSON_FORMAT:
            return {"ok": False, "status": "not_v2", "last_modified": ""}
        previous = _clean((current.get("method_metadata") or {}).get("last_modified"))
        if previous == stamped:
            return {"ok": True, "status": "unchanged", "last_modified": stamped}
        updated = _contract_call(stamp_last_modified, current, stamped)
        # The persisted projection is a no-op on an already-persisted payload,
        # so this rewrites the file exactly as its own writer would have.
        changed = _commit_text_files({method_path: _method_json_text(updated)})
    return {
        "ok": True,
        "status": "stamped" if changed else "unchanged",
        "last_modified": stamped,
        "previous_last_modified": previous,
    }
