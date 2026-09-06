"""Load, save, and eagerly refresh self-contained Cape Cod methods."""
from __future__ import annotations

import getpass
import json
import math
import os
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, List, Mapping, Tuple

import pandas as pd
from fastapi import HTTPException

from arcrho_api.cape_cod_contract import (
    CC_JSON_FORMAT,
    CapeCodContractError,
    apply_owned_patch,
    build_cape_cod_output_sidecar,
    cape_cod_output_variants,
    cape_cod_precedent_names,
    cape_cod_ultimates_triangle,
    method_revisions,
    normalize_cape_cod_method,
    recalculate_cape_cod_method,
)
from arcrho_api.io import persisted_json_text
from arcrho_api.sidecar_audit_contract import AUDIT_ACTION_AUTO_REFRESH
from arcrho_api.timestamps import utc_now_text
from app_server import config
from app_server.helpers import sanitize_dataset_file_name
from app_server.services import (
    dataset_instance_index_service,
    dataset_sidecar_status_service,
    dependent_propagation_service,
    precedent_cache_service,
    user_identity_service,
)


READ_MAX_WORKERS = 4
MAX_REFRESH_VISITS_PER_DATASET = 4
SOURCE_ROLES = ("latest", "exposure", "prior_ultimate")
_ROLE_LABELS = {
    "latest": "Latest",
    "exposure": "Exposure",
    "prior_ultimate": "Prior Ultimate",
}
_ROLE_NAME_KEYS = {
    "latest": "latest_dataset",
    "exposure": "exposure_dataset",
    "prior_ultimate": "prior_ultimate_dataset",
}
SnapshotCacheKey = Tuple[str, str, int, Tuple[str, ...]]
_READ_EXECUTOR = ThreadPoolExecutor(
    max_workers=READ_MAX_WORKERS,
    thread_name_prefix="arcrho-cc-read",
)


def _clean(value: Any) -> str:
    return str(value if value is not None else "").strip()


def _key(value: Any) -> str:
    return " ".join(_clean(value).lower().split())


def _now() -> str:
    return utc_now_text()


def _lock(project_name: str, reserving_class: str) -> threading.RLock:
    return dataset_sidecar_status_service.reserving_class_io_lock(project_name, reserving_class)


def _method_path(project_name: str, reserving_class: str, method_name: str) -> str:
    return dataset_sidecar_status_service.method_json_path(
        project_name,
        reserving_class,
        dataset_sidecar_status_service.METHOD_TYPE_CAPE_COD,
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
        raise HTTPException(423, f"Cape Cod file is locked or inaccessible: {os.path.basename(path)}") from exc
    except (OSError, json.JSONDecodeError) as exc:
        raise HTTPException(500, f"Invalid Cape Cod JSON: {os.path.basename(path)}: {exc}") from exc
    return payload if isinstance(payload, dict) else {}


def _json_text(payload: Mapping[str, Any]) -> str:
    return persisted_json_text(payload)


def _read_bytes_if_file(path: str) -> bytes | None:
    if not os.path.isfile(path):
        return None
    with open(path, "rb") as handle:
        return handle.read()


def _commit_text_files(files: Mapping[str, str], *, last_paths: Iterable[str] = ()) -> List[str]:
    """Replace one Cape Cod publication atomically and restore all prior bytes on failure."""

    last_keys = {os.path.normcase(os.path.abspath(path)) for path in last_paths}
    paths = list(files)
    read_futures = {path: _READ_EXECUTOR.submit(_read_bytes_if_file, path) for path in paths}
    changed = {
        path: files[path]
        for path in paths
        if read_futures[path].result() != files[path].encode("utf-8")
    }
    ordered_paths = sorted(
        changed,
        key=lambda path: (
            os.path.normcase(os.path.abspath(path)) in last_keys,
            os.path.normcase(path),
        ),
    )
    staged: Dict[str, str] = {}
    backups: Dict[str, bytes | None] = {
        path: read_futures[path].result() for path in ordered_paths
    }
    replaced: List[str] = []
    try:
        for path in ordered_paths:
            os.makedirs(os.path.dirname(path), exist_ok=True)
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
            raise RuntimeError(f"{exc}; Cape Cod rollback failed: {'; '.join(rollback_errors)}") from exc
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
    except CapeCodContractError as exc:
        raise HTTPException(422, str(exc)) from exc
    if not isinstance(result, dict):
        raise HTTPException(500, "Canonical Cape Cod calculation returned an invalid payload.")
    return result


def _details(payload: Mapping[str, Any]) -> Dict[str, Any]:
    value = payload.get("details_tab") if isinstance(payload, Mapping) else None
    return value if isinstance(value, dict) else {}


def _method_tab(payload: Mapping[str, Any]) -> Dict[str, Any]:
    value = payload.get("method_tab") if isinstance(payload, Mapping) else None
    return value if isinstance(value, dict) else {}


def _identity(payload: Mapping[str, Any]) -> Tuple[str, str]:
    details = _details(payload)
    method_name = _clean(details.get("name"))
    output_dataset = method_name
    if not method_name:
        raise HTTPException(422, "Cape Cod method name is required.")
    return method_name, output_dataset


def _unique_names(values: Iterable[Any]) -> List[str]:
    output: List[str] = []
    seen = set()
    for value in values:
        name = _clean(value)
        normalized = _key(name)
        if name and normalized not in seen:
            seen.add(normalized)
            output.append(name)
    return output


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


def _read_source_snapshot_from_sidecar(
    project_name: str,
    reserving_class: str,
    requested_name: str,
    sidecar: Mapping[str, Any],
    *,
    role: str,
    origin_length: int,
    origin_labels: Iterable[Any],
    allow_review_needed: bool = False,
) -> Dict[str, Any]:
    if not sidecar:
        raise HTTPException(404, f"Cape Cod precedent sidecar is missing: {requested_name}")
    status = dataset_sidecar_status_service.normalize_status(sidecar.get("status"))
    method_type = dataset_sidecar_status_service.normalize_method_type(
        sidecar.get("method_type"), sidecar.get("source_kind")
    )
    if not allow_review_needed \
            and method_type != dataset_sidecar_status_service.METHOD_TYPE_NONE \
            and status == dataset_sidecar_status_service.STATUS_REVIEW_NEEDED:
        raise HTTPException(409, f"Cape Cod precedent requires review: {requested_name}")
    data_format = _clean(sidecar.get("data_format")).lower()
    role_label = _ROLE_LABELS.get(role, role.replace("_", " ").title())
    if role == "latest" and data_format != "triangle":
        raise HTTPException(422, f"Cape Cod Latest source '{requested_name}' must be a Triangle dataset.")
    if role != "latest" and data_format != "vector":
        raise HTTPException(422, f"Cape Cod {role_label} source '{requested_name}' must be a Vector dataset.")
    # Stored, not displayed: the CSV opened below is the sidecar's own. A
    # precedent stored finer than the method's period is brought to it, an
    # Engine-generated one by a rebuild and a hand-entered one by a roll-up.
    try:
        csv_path, needs_rollup = precedent_cache_service.precedent_source(
            project_name, reserving_class, requested_name, sidecar, origin_length
        )
    except RuntimeError as exc:
        raise HTTPException(422, f"Cape Cod precedent '{requested_name}' {exc}.") from exc
    if not csv_path:
        raise HTTPException(422, f"Cape Cod precedent '{requested_name}' does not identify its cache CSV.")
    try:
        frame = pd.read_csv(csv_path, header=None).astype(object)
    except FileNotFoundError as exc:
        raise HTTPException(404, f"Cape Cod precedent CSV is missing: {requested_name}") from exc
    except PermissionError as exc:
        raise HTTPException(423, f"Cape Cod precedent CSV is locked: {requested_name}") from exc
    except Exception as exc:
        raise HTTPException(422, f"Cape Cod precedent CSV is invalid: {requested_name}: {exc}") from exc
    frame = frame.where(pd.notnull(frame), None)
    raw_values = frame.values.tolist()
    if needs_rollup:
        try:
            raw_values = precedent_cache_service.rollup_rows(project_name, sidecar, raw_values, origin_length)
        except ValueError as exc:
            raise HTTPException(
                422,
                f"Cape Cod precedent '{requested_name}' could not be rolled up to the method's period: {exc}",
            ) from exc
    method_origin_labels = [str(item if item is not None else "") for item in origin_labels]
    if not method_origin_labels:
        raise HTTPException(422, "Cape Cod method origin labels are required before loading precedents.")
    if len(raw_values) != len(method_origin_labels):
        raise HTTPException(
            422,
            f"Cape Cod precedent '{requested_name}' has {len(raw_values)} rows; "
            f"expected {len(method_origin_labels)}.",
        )
    snapshot = {
        "name": _clean(sidecar.get("dataset_name")) or requested_name,
        "origin_labels": method_origin_labels,
        "values": raw_values,
        "mask": [[value is not None for value in row] for row in raw_values],
    }
    if role == "prior_ultimate" and method_type == dataset_sidecar_status_service.METHOD_TYPE_DFM:
        snapshot["percentage_developed"] = _prior_ultimate_percentage_developed(
            project_name, reserving_class, sidecar, requested_name, len(method_origin_labels)
        )
    return snapshot


def _prior_ultimate_percentage_developed(
    project_name: str,
    reserving_class: str,
    sidecar: Mapping[str, Any],
    requested_name: str,
    row_count: int,
) -> List[Any]:
    """Read the development pattern behind a DFM-published prior ultimate.

    The percentage developed belongs to the DFM's selected development factors,
    so it is read from the DFM method rather than divided out of the published
    ultimates: dividing cannot describe an origin whose latest figure is zero.
    A prior ultimate with no DFM behind it carries no pattern, and Cape Cod
    falls back to the ratio for it.

    The DFM method is read the way a Bootstrap reads its own DFM precedent: the
    dependency graph edge stays on the dataset the DFM publishes, and only the
    embedded snapshot comes from the method JSON.
    """

    method_name = _clean(sidecar.get("method_name")) or requested_name
    values = dataset_instance_index_service.development_pattern_values(
        project_name, reserving_class, method_name
    )
    if len(values) != row_count:
        raise HTTPException(
            422,
            f"Cape Cod Prior Ultimate pattern '{method_name}' has {len(values)} origins; "
            f"expected {row_count}.",
        )
    return values


def _read_sidecars(
    project_name: str,
    reserving_class: str,
    names: Iterable[Any],
    cache: Dict[str, Dict[str, Any]] | None = None,
) -> Dict[str, Dict[str, Any]]:
    snapshot = cache if cache is not None else {}
    unique = _unique_names(names)
    pending = [name for name in unique if _key(name) not in snapshot]
    futures = {
        name: _READ_EXECUTOR.submit(_read_json, _sidecar_path(project_name, reserving_class, name))
        for name in pending
    }
    for name in pending:
        snapshot[_key(name)] = futures[name].result()
    return {name: snapshot.get(_key(name), {}) for name in unique}


def _source_snapshots(
    project_name: str,
    reserving_class: str,
    method: Mapping[str, Any],
    roles: Iterable[str],
    *,
    allow_review_needed: bool = False,
    sidecar_cache: Dict[str, Dict[str, Any]] | None = None,
    snapshot_cache: Dict[SnapshotCacheKey, Dict[str, Any]] | None = None,
) -> Dict[str, Any]:
    tab = _method_tab(method)
    details = _details(method)
    origin_length = int(details.get("origin_length") or 12)
    origin_labels = [
        str(item if item is not None else "")
        for item in tab.get("origin_labels", [])
    ]
    requested_roles = set(roles)
    role_names: List[Tuple[str, str]] = [
        (role, _clean(tab.get(_ROLE_NAME_KEYS[role])))
        for role in SOURCE_ROLES
        if role in requested_roles
    ]
    role_names = [(role, name) for role, name in role_names if name]
    sidecars = _read_sidecars(
        project_name,
        reserving_class,
        [name for _role, name in role_names],
        sidecar_cache,
    )
    snapshots = snapshot_cache if snapshot_cache is not None else {}
    origin_axis = tuple(origin_labels)
    futures: Dict[SnapshotCacheKey, Any] = {}
    for role, name in role_names:
        cache_key = (_key(name), role, origin_length, origin_axis)
        if cache_key not in snapshots and cache_key not in futures:
            futures[cache_key] = _READ_EXECUTOR.submit(
                _read_source_snapshot_from_sidecar,
                project_name,
                reserving_class,
                name,
                sidecars.get(name) or {},
                role=role,
                origin_length=origin_length,
                origin_labels=origin_labels,
                allow_review_needed=allow_review_needed,
            )
    for cache_key, future in futures.items():
        snapshots[cache_key] = future.result()
    return {
        role: snapshots[(_key(name), role, origin_length, origin_axis)]
        for role, name in role_names
    }


def _recalculate_with_sources(
    project_name: str,
    reserving_class: str,
    payload: Mapping[str, Any],
    roles: Iterable[str],
    *,
    changed_precedents: Iterable[str],
    allow_review_needed: bool = False,
    sidecar_cache: Dict[str, Dict[str, Any]] | None = None,
    snapshot_cache: Dict[SnapshotCacheKey, Dict[str, Any]] | None = None,
) -> Dict[str, Any]:
    snapshots = _source_snapshots(
        project_name,
        reserving_class,
        payload,
        roles,
        allow_review_needed=allow_review_needed,
        sidecar_cache=sidecar_cache,
        snapshot_cache=snapshot_cache,
    )
    return _contract_call(
        recalculate_cape_cod_method,
        payload,
        source_snapshots=snapshots,
        changed_precedents=changed_precedents,
        timestamp=_now(),
    )


def _latest_triangle_rows(snapshot: Mapping[str, Any]) -> List[List[Any]]:
    """Observed cumulative Latest rows restricted to the regular n - i shape.

    Row ``i`` keeps its first ``n - i`` cells (oldest origin first); a source
    row that is shorter than ``n - i`` stays short so the canonical contract
    rejects the irregular triangle.
    """

    values = snapshot.get("values") if isinstance(snapshot.get("values"), list) else []
    row_count = len(values)
    rows: List[List[Any]] = []
    for index, raw in enumerate(values):
        row = list(raw) if isinstance(raw, list) else [raw]
        rows.append(row[: max(0, row_count - index)])
    return rows


def _ultimates_triangle(
    project_name: str,
    reserving_class: str,
    method: Mapping[str, Any],
    *,
    sidecar_cache: Dict[str, Dict[str, Any]] | None = None,
    snapshot_cache: Dict[SnapshotCacheKey, Dict[str, Any]] | None = None,
) -> List[List[Any]] | None:
    """Return the as-if diagnostic ultimates triangle, or None when unavailable.

    The triangle is derived display state, so an unavailable, locked, or
    irregular Latest source must degrade this field to None rather than fail
    the surrounding load/save/refresh response.
    """

    try:
        snapshots = _source_snapshots(
            project_name,
            reserving_class,
            method,
            {"latest"},
            allow_review_needed=True,
            sidecar_cache=sidecar_cache,
            snapshot_cache=snapshot_cache,
        )
        latest = snapshots.get("latest")
        if not isinstance(latest, Mapping):
            return None
        return cape_cod_ultimates_triangle(method, _latest_triangle_rows(latest))
    except (HTTPException, CapeCodContractError):
        return None


def _csv_text(values: Iterable[Any]) -> str:
    rows: List[str] = []
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


def _output_files(
    project_name: str,
    reserving_class: str,
    payload: Mapping[str, Any],
) -> Dict[str, str]:
    method_name, output_dataset = _identity(payload)
    del method_name
    data_dir = config.get_project_dataset_cache_dir(project_name, reserving_class)
    safe_name = sanitize_dataset_file_name(output_dataset)
    return {
        os.path.join(data_dir, f"{safe_name}@{period_length}.csv"): _csv_text(values)
        for period_length, values in cape_cod_output_variants(payload).items()
    }


def _output_paths(
    project_name: str,
    reserving_class: str,
    payload: Mapping[str, Any],
) -> List[str]:
    """Project output filenames from geometry without recalculating method values."""

    _method_name, output_dataset = _identity(payload)
    origin_length = int(_details(payload).get("origin_length") or 12)
    periods = [origin_length]
    periods.extend(
        target
        for target in (3, 6, 12)
        if target > origin_length and target % origin_length == 0
    )
    data_dir = config.get_project_dataset_cache_dir(project_name, reserving_class)
    safe_name = sanitize_dataset_file_name(output_dataset)
    return [os.path.join(data_dir, f"{safe_name}@{period}.csv") for period in periods]


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

    _method_name, output_dataset = _identity(payload)
    origin_length = int(_details(payload).get("origin_length") or 12)
    output_files = _output_files(project_name, reserving_class, payload)
    primary = next(
        path for path in output_files if path.endswith(f"@{origin_length}.csv")
    )
    canonical_existing: Dict[str, Any] = dict(existing)
    if not existing:
        graph_seed = {
            "dataset_name": output_dataset,
            "dataset_type": _clean(_details(payload).get("output_type")) or output_dataset,
            "project_name": project_name,
            "reserving_class": reserving_class,
            "source_kind": "cape_cod",
            "method_type": dataset_sidecar_status_service.METHOD_TYPE_CAPE_COD,
            "precedents": dataset_sidecar_status_service.name_entries(
                cape_cod_precedent_names(payload)
            ),
            "dependents": [],
        }
        calculated_dataset_service.apply_sidecar_graph_fields(
            graph_seed,
            project_name,
            graph_seed["dataset_type"],
        )
        canonical_existing = graph_seed
    return _contract_call(
        build_cape_cod_output_sidecar,
        payload,
        project_name=project_name,
        reserving_class=reserving_class,
        csv_file=os.path.basename(primary),
        existing=canonical_existing,
        existing_record=bool(existing),
        dependents=canonical_existing.get("dependents"),
        notes=notes,
        timestamp=_now(),
        user=user_identity_service.get_current_display_name() or getpass.getuser(),
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
    new_precedents = cape_cod_precedent_names(payload)
    graph_changed = {_key(item) for item in old_precedents} != {_key(item) for item in new_precedents}
    graph_updated = False
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
                    str(exc.detail).replace("Result Selection", "Cape Cod"),
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
        files = {_method_path(project_name, reserving_class, method_name): _json_text(payload)}
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
    method: Mapping[str, Any],
    sidecar: Mapping[str, Any],
) -> None:
    method_name, output_dataset = _identity(method)
    if _key(method_name) != _key(requested_method_name):
        raise HTTPException(409, "Cape Cod method identity does not match the requested method.")
    if _key(sidecar.get("dataset_name")) != _key(output_dataset):
        raise HTTPException(409, "Cape Cod sidecar identity does not match the method JSON.")
    sidecar_method = _clean(sidecar.get("method_name")) or output_dataset
    if _key(sidecar_method) != _key(method_name):
        raise HTTPException(409, "Cape Cod sidecar is owned by a different method.")
    if dataset_sidecar_status_service.normalize_method_type(
        sidecar.get("method_type"), sidecar.get("source_kind")
    ) != dataset_sidecar_status_service.METHOD_TYPE_CAPE_COD:
        raise HTTPException(409, "Cape Cod output sidecar does not identify a Cape Cod output.")
    if _clean(sidecar.get("data_format")).lower() != "vector":
        raise HTTPException(409, "Cape Cod output sidecar must identify a Vector dataset.")
    # Stored, not displayed: the check is that the CSV this method wrote
    # holds its own periods.
    sidecar_period = precedent_cache_service.source_period(sidecar)
    method_period = int(_details(method).get("origin_length") or 0)
    if sidecar_period != method_period:
        raise HTTPException(409, "Cape Cod method and output sidecar origin lengths do not match.")
    method_origins = [str(item) for item in _method_tab(method).get("origin_labels", [])]
    sidecar_origins = (
        [str(item) for item in sidecar.get("origin_labels", [])]
        if isinstance(sidecar.get("origin_labels"), list)
        else []
    )
    if sidecar_origins != method_origins:
        raise HTTPException(409, "Cape Cod method and output sidecar origin labels do not match.")
    method_precedents = {_key(item) for item in cape_cod_precedent_names(method)}
    sidecar_precedents = {
        _key(item) for item in dataset_sidecar_status_service.entry_names(sidecar.get("precedents"))
    }
    if method_precedents != sidecar_precedents:
        raise HTTPException(409, "Cape Cod method and output sidecar precedents do not match.")
    if _clean(sidecar.get("publication_revision")) != _revision_response(method)["publication_revision"]:
        raise HTTPException(409, "Cape Cod method and output sidecar publication revisions do not match.")


def _method_response(
    project_name: str,
    reserving_class: str,
    method: Mapping[str, Any],
    sidecar: Mapping[str, Any],
    *,
    changed_paths: Iterable[str] = (),
    sidecar_cache: Dict[str, Dict[str, Any]] | None = None,
    snapshot_cache: Dict[SnapshotCacheKey, Dict[str, Any]] | None = None,
) -> Dict[str, Any]:
    method_name, output_dataset = _identity(method)
    origin_length = int(_details(method).get("origin_length") or 12)
    output_paths = sorted(_output_paths(project_name, reserving_class, method), key=os.path.normcase)
    return {
        "ok": True,
        "project_name": project_name,
        "reserving_class": reserving_class,
        "method_name": method_name,
        "output_dataset": output_dataset,
        "method": dict(method),
        **_revision_response(method),
        "sidecar": _sidecar_response(sidecar, exists=bool(sidecar)),
        "changed_paths": sorted(changed_paths, key=os.path.normcase),
        "aggregated_csv_paths": [
            path for path in output_paths if not path.endswith(f"@{origin_length}.csv")
        ],
        "ultimates_triangle": _ultimates_triangle(
            project_name,
            reserving_class,
            method,
            sidecar_cache=sidecar_cache,
            snapshot_cache=snapshot_cache,
        ),
    }


def load_cape_cod_method(
    project_name: str,
    reserving_class: str,
    method_name: str,
) -> Dict[str, Any]:
    project = _clean(project_name)
    reserving = _clean(reserving_class)
    name = _clean(method_name)
    if not project or not reserving or not name:
        raise HTTPException(400, "project_name, reserving_class, and method_name are required.")
    method_path = _method_path(project, reserving, name)
    sidecar_path = _sidecar_path(project, reserving, name)
    with dataset_sidecar_status_service.sidecar_write_lock(sidecar_path):
        method_future = _READ_EXECUTOR.submit(_read_json, method_path)
        sidecar_future = _READ_EXECUTOR.submit(_read_json, sidecar_path)
        method = method_future.result()
        sidecar = sidecar_future.result()
        if not method:
            raise HTTPException(404, f"Cape Cod method not found: {name}")
        if not sidecar:
            raise HTTPException(409, "Cape Cod requires both its method JSON and output sidecar.")
        json_format = _clean(method.get("json_format"))
        if json_format != CC_JSON_FORMAT:
            raise HTTPException(422, f"Unsupported Cape Cod JSON format: {json_format or '(missing)'}.")
        normalized = _contract_call(
            normalize_cape_cod_method,
            method,
            require_complete=True,
        )
        _validate_pair(name, normalized, sidecar)
        return _method_response(project, reserving, normalized, sidecar)


def _roles_for_save(
    current: Mapping[str, Any] | None,
    merged: Mapping[str, Any],
) -> set[str]:
    if not current:
        return set(SOURCE_ROLES)
    current_details = _details(current)
    next_details = _details(merged)
    current_tab = _method_tab(current)
    next_tab = _method_tab(merged)
    latest_changed = _key(current_tab.get("latest_dataset")) != _key(next_tab.get("latest_dataset"))
    geometry_changed = int(current_details.get("origin_length") or 12) != int(
        next_details.get("origin_length") or 12
    )
    if latest_changed or geometry_changed:
        return set(SOURCE_ROLES)
    return {
        role
        for role in ("exposure", "prior_ultimate")
        if _key(current_tab.get(_ROLE_NAME_KEYS[role])) != _key(next_tab.get(_ROLE_NAME_KEYS[role]))
    }


def save_cape_cod_method(
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
    incoming = _contract_call(
        normalize_cape_cod_method,
        method,
        require_complete=False,
    )
    method_name, output_dataset = _identity(incoming)
    method_path = _method_path(project, reserving, method_name)
    sidecar_path = _sidecar_path(project, reserving, output_dataset)
    sidecar_cache: Dict[str, Dict[str, Any]] = {}
    snapshot_cache: Dict[SnapshotCacheKey, Dict[str, Any]] = {}
    with _lock(project, reserving), dataset_sidecar_status_service.sidecar_write_lock(sidecar_path):
        current_future = _READ_EXECUTOR.submit(_read_json, method_path)
        sidecar_future = _READ_EXECUTOR.submit(_read_json, sidecar_path)
        current = current_future.result()
        existing_sidecar = sidecar_future.result()
        if current:
            if _clean(current.get("json_format")) != CC_JSON_FORMAT:
                raise HTTPException(409, "Cape Cod changed on disk; reload it before saving.")
            current = _contract_call(
                normalize_cape_cod_method,
                current,
                require_complete=True,
            )
            current_name, current_output = _identity(current)
            if _key(current_name) != _key(method_name) or _key(current_output) != _key(output_dataset):
                raise HTTPException(409, "An existing Cape Cod cannot change its method or output identity during Save.")
            current_revisions = _revision_response(current)
            if expected_owned_revision is not None \
                    and _clean(expected_owned_revision) != current_revisions["owned_revision"]:
                raise HTTPException(409, "Cape Cod owned settings changed on disk; reload before saving.")
            merged = _contract_call(apply_owned_patch, current, method, timestamp=_now())
        else:
            if expected_owned_revision is not None and _clean(expected_owned_revision):
                raise HTTPException(409, "Cape Cod was removed on disk; reload before saving.")
            merged = incoming
        if existing_sidecar:
            owner = _clean(existing_sidecar.get("method_name")) or _clean(
                existing_sidecar.get("dataset_name")
            )
            owner_type = dataset_sidecar_status_service.normalize_method_type(
                existing_sidecar.get("method_type"), existing_sidecar.get("source_kind")
            )
            if _key(owner) != _key(method_name) \
                    or owner_type != dataset_sidecar_status_service.METHOD_TYPE_CAPE_COD:
                raise HTTPException(
                    409,
                    f"Output dataset '{output_dataset}' is already owned by '{owner}'. Choose a unique Cape Cod name.",
                )
        roles = _roles_for_save(current or None, merged)
        if roles:
            refreshed = _recalculate_with_sources(
                project,
                reserving,
                merged,
                roles,
                changed_precedents=cape_cod_precedent_names(merged),
                allow_review_needed=True,
                sidecar_cache=sidecar_cache,
                snapshot_cache=snapshot_cache,
            )
        else:
            refreshed = _contract_call(
                recalculate_cape_cod_method,
                merged,
                timestamp=_now(),
                update_refresh_timestamp=False,
            )
        previous_publication = _revision_response(current)["publication_revision"] if current else ""
        next_publication = _revision_response(refreshed)["publication_revision"]
        publication_changed = not current or previous_publication != next_publication
        published_sidecar, changed_paths = _publish(
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
            published_sidecar,
            changed_paths=changed_paths,
            sidecar_cache=sidecar_cache,
            snapshot_cache=snapshot_cache,
        )
    response["derived_rebased"] = bool(
        current
        and expected_derived_revision is not None
        and _clean(expected_derived_revision) != _revision_response(current)["derived_revision"]
    )
    response["unreviewed_precedents"] = dataset_sidecar_status_service.review_needed_precedent_names(
        project,
        reserving,
        cape_cod_precedent_names(refreshed),
    )
    response["unreviewed_precedent_count"] = len(response["unreviewed_precedents"])
    if publication_changed:
        response["propagation"] = dependent_propagation_service.enqueue_marked_save_propagation(
            project,
            reserving,
            output_dataset,
            _clean(_details(refreshed).get("output_type")) or output_dataset,
        )
    else:
        response["propagation"] = dependent_propagation_service.unchanged_propagation()
    response["propagation_ok"] = bool(response["propagation"].get("ok"))
    response["calculated_updates"] = response["propagation"]
    response["index_ok"] = bool(response["propagation"].get("index_ok", True))
    response["index_error"] = _clean(response["propagation"].get("index_error"))
    return response


def save_propagation_roots(
    project_name: str,
    reserving_class: str,
    method: Dict[str, Any],
    **_ignored: Any,
) -> List[Tuple[str, str]]:
    """Return the changed roots ``save_cape_cod_method`` would propagate from.

    The two-step save plans the dependent closure before anything is written,
    and the roots must be derived exactly the way the save derives them. Save
    refuses to change a Cape Cod's output identity (409), so the incoming
    payload's identity is the identity the save will publish.
    """

    incoming = _contract_call(
        normalize_cape_cod_method,
        method,
        require_complete=False,
    )
    _method_name, output_dataset = _identity(incoming)
    output_type = _clean(_details(incoming).get("output_type")) or output_dataset
    return [(output_dataset, output_type)]


def _mark_review_needed(
    project_name: str,
    reserving_class: str,
    output_dataset: str,
) -> Dict[str, Any]:
    sidecar_path = _sidecar_path(project_name, reserving_class, output_dataset)
    with dataset_sidecar_status_service.sidecar_write_lock(sidecar_path):
        sidecar = _read_json(sidecar_path)
        if not sidecar:
            return {}
        if dataset_sidecar_status_service.normalize_status(sidecar.get("status")) \
                == dataset_sidecar_status_service.STATUS_REVIEW_NEEDED:
            return sidecar
        sidecar["status"] = dataset_sidecar_status_service.STATUS_REVIEW_NEEDED
        _commit_text_files({sidecar_path: _json_text(sidecar)})
        return sidecar


def _refresh_one(
    project_name: str,
    reserving_class: str,
    output_dataset: str,
    sidecar: Mapping[str, Any],
    changed_names: Iterable[str],
    *,
    blocked_precedent_keys: set[str],
    sidecar_cache: Dict[str, Dict[str, Any]],
    snapshot_cache: Dict[SnapshotCacheKey, Dict[str, Any]],
    method_payload: Mapping[str, Any] | None = None,
) -> Dict[str, Any]:
    method_name = _clean(sidecar.get("method_name")) or output_dataset
    method = dict(method_payload) if isinstance(method_payload, Mapping) else _read_json(
        _method_path(project_name, reserving_class, method_name)
    )
    if not method:
        raise RuntimeError("Cape Cod method JSON is missing.")
    if _clean(method.get("json_format")) != CC_JSON_FORMAT:
        raise RuntimeError("Cape Cod automatic refresh requires canonical v1 JSON.")
    method = _contract_call(
        normalize_cape_cod_method,
        method,
        require_complete=True,
    )
    tab = _method_tab(method)
    precedent_names = cape_cod_precedent_names(method)
    blocked = [name for name in precedent_names if _key(name) in blocked_precedent_keys]
    if blocked:
        # ``blocked_precedent_keys`` holds precedents whose refresh failed,
        # not merely review-flagged ones; the message must say so.
        raise RuntimeError("Required Cape Cod precedent could not be refreshed: " + ", ".join(blocked))
    changed_keys = {_key(name) for name in changed_names if _key(name)}
    matched = [name for name in precedent_names if _key(name) in changed_keys]
    if not matched:
        return {
            "ok": True,
            "dataset_name": output_dataset,
            "skipped": True,
            "reason": "stale_reverse_dependency_edge",
        }
    if _key(_clean(tab.get("latest_dataset"))) in changed_keys:
        roles = set(SOURCE_ROLES)
    else:
        roles = {
            role
            for role in ("exposure", "prior_ultimate")
            if _key(_clean(tab.get(_ROLE_NAME_KEYS[role]))) in changed_keys
        }
    if not roles:
        return {
            "ok": True,
            "dataset_name": output_dataset,
            "skipped": True,
            "reason": "stale_reverse_dependency_edge",
        }
    refreshed = _recalculate_with_sources(
        project_name,
        reserving_class,
        method,
        roles,
        changed_precedents=matched,
        allow_review_needed=True,
        sidecar_cache=sidecar_cache,
        snapshot_cache=snapshot_cache,
    )
    before_revisions = _revision_response(method)
    after_revisions = _revision_response(refreshed)
    if before_revisions["derived_revision"] == after_revisions["derived_revision"]:
        # Preserve method bytes when a source save did not change the embedded
        # snapshot; only a Review Needed sidecar may need restoration.
        refreshed["method_metadata"]["data_refreshed"] = method["method_metadata"]["data_refreshed"]
    output_changed = (
        before_revisions["publication_revision"]
        != after_revisions["publication_revision"]
    )
    before_text = _json_text(method)
    after_text = _json_text(refreshed)
    updated_sidecar, changed_paths = _publish(
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
        "dataset_type": _clean(_details(refreshed).get("output_type")) or output_dataset,
        "updated": before_text != after_text,
        "output_changed": output_changed,
        "status_refreshed": (
            dataset_sidecar_status_service.normalize_status(sidecar.get("status"))
            == dataset_sidecar_status_service.STATUS_REVIEW_NEEDED
            and dataset_sidecar_status_service.normalize_status(updated_sidecar.get("status"))
            == dataset_sidecar_status_service.STATUS_CURRENT
        ),
        "method": refreshed,
        "sidecar": updated_sidecar,
        "changed_paths": changed_paths,
    }


def _cascade_names(report: Mapping[str, Any]) -> Tuple[List[str], List[str]]:
    fresh: List[str] = []
    failed: List[str] = []
    fresh.extend(
        _clean(item.get("dataset_type_name"))
        for item in report.get("updated", [])
        if isinstance(item, Mapping) and _clean(item.get("dataset_type_name"))
    )
    failed.extend(
        _clean(item.get("dataset_type_name"))
        for item in report.get("skipped", [])
        if isinstance(item, Mapping) and _clean(item.get("dataset_type_name"))
    )
    for field in ("dfm_updates", "result_selection_updates", "bornhuetter_ferguson_updates"):
        domain = report.get(field) if isinstance(report.get(field), Mapping) else {}
        for result_field in ("updated", "status_refreshed"):
            fresh.extend(
                _clean(item.get("dataset_name") or item.get("dataset_type"))
                for item in domain.get(result_field, [])
                if isinstance(item, Mapping)
                and _clean(item.get("dataset_name") or item.get("dataset_type"))
            )
        failed.extend(
            _clean(item.get("dataset_name") or item.get("dataset_type"))
            for item in domain.get("errors", [])
            if isinstance(item, Mapping)
            and _clean(item.get("dataset_name") or item.get("dataset_type"))
        )
        if field == "result_selection_updates":
            fresh.extend(
                _clean(name)
                for name in domain.get("downstream_fresh_names", [])
                if _clean(name)
            )
            failed.extend(
                _clean(name)
                for name in domain.get("downstream_blocked_names", [])
                if _clean(name)
            )
    return _unique_names(fresh), _unique_names(failed)


def _refresh_downstream_domains(
    project_name: str,
    reserving_class: str,
    output_name: str,
    output_type: str,
    *,
    finalize_method_review_status: bool = True,
) -> Dict[str, Any]:
    from app_server.services import calculated_dataset_service

    return calculated_dataset_service.recalculate_dependents(
        project_name,
        reserving_class,
        output_name,
        output_type,
        include_cape_cod=False,
        include_bootstrap=False,
        finalize_method_review_status=finalize_method_review_status,
        rebuild_index=False,
    )


def refresh_cape_cod_method(
    project_name: str,
    reserving_class: str,
    method_name: str,
) -> Dict[str, Any]:
    project = _clean(project_name)
    reserving = _clean(reserving_class)
    name = _clean(method_name)
    if not project or not reserving or not name:
        raise HTTPException(400, "project_name, reserving_class, and method_name are required.")
    dependent_propagation_service.require_reserving_class_writable(project, reserving)
    output_name = name
    sidecar_path = _sidecar_path(project, reserving, output_name)
    sidecar_cache: Dict[str, Dict[str, Any]] = {}
    snapshot_cache: Dict[SnapshotCacheKey, Dict[str, Any]] = {}
    with _lock(project, reserving), dataset_sidecar_status_service.sidecar_write_lock(sidecar_path):
        method_future = _READ_EXECUTOR.submit(
            _read_json,
            _method_path(project, reserving, name),
        )
        sidecar_future = _READ_EXECUTOR.submit(_read_json, sidecar_path)
        method = method_future.result()
        sidecar = sidecar_future.result()
        if not method:
            raise HTTPException(404, f"Cape Cod method not found: {name}")
        if not sidecar:
            raise HTTPException(409, "Cape Cod output sidecar is missing.")
        if _clean(method.get("json_format")) != CC_JSON_FORMAT:
            raise HTTPException(422, "Cape Cod refresh requires canonical v1 JSON.")
        was_review_needed = (
            dataset_sidecar_status_service.normalize_status(sidecar.get("status"))
            == dataset_sidecar_status_service.STATUS_REVIEW_NEEDED
        )
        try:
            result = _refresh_one(
                project,
                reserving,
                output_name,
                sidecar,
                cape_cod_precedent_names(method),
                blocked_precedent_keys=set(),
                sidecar_cache=sidecar_cache,
                snapshot_cache=snapshot_cache,
                method_payload=method,
            )
            if was_review_needed:
                _mark_review_needed(project, reserving, output_name)
                result["sidecar"] = _read_json(sidecar_path) or result.get("sidecar") or sidecar
                result["status_refreshed"] = False
        except Exception:
            _mark_review_needed(project, reserving, output_name)
            dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
                project,
                reserving,
                [output_name],
            )
            raise
    response = _method_response(
        project,
        reserving,
        result.get("method") or method,
        result.get("sidecar") or sidecar,
        changed_paths=result.get("changed_paths") or [],
        sidecar_cache=sidecar_cache,
        snapshot_cache=snapshot_cache,
    )
    response.update({
        "updated": bool(result.get("updated")),
        "output_changed": bool(result.get("output_changed")),
        "status_refreshed": bool(result.get("status_refreshed")),
    })
    if response["output_changed"] or response["status_refreshed"]:
        response["propagation"] = dependent_propagation_service.enqueue_marked_save_propagation(
            project,
            reserving,
            output_name,
            _clean(result.get("dataset_type")) or output_name,
        )
    else:
        response["propagation"] = {
            "ok": True,
            "skipped": True,
            "reason": "publication_unchanged",
        }
    response["propagation_ok"] = bool(response["propagation"].get("ok"))
    response["calculated_updates"] = response["propagation"]
    response["index_ok"] = bool(response["propagation"].get("index_ok", True))
    response["index_error"] = _clean(response["propagation"].get("index_error"))
    return response


def refresh_dependents(
    project_name: str,
    reserving_class: str,
    changed_dataset_names: Iterable[Any],
    *,
    rebuild_index: bool = True,
    blocked_precedent_names: Iterable[Any] = (),
    finalize_method_review_status: bool = True,
) -> Dict[str, Any]:
    """Refresh Cape Cod reverse-edge branches and feed changed outputs through other domains."""

    project = _clean(project_name)
    reserving = _clean(reserving_class)
    changed_names = _unique_names(changed_dataset_names)
    queue = list(changed_names)
    blocked_keys = {_key(name) for name in blocked_precedent_names if _key(name)}
    sidecar_cache: Dict[str, Dict[str, Any]] = {}
    snapshot_cache: Dict[SnapshotCacheKey, Dict[str, Any]] = {}
    visit_counts: Dict[str, int] = {}
    updated: List[Dict[str, Any]] = []
    status_refreshed: List[Dict[str, Any]] = []
    skipped: List[Dict[str, Any]] = []
    errors: List[Dict[str, Any]] = []
    index_error = ""
    with _lock(project, reserving):
        while queue:
            frontier = _unique_names(queue)
            queue = []
            allowed_frontier: List[str] = []
            for name in frontier:
                normalized = _key(name)
                visit_counts[normalized] = visit_counts.get(normalized, 0) + 1
                if visit_counts[normalized] > MAX_REFRESH_VISITS_PER_DATASET:
                    errors.append({
                        "dataset_name": name,
                        "reason": "Cape Cod dependency refresh did not converge.",
                    })
                    continue
                allowed_frontier.append(name)
            if not allowed_frontier:
                continue
            source_sidecars = _read_sidecars(
                project,
                reserving,
                allowed_frontier,
                sidecar_cache,
            )
            dependent_sources: Dict[str, Dict[str, Dict[str, Any]]] = {}
            for source_name in allowed_frontier:
                source_sidecar = source_sidecars.get(source_name) or {}
                for dependent_name in dataset_sidecar_status_service.entry_names(
                    source_sidecar.get("dependents")
                ):
                    dependent_sources.setdefault(dependent_name, {})[source_name] = source_sidecar
            if not dependent_sources:
                continue
            dependent_sidecars = _read_sidecars(
                project,
                reserving,
                dependent_sources,
                sidecar_cache,
            )
            method_paths_by_dependent: Dict[str, str] = {}
            for dependent_name, sidecar in dependent_sidecars.items():
                if dataset_sidecar_status_service.normalize_method_type(
                    sidecar.get("method_type"), sidecar.get("source_kind")
                ) != dataset_sidecar_status_service.METHOD_TYPE_CAPE_COD:
                    continue
                method_name = _clean(sidecar.get("method_name")) or dependent_name
                method_paths_by_dependent[dependent_name] = _method_path(
                    project,
                    reserving,
                    method_name,
                )
            method_futures = {
                path: _READ_EXECUTOR.submit(_read_json, path)
                for path in set(method_paths_by_dependent.values())
            }
            dependent_methods = {
                dependent_name: method_futures[path].result()
                for dependent_name, path in method_paths_by_dependent.items()
            }
            for dependent_name in sorted(dependent_sources, key=lambda item: (_key(item), item)):
                sidecar = dependent_sidecars.get(dependent_name) or {}
                if not sidecar:
                    errors.append({
                        "dataset_name": dependent_name,
                        "reason": "dependency_sidecar_missing",
                    })
                    blocked_keys.add(_key(dependent_name))
                    continue
                method_type = dataset_sidecar_status_service.normalize_method_type(
                    sidecar.get("method_type"), sidecar.get("source_kind")
                )
                if method_type != dataset_sidecar_status_service.METHOD_TYPE_CAPE_COD:
                    skipped.append({
                        "dataset_name": dependent_name,
                        "reason": "non_cape_cod_dependent_handled_by_central_cascade",
                    })
                    continue
                try:
                    with dataset_sidecar_status_service.sidecar_write_lock(
                        _sidecar_path(project, reserving, dependent_name)
                    ):
                        result = _refresh_one(
                            project,
                            reserving,
                            dependent_name,
                            sidecar,
                            dependent_sources[dependent_name],
                            blocked_precedent_keys=blocked_keys,
                            sidecar_cache=sidecar_cache,
                            snapshot_cache=snapshot_cache,
                            method_payload=dependent_methods.get(dependent_name, {}),
                        )
                except Exception as exc:
                    blocked_keys.add(_key(dependent_name))
                    review_sidecar = _mark_review_needed(project, reserving, dependent_name)
                    if review_sidecar:
                        sidecar_cache[_key(dependent_name)] = review_sidecar
                    touched = dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
                        project,
                        reserving,
                        [dependent_name],
                    )
                    for item in touched:
                        sidecar_cache.pop(_key(item.get("dataset_name")), None)
                    errors.append({"dataset_name": dependent_name, "reason": str(exc)})
                    queue.append(dependent_name)
                    continue
                refreshed_sidecar = result.get("sidecar") or sidecar
                sidecar_cache[_key(dependent_name)] = refreshed_sidecar
                if result.get("updated"):
                    updated.append({
                        "dataset_name": dependent_name,
                        "dataset_type": result.get("dataset_type") or dependent_name,
                        "output_changed": bool(result.get("output_changed")),
                    })
                if result.get("status_refreshed"):
                    status_refreshed.append({"dataset_name": dependent_name})
                if not result.get("updated"):
                    skipped.append({
                        "dataset_name": dependent_name,
                        "reason": result.get("reason") or (
                            "status_refreshed" if result.get("status_refreshed") else "not_updated"
                        ),
                    })
                if not result.get("output_changed") and not result.get("status_refreshed"):
                    continue
                blocked_keys.discard(_key(dependent_name))
                touched = dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
                    project,
                    reserving,
                    [dependent_name],
                )
                for item in touched:
                    sidecar_cache.pop(_key(item.get("dataset_name")), None)
                queue.append(dependent_name)
                try:
                    cascade = _refresh_downstream_domains(
                        project,
                        reserving,
                        dependent_name,
                        _clean(result.get("dataset_type")) or dependent_name,
                        finalize_method_review_status=False,
                    )
                    fresh_names, failed_names = _cascade_names(cascade)
                    queue.extend(fresh_names)
                    queue.extend(failed_names)
                    blocked_keys.difference_update(_key(name) for name in fresh_names)
                    blocked_keys.update(_key(name) for name in failed_names)
                    for name in [*fresh_names, *failed_names]:
                        sidecar_cache.pop(_key(name), None)
                    if not cascade.get("ok", True):
                        from app_server.services import calculated_dataset_service

                        reasons = calculated_dataset_service.cascade_failure_reasons(cascade)
                        errors.append({
                            "dataset_name": dependent_name,
                            "reason": "Downstream refresh failed after Cape Cod publication"
                            + (": " + "; ".join(reasons) if reasons else "."),
                            "cascade": cascade,
                        })
                except Exception as exc:
                    sidecar_cache.clear()
                    snapshot_cache.clear()
                    errors.append({
                        "dataset_name": dependent_name,
                        "reason": f"Downstream refresh failed after Cape Cod publication: {exc}",
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
        by_key: Dict[str, Dict[str, Any]] = {}
        for item in items:
            name = _clean(item.get("dataset_name"))
            if name:
                by_key.setdefault(_key(name), item)
        return [by_key[key] for key in sorted(by_key)]

    return {
        "ok": not errors,
        "project_name": project,
        "reserving_class": reserving,
        "changed_dataset_names": changed_names,
        "updated": unique_updates(updated),
        "status_refreshed": unique_updates(status_refreshed),
        "skipped": skipped,
        "errors": errors,
        "review_status_updates": review_status_updates,
        "index_ok": not index_error,
        "index_error": index_error,
    }
