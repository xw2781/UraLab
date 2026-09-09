"""Load, save, and eagerly refresh self-contained Bootstrap methods.

A Bootstrap is the first ArcRho method whose data precedent is another *method*:
it re-fits a DFM to simulated pseudo triangles.  The reserving-class dependency
graph is keyed by dataset name, so the DFM method is resolved to the dataset it
publishes for every graph operation, while the numbers themselves are read from
the DFM method JSON through ``bootstrap_contract.dfm_snapshot_from_method``.
"""
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

from arcrho_api.bootstrap_contract import (
    BST_JSON_FORMAT,
    BootstrapContractError,
    apply_owned_patch,
    bootstrap_output_variants,
    build_bootstrap_output_sidecar,
    dfm_snapshot_from_method,
    method_revisions,
    normalize_bootstrap_method,
    recalculate_bootstrap_method,
    snapshot_revision,
)
from arcrho_api.dfm_contract import DFM_JSON_FORMAT
from arcrho_api.io import persisted_json_text
from arcrho_api.sidecar_audit_contract import AUDIT_ACTION_AUTO_REFRESH
from arcrho_api.timestamps import utc_now_text
from app_server import config
from app_server.helpers import sanitize_dataset_file_name
from app_server.services import (
    dataset_sidecar_status_service,
    dependent_propagation_service,
    precedent_cache_service,
    user_identity_service,
)


READ_MAX_WORKERS = 4
MAX_REFRESH_VISITS_PER_DATASET = 4
SOURCE_ROLES = ("dfm", "target_ultimate")
_ROLE_LABELS = {
    "dfm": "DFM",
    "target_ultimate": "Target Ultimate",
}
DFM_JSON_FORMATS = (DFM_JSON_FORMAT,)
SnapshotCacheKey = Tuple[str, str, int, Tuple[str, ...]]
_READ_EXECUTOR = ThreadPoolExecutor(
    max_workers=READ_MAX_WORKERS,
    thread_name_prefix="arcrho-bst-read",
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
        dataset_sidecar_status_service.METHOD_TYPE_BOOTSTRAP,
        method_name,
    )


def _dfm_method_path(project_name: str, reserving_class: str, dfm_method_name: str) -> str:
    return dataset_sidecar_status_service.method_json_path(
        project_name,
        reserving_class,
        dataset_sidecar_status_service.METHOD_TYPE_DFM,
        dfm_method_name,
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
        raise HTTPException(423, f"Bootstrap file is locked or inaccessible: {os.path.basename(path)}") from exc
    except (OSError, json.JSONDecodeError) as exc:
        raise HTTPException(500, f"Invalid Bootstrap JSON: {os.path.basename(path)}: {exc}") from exc
    return payload if isinstance(payload, dict) else {}


def _json_text(payload: Mapping[str, Any]) -> str:
    return persisted_json_text(payload)


def _read_bytes_if_file(path: str) -> bytes | None:
    if not os.path.isfile(path):
        return None
    with open(path, "rb") as handle:
        return handle.read()


def _commit_text_files(files: Mapping[str, str], *, last_paths: Iterable[str] = ()) -> List[str]:
    """Replace one Bootstrap publication atomically and restore all prior bytes on failure."""

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
            raise RuntimeError(f"{exc}; Bootstrap rollback failed: {'; '.join(rollback_errors)}") from exc
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
    except BootstrapContractError as exc:
        raise HTTPException(422, str(exc)) from exc
    if not isinstance(result, dict):
        raise HTTPException(500, "Canonical Bootstrap calculation returned an invalid payload.")
    return result


def _details(payload: Mapping[str, Any]) -> Dict[str, Any]:
    value = payload.get("details_tab") if isinstance(payload, Mapping) else None
    return value if isinstance(value, dict) else {}


def _results_tab(payload: Mapping[str, Any]) -> Dict[str, Any]:
    value = payload.get("results_tab") if isinstance(payload, Mapping) else None
    return value if isinstance(value, dict) else {}


def _identity(payload: Mapping[str, Any]) -> Tuple[str, str]:
    details = _details(payload)
    method_name = _clean(details.get("name"))
    output_dataset = method_name
    if not method_name:
        raise HTTPException(422, "Bootstrap method name is required.")
    return method_name, output_dataset


def _role_names(payload: Mapping[str, Any]) -> List[Tuple[str, str]]:
    """Return the configured (role, name) pairs, blanks removed, in role order."""

    return [
        (role, name)
        for role, name in (
            ("dfm", _clean(_details(payload).get("dfm_method"))),
            ("target_ultimate", _clean(_results_tab(payload).get("target_ultimate"))),
        )
        if name
    ]


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


# ---------------------------------------------------------------------------
# The method-to-method edge
# ---------------------------------------------------------------------------


def _read_dfm_payload(project_name: str, reserving_class: str, dfm_method_name: str) -> Dict[str, Any]:
    payload = _read_json(_dfm_method_path(project_name, reserving_class, dfm_method_name))
    if not payload:
        raise HTTPException(404, f"Bootstrap DFM precedent is missing: {dfm_method_name}")
    json_format = _clean(payload.get("json_format") or payload.get("json_format")).lower()
    if json_format not in DFM_JSON_FORMATS:
        raise HTTPException(
            422,
            f"Bootstrap DFM precedent '{dfm_method_name}' uses an unsupported DFM JSON format: "
            f"{json_format or '(missing)'}.",
        )
    return payload


def _dfm_output_dataset(dfm_payload: Mapping[str, Any]) -> str:
    details = dfm_payload.get("details_tab") if isinstance(dfm_payload, Mapping) else None
    details = details if isinstance(details, Mapping) else {}
    return _clean(details.get("output_dataset")) or _clean(details.get("name"))


def _resolve_dfm_output_dataset(
    project_name: str,
    reserving_class: str,
    dfm_method_name: str,
    *,
    dfm_cache: Dict[str, Dict[str, Any]] | None = None,
) -> str:
    """Resolve a DFM method name to the dataset name the graph knows it by.

    The Bootstrap stores a *method* name, but every reverse `dependents` edge,
    cycle check, and Review Needed lookup in the reserving class is keyed by
    dataset name.  A DFM that publishes under its own name resolves to itself.
    """

    name = _clean(dfm_method_name)
    if not name:
        return ""
    cache = dfm_cache if dfm_cache is not None else {}
    cached = cache.get(_key(name))
    if cached is None:
        cached = _read_dfm_payload(project_name, reserving_class, name)
        cache[_key(name)] = cached
    return _dfm_output_dataset(cached) or name


def _precedent_dataset_names(
    project_name: str,
    reserving_class: str,
    method: Mapping[str, Any],
    *,
    dfm_cache: Dict[str, Dict[str, Any]] | None = None,
) -> List[str]:
    """Return the Bootstrap precedents as the dependency graph names them."""

    names: List[str] = []
    for role, name in _role_names(method):
        if role == "dfm":
            names.append(
                _resolve_dfm_output_dataset(
                    project_name, reserving_class, name, dfm_cache=dfm_cache
                )
            )
        else:
            names.append(name)
    return _unique_names(names)


def _assert_precedent_available(
    requested_name: str,
    sidecar: Mapping[str, Any],
    *,
    allow_review_needed: bool,
) -> None:
    if not sidecar:
        raise HTTPException(404, f"Bootstrap precedent sidecar is missing: {requested_name}")
    status = dataset_sidecar_status_service.normalize_status(sidecar.get("status"))
    method_type = dataset_sidecar_status_service.normalize_method_type(
        sidecar.get("method_type"), sidecar.get("source_kind")
    )
    if not allow_review_needed \
            and method_type != dataset_sidecar_status_service.METHOD_TYPE_NONE \
            and status == dataset_sidecar_status_service.STATUS_REVIEW_NEEDED:
        raise HTTPException(409, f"Bootstrap precedent requires review: {requested_name}")


def _read_dfm_snapshot(
    project_name: str,
    reserving_class: str,
    dfm_method_name: str,
    dfm_dataset_name: str,
    sidecar: Mapping[str, Any],
    *,
    allow_review_needed: bool = False,
    dfm_cache: Dict[str, Dict[str, Any]] | None = None,
) -> Dict[str, Any]:
    """Project the precedent DFM method JSON onto the embedded Bootstrap snapshot."""

    _assert_precedent_available(
        dfm_dataset_name, sidecar, allow_review_needed=allow_review_needed
    )
    method_type = dataset_sidecar_status_service.normalize_method_type(
        sidecar.get("method_type"), sidecar.get("source_kind")
    )
    if method_type != dataset_sidecar_status_service.METHOD_TYPE_DFM:
        raise HTTPException(
            422,
            f"Bootstrap DFM source '{dfm_method_name}' must be a DFM method.",
        )
    cache = dfm_cache if dfm_cache is not None else {}
    payload = cache.get(_key(dfm_method_name))
    if payload is None:
        payload = _read_dfm_payload(project_name, reserving_class, dfm_method_name)
        cache[_key(dfm_method_name)] = payload
    return _contract_call(dfm_snapshot_from_method, payload)


def _read_target_snapshot(
    project_name: str,
    reserving_class: str,
    requested_name: str,
    sidecar: Mapping[str, Any],
    *,
    origin_length: int,
    origin_labels: Iterable[Any],
    allow_review_needed: bool = False,
) -> Dict[str, Any]:
    """Read the target ultimate Vector dataset the simulated reserves scale onto.

    ``origin_labels`` is the axis the Bootstrap inherits from its DFM.  The
    canonical contract remaps the target onto that axis by label, so this
    returns the target's own labels whenever it declares them.
    """

    _assert_precedent_available(
        requested_name, sidecar, allow_review_needed=allow_review_needed
    )
    if _clean(sidecar.get("data_format")).lower() != "vector":
        raise HTTPException(
            422,
            f"Bootstrap Target Ultimate source '{requested_name}' must be a Vector dataset.",
        )
    # Stored, not displayed: the CSV opened below is the sidecar's own. A
    # precedent stored finer than the method's period is brought to it, an
    # Engine-generated one by a rebuild and a hand-entered one by a roll-up.
    try:
        csv_path, needs_rollup = precedent_cache_service.precedent_source(
            project_name, reserving_class, requested_name, sidecar, origin_length
        )
    except RuntimeError as exc:
        raise HTTPException(422, f"Bootstrap precedent '{requested_name}' {exc}.") from exc
    if not csv_path:
        raise HTTPException(422, f"Bootstrap precedent '{requested_name}' does not identify its cache CSV.")
    try:
        frame = pd.read_csv(csv_path, header=None, float_precision="round_trip").astype(object)
    except FileNotFoundError as exc:
        raise HTTPException(404, f"Bootstrap precedent CSV is missing: {requested_name}") from exc
    except PermissionError as exc:
        raise HTTPException(423, f"Bootstrap precedent CSV is locked: {requested_name}") from exc
    except Exception as exc:
        raise HTTPException(422, f"Bootstrap precedent CSV is invalid: {requested_name}: {exc}") from exc
    frame = frame.where(pd.notnull(frame), None)
    raw_values = frame.values.tolist()
    if needs_rollup:
        try:
            raw_values = precedent_cache_service.rollup_rows(project_name, sidecar, raw_values, origin_length)
        except ValueError as exc:
            raise HTTPException(
                422,
                f"Bootstrap precedent '{requested_name}' could not be rolled up to the method's period: {exc}",
            ) from exc
    expected_labels = [str(item if item is not None else "") for item in origin_labels]
    if expected_labels and len(raw_values) != len(expected_labels):
        raise HTTPException(
            422,
            f"Bootstrap precedent '{requested_name}' has {len(raw_values)} rows; "
            f"expected {len(expected_labels)}.",
        )
    declared = sidecar.get("origin_labels")
    declared = [str(item if item is not None else "") for item in declared] if isinstance(declared, list) else []
    target_labels = declared if len(declared) == len(raw_values) else expected_labels
    if len(target_labels) != len(raw_values):
        raise HTTPException(
            422,
            f"Bootstrap precedent '{requested_name}' does not identify one origin label per row.",
        )
    return {
        "name": _clean(sidecar.get("dataset_name")) or requested_name,
        "origin_labels": target_labels,
        # The canonical contract reads a flat vector, one value per origin.
        "values": [row[0] if isinstance(row, list) and row else None for row in raw_values],
    }


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
    dfm_cache: Dict[str, Dict[str, Any]] | None = None,
) -> Dict[str, Any]:
    details = _details(method)
    results = _results_tab(method)
    origin_length = int(details.get("origin_length") or 12)
    origin_labels = [
        str(item if item is not None else "")
        for item in results.get("origin_labels", [])
    ]
    requested_roles = set(roles)
    role_names = [
        (role, name) for role, name in _role_names(method) if role in requested_roles
    ]
    graph_names = {
        role: (
            _resolve_dfm_output_dataset(project_name, reserving_class, name, dfm_cache=dfm_cache)
            if role == "dfm"
            else name
        )
        for role, name in role_names
    }
    # One batched sidecar read for every role; only the target CSV read below
    # has a true data dependency on the DFM, so it is the only sequential step.
    sidecars = _read_sidecars(
        project_name,
        reserving_class,
        list(graph_names.values()),
        sidecar_cache,
    )
    snapshots = snapshot_cache if snapshot_cache is not None else {}
    origin_axis = tuple(origin_labels)
    resolved: Dict[str, Dict[str, Any]] = {}
    for role, name in role_names:
        if role != "dfm":
            continue
        cache_key = (_key(name), role, origin_length, origin_axis)
        if cache_key not in snapshots:
            snapshots[cache_key] = _read_dfm_snapshot(
                project_name,
                reserving_class,
                name,
                graph_names[role],
                sidecars.get(graph_names[role]) or {},
                allow_review_needed=allow_review_needed,
                dfm_cache=dfm_cache,
            )
        resolved[role] = snapshots[cache_key]

    # A Bootstrap inherits its origin axis and origin length from its DFM, so a
    # freshly read DFM snapshot — not the method's own stale copy — defines the
    # axis the target ultimate must line up with.  On a first save the method
    # has no axis at all until this point.
    dfm_snapshot = resolved.get("dfm")
    if dfm_snapshot:
        origin_labels = [str(item) for item in dfm_snapshot.get("origin_labels") or []]
        origin_length = int(dfm_snapshot.get("origin_length") or origin_length)

    for role, name in role_names:
        if role == "dfm":
            continue
        cache_key = (_key(name), role, origin_length, tuple(origin_labels))
        if cache_key not in snapshots:
            snapshots[cache_key] = _read_target_snapshot(
                project_name,
                reserving_class,
                graph_names[role],
                sidecars.get(graph_names[role]) or {},
                origin_length=origin_length,
                origin_labels=origin_labels,
                allow_review_needed=allow_review_needed,
            )
        resolved[role] = snapshots[cache_key]
    return resolved


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
    dfm_cache: Dict[str, Dict[str, Any]] | None = None,
) -> Dict[str, Any]:
    snapshots = _source_snapshots(
        project_name,
        reserving_class,
        payload,
        roles,
        allow_review_needed=allow_review_needed,
        sidecar_cache=sidecar_cache,
        snapshot_cache=snapshot_cache,
        dfm_cache=dfm_cache,
    )
    return _contract_call(
        recalculate_bootstrap_method,
        payload,
        dfm_snapshot=snapshots.get("dfm"),
        target_snapshot=snapshots.get("target_ultimate"),
        changed_precedents=changed_precedents,
        timestamp=_now(),
    )


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
    _method_name, output_dataset = _identity(payload)
    data_dir = config.get_project_dataset_cache_dir(project_name, reserving_class)
    safe_name = sanitize_dataset_file_name(output_dataset)
    return {
        os.path.join(data_dir, f"{safe_name}@{period_length}.csv"): _csv_text(values)
        for period_length, values in bootstrap_output_variants(payload).items()
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
    precedents: Iterable[str],
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
    graph_precedents = list(precedents)
    canonical_existing: Dict[str, Any] = dict(existing)
    if not existing:
        graph_seed = {
            "dataset_name": output_dataset,
            "dataset_type": _clean(_details(payload).get("output_type")) or output_dataset,
            "project_name": project_name,
            "reserving_class": reserving_class,
            "source_kind": "bootstrap",
            "method_type": dataset_sidecar_status_service.METHOD_TYPE_BOOTSTRAP,
            "precedents": dataset_sidecar_status_service.name_entries(graph_precedents),
            "dependents": [],
        }
        calculated_dataset_service.apply_sidecar_graph_fields(
            graph_seed,
            project_name,
            graph_seed["dataset_type"],
        )
        canonical_existing = graph_seed
    return _contract_call(
        build_bootstrap_output_sidecar,
        payload,
        project_name=project_name,
        reserving_class=reserving_class,
        csv_file=os.path.basename(primary),
        precedents=graph_precedents,
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
    dfm_cache: Dict[str, Dict[str, Any]] | None = None,
) -> Tuple[Dict[str, Any], List[str]]:
    method_name, output_dataset = _identity(payload)
    sidecar_path = _sidecar_path(project_name, reserving_class, output_dataset)
    new_precedents = _precedent_dataset_names(
        project_name, reserving_class, payload, dfm_cache=dfm_cache
    )
    sidecar = _build_sidecar(
        project_name,
        reserving_class,
        payload,
        existing_sidecar,
        precedents=new_precedents,
        notes=notes,
        changed=changed,
        automatic=automatic,
    )
    old_precedents = dataset_sidecar_status_service.entry_names(existing_sidecar.get("precedents"))
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
                    str(exc.detail).replace("Result Selection", "Bootstrap"),
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
    project_name: str,
    reserving_class: str,
    requested_method_name: str,
    method: Mapping[str, Any],
    sidecar: Mapping[str, Any],
    *,
    dfm_cache: Dict[str, Dict[str, Any]] | None = None,
) -> None:
    method_name, output_dataset = _identity(method)
    if _key(method_name) != _key(requested_method_name):
        raise HTTPException(409, "Bootstrap method identity does not match the requested method.")
    if _key(sidecar.get("dataset_name")) != _key(output_dataset):
        raise HTTPException(409, "Bootstrap sidecar identity does not match the method JSON.")
    sidecar_method = _clean(sidecar.get("method_name")) or output_dataset
    if _key(sidecar_method) != _key(method_name):
        raise HTTPException(409, "Bootstrap sidecar is owned by a different method.")
    if dataset_sidecar_status_service.normalize_method_type(
        sidecar.get("method_type"), sidecar.get("source_kind")
    ) != dataset_sidecar_status_service.METHOD_TYPE_BOOTSTRAP:
        raise HTTPException(409, "Bootstrap output sidecar does not identify a Bootstrap output.")
    if _clean(sidecar.get("data_format")).lower() != "vector":
        raise HTTPException(409, "Bootstrap output sidecar must identify a Vector dataset.")
    # Stored, not displayed: the check is that the CSV this method wrote
    # holds its own periods.
    sidecar_period = precedent_cache_service.source_period(sidecar)
    method_period = int(_details(method).get("origin_length") or 0)
    if sidecar_period != method_period:
        raise HTTPException(409, "Bootstrap method and output sidecar origin lengths do not match.")
    method_origins = [str(item) for item in _results_tab(method).get("origin_labels", [])]
    sidecar_origins = (
        [str(item) for item in sidecar.get("origin_labels", [])]
        if isinstance(sidecar.get("origin_labels"), list)
        else []
    )
    if sidecar_origins != method_origins:
        raise HTTPException(409, "Bootstrap method and output sidecar origin labels do not match.")
    method_precedents = {
        _key(item)
        for item in _precedent_dataset_names(
            project_name, reserving_class, method, dfm_cache=dfm_cache
        )
    }
    sidecar_precedents = {
        _key(item) for item in dataset_sidecar_status_service.entry_names(sidecar.get("precedents"))
    }
    if method_precedents != sidecar_precedents:
        raise HTTPException(409, "Bootstrap method and output sidecar precedents do not match.")
    if _clean(sidecar.get("publication_revision")) != _revision_response(method)["publication_revision"]:
        raise HTTPException(409, "Bootstrap method and output sidecar publication revisions do not match.")


def _method_response(
    project_name: str,
    reserving_class: str,
    method: Mapping[str, Any],
    sidecar: Mapping[str, Any],
    *,
    changed_paths: Iterable[str] = (),
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
    }


def load_bootstrap_method(
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
            raise HTTPException(404, f"Bootstrap method not found: {name}")
        if not sidecar:
            raise HTTPException(409, "Bootstrap requires both its method JSON and output sidecar.")
        json_format = _clean(method.get("json_format"))
        if json_format != BST_JSON_FORMAT:
            raise HTTPException(422, f"Unsupported Bootstrap JSON format: {json_format or '(missing)'}.")
        normalized = _contract_call(
            normalize_bootstrap_method,
            method,
            require_complete=True,
        )
        _validate_pair(project, reserving, name, normalized, sidecar)
        return _method_response(project, reserving, normalized, sidecar)


def _roles_for_save(
    current: Mapping[str, Any] | None,
    merged: Mapping[str, Any],
) -> set[str]:
    if not current:
        return set(SOURCE_ROLES)
    current_names = dict(_role_names(current))
    next_names = dict(_role_names(merged))
    return {
        role
        for role in SOURCE_ROLES
        if _key(current_names.get(role)) != _key(next_names.get(role))
    }


def save_bootstrap_method(
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
        normalize_bootstrap_method,
        method,
        require_complete=False,
    )
    method_name, output_dataset = _identity(incoming)
    method_path = _method_path(project, reserving, method_name)
    sidecar_path = _sidecar_path(project, reserving, output_dataset)
    sidecar_cache: Dict[str, Dict[str, Any]] = {}
    snapshot_cache: Dict[SnapshotCacheKey, Dict[str, Any]] = {}
    dfm_cache: Dict[str, Dict[str, Any]] = {}
    with _lock(project, reserving), dataset_sidecar_status_service.sidecar_write_lock(sidecar_path):
        current_future = _READ_EXECUTOR.submit(_read_json, method_path)
        sidecar_future = _READ_EXECUTOR.submit(_read_json, sidecar_path)
        current = current_future.result()
        existing_sidecar = sidecar_future.result()
        if current:
            if _clean(current.get("json_format")) != BST_JSON_FORMAT:
                raise HTTPException(409, "Bootstrap changed on disk; reload it before saving.")
            current = _contract_call(
                normalize_bootstrap_method,
                current,
                require_complete=True,
            )
            current_name, current_output = _identity(current)
            if _key(current_name) != _key(method_name) or _key(current_output) != _key(output_dataset):
                raise HTTPException(409, "An existing Bootstrap cannot change its method or output identity during Save.")
            current_revisions = _revision_response(current)
            if expected_owned_revision is not None \
                    and _clean(expected_owned_revision) != current_revisions["owned_revision"]:
                raise HTTPException(409, "Bootstrap owned settings changed on disk; reload before saving.")
            merged = _contract_call(apply_owned_patch, current, method, timestamp=_now())
        else:
            if expected_owned_revision is not None and _clean(expected_owned_revision):
                raise HTTPException(409, "Bootstrap was removed on disk; reload before saving.")
            merged = incoming
        if existing_sidecar:
            owner = _clean(existing_sidecar.get("method_name")) or _clean(
                existing_sidecar.get("dataset_name")
            )
            owner_type = dataset_sidecar_status_service.normalize_method_type(
                existing_sidecar.get("method_type"), existing_sidecar.get("source_kind")
            )
            if _key(owner) != _key(method_name) \
                    or owner_type != dataset_sidecar_status_service.METHOD_TYPE_BOOTSTRAP:
                raise HTTPException(
                    409,
                    f"Output dataset '{output_dataset}' is already owned by '{owner}'. Choose a unique Bootstrap name.",
                )
        roles = _roles_for_save(current or None, merged)
        if roles:
            refreshed = _recalculate_with_sources(
                project,
                reserving,
                merged,
                roles,
                changed_precedents=_precedent_dataset_names(
                    project, reserving, merged, dfm_cache=dfm_cache
                ),
                allow_review_needed=True,
                sidecar_cache=sidecar_cache,
                snapshot_cache=snapshot_cache,
                dfm_cache=dfm_cache,
            )
        else:
            refreshed = _contract_call(
                recalculate_bootstrap_method,
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
            dfm_cache=dfm_cache,
        )
        response = _method_response(
            project,
            reserving,
            refreshed,
            published_sidecar,
            changed_paths=changed_paths,
        )
        graph_precedents = _precedent_dataset_names(
            project, reserving, refreshed, dfm_cache=dfm_cache
        )
    response["derived_rebased"] = bool(
        current
        and expected_derived_revision is not None
        and _clean(expected_derived_revision) != _revision_response(current)["derived_revision"]
    )
    response["unreviewed_precedents"] = dataset_sidecar_status_service.review_needed_precedent_names(
        project,
        reserving,
        graph_precedents,
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
    """Return the roots ``save_bootstrap_method`` would propagate from.

    The two-step save plans the dependent closure before anything is written,
    and the roots must be derived exactly the way the save derives them. Save
    refuses to change a Bootstrap method's output identity (409), so the
    incoming payload's identity is the identity the save will publish.
    """

    incoming = _contract_call(
        normalize_bootstrap_method,
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
        raise RuntimeError("Bootstrap method JSON is missing.")
    if _clean(method.get("json_format")) != BST_JSON_FORMAT:
        raise RuntimeError("Bootstrap automatic refresh requires canonical v1 JSON.")
    method = _contract_call(
        normalize_bootstrap_method,
        method,
        require_complete=True,
    )
    dfm_cache: Dict[str, Dict[str, Any]] = {}
    role_names = _role_names(method)
    graph_names = {
        role: (
            _resolve_dfm_output_dataset(project_name, reserving_class, name, dfm_cache=dfm_cache)
            if role == "dfm"
            else name
        )
        for role, name in role_names
    }
    precedent_names = _unique_names(graph_names.values())
    blocked = [name for name in precedent_names if _key(name) in blocked_precedent_keys]
    if blocked:
        # ``blocked_precedent_keys`` holds precedents whose refresh failed,
        # not merely review-flagged ones; the message must say so.
        raise RuntimeError("Required Bootstrap precedent could not be refreshed: " + ", ".join(blocked))
    changed_keys = {_key(name) for name in changed_names if _key(name)}
    matched = [name for name in precedent_names if _key(name) in changed_keys]
    if not matched:
        return {
            "ok": True,
            "dataset_name": output_dataset,
            "skipped": True,
            "reason": "stale_reverse_dependency_edge",
        }
    roles = {role for role, _name in role_names if _key(graph_names[role]) in changed_keys}
    if not roles:
        return {
            "ok": True,
            "dataset_name": output_dataset,
            "skipped": True,
            "reason": "stale_reverse_dependency_edge",
        }
    if roles == {"dfm"}:
        # A 10,000-simulation run costs about 1.4 s under the reserving-class
        # lock, so a DFM save that left the observed triangle and the selected
        # ratios untouched must not trigger one.
        snapshots = _source_snapshots(
            project_name,
            reserving_class,
            method,
            roles,
            allow_review_needed=True,
            sidecar_cache=sidecar_cache,
            snapshot_cache=snapshot_cache,
            dfm_cache=dfm_cache,
        )
        incoming = snapshots.get("dfm") or {}
        if snapshot_revision(incoming) == _clean(_details(method).get("dfm_source_revision")):
            return {
                "ok": True,
                "dataset_name": output_dataset,
                "skipped": True,
                "reason": "dfm_snapshot_unchanged",
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
        dfm_cache=dfm_cache,
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
        dfm_cache=dfm_cache,
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
    for field in (
        "dfm_updates",
        "result_selection_updates",
        "bornhuetter_ferguson_updates",
        "cape_cod_updates",
    ):
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
        include_bootstrap=False,
        finalize_method_review_status=finalize_method_review_status,
        rebuild_index=False,
    )


def refresh_bootstrap_method(
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
            raise HTTPException(404, f"Bootstrap method not found: {name}")
        if not sidecar:
            raise HTTPException(409, "Bootstrap output sidecar is missing.")
        if _clean(method.get("json_format")) != BST_JSON_FORMAT:
            raise HTTPException(422, "Bootstrap refresh requires canonical v1 JSON.")
        was_review_needed = (
            dataset_sidecar_status_service.normalize_status(sidecar.get("status"))
            == dataset_sidecar_status_service.STATUS_REVIEW_NEEDED
        )
        try:
            normalized = _contract_call(
                normalize_bootstrap_method,
                method,
                require_complete=True,
            )
            result = _refresh_one(
                project,
                reserving,
                output_name,
                sidecar,
                _precedent_dataset_names(project, reserving, normalized),
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
    """Refresh Bootstrap reverse-edge branches and feed changed outputs through other domains."""

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
                        "reason": "Bootstrap dependency refresh did not converge.",
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
                ) != dataset_sidecar_status_service.METHOD_TYPE_BOOTSTRAP:
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
                if method_type != dataset_sidecar_status_service.METHOD_TYPE_BOOTSTRAP:
                    skipped.append({
                        "dataset_name": dependent_name,
                        "reason": "non_bootstrap_dependent_handled_by_central_cascade",
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
                            "reason": "Downstream refresh failed after Bootstrap publication"
                            + (": " + "; ".join(reasons) if reasons else "."),
                            "cascade": cascade,
                        })
                except Exception as exc:
                    sidecar_cache.clear()
                    snapshot_cache.clear()
                    errors.append({
                        "dataset_name": dependent_name,
                        "reason": f"Downstream refresh failed after Bootstrap publication: {exc}",
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
