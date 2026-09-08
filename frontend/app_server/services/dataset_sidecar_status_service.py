"""Dataset sidecar method dependency status helpers."""
from __future__ import annotations

import json
import os
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from contextlib import ExitStack
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, List, Set

from arcrho_api.dataset_link_contract import expand_sidecar_links
from arcrho_api.io import persisted_json_text
from arcrho_api.sidecar_core_contract import dependency_entries, dependency_names, finalize_sidecar
from arcrho_api.timestamps import utc_now_text
from app_server import config
from app_server.helpers import _canon_dataset_name, sanitize_dataset_file_name

METHOD_TYPE_NONE = "None"
METHOD_TYPE_DFM = "DFM"
METHOD_TYPE_RESULT_SELECTION = "Result Selection"
METHOD_TYPE_BORN_HUETTER_FERGUSON = "Bornhuetter Ferguson"
METHOD_TYPE_CAPE_COD = "Cape Cod"
METHOD_TYPE_BOOTSTRAP = "Bootstrap"
METHOD_TYPE_BERQUIST_SHERMAN_SR = "B&S Settlement Rate Adjustment"
METHOD_TYPE_BERQUIST_SHERMAN_CRA = "B&S Case Reserve Adequacy Adjustment"
SOURCE_KIND_BERQUIST_SHERMAN_SR = "berquist_sherman_sr"
SOURCE_KIND_BERQUIST_SHERMAN_CRA = "berquist_sherman_cra"
STATUS_CURRENT = 0
STATUS_REVIEW_NEEDED = 2
_SIDECAR_READ_EXECUTOR = ThreadPoolExecutor(max_workers=6, thread_name_prefix="arcrho-graph-read")
_SIDECAR_WRITE_LOCKS_GUARD = threading.Lock()
_SIDECAR_WRITE_LOCKS: Dict[str, threading.RLock] = {}
_RESERVING_CLASS_LOCKS_GUARD = threading.Lock()
_RESERVING_CLASS_LOCKS: Dict[str, threading.RLock] = {}

_CANONICAL_METHOD_TYPES = {
    METHOD_TYPE_DFM.lower(): METHOD_TYPE_DFM,
    METHOD_TYPE_RESULT_SELECTION.lower(): METHOD_TYPE_RESULT_SELECTION,
    METHOD_TYPE_BORN_HUETTER_FERGUSON.lower(): METHOD_TYPE_BORN_HUETTER_FERGUSON,
    METHOD_TYPE_CAPE_COD.lower(): METHOD_TYPE_CAPE_COD,
    METHOD_TYPE_BOOTSTRAP.lower(): METHOD_TYPE_BOOTSTRAP,
    METHOD_TYPE_BERQUIST_SHERMAN_SR.lower(): METHOD_TYPE_BERQUIST_SHERMAN_SR,
    METHOD_TYPE_BERQUIST_SHERMAN_CRA.lower(): METHOD_TYPE_BERQUIST_SHERMAN_CRA,
}
_METHOD_TYPE_BY_SOURCE_KIND = {
    "dfm": METHOD_TYPE_DFM,
    "result_selection": METHOD_TYPE_RESULT_SELECTION,
    "bornhuetter_ferguson": METHOD_TYPE_BORN_HUETTER_FERGUSON,
    "cape_cod": METHOD_TYPE_CAPE_COD,
    "bootstrap": METHOD_TYPE_BOOTSTRAP,
    SOURCE_KIND_BERQUIST_SHERMAN_SR: METHOD_TYPE_BERQUIST_SHERMAN_SR,
    SOURCE_KIND_BERQUIST_SHERMAN_CRA: METHOD_TYPE_BERQUIST_SHERMAN_CRA,
}
METHOD_JSON_FILENAME_PREFIX_BY_TYPE = {
    METHOD_TYPE_DFM: "DFM@",
    METHOD_TYPE_RESULT_SELECTION: "RS@",
    METHOD_TYPE_BORN_HUETTER_FERGUSON: "BF@",
    METHOD_TYPE_CAPE_COD: "CC@",
    METHOD_TYPE_BOOTSTRAP: "BST@",
    METHOD_TYPE_BERQUIST_SHERMAN_SR: "BSSR@",
    METHOD_TYPE_BERQUIST_SHERMAN_CRA: "BSCRA@",
}


def _clean_text(value: Any) -> str:
    return str(value if value is not None else "").strip()


def normalize_method_type(value: Any = "", source_kind: Any = "") -> str:
    text = _clean_text(value)
    if text and text.lower() not in {"none", "null"}:
        normalized = text.lower().replace("_", " ")
        return _CANONICAL_METHOD_TYPES.get(normalized, text)
    source = _clean_text(source_kind).lower()
    return _METHOD_TYPE_BY_SOURCE_KIND.get(source, METHOD_TYPE_NONE)


def normalize_status(value: Any) -> int:
    try:
        status = int(value)
    except (TypeError, ValueError):
        return STATUS_CURRENT
    return STATUS_REVIEW_NEEDED if status == STATUS_REVIEW_NEEDED else STATUS_CURRENT


def name_entries(names: Iterable[Any]) -> List[Dict[str, str]]:
    """Persisted dependency entries for plain names (``arcrho_api.sidecar_core_contract``)."""
    return dependency_entries(list(names or []))


def entry_names(entries: Any) -> List[str]:
    """The dataset names of persisted dependency entries, in order."""
    return dependency_names(entries)


def review_needed_precedent_names(
    project_name: str,
    reserving_class: str,
    precedents: Any,
) -> List[str]:
    """Return method-backed precedents awaiting human review in input order."""
    names = entry_names(precedents)
    futures = {
        name: _SIDECAR_READ_EXECUTOR.submit(
            read_sidecar,
            sidecar_path(project_name, reserving_class, name),
        )
        for name in names
    }
    review_needed: List[str] = []
    for name in names:
        payload = futures[name].result()
        if not payload:
            continue
        method_type = normalize_method_type(
            payload.get("method_type"),
            payload.get("source_kind"),
        )
        if method_type != METHOD_TYPE_NONE \
                and normalize_status(payload.get("status")) == STATUS_REVIEW_NEEDED:
            review_needed.append(name)
    return review_needed


def merge_name_entries(*entry_lists: Any) -> List[Dict[str, str]]:
    names: List[str] = []
    for entries in entry_lists:
        names.extend(entry_names(entries))
    return name_entries(names)


def sidecar_path(project_name: str, reserving_class: str, dataset_name: str) -> str:
    return os.path.join(
        config.get_project_dataset_sidecar_dir(project_name, reserving_class),
        f"{sanitize_dataset_file_name(dataset_name)}.json",
    )


def method_json_path(
    project_name: str,
    reserving_class: str,
    method_type: Any,
    method_name: str,
) -> str:
    """Return the canonical ``<PREFIX>@<name>.json`` path for a method type.

    ``method_type`` may be a canonical method type ("DFM") or a source kind
    ("dfm", "berquist_sherman_sr").
    """
    prefix = METHOD_JSON_FILENAME_PREFIX_BY_TYPE.get(
        normalize_method_type(method_type)
    ) or METHOD_JSON_FILENAME_PREFIX_BY_TYPE.get(
        normalize_method_type("", method_type)
    )
    if not prefix:
        raise ValueError(f"Unknown method type: {method_type}")
    filename = f"{prefix}{sanitize_dataset_file_name(method_name, 'Name')}.json"
    return os.path.join(
        config.get_project_method_data_dir(project_name, reserving_class),
        filename,
    )


def sidecar_write_lock(path: str) -> threading.RLock:
    key = os.path.normcase(os.path.abspath(path))
    with _SIDECAR_WRITE_LOCKS_GUARD:
        return _SIDECAR_WRITE_LOCKS.setdefault(key, threading.RLock())


def replace_staged_file(source: str, target: str) -> None:
    """``os.replace`` that rides out a concurrent reader's sharing violation.

    On Windows a reader holding the target open without FILE_SHARE_DELETE
    makes the swap fail with WinError 5 for the duration of the read — on the
    server host that reader is typically another Engine walking the class, or
    a client re-reading the sidecar over the share. Such reads last
    milliseconds, so a save must wait through a few of them rather than fail;
    one save died exactly this way in production. The last attempt re-raises,
    so a genuinely locked file still fails loudly.
    """

    for attempt in range(10):
        try:
            os.replace(source, target)
            return
        except PermissionError:
            if attempt == 9:
                raise
            time.sleep(0.15)


def reserving_class_io_lock(project_name: str, reserving_class: str) -> threading.RLock:
    key = f"{_clean_text(project_name).casefold()}\0{_clean_text(reserving_class).casefold()}"
    with _RESERVING_CLASS_LOCKS_GUARD:
        return _RESERVING_CLASS_LOCKS.setdefault(key, threading.RLock())


def read_sidecar(path: str) -> Dict[str, Any]:
    try:
        with open(path, "r", encoding="utf-8") as fh:
            payload = json.load(fh)
        return expand_sidecar_links(payload) if isinstance(payload, dict) else {}
    except Exception:
        return {}


def read_sidecar_strict(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as fh:
        payload = json.load(fh)
    if not isinstance(payload, dict):
        raise ValueError(f"Dataset sidecar must contain a JSON object: {os.path.basename(path)}")
    return expand_sidecar_links(payload)


def write_sidecar(path: str, payload: Dict[str, Any]) -> None:
    """Write one sidecar; ``audit_log`` lands last under the one audit policy."""

    with sidecar_write_lock(path):
        tmp_path = f"{path}.{uuid.uuid4()}.tmp"
        os.makedirs(os.path.dirname(path), exist_ok=True)
        try:
            with open(tmp_path, "w", encoding="utf-8", newline="\n") as fh:
                fh.write(persisted_json_text(finalize_sidecar(payload)))
            os.replace(tmp_path, path)
        finally:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except OSError:
                pass


def _timestamp_from_text(value: Any) -> float:
    text = _clean_text(value)
    if not text:
        return 0.0
    numeric = None
    try:
        numeric = float(text)
    except ValueError:
        numeric = None
    if numeric and numeric > 0:
        return numeric / 1000.0 if numeric > 1000000000000 else numeric
    try:
        parsed = text.replace("Z", "+00:00")
        return datetime.fromisoformat(parsed).timestamp()
    except ValueError:
        return 0.0


def sidecar_timestamp(path: str, payload: Dict[str, Any]) -> float:
    for key in ("updated_at", "updated", "modified_at", "modified", "last_modified"):
        ts = _timestamp_from_text(payload.get(key))
        if ts > 0:
            return ts
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0.0


def now_utc_iso() -> str:
    return utc_now_text()


def compute_status(project_name: str, reserving_class: str, dataset_name: str, payload: Dict[str, Any], path: str = "") -> int:
    method_type = normalize_method_type(payload.get("method_type"), payload.get("source_kind"))
    if method_type == METHOD_TYPE_NONE:
        return STATUS_CURRENT
    sidecar_file = path or sidecar_path(project_name, reserving_class, dataset_name)
    current_ts = sidecar_timestamp(sidecar_file, payload)
    if current_ts <= 0:
        return normalize_status(payload.get("status"))
    for precedent_name in entry_names(payload.get("precedents")):
        dep_path = sidecar_path(project_name, reserving_class, precedent_name)
        dep_payload = read_sidecar(dep_path)
        if not dep_payload:
            return STATUS_REVIEW_NEEDED
        dep_method_type = normalize_method_type(
            dep_payload.get("method_type"),
            dep_payload.get("source_kind"),
        )
        if dep_method_type != METHOD_TYPE_NONE and normalize_status(dep_payload.get("status")) == STATUS_REVIEW_NEEDED:
            return STATUS_REVIEW_NEEDED
        if sidecar_timestamp(dep_path, dep_payload) > current_ts + 0.000001:
            return STATUS_REVIEW_NEEDED
    return STATUS_CURRENT


def apply_status_fields(
    payload: Dict[str, Any],
    project_name: str,
    reserving_class: str,
    dataset_name: str,
    *,
    path: str = "",
    method_type: Any = "",
    force_status: int | None = None,
) -> Dict[str, Any]:
    payload["method_type"] = normalize_method_type(method_type or payload.get("method_type"), payload.get("source_kind"))
    if payload["method_type"] == METHOD_TYPE_NONE:
        payload["status"] = STATUS_CURRENT
    elif force_status is not None:
        payload["status"] = normalize_status(force_status)
    else:
        payload["status"] = compute_status(project_name, reserving_class, dataset_name, payload, path)
    return payload


def _remove_dependent(payload: Dict[str, Any], dependent_name: str) -> bool:
    old = entry_names(payload.get("dependents"))
    key = _canon_dataset_name(dependent_name)
    next_names = [name for name in old if _canon_dataset_name(name) != key]
    if len(next_names) == len(old):
        return False
    payload["dependents"] = name_entries(next_names)
    return True


def _add_dependent(payload: Dict[str, Any], dependent_name: str) -> bool:
    old = entry_names(payload.get("dependents"))
    key = _canon_dataset_name(dependent_name)
    if not key or key in {_canon_dataset_name(name) for name in old}:
        return False
    payload["dependents"] = name_entries([*old, dependent_name])
    return True


def update_precedent_dependents(
    project_name: str,
    reserving_class: str,
    dependent_name: str,
    old_precedents: Iterable[Any],
    new_precedents: Iterable[Any],
    *,
    require_new_precedents: bool = False,
) -> List[str]:
    old_by_key = {_canon_dataset_name(name): _clean_text(name) for name in old_precedents if _canon_dataset_name(name)}
    new_by_key = {_canon_dataset_name(name): _clean_text(name) for name in new_precedents if _canon_dataset_name(name)}
    names_by_key = {**old_by_key, **new_by_key}
    paths_by_key = {
        key: sidecar_path(project_name, reserving_class, source_name)
        for key, source_name in names_by_key.items()
    }
    ordered_keys = sorted(paths_by_key, key=lambda item: os.path.normcase(paths_by_key[item]))
    staged: Dict[str, str] = {}
    backups: Dict[str, bytes] = {}
    replaced: List[str] = []
    touched: List[str] = []
    with ExitStack() as stack:
        for key in ordered_keys:
            stack.enter_context(sidecar_write_lock(paths_by_key[key]))
        futures = {
            key: _SIDECAR_READ_EXECUTOR.submit(read_sidecar_strict, paths_by_key[key])
            for key in ordered_keys
            if os.path.isfile(paths_by_key[key])
        }
        payloads: Dict[str, Dict[str, Any]] = {}
        for key in ordered_keys:
            path = paths_by_key[key]
            if key not in futures:
                if require_new_precedents and key in new_by_key:
                    raise FileNotFoundError(f"Dataset sidecar is missing for Result Selection precedent '{new_by_key[key]}'.")
                continue
            try:
                payloads[key] = futures[key].result()
            except Exception as exc:
                if require_new_precedents and key in new_by_key:
                    raise RuntimeError(
                        f"Unable to register Result Selection precedent '{new_by_key[key]}': {exc}"
                    ) from exc
                continue

        updates: Dict[str, Dict[str, Any]] = {}
        for key in ordered_keys:
            payload = payloads.get(key)
            if not payload:
                continue
            changed = _remove_dependent(payload, dependent_name) if key not in new_by_key else _add_dependent(payload, dependent_name)
            if not changed:
                continue
            source_name = names_by_key[key]
            path = paths_by_key[key]
            apply_status_fields(payload, project_name, reserving_class, source_name, path=path)
            updates[path] = payload
            touched.append(source_name)

        try:
            for path, payload in updates.items():
                with open(path, "rb") as fh:
                    backups[path] = fh.read()
                temporary = f"{path}.{uuid.uuid4()}.tmp"
                with open(temporary, "w", encoding="utf-8", newline="\n") as fh:
                    fh.write(persisted_json_text(finalize_sidecar(payload)))
                staged[path] = temporary
            for path in sorted(updates, key=os.path.normcase):
                replace_staged_file(staged.pop(path), path)
                replaced.append(path)
        except Exception as exc:
            rollback_errors = []
            for path in reversed(replaced):
                try:
                    temporary = f"{path}.{uuid.uuid4()}.rollback"
                    with open(temporary, "wb") as fh:
                        fh.write(backups[path])
                    os.replace(temporary, path)
                except OSError as rollback_exc:
                    rollback_errors.append(f"{os.path.basename(path)}: {rollback_exc}")
            if rollback_errors:
                raise RuntimeError(f"{exc}; dependency graph rollback failed: {'; '.join(rollback_errors)}") from exc
            raise
        finally:
            for temporary in staged.values():
                try:
                    os.remove(temporary)
                except OSError:
                    pass
    return touched


def read_sidecars(project_name: str, reserving_class: str, dataset_names: Iterable[Any]) -> Dict[str, Dict[str, Any]]:
    """Read many sidecars at once, keyed by canonical name.

    Project data can live on a mapped network drive where every open is a
    round trip, so the reads go through the bounded pool instead of a
    per-file awaited loop. Missing or unreadable sidecars come back absent.
    """

    names: List[str] = []
    seen: Set[str] = set()
    for raw in dataset_names or []:
        name = _clean_text(raw)
        key = _canon_dataset_name(name)
        if not key or key in seen:
            continue
        seen.add(key)
        names.append(name)
    futures = {
        _canon_dataset_name(name): _SIDECAR_READ_EXECUTOR.submit(
            read_sidecar, sidecar_path(project_name, reserving_class, name)
        )
        for name in names
    }
    out: Dict[str, Dict[str, Any]] = {}
    for key, future in futures.items():
        payload = future.result()
        if payload:
            out[key] = payload
    return out


def dependent_closure(
    project_name: str,
    reserving_class: str,
    changed_dataset_names: Iterable[Any],
) -> List[Dict[str, str]]:
    """Name every object reachable from the changed roots, nearest tier first.

    Walks the same sidecar ``dependents`` edges as
    ``_refresh_method_statuses_for_dependents_unlocked`` below, but reads only
    — it takes no write locks and marks nothing, so the two-step save can show
    the user what a save would reach before anything is written. Each tier is
    read through the bounded pool rather than one file at a time, because the
    same walk on a Client PC would otherwise pay one network round trip per
    node.
    """

    roots: List[str] = []
    queued: Set[str] = set()
    for raw in changed_dataset_names or []:
        name = _clean_text(raw)
        key = _canon_dataset_name(name)
        if not key or key in queued:
            continue
        queued.add(key)
        roots.append(name)

    out: List[Dict[str, str]] = []
    reported: Set[str] = set(queued)
    frontier = roots
    while frontier:
        payloads = read_sidecars(project_name, reserving_class, frontier)
        next_frontier: List[str] = []
        for source_name in frontier:
            payload = payloads.get(_canon_dataset_name(source_name))
            if not payload:
                continue
            for dependent_name in entry_names(payload.get("dependents")):
                key = _canon_dataset_name(dependent_name)
                if not key or key in reported:
                    continue
                reported.add(key)
                next_frontier.append(dependent_name)
        if not next_frontier:
            break
        tier_payloads = read_sidecars(project_name, reserving_class, next_frontier)
        for dependent_name in next_frontier:
            dependent_payload = tier_payloads.get(_canon_dataset_name(dependent_name), {})
            out.append({
                "dataset_name": dependent_name,
                "method_type": normalize_method_type(
                    dependent_payload.get("method_type"),
                    dependent_payload.get("source_kind"),
                ),
                "source_kind": _clean_text(dependent_payload.get("source_kind")),
            })
        frontier = next_frontier
    return out


def graph_signature(
    project_name: str,
    reserving_class: str,
    dataset_names: Iterable[Any],
) -> List[List[Any]]:
    """Project the sidecar state that decides a dependent walk's shape.

    The two-step save fingerprints this over the roots plus their closure, so
    a precedent rewired, an object refreshed, or a status flipped between the
    plan the user reviewed and the commit is caught under the lease.
    """

    payloads = read_sidecars(project_name, reserving_class, dataset_names)
    return sorted(
        (
            [
                key,
                _clean_text(payload.get("updated_at")),
                normalize_status(payload.get("status")),
                normalize_method_type(payload.get("method_type"), payload.get("source_kind")),
                sorted(
                    _canon_dataset_name(name)
                    for name in entry_names(payload.get("precedents"))
                ),
                sorted(
                    _canon_dataset_name(name)
                    for name in entry_names(payload.get("dependents"))
                ),
            ]
            for key, payload in payloads.items()
        ),
        key=lambda item: item[0],
    )


def _refresh_method_statuses_for_dependents_unlocked(
    project_name: str,
    reserving_class: str,
    changed_dataset_names: Iterable[Any],
    *,
    direct_only: bool = False,
) -> List[Dict[str, Any]]:
    touched: List[Dict[str, Any]] = []
    processed_sources: Set[str] = set()
    queued_sources: Set[str] = set()
    marked_dependents: Set[str] = set()
    method_dependent_keys: Set[str] = set()
    queue: List[str] = []
    for source_name in changed_dataset_names or []:
        clean_name = _clean_text(source_name)
        source_key = _canon_dataset_name(clean_name)
        if not source_key or source_key in queued_sources:
            continue
        queued_sources.add(source_key)
        queue.append(clean_name)

    while queue:
        source_name = queue.pop(0)
        source_key = _canon_dataset_name(source_name)
        if not source_key or source_key in processed_sources:
            continue
        processed_sources.add(source_key)
        source_path = sidecar_path(project_name, reserving_class, source_name)
        with sidecar_write_lock(source_path):
            source_payload = read_sidecar(source_path)
        if not source_payload:
            continue
        for dependent_name in entry_names(source_payload.get("dependents")):
            dep_key = _canon_dataset_name(dependent_name)
            if not dep_key:
                continue
            if dep_key not in marked_dependents:
                marked_dependents.add(dep_key)
                dep_path = sidecar_path(project_name, reserving_class, dependent_name)
                with sidecar_write_lock(dep_path):
                    dep_payload = read_sidecar(dep_path)
                    if dep_payload:
                        method_type = normalize_method_type(dep_payload.get("method_type"), dep_payload.get("source_kind"))
                        if method_type != METHOD_TYPE_NONE:
                            method_dependent_keys.add(dep_key)
                            before = normalize_status(dep_payload.get("status"))
                            dep_payload["method_type"] = method_type
                            dep_payload["status"] = STATUS_REVIEW_NEEDED
                            after = normalize_status(dep_payload.get("status"))
                            if after != before:
                                write_sidecar(dep_path, dep_payload)
                                touched.append({"dataset_name": dependent_name, "status": after})
            # direct_only stops expanding at the first method tier: plain
            # vectors between a root and its nearest methods are walked
            # through, but a marked method's own downstream stays for the
            # Engine job's closure marking.
            if direct_only and dep_key in method_dependent_keys:
                continue
            if dep_key not in queued_sources:
                queued_sources.add(dep_key)
                queue.append(dependent_name)
    return touched


def refresh_method_statuses_for_dependents(
    project_name: str,
    reserving_class: str,
    changed_dataset_names: Iterable[Any],
    *,
    direct_only: bool = False,
) -> List[Dict[str, Any]]:
    # direct_only marks only the first method tier reachable from the changed
    # roots (walking through plain vectors but never past a marked method)
    # instead of the whole reachable closure. Save paths use it so a Client
    # PC save pays a handful of SMB round trips; the Engine job re-marks the
    # full closure on claim, where the same walk runs against local disk.
    with reserving_class_io_lock(project_name, reserving_class):
        return _refresh_method_statuses_for_dependents_unlocked(
            project_name,
            reserving_class,
            changed_dataset_names,
            direct_only=direct_only,
        )
