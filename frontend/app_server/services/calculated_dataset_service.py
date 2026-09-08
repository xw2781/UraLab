from __future__ import annotations

import ast
import getpass
import hashlib
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from typing import TYPE_CHECKING, Any, Callable, Dict, List, Mapping, NamedTuple, Sequence, Set, Tuple

import numpy as np
import pandas as pd

from arcrho_api.dataset_display_contract import normalize_show_subtotal
from arcrho_api.dataset_link_contract import link_precedent_names
from arcrho_api.sidecar_core_contract import stored_length_fields, stored_lengths
from arcrho_api.timestamps import utc_now_text
from app_server import config
from app_server.helpers import (
    _canon_dataset_name,
    build_dataset_cache_file_name,
    sanitize_dataset_file_name,
)
from app_server.services import (
    class_folder_scan_cache,
    dataset_instance_index_service,
    dataset_number_format_service,
    dataset_sidecar_status_service,
    dataset_types_service,
    file_read_cache,
    precedent_cache_service,
    runtime_cache_provenance_service,
    user_identity_service,
)

if TYPE_CHECKING:
    from app_server.services import dataset_types_plan_service

_METHOD_DEPENDENT_READ_EXECUTOR = ThreadPoolExecutor(
    max_workers=6,
    thread_name_prefix="arcrho-method-dependent-read",
)

def _clean_text(value: Any) -> str:
    return str(value if value is not None else "").strip()


def _bool_value(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    return str(value or "").strip().lower() in {"true", "1", "yes", "y"}


def _now_utc_iso() -> str:
    return utc_now_text()


def _current_user_name() -> str:
    display_name = user_identity_service.get_current_display_name()
    if display_name:
        return display_name
    for value in (os.environ.get("USERNAME"), os.environ.get("USER")):
        text = _clean_text(value)
        if text:
            return text
    try:
        return _clean_text(getpass.getuser()) or "calculated"
    except Exception:
        return "calculated"


def _dataset_type_rows(project_name: str) -> List[Dict[str, Any]]:
    data = dataset_types_service.load_dataset_types_data(
        project_name,
        strict=True,
    )
    rows = []
    for row in data.get("rows") or []:
        if not isinstance(row, list):
            continue
        name = _clean_text(row[0] if len(row) > 0 else "")
        if not name:
            continue
        rows.append({
            "name": name,
            "data_format": _clean_text(row[1] if len(row) > 1 else "Triangle") or "Triangle",
            "category": _clean_text(row[2] if len(row) > 2 else ""),
            "calculated": _bool_value(row[3] if len(row) > 3 else False),
            "formula": _clean_text(row[4] if len(row) > 4 else ""),
            "source": _clean_text(row[5] if len(row) > 5 else ""),
            "generated": _bool_value(row[6] if len(row) > 6 else False),
        })
    return rows


def _dataset_type_name_by_key(project_name: str) -> Dict[str, str]:
    return {
        _canon_dataset_name(row["name"]): row["name"]
        for row in _dataset_type_rows(project_name)
        if _canon_dataset_name(row.get("name"))
    }


def _formula_components(formula: str, known_names: List[str]) -> List[str]:
    text = _clean_text(formula)
    if not text:
        return []

    out: List[str] = []
    seen: Set[str] = set()
    masked_parts = []
    last = 0
    for match in re.finditer(r'"([^"]+)"', text):
        token = _clean_text(match.group(1))
        key = _canon_dataset_name(token)
        if token and key and key not in seen:
            seen.add(key)
            out.append(token)
        masked_parts.append(text[last:match.start()])
        masked_parts.append(" ")
        last = match.end()
    masked_parts.append(text[last:])
    unquoted_text = "".join(masked_parts)

    for name in sorted(
        {str(item or "").strip() for item in known_names if str(item or "").strip()},
        key=len,
        reverse=True,
    ):
        key = _canon_dataset_name(name)
        if not key or key in seen:
            continue
        pattern = re.compile(rf"(?<![A-Za-z0-9_]){re.escape(name)}(?![A-Za-z0-9_])", flags=re.IGNORECASE)
        if pattern.search(unquoted_text):
            seen.add(key)
            out.append(name)
    return out


def _calculated_dataset_contract_from_rows(
    rows: List[Dict[str, Any]],
    dataset_type_name: str,
) -> Dict[str, Any] | None:
    known_names = [row["name"] for row in rows]
    target_key = _canon_dataset_name(dataset_type_name)
    for row in rows:
        if _canon_dataset_name(row["name"]) != target_key:
            continue
        if not row.get("calculated") or row.get("generated") or not _clean_text(row.get("formula")):
            return None
        precedents = _formula_components(row["formula"], known_names)
        rows_by_key = {
            _canon_dataset_name(item.get("name")): dict(item)
            for item in rows
            if _canon_dataset_name(item.get("name"))
        }
        return {
            **row,
            "precedents": precedents,
            "precedent_contracts": {
                _canon_dataset_name(name): rows_by_key.get(
                    _canon_dataset_name(name),
                    {},
                )
                for name in precedents
            },
        }
    return None


def calculated_dataset_contract(
    project_name: str,
    dataset_type_name: str,
) -> Dict[str, Any] | None:
    contract = _calculated_dataset_contract_from_rows(
        _dataset_type_rows(project_name),
        dataset_type_name,
    )
    return dict(contract) if contract else None


def _direct_precedent_names(project_name: str, dataset_type_name: str) -> List[str]:
    contract = calculated_dataset_contract(project_name, dataset_type_name)
    return list(contract.get("precedents") or []) if contract else []


def calculated_dataset_dependency_names(
    project_name: str,
    dataset_type_name: str,
) -> List[str] | None:
    """Return direct inputs for an app-calculated type, or ``None`` otherwise."""
    contract = calculated_dataset_contract(project_name, dataset_type_name)
    return list(contract.get("precedents") or []) if contract else None


def _direct_dependent_names(project_name: str, dataset_type_name: str) -> List[str]:
    rows = _dataset_type_rows(project_name)
    known_names = [row["name"] for row in rows]
    target_key = _canon_dataset_name(dataset_type_name)
    out: List[str] = []
    seen: Set[str] = set()
    if not target_key:
        return out
    for row in rows:
        if not row.get("calculated") or row.get("generated") or not _clean_text(row.get("formula")):
            continue
        components = _formula_components(row["formula"], known_names)
        component_keys = {_canon_dataset_name(component) for component in components}
        if target_key not in component_keys:
            continue
        dep_key = _canon_dataset_name(row["name"])
        if dep_key and dep_key not in seen:
            seen.add(dep_key)
            out.append(row["name"])
    return out


def _name_entries(names: List[str]) -> List[Dict[str, str]]:
    return dataset_sidecar_status_service.name_entries(names)


def _precedent_entries(
    project_name: str,
    dataset_type_name: str,
    dependency_info: List[Dict[str, Any]] | None = None,
) -> List[Dict[str, Any]]:
    # The persisted graph is location-independent: a precedent is named, never
    # pathed. ``dependency_info`` (paths, stats) is runtime state for the
    # calculation that produced the output and is not written to the sidecar.
    return _name_entries(_direct_precedent_names(project_name, dataset_type_name))


_DEPENDENCY_RECORD_FIELDS = (
    "source_kind",
    "data_format",
    "origin_length",
    "development_length",
    "period_length",
    "cumulative",
    "calendar",
    "size",
    "mtime_ns",
    "sha256",
    "input_path",
    "input_size",
    "input_mtime_ns",
    "input_sha256",
)


def _dependency_fingerprints(dependency_info: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Project the components a calculation read into its technical record.

    Each entry names the dependency and the exact file it was read from with
    that file's fingerprint, plus the input a DFM method output was rebuilt
    from. A live-preview component has no file and is left out.
    """
    entries: List[Dict[str, Any]] = []
    for item in dependency_info:
        if not isinstance(item, dict):
            continue
        path = _clean_text(item.get("path"))
        if not path:
            continue
        entry: Dict[str, Any] = {
            "dataset_type": _clean_text(item.get("dataset_type_name")),
            "dataset_name": _clean_text(item.get("dataset_name") or item.get("dataset_type_name")),
            "path": path,
        }
        for key in _DEPENDENCY_RECORD_FIELDS:
            value = item.get(key)
            if value not in (None, ""):
                entry[key] = value
        entries.append(entry)
    return entries


def sidecar_graph_fields(
    project_name: str,
    dataset_type_name: str,
    dependency_info: List[Dict[str, Any]] | None = None,
    reserving_class: str = "",
) -> Dict[str, Any]:
    dependent_names = _direct_dependent_names(project_name, dataset_type_name)
    rc = _clean_text(reserving_class)
    if rc:
        existing_keys = _existing_dataset_keys(project_name, rc)
        dependent_names = [
            name
            for name in dependent_names
            if _canon_dataset_name(name) in existing_keys
        ]
    return {
        "precedents": _precedent_entries(project_name, dataset_type_name, dependency_info),
        "dependents": _name_entries(dependent_names),
    }


def apply_sidecar_graph_fields(
    payload: Dict[str, Any],
    project_name: str = "",
    dataset_type_name: str = "",
    dependency_info: List[Dict[str, Any]] | None = None,
) -> Dict[str, Any]:
    project = _clean_text(project_name or payload.get("project_name"))
    dataset_type = _clean_text(
        dataset_type_name
        or payload.get("dataset_type")
        or payload.get("dataset_name")
    )
    if not project or not dataset_type:
        payload["precedents"] = []
        payload["dependents"] = []
        payload.pop("dependencies", None)
        return payload

    rows_by_key = {
        _canon_dataset_name(row["name"]): row
        for row in _dataset_type_rows(project)
        if _canon_dataset_name(row.get("name"))
    }
    row = rows_by_key.get(_canon_dataset_name(dataset_type))
    if row:
        formula = _clean_text(row.get("formula"))
        payload["calculated"] = bool(row.get("calculated") and not row.get("generated") and formula)

    existing_dependent_names = dataset_sidecar_status_service.entry_names(payload.get("dependents"))
    existing_precedents = payload.get("precedents")
    owning_method_type = dataset_sidecar_status_service.normalize_method_type(
        payload.get("method_type"),
        payload.get("source_kind"),
    )
    graph_fields = sidecar_graph_fields(
        project,
        dataset_type,
        dependency_info,
        _clean_text(payload.get("reserving_class")),
    )
    reserving_class = _clean_text(payload.get("reserving_class"))
    own_dataset_name = _clean_text(payload.get("dataset_name")) or dataset_type
    own_key = _canon_dataset_name(own_dataset_name)
    preserved_method_dependents: List[str] = []
    if reserving_class and existing_dependent_names:
        futures = {
            name: _METHOD_DEPENDENT_READ_EXECUTOR.submit(
                dataset_sidecar_status_service.read_sidecar,
                dataset_sidecar_status_service.sidecar_path(project, reserving_class, name),
            )
            for name in existing_dependent_names
        }
        for name in existing_dependent_names:
            dependent = futures[name].result()
            # A dependent survives the type-graph rebuild when it is a method
            # (the graph cannot re-derive a method edge) or when its own
            # sidecar still names this dataset as a precedent -- the
            # instance-level edge an ArcRho cell link put there.
            if not dependent or dataset_sidecar_status_service.normalize_method_type(
                dependent.get("method_type"),
                dependent.get("source_kind"),
            ) != dataset_sidecar_status_service.METHOD_TYPE_NONE or own_key in {
                _canon_dataset_name(entry_name)
                for entry_name in dataset_sidecar_status_service.entry_names(dependent.get("precedents"))
            }:
                preserved_method_dependents.append(name)
    graph_fields["dependents"] = dataset_sidecar_status_service.merge_name_entries(
        graph_fields.get("dependents"),
        dataset_sidecar_status_service.name_entries(preserved_method_dependents),
    )
    if owning_method_type != dataset_sidecar_status_service.METHOD_TYPE_NONE:
        graph_fields["precedents"] = existing_precedents if isinstance(existing_precedents, list) else []
    else:
        # ArcRho cell links are instance-level precedent edges on top of the
        # dataset-type formula graph: the dependent-propagation walk follows
        # them to re-evaluate the linked cells when a source dataset changes.
        # Excel references contribute no edge, and a self-reference never
        # records one.
        linked_names = [
            name
            for name in link_precedent_names(
                payload.get("internal_links"),
                payload.get("formula_links"),
            )
            if _canon_dataset_name(name) != own_key
        ]
        if linked_names:
            graph_fields["precedents"] = dataset_sidecar_status_service.merge_name_entries(
                graph_fields.get("precedents"),
                dataset_sidecar_status_service.name_entries(linked_names),
            )
    payload.update(graph_fields)
    payload.pop("dependencies", None)
    return payload


def _replace_formula_refs(formula: str, known_names: List[str]) -> Tuple[str, Dict[str, str]]:
    text = _clean_text(formula)
    refs: Dict[str, str] = {}
    by_key = {
        _canon_dataset_name(name): str(name or "").strip()
        for name in known_names
        if _canon_dataset_name(name)
    }

    def new_var(name: str) -> str:
        key = _canon_dataset_name(name)
        if not key:
            return ""
        for var, ref_name in refs.items():
            if _canon_dataset_name(ref_name) == key:
                return var
        var = f"_d{len(refs)}"
        refs[var] = by_key.get(key) or name
        return var

    def quoted_repl(match: re.Match[str]) -> str:
        var = new_var(_clean_text(match.group(1)))
        return var or "0"

    expr = re.sub(r'"([^"]+)"', quoted_repl, text)
    for name in sorted(
        {str(item or "").strip() for item in known_names if str(item or "").strip()},
        key=len,
        reverse=True,
    ):
        key = _canon_dataset_name(name)
        if not key:
            continue
        pattern = re.compile(rf"(?<![A-Za-z0-9_]){re.escape(name)}(?![A-Za-z0-9_])", flags=re.IGNORECASE)
        if pattern.search(expr):
            var = new_var(name)
            expr = pattern.sub(var, expr)
    return expr, refs


def _eval_ast(node: ast.AST, values: Dict[str, np.ndarray]) -> np.ndarray | float:
    if isinstance(node, ast.Expression):
        return _eval_ast(node.body, values)
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return float(node.value)
    if isinstance(node, ast.Name):
        if node.id not in values:
            raise ValueError(f"Unknown formula variable: {node.id}")
        return values[node.id]
    if isinstance(node, ast.UnaryOp):
        operand = _eval_ast(node.operand, values)
        if isinstance(node.op, ast.USub):
            return -operand
        if isinstance(node.op, ast.UAdd):
            return operand
    if isinstance(node, ast.BinOp):
        left = _eval_ast(node.left, values)
        right = _eval_ast(node.right, values)
        if isinstance(node.op, ast.Add):
            return left + right
        if isinstance(node.op, ast.Sub):
            return left - right
        if isinstance(node.op, ast.Mult):
            return left * right
        if isinstance(node.op, ast.Div):
            return left / right
    raise ValueError("Formula contains unsupported syntax.")


def _read_sidecar(path: str) -> Dict[str, Any]:
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def _csv_base_name(path: str) -> str:
    stem = os.path.splitext(os.path.basename(path))[0]
    return dataset_instance_index_service._normalize_cached_dataset_name(stem)


def _cached_csv_data_format(path: str, sidecar: Dict[str, Any]) -> str:
    stem = os.path.splitext(os.path.basename(path))[0]
    parts = stem.split("@")
    if (
        len(parts) >= 5
        and parts[-4].strip().isdigit()
        and parts[-3].strip().isdigit()
        and parts[-2].strip().lower() in {"cum", "inc"}
        and parts[-1].strip().lower() in {"dev", "cal"}
    ):
        return "Triangle"
    if len(parts) >= 2 and parts[-1].strip().isdigit():
        return "Vector"
    return _clean_text(sidecar.get("data_format"))


class _DatasetCacheScan(NamedTuple):
    """One observation of a reserving class's cached CSV folder and sidecars.

    Dependency resolution asks the same folder about several dependencies in a
    row, so callers enumerate once and hand this snapshot down instead of
    re-reading every sidecar per dependency. ``mtime`` comes from the directory
    listing that found the file, so no path is stat-ed twice. ``csv_stats``
    keeps the listing's ``(mtime_ns, size)`` identity per normalized CSV path
    so component value reads can validate the in-memory matrix cache without
    another stat.
    """

    exists: bool
    csv_files: Tuple[Tuple[str, float], ...]
    sidecars: Dict[str, Dict[str, Any]]
    csv_stats: Dict[str, Tuple[int, int]] = {}


class _MethodFolderScan(NamedTuple):
    """One observation of a reserving class's method JSON folder."""

    exists: bool
    method_files: Tuple[Tuple[str, float], ...]
    payloads: Dict[str, Dict[str, Any]]


def _scan_dataset_cache_folder(project_name: str, reserving_class: str) -> _DatasetCacheScan:
    return _scan_dataset_cache_folder_at(
        config.get_project_dataset_cache_dir(project_name, reserving_class)
    )


def _scan_dataset_cache_folder_at(folder: str) -> _DatasetCacheScan:
    """Observe one cached CSV folder given its path (the runtime derives it from a CSV)."""
    exists, csv_entries = class_folder_scan_cache.scan_files_with_stats(folder, ".csv")
    if not exists:
        return _DatasetCacheScan(exists=False, csv_files=(), sidecars={}, csv_stats={})
    # Sidecars are validated against their own folder's listing, so repeat
    # scans cost two directory enumerations instead of one read per file. The
    # folder is derived from the cache folder exactly like the per-CSV mapping
    # in _dataset_sidecar_path_for_cached_csv resolves it.
    if os.path.basename(folder).lower() == config.DATASET_CACHE_DIR.lower():
        sidecar_folder = os.path.join(os.path.dirname(folder), config.DATASET_SIDECAR_DIR)
    else:
        sidecar_folder = os.path.join(folder, config.DATASET_SIDECAR_DIR)
    _sidecar_folder_exists, sidecar_entries = class_folder_scan_cache.scan_files_with_stats(
        sidecar_folder,
        ".json",
    )
    sidecars = class_folder_scan_cache.read_json_files_cached(
        (
            dataset_instance_index_service._dataset_sidecar_path_for_cached_csv(entry.path)
            for entry in csv_entries
        ),
        class_folder_scan_cache.stats_by_normcase_path(sidecar_entries),
    )
    return _DatasetCacheScan(
        exists=True,
        csv_files=tuple((entry.path, entry.mtime) for entry in csv_entries),
        sidecars=sidecars,
        csv_stats=class_folder_scan_cache.stats_by_normcase_path(csv_entries),
    )


def _scan_dfm_method_folder(project_name: str, reserving_class: str) -> _MethodFolderScan:
    return _scan_dfm_method_folder_at(
        config.get_project_method_data_dir(project_name, reserving_class)
    )


def _scan_dfm_method_folder_at(folder: str) -> _MethodFolderScan:
    """Observe one DFM method folder given its path."""
    exists, method_entries = class_folder_scan_cache.scan_files_with_stats(
        folder, ".json", name_prefix="DFM@"
    )
    if not exists:
        return _MethodFolderScan(exists=False, method_files=(), payloads={})
    payloads = class_folder_scan_cache.read_json_files_cached(
        (entry.path for entry in method_entries),
        class_folder_scan_cache.stats_by_normcase_path(method_entries),
    )
    return _MethodFolderScan(
        exists=True,
        method_files=tuple((entry.path, entry.mtime) for entry in method_entries),
        payloads=payloads,
    )


def _sidecar_for_csv(path: str, scan: _DatasetCacheScan | None = None) -> Dict[str, Any]:
    sidecar_path = dataset_instance_index_service._dataset_sidecar_path_for_cached_csv(path)
    payload = scan.sidecars.get(sidecar_path) if scan is not None else None
    if payload is None:
        payload = _read_sidecar(sidecar_path)
    return {**payload, "_sidecar_path": sidecar_path}


def _json_tab(source: Dict[str, Any], key: str) -> Dict[str, Any]:
    value = source.get(key) if isinstance(source, dict) else None
    return value if isinstance(value, dict) else {}


def _method_output_names(payload: Dict[str, Any], path: str = "") -> Set[str]:
    names: Set[str] = set()
    details = _json_tab(payload, "details_tab")
    for key in ("output_type", "name"):
        text = _clean_text(details.get(key))
        if text:
            names.add(text)
    stem = os.path.splitext(os.path.basename(path))[0]
    if stem.startswith("DFM@"):
        text = _clean_text(config.decode_filename_segment(stem[len("DFM@"):]))
        if text:
            names.add(text)
    return names


def _candidate_dfm_methods(
    project_name: str,
    reserving_class: str,
    dataset_type_name: str,
    scan: _MethodFolderScan | None = None,
) -> List[Dict[str, Any]]:
    folder = config.get_project_method_data_dir(project_name, reserving_class)
    dep_key = _canon_dataset_name(dataset_type_name)
    if not dep_key:
        return []
    scan = scan if scan is not None else _scan_dfm_method_folder(project_name, reserving_class)
    if not scan.exists:
        return []

    out: List[Dict[str, Any]] = []
    seen: Set[str] = set()

    def add_candidate(path: str, mtime: float) -> None:
        norm = os.path.abspath(path)
        if norm in seen:
            return
        payload = scan.payloads.get(path)
        if payload is None:
            return
        seen.add(norm)
        names = _method_output_names(payload, path)
        if dep_key not in {_canon_dataset_name(name) for name in names}:
            return
        details = _json_tab(payload, "details_tab")
        output_type = _clean_text(details.get("output_type"))
        method_name = _clean_text(details.get("name"))
        score = 0
        if _canon_dataset_name(output_type) == dep_key:
            score += 8
        if _canon_dataset_name(method_name) == dep_key:
            score += 4
        out.append({
            "path": path,
            "payload": payload,
            "score": score,
            "mtime": mtime,
        })

    # The exact-name file keeps its historical priority among equal scores.
    direct_path = os.path.join(folder, f"DFM@{sanitize_dataset_file_name(dataset_type_name)}.json")
    mtime_by_path = dict(scan.method_files)
    if direct_path in mtime_by_path:
        add_candidate(direct_path, mtime_by_path[direct_path])
    for path, mtime in scan.method_files:
        add_candidate(path, mtime)

    out.sort(key=lambda item: (int(item.get("score") or 0), float(item.get("mtime") or 0)), reverse=True)
    best_score = int(out[0].get("score") or 0) if out else 0
    return [item for item in out if int(item.get("score") or 0) == best_score]


def _path_in_dir(path: str, folder: str) -> bool:
    try:
        child = os.path.normcase(os.path.realpath(os.path.abspath(path)))
        parent = os.path.normcase(os.path.realpath(os.path.abspath(folder)))
        return os.path.commonpath([child, parent]) == parent
    except Exception:
        return False


def _existing_path_in_dir(path: str, folder: str) -> str:
    recorded_path = _clean_text(path)
    if not recorded_path:
        return ""
    candidates = [recorded_path]
    relocated_path = os.path.join(
        folder,
        os.path.basename(recorded_path),
    )
    if os.path.normcase(relocated_path) != os.path.normcase(recorded_path):
        candidates.append(relocated_path)
    for candidate in candidates:
        if _path_in_dir(candidate, folder) and os.path.isfile(candidate):
            return candidate
    return ""


def _finite_float(value: Any) -> float | None:
    try:
        out = float(value)
    except (TypeError, ValueError):
        return None
    return out if np.isfinite(out) else None


def _read_numeric_csv(path: str) -> np.ndarray:
    df = pd.read_csv(path, header=None, dtype="float64", keep_default_na=True, float_precision="round_trip")
    return df.to_numpy(dtype="float64")


def _read_dfm_input_triangle(
    project_name: str,
    reserving_class: str,
    payload: Dict[str, Any],
    target_settings: Dict[str, Any],
    exact_input_path: str = "",
    scan: _DatasetCacheScan | None = None,
) -> Tuple[np.ndarray | None, str, str]:
    data_tab = _json_tab(payload, "data_tab")
    details = _json_tab(payload, "details_tab")
    dataset_folder = config.get_project_dataset_cache_dir(project_name, reserving_class)
    if exact_input_path:
        if (
            not _path_in_dir(exact_input_path, dataset_folder)
            or not os.path.isfile(exact_input_path)
        ):
            return None, exact_input_path, "Recorded DFM input triangle path is invalid."
        try:
            return _read_numeric_csv(exact_input_path), exact_input_path, ""
        except Exception as exc:
            return None, exact_input_path, str(exc)

    path = _clean_text(data_tab.get("input data triangle csv path"))
    if path and os.path.isfile(path) and _path_in_dir(path, dataset_folder):
        try:
            return _read_numeric_csv(path), path, ""
        except Exception as exc:
            return None, path, str(exc)

    input_name = _clean_text(details.get("input_triangle"))
    if not input_name:
        return None, "", "DFM method is missing an input triangle name."
    candidates = _candidate_csvs(project_name, reserving_class, input_name, target_settings, scan=scan)
    if not candidates:
        return None, "", f"Missing DFM input triangle: {input_name}"
    if len(candidates) > 1:
        return None, "", f"Ambiguous DFM input triangle: {input_name}"
    path = str(candidates[0]["path"])
    try:
        return _read_numeric_csv(path), path, ""
    except Exception as exc:
        return None, path, str(exc)


def _selected_dfm_ratio_values(payload: Dict[str, Any], dev_count: int) -> List[float]:
    ratios_tab = _json_tab(payload, "ratios_tab")
    formulas = _json_tab(ratios_tab, "average_formulas")
    selected = formulas.get("selected") if isinstance(formulas.get("selected"), list) else []
    values = formulas.get("values") if isinstance(formulas.get("values"), list) else []
    ratio_count = max(0, int(dev_count or 0))
    out: List[float] = []
    for col in range(ratio_count):
        if dev_count and col >= dev_count - 1:
            out.append(1.0)
            continue
        selected_row = None
        for row_index, row in enumerate(selected):
            row_values = row if isinstance(row, list) else []
            if col < len(row_values) and _finite_float(row_values[col]) == 1.0:
                selected_row = row_index
                break
        if selected_row is None:
            selected_row = 0
        source_row = values[selected_row] if selected_row < len(values) and isinstance(values[selected_row], list) else []
        ratio = _finite_float(source_row[col] if col < len(source_row) else None)
        out.append(ratio if ratio is not None else 1.0)
    return out


def _cumulative_factors(ratio_values: List[float]) -> List[float | None]:
    cumulative: List[float | None] = [None] * len(ratio_values)
    running: float | None = None
    for index in range(len(ratio_values) - 1, -1, -1):
        value = ratio_values[index]
        if not np.isfinite(value):
            cumulative[index] = None
            running = None
            continue
        if index == len(ratio_values) - 1:
            running = value
        elif running is not None and np.isfinite(running):
            running = value * running
        else:
            cumulative[index] = None
            running = None
            continue
        cumulative[index] = running
    return cumulative


def _build_dfm_method_vector(
    project_name: str,
    reserving_class: str,
    payload: Dict[str, Any],
    target_settings: Dict[str, Any],
    exact_input_path: str = "",
    scan: _DatasetCacheScan | None = None,
) -> Tuple[np.ndarray | None, str, str]:
    data_tab = _json_tab(payload, "data_tab")
    input_values, input_path, error = _read_dfm_input_triangle(
        project_name,
        reserving_class,
        payload,
        target_settings,
        exact_input_path=exact_input_path,
        scan=scan,
    )
    if error:
        return None, input_path, error
    if input_values is None or input_values.ndim != 2:
        return None, input_path, "DFM input triangle could not be loaded."

    dev_labels = data_tab.get("development_labels") if isinstance(data_tab.get("development_labels"), list) else []
    origin_labels = data_tab.get("origin_labels") if isinstance(data_tab.get("origin_labels"), list) else []
    dev_count = len(dev_labels) or input_values.shape[1]
    if dev_count <= 0:
        return None, input_path, "DFM method is missing development periods."

    cumulative = _cumulative_factors(_selected_dfm_ratio_values(payload, dev_count))
    row_count = len(origin_labels) or input_values.shape[0]
    out: List[float] = []
    for row_index in range(row_count):
        latest_value = None
        latest_col = None
        if row_index < input_values.shape[0]:
            max_col = min(dev_count - 1, input_values.shape[1] - 1)
            for col in range(max_col, -1, -1):
                value = _finite_float(input_values[row_index, col])
                if value is not None:
                    latest_value = value
                    latest_col = col
                    break
        factor = cumulative[latest_col] if latest_col is not None and latest_col < len(cumulative) else None
        out.append(latest_value * factor if latest_value is not None and factor is not None else np.nan)
    return np.asarray(out, dtype="float64").reshape((-1, 1)), input_path, ""


def _candidate_csvs(
    project_name: str,
    reserving_class: str,
    dataset_type_name: str,
    target_settings: Dict[str, Any],
    expected_data_format: str = "",
    scan: _DatasetCacheScan | None = None,
) -> List[Dict[str, Any]]:
    dep_key = _canon_dataset_name(dataset_type_name)
    out: List[Dict[str, Any]] = []
    scan = scan if scan is not None else _scan_dataset_cache_folder(project_name, reserving_class)
    if not scan.exists:
        return []
    for path, mtime in scan.csv_files:
        sidecar = _sidecar_for_csv(path, scan)
        candidate_data_format = _cached_csv_data_format(path, sidecar)
        if (
            _clean_text(expected_data_format)
            and candidate_data_format
            and candidate_data_format.lower() != _clean_text(expected_data_format).lower()
        ):
            continue
        dataset_name = _clean_text(sidecar.get("dataset_name") or _csv_base_name(path))
        type_name = _clean_text(sidecar.get("dataset_type") or _csv_base_name(path))
        if dep_key not in {_canon_dataset_name(dataset_name), _canon_dataset_name(type_name), _canon_dataset_name(_csv_base_name(path))}:
            continue
        score = 0
        if _canon_dataset_name(type_name) == dep_key:
            score += 8
        if _canon_dataset_name(dataset_name) == dep_key:
            score += 4
        for key in ("origin_length", "development_length"):
            if str(sidecar.get(key) or "").strip() and str(sidecar.get(key)) == str(target_settings.get(key)):
                score += 1
        for key in ("cumulative", "calendar"):
            if key in sidecar and bool(sidecar.get(key)) == bool(target_settings.get(key)):
                score += 1
        out.append({
            "path": path,
            "sidecar": sidecar,
            "data_format": candidate_data_format,
            "score": score,
            "mtime": mtime,
        })
    out.sort(key=lambda item: (int(item.get("score") or 0), float(item.get("mtime") or 0)), reverse=True)
    best_score = int(out[0].get("score") or 0) if out else 0
    return [item for item in out if int(item.get("score") or 0) == best_score]


def _target_paths(
    project_name: str,
    reserving_class: str,
    dataset_name: str,
    settings: Dict[str, Any],
    data_format: Any = "",
) -> Tuple[str, str]:
    folder = config.get_project_dataset_cache_dir(project_name, reserving_class)
    csv_name = build_dataset_cache_file_name(
        dataset_name,
        data_format or "Triangle",
        settings.get("origin_length") or 12,
        settings.get("development_length") or 12,
        settings.get("cumulative", True),
        settings.get("calendar", False),
    )
    csv_path = os.path.join(folder, f"{csv_name}.csv")
    sidecar_path = os.path.join(
        config.get_project_dataset_sidecar_dir(project_name, reserving_class),
        f"{sanitize_dataset_file_name(dataset_name)}.json",
    )
    return csv_path, sidecar_path


def _load_component_matrix(path: str) -> Tuple[np.ndarray, Dict[str, Any]]:
    """Read one component CSV and its content fingerprint (cache-miss loader)."""

    df = pd.read_csv(path, header=None, dtype="float64", keep_default_na=True, float_precision="round_trip")
    return df.to_numpy(dtype="float64"), runtime_cache_provenance_service.file_fingerprint(path)


def _component_at_target_shape(
    project_name: str,
    sidecar: Mapping[str, Any],
    values: np.ndarray,
    target_settings: Mapping[str, Any],
) -> np.ndarray:
    """Aggregate a hand-entered component to the shape the formula runs at.

    A formula is evaluated at its output's own display shape, and a
    hand-entered precedent may be stored finer than that: its CSV then holds a
    column for every stored period, and every cell the coarse view does not
    read is a cumulative 0. Rolling it up here is the same in-memory read the
    methods' precedent resolver already does, so the formula sees the values
    the Dataset window shows rather than the finer grid underneath them.
    """

    target_origin = int(target_settings.get("origin_length") or 12)
    target_development = int(target_settings.get("development_length") or 12)
    if (target_origin, target_development) == stored_lengths(sidecar):
        return values
    if precedent_cache_service.rollup_reason(sidecar, target_origin, target_development):
        return values
    rows = precedent_cache_service.rollup_rows(
        project_name,
        sidecar,
        values.tolist(),
        target_origin,
        target_development,
    )
    return np.array(
        [[np.nan if value is None else float(value) for value in row] for row in rows],
        dtype="float64",
    )


def _existing_target_settings(project_name: str, reserving_class: str, dataset_name: str) -> Dict[str, Any]:
    sidecar_path = os.path.join(
        config.get_project_dataset_sidecar_dir(project_name, reserving_class),
        f"{sanitize_dataset_file_name(dataset_name)}.json",
    )
    # The mtime-validated cache turns the repeat read into one stat round trip.
    try:
        payload = file_read_cache.read_json_file_cached(sidecar_path)
    except Exception:
        payload = {}
    if not isinstance(payload, dict):
        payload = {}
    # Display, not stored: these settings are what the output is regenerated
    # at, and a calculated output is written at the shape it is shown at.
    return {
        "origin_length": int(payload.get("origin_length") or 12),
        "development_length": int(payload.get("development_length") or 12),
        "cumulative": bool(payload.get("cumulative", True)),
        "calendar": bool(payload.get("calendar", False)),
        "created": _clean_text(payload.get("created")),
    }


def _load_components(
    project_name: str,
    reserving_class: str,
    components: List[str],
    target_settings: Dict[str, Any],
    component_overrides: Dict[str, np.ndarray] | None = None,
    component_paths: Dict[str, str] | None = None,
    component_formats: Dict[str, str] | None = None,
    component_method_sources: Dict[str, Dict[str, str]] | None = None,
) -> Tuple[Dict[str, np.ndarray], List[Dict[str, Any]], List[str]]:
    values: Dict[str, np.ndarray] = {}
    dependency_info: List[Dict[str, Any]] = []
    errors: List[str] = []
    overrides = component_overrides or {}
    exact_paths = component_paths or {}
    expected_formats = component_formats or {}
    exact_method_sources = component_method_sources or {}
    # Every component resolves against the same two folders and this loop only
    # reads, so each folder is enumerated at most once per call instead of once
    # per component. On a network drive that removes one full sidecar sweep per
    # extra dependency.
    dataset_scan: _DatasetCacheScan | None = None
    method_scan: _MethodFolderScan | None = None

    def cached_dataset_scan() -> _DatasetCacheScan:
        nonlocal dataset_scan
        if dataset_scan is None:
            dataset_scan = _scan_dataset_cache_folder(project_name, reserving_class)
        return dataset_scan

    def cached_method_scan() -> _MethodFolderScan:
        nonlocal method_scan
        if method_scan is None:
            method_scan = _scan_dfm_method_folder(project_name, reserving_class)
        return method_scan

    for index, component in enumerate(components):
        component_key = _canon_dataset_name(component)
        expected_format = _clean_text(expected_formats.get(component_key))
        override = overrides.get(_canon_dataset_name(component))
        if override is not None:
            arr = np.asarray(override, dtype="float64")
            if arr.ndim == 1:
                arr = arr.reshape((-1, 1))
            if arr.ndim != 2:
                errors.append(f"Unsupported live preview shape for dependency: {component}")
                continue
            values[f"_d{index}"] = arr
            dependency_info.append({
                "dataset_type_name": component,
                "path": "",
                "source_kind": "live_preview",
            })
            continue
        exact_method_source = exact_method_sources.get(component_key)
        exact_method = (
            exact_method_source
            if isinstance(exact_method_source, dict)
            else {}
        )
        exact_method_path = _clean_text(exact_method.get("path"))
        exact_method_input_path = _clean_text(exact_method.get("input_path"))
        validated_method_input_path = ""
        exact_path = _clean_text(exact_paths.get(component_key))
        candidates: List[Dict[str, Any]]
        method_candidates: List[Dict[str, Any]] | None = None
        if exact_method_path:
            method_folder = config.get_project_method_data_dir(
                project_name,
                reserving_class,
            )
            dataset_folder = config.get_project_dataset_cache_dir(
                project_name,
                reserving_class,
            )
            resolved_method_path = _existing_path_in_dir(
                exact_method_path,
                method_folder,
            )
            method_payload = (
                _read_sidecar(resolved_method_path)
                if resolved_method_path
                else {}
            )
            method_output_matches = component_key in {
                _canon_dataset_name(name)
                for name in _method_output_names(
                    method_payload,
                    resolved_method_path,
                )
            }
            if resolved_method_path and method_output_matches:
                if exact_method_input_path:
                    data_tab = _json_tab(method_payload, "data_tab")
                    details = _json_tab(method_payload, "details_tab")
                    current_input_path = _clean_text(
                        data_tab.get("input data triangle csv path")
                    )
                    resolved_recorded_input = _existing_path_in_dir(
                        exact_method_input_path,
                        dataset_folder,
                    )
                    resolved_current_input = _existing_path_in_dir(
                        current_input_path,
                        dataset_folder,
                    )
                    if (
                        resolved_recorded_input
                        and resolved_current_input
                        and os.path.normcase(os.path.realpath(resolved_recorded_input))
                        == os.path.normcase(os.path.realpath(resolved_current_input))
                    ):
                        validated_method_input_path = resolved_current_input
                    elif resolved_recorded_input and not current_input_path:
                        input_name = _clean_text(details.get("input_triangle"))
                        input_sidecar = _sidecar_for_csv(resolved_recorded_input, dataset_scan)
                        exact_input_names = {
                            _canon_dataset_name(input_sidecar.get("dataset_name")),
                            _canon_dataset_name(input_sidecar.get("dataset_type")),
                            _canon_dataset_name(
                                _csv_base_name(resolved_recorded_input)
                            ),
                        }
                        if _canon_dataset_name(input_name) in exact_input_names:
                            validated_method_input_path = resolved_recorded_input
                method_candidates = [{
                    "path": resolved_method_path,
                    "payload": method_payload,
                    "score": 1,
                    "mtime": os.stat(resolved_method_path).st_mtime,
                }]
                candidates = []
            else:
                candidates = _candidate_csvs(
                    project_name,
                    reserving_class,
                    component,
                    target_settings,
                    expected_data_format=expected_format,
                    scan=cached_dataset_scan(),
                )
        elif exact_path:
            # An exact path needs one sidecar, so reuse a folder observation only
            # when an earlier component already paid for it.
            sidecar = _sidecar_for_csv(exact_path, dataset_scan)
            exact_data_format = _cached_csv_data_format(exact_path, sidecar)
            exact_names = {
                _canon_dataset_name(sidecar.get("dataset_name")),
                _canon_dataset_name(sidecar.get("dataset_type")),
                _canon_dataset_name(_csv_base_name(exact_path)),
            }
            if (
                not os.path.isfile(exact_path)
                or component_key not in exact_names
                or (
                    expected_format
                    and exact_data_format
                    and exact_data_format.lower() != expected_format.lower()
                )
            ):
                errors.append(f"Invalid exact dependency path: {component}")
                continue
            candidates = [{
                "path": exact_path,
                "sidecar": sidecar,
                "data_format": exact_data_format,
                "score": 1,
                "mtime": os.stat(exact_path).st_mtime,
            }]
        else:
            candidates = _candidate_csvs(
                project_name,
                reserving_class,
                component,
                target_settings,
                expected_data_format=expected_format,
                scan=cached_dataset_scan(),
            )
        if not candidates:
            if method_candidates is None:
                method_candidates = _candidate_dfm_methods(
                    project_name,
                    reserving_class,
                    component,
                    scan=cached_method_scan(),
                )
            if not method_candidates:
                errors.append(f"Missing dependency: {component}")
                continue
            if len(method_candidates) > 1:
                errors.append(f"Ambiguous DFM dependency: {component}")
                continue
            method_item = method_candidates[0]
            method_path = str(method_item["path"])
            arr, input_path, error = _build_dfm_method_vector(
                project_name,
                reserving_class,
                method_item.get("payload") if isinstance(method_item.get("payload"), dict) else {},
                target_settings,
                exact_input_path=validated_method_input_path,
                scan=dataset_scan,
            )
            if error or arr is None:
                errors.append(f"Failed to rebuild DFM dependency {component}: {error or 'unknown error'}")
                continue
            var = f"_d{index}"
            values[var] = arr
            fingerprint = runtime_cache_provenance_service.file_fingerprint(method_path)
            dependency_entry = {
                "dataset_type_name": component,
                "path": method_path,
                "source_kind": "dfm_method",
                "input_path": input_path,
                "mtime": fingerprint["mtime_ns"] / 1_000_000_000,
                **fingerprint,
            }
            if input_path:
                input_fingerprint = runtime_cache_provenance_service.file_fingerprint(
                    input_path
                )
                dependency_entry["input_mtime_ns"] = input_fingerprint["mtime_ns"]
                dependency_entry["input_size"] = input_fingerprint["size"]
                dependency_entry["input_sha256"] = input_fingerprint["sha256"]
            dependency_info.append(dependency_entry)
            continue
        if len(candidates) > 1:
            errors.append(f"Ambiguous dependency: {component}")
            continue
        item = candidates[0]
        path = str(item["path"])
        try:
            arr, fingerprint = class_folder_scan_cache.read_matrix_cached(
                path,
                _load_component_matrix,
                stat_hint=(
                    dataset_scan.csv_stats.get(os.path.normcase(path))
                    if dataset_scan is not None
                    else None
                ),
            )
        except Exception as exc:
            errors.append(f"Failed to read dependency {component}: {exc}")
            continue
        sidecar = item.get("sidecar") if isinstance(item.get("sidecar"), dict) else {}
        try:
            arr = _component_at_target_shape(project_name, sidecar, arr, target_settings)
        except ValueError as exc:
            errors.append(f"Failed to read dependency {component}: {exc}")
            continue
        var = f"_d{index}"
        values[var] = arr
        # Stored, not displayed: this records what was read from ``path``.
        stored_origin, stored_development = stored_lengths(sidecar)
        dependency_info.append({
            "dataset_type_name": component,
            "dataset_name": _clean_text(sidecar.get("dataset_name") or component),
            "path": path,
            "source_kind": _clean_text(sidecar.get("source_kind")),
            "data_format": _clean_text(sidecar.get("data_format")),
            "stored_origin_length": stored_origin,
            "stored_development_length": stored_development,
            "cumulative": sidecar.get("cumulative"),
            "calendar": sidecar.get("calendar"),
            "mtime": fingerprint["mtime_ns"] / 1_000_000_000,
            **fingerprint,
        })
    return values, dependency_info, errors


def _array_from_preview_values(
    values: List[List[Any]] | None,
    mask: List[List[bool]] | None = None,
) -> np.ndarray:
    rows = values if isinstance(values, list) else []
    out: List[List[float]] = []
    for r, row in enumerate(rows):
        source_row = row if isinstance(row, list) else []
        mask_row = mask[r] if isinstance(mask, list) and r < len(mask) and isinstance(mask[r], list) else None
        converted: List[float] = []
        for c, value in enumerate(source_row):
            if mask_row is not None and c < len(mask_row) and not bool(mask_row[c]):
                converted.append(np.nan)
                continue
            number = _finite_float(value)
            converted.append(number if number is not None else np.nan)
        out.append(converted)
    return np.asarray(out, dtype="float64")


def _jsonable_matrix(arr: np.ndarray) -> List[List[float | None]]:
    matrix = np.asarray(arr, dtype="float64")
    if matrix.ndim == 1:
        matrix = matrix.reshape((-1, 1))
    return [
        [float(value) if np.isfinite(value) else None for value in row]
        for row in matrix.tolist()
    ]


def _matrix_mask(arr: np.ndarray) -> List[List[bool]]:
    matrix = np.asarray(arr, dtype="float64")
    if matrix.ndim == 1:
        matrix = matrix.reshape((-1, 1))
    return [
        [bool(np.isfinite(value)) for value in row]
        for row in matrix.tolist()
    ]


def _latest_diagonal_or_vector_values(arr: np.ndarray, data_format: str) -> List[float | None]:
    matrix = np.asarray(arr, dtype="float64")
    if matrix.ndim == 1:
        matrix = matrix.reshape((-1, 1))
    if _clean_text(data_format).lower() == "vector":
        return [
            float(row[0]) if len(row) and np.isfinite(row[0]) else None
            for row in matrix
        ]
    out: List[float | None] = []
    for row in matrix:
        picked: float | None = None
        for value in reversed(row):
            if np.isfinite(value):
                picked = float(value)
                break
        out.append(picked)
    return out


def _calculated_rows_by_key(project_name: str) -> Dict[str, Dict[str, Any]]:
    return {
        _canon_dataset_name(row["name"]): row
        for row in _dataset_type_rows(project_name)
        if row.get("calculated") and not row.get("generated") and _clean_text(row.get("formula"))
    }


def _dependency_map(project_name: str, rows: List[Dict[str, Any]] | None = None) -> Dict[str, Set[str]]:
    rows = rows if rows is not None else _dataset_type_rows(project_name)
    known_names = [row["name"] for row in rows]
    out: Dict[str, Set[str]] = {}
    for row in rows:
        if not row.get("calculated") or row.get("generated") or not _clean_text(row.get("formula")):
            continue
        target_key = _canon_dataset_name(row["name"])
        for component in _formula_components(row["formula"], known_names):
            comp_key = _canon_dataset_name(component)
            if comp_key:
                out.setdefault(comp_key, set()).add(target_key)
    return out


def _target_dependency_map(project_name: str, rows: List[Dict[str, Any]] | None = None) -> Dict[str, Set[str]]:
    rows = rows if rows is not None else _dataset_type_rows(project_name)
    known_names = [row["name"] for row in rows]
    out: Dict[str, Set[str]] = {}
    for row in rows:
        if not row.get("calculated") or row.get("generated") or not _clean_text(row.get("formula")):
            continue
        target_key = _canon_dataset_name(row["name"])
        deps = {
            _canon_dataset_name(component)
            for component in _formula_components(row["formula"], known_names)
            if _canon_dataset_name(component)
        }
        out[target_key] = deps
    return out


def _existing_dataset_keys(project_name: str, reserving_class: str) -> Set[str]:
    index = dataset_instance_index_service.get_index(project_name, reserving_class, refresh=False)
    keys: Set[str] = set()
    for item in index.get("files", []) if isinstance(index.get("files"), list) else []:
        for value in [item.get("name"), item.get("dataset_type")]:
            key = _canon_dataset_name(value)
            if key:
                keys.add(key)
    return keys


def _downstream_keys(
    project_name: str,
    changed_names: List[str],
    rows: List[Dict[str, Any]] | None = None,
) -> List[str]:
    dep_map = _dependency_map(project_name, rows)
    seen: Set[str] = set()
    out: List[str] = []
    queue = [_canon_dataset_name(name) for name in changed_names if _canon_dataset_name(name)]
    while queue:
        key = queue.pop(0)
        for target in sorted(dep_map.get(key, set())):
            if target in seen:
                continue
            seen.add(target)
            out.append(target)
            queue.append(target)

    target_set = set(out)
    deps_by_target = _target_dependency_map(project_name, rows)
    ordered: List[str] = []
    visiting: Set[str] = set()
    visited: Set[str] = set()

    def visit(key: str) -> None:
        if key in visited:
            return
        if key in visiting:
            return
        visiting.add(key)
        for dep in sorted(deps_by_target.get(key, set())):
            if dep in target_set:
                visit(dep)
        visiting.remove(key)
        visited.add(key)
        ordered.append(key)

    for key in out:
        visit(key)
    return ordered


def _existing_downstream_keys(
    project_name: str,
    reserving_class: str,
    changed_names: List[str],
    rows: List[Dict[str, Any]] | None = None,
) -> List[str]:
    existing_keys = _existing_dataset_keys(project_name, reserving_class)
    if not existing_keys:
        return []
    return [
        key
        for key in _downstream_keys(project_name, changed_names, rows)
        if key in existing_keys
    ]


def existing_downstream_dataset_types(
    project_name: str,
    reserving_class: str,
    changed_names: List[str],
) -> List[Dict[str, Any]]:
    """Name every calculated dataset a walk from ``changed_names`` can reach.

    Read-only view of the same formula graph ``_recalculate_dependents_impl``
    walks, in the same dependency order, restricted to instances that exist in
    this reserving class. The two-step save shows this list before it writes
    anything; it is a superset of what a walk ultimately rewrites, because
    only the walk itself can tell whether an intermediate output changed.
    """

    rows = _dataset_type_rows(project_name)
    rows_by_key = {_canon_dataset_name(row.get("name")): row for row in rows}
    out: List[Dict[str, Any]] = []
    for key in _existing_downstream_keys(project_name, reserving_class, changed_names, rows):
        row = rows_by_key.get(key)
        if not row:
            continue
        out.append({
            "dataset_name": str(row.get("name") or ""),
            "data_format": str(row.get("data_format") or ""),
            "formula": str(row.get("formula") or ""),
        })
    return out


def dataset_type_graph_signature(project_name: str) -> List[List[Any]]:
    """Project the dataset-type rows that decide the calculated dependency graph.

    The two-step save fingerprints this so a formula edited between the plan
    the user reviewed and the commit is caught before the save lands.
    """

    return sorted(
        (
            [
                _canon_dataset_name(row.get("name")),
                bool(row.get("calculated")),
                bool(row.get("generated")),
                _clean_text(row.get("formula")),
            ]
            for row in _dataset_type_rows(project_name)
            if _canon_dataset_name(row.get("name"))
        ),
        key=lambda item: item[0],
    )


def _recalculate_dataset_impl(
    project_name: str,
    reserving_class: str,
    dataset_type_name: str,
    *,
    component_paths: Dict[str, str] | None = None,
    component_method_sources: Dict[str, Dict[str, str]] | None = None,
    dataset_type_rows: List[Dict[str, Any]] | None = None,
    mark_dependents_review: bool = True,
) -> Dict[str, Any]:
    if dataset_type_rows is None:
        rows_by_key = _calculated_rows_by_key(project_name)
        all_rows = _dataset_type_rows(project_name)
    else:
        all_rows = dataset_type_rows
        rows_by_key = {
            _canon_dataset_name(item.get("name")): item
            for item in all_rows
            if item.get("calculated") and not item.get("generated") and _clean_text(item.get("formula"))
        }
    row = rows_by_key.get(_canon_dataset_name(dataset_type_name))
    if not row:
        return {"ok": False, "dataset_type_name": dataset_type_name, "skipped": True, "reason": "not_calculated"}

    known_names = [item["name"] for item in all_rows]
    expr, refs = _replace_formula_refs(row["formula"], known_names)
    ordered_components = [refs[key] for key in sorted(refs.keys(), key=lambda item: int(item[2:]))]

    settings = _existing_target_settings(project_name, reserving_class, row["name"])
    values, precedents, errors = _load_components(
        project_name,
        reserving_class,
        ordered_components,
        settings,
        component_paths=component_paths,
        component_method_sources=component_method_sources,
        component_formats={
            _canon_dataset_name(item.get("name")): _clean_text(item.get("data_format"))
            for item in all_rows
            if _canon_dataset_name(item.get("name"))
        },
    )
    if errors:
        missing_prefix = "Missing dependency: "
        return {
            "ok": False,
            "dataset_type_name": row["name"],
            "skipped": True,
            "reason": "dependency_error",
            "errors": errors,
            "missing_dependencies": [
                error[len(missing_prefix):]
                for error in errors
                if error.startswith(missing_prefix)
            ],
        }

    eval_values: Dict[str, np.ndarray] = {}
    for var, ref_name in refs.items():
        try:
            idx = ordered_components.index(ref_name)
        except ValueError:
            return {"ok": False, "dataset_type_name": row["name"], "skipped": True, "reason": f"missing_reference:{ref_name}"}
        eval_values[var] = values[f"_d{idx}"]

    try:
        parsed = ast.parse(expr, mode="eval")
        with np.errstate(divide="ignore", invalid="ignore"):
            result = _eval_ast(parsed, eval_values)
    except Exception as exc:
        return {
            "ok": False,
            "dataset_type_name": row["name"],
            "skipped": True,
            "reason": "formula_error",
            "errors": [str(exc)],
        }

    arr = np.asarray(result, dtype="float64")
    if arr.ndim == 0:
        first = next(iter(eval_values.values()), np.zeros((1, 1), dtype="float64"))
        arr = np.full(first.shape, float(arr), dtype="float64")
    if arr.ndim == 1:
        arr = arr.reshape((-1, 1))
    if arr.ndim != 2:
        return {"ok": False, "dataset_type_name": row["name"], "skipped": True, "reason": "unsupported_result_shape"}

    csv_path, sidecar_path = _target_paths(
        project_name,
        reserving_class,
        row["name"],
        settings,
        row.get("data_format") or "Triangle",
    )
    now = _now_utc_iso()
    existing_sidecar = _read_sidecar(sidecar_path)
    created = existing_sidecar.get("created") or settings.get("created") or now
    user_name = _current_user_name()
    action_value = "Update" if existing_sidecar else "Insert"
    default_format_settings = dataset_number_format_service.dataset_type_number_format_settings(
        row["name"],
    )
    number_format = dataset_number_format_service.normalize_number_format(
        existing_sidecar.get("number_format") or default_format_settings["number_format"]
    )
    decimal_places = existing_sidecar.get("decimal_places")
    if decimal_places is None:
        decimal_places = dataset_number_format_service.number_format_decimal_places(number_format)
    payload = {
        **({"audit_log": existing_sidecar.get("audit_log")} if existing_sidecar else {}),
        "dataset_name": row["name"],
        "dataset_type": row["name"],
        "reserving_class": reserving_class,
        "project_name": project_name,
        "source_kind": "calculated",
        "data_format": row.get("data_format") or "Triangle",
        "origin_length": settings.get("origin_length") or 12,
        "development_length": settings.get("development_length") or 12,
        # The formula was evaluated at these lengths and the CSV written at
        # them, so they are the shape this output is stored at.
        **stored_length_fields(
            row.get("data_format") or "Triangle",
            settings.get("origin_length") or 12,
            settings.get("development_length") or 12,
        ),
        "cumulative": bool(settings.get("cumulative", True)),
        "calendar": bool(settings.get("calendar", False)),
        "show_subtotal": normalize_show_subtotal(existing_sidecar.get("show_subtotal")),
        "csv_file": os.path.basename(csv_path),
        "created": created,
        "updated_at": now,
        "modified_by": user_name,
        "calculated": True,
        "method_type": dataset_sidecar_status_service.METHOD_TYPE_NONE,
        "status": dataset_sidecar_status_service.STATUS_CURRENT,
        "dependents": existing_sidecar.get("dependents", []),
        "number_format": number_format,
        "decimal_places": dataset_number_format_service.normalize_decimal_places(
            decimal_places,
            default_format_settings["decimal_places"],
        ),
    }
    apply_sidecar_graph_fields(payload, project_name, row["name"], precedents)
    from app_server.services.dataset_service import (
        _append_dataset_audit_entry,
        _write_dataset_csv_and_sidecar,
    )

    _append_dataset_audit_entry(payload, action_value, event_date=now, user_name=user_name)

    os.makedirs(os.path.dirname(csv_path), exist_ok=True)
    os.makedirs(os.path.dirname(sidecar_path), exist_ok=True)
    _write_dataset_csv_and_sidecar(pd.DataFrame(arr), csv_path, sidecar_path, payload)
    # The sidecar names the precedents and nothing else. What this output was
    # built from -- the formula and each dependency's file and fingerprint --
    # is the technical record beside the CSV, which the exact-cache check reads.
    runtime_cache_provenance_service.record_calculated(
        csv_path,
        identity=runtime_cache_provenance_service.calculated_cache_identity(
            csv_path,
            project_name=project_name,
            reserving_class=reserving_class,
            dataset_name=row["name"],
            dataset_type=row["name"],
        ),
        formula=row["formula"],
        dependencies=_dependency_fingerprints(precedents),
    )
    status_updates = (
        dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
            project_name,
            reserving_class,
            [row["name"]],
        )
        if mark_dependents_review
        else []
    )
    config.DATASETS["arcrhotri_" + hashlib.sha1(csv_path.encode("utf-8")).hexdigest()[:16]] = csv_path

    return {
        "ok": True,
        "dataset_type_name": row["name"],
        "path": csv_path,
        "sidecar_path": sidecar_path,
        "precedents": payload.get("precedents", []),
        "dependents": payload.get("dependents", []),
        "status_updates": status_updates,
    }


def recalculate_dataset(
    project_name: str,
    reserving_class: str,
    dataset_type_name: str,
    *,
    component_paths: Dict[str, str] | None = None,
    component_method_sources: Dict[str, Dict[str, str]] | None = None,
    dataset_type_rows: List[Dict[str, Any]] | None = None,
    mark_dependents_review: bool = True,
) -> Dict[str, Any]:
    with dataset_sidecar_status_service.reserving_class_io_lock(project_name, reserving_class):
        return _recalculate_dataset_impl(
            project_name,
            reserving_class,
            dataset_type_name,
            component_paths=component_paths,
            component_method_sources=component_method_sources,
            dataset_type_rows=dataset_type_rows,
            mark_dependents_review=mark_dependents_review,
        )


def _refresh_link_driven_dependents(
    project_name: str,
    reserving_class: str,
    root_names: Sequence[str],
    visited_keys: Set[str],
    link_updates: Dict[str, List[Any]],
) -> List[str]:
    """Refresh every link-driven dataset a set of fresh roots reaches.

    Follows the roots' persisted ``dependents`` edges to plain input datasets
    whose cells are driven by ArcRho links, re-evaluates each one through
    ``dataset_link_refresh_service``, and chains onward through link-driven
    dependents of what it refreshed. ``visited_keys`` carries everything this
    walk already rewrote or visited, which is the cycle guard: a dataset is
    refreshed at most once per walk, so a link that points back upstream (a
    candidate ultimate reading the Result Selection it feeds, say) converges
    instead of looping. Returns the datasets whose values changed; failures
    and Excel keep-stale warnings accumulate into ``link_updates``.
    """

    from app_server.services import dataset_link_refresh_service

    queue: List[str] = [_clean_text(name) for name in root_names if _clean_text(name)]
    fresh: List[str] = []
    walked_roots: Set[str] = set()
    while queue:
        root = queue.pop(0)
        root_key = _canon_dataset_name(root)
        if not root_key or root_key in walked_roots:
            continue
        walked_roots.add(root_key)
        root_payload = dataset_sidecar_status_service.read_sidecar(
            dataset_sidecar_status_service.sidecar_path(project_name, reserving_class, root)
        )
        for name in dataset_sidecar_status_service.entry_names((root_payload or {}).get("dependents")):
            key = _canon_dataset_name(name)
            if not key or key in visited_keys:
                continue
            dependent = dataset_sidecar_status_service.read_sidecar(
                dataset_sidecar_status_service.sidecar_path(project_name, reserving_class, name)
            )
            if not dependent:
                continue
            if dataset_sidecar_status_service.normalize_method_type(
                dependent.get("method_type"), dependent.get("source_kind")
            ) != dataset_sidecar_status_service.METHOD_TYPE_NONE:
                continue
            if _clean_text(dependent.get("source_kind")).casefold() != "input":
                continue
            if not (dependent.get("internal_links") or dependent.get("formula_links")):
                continue
            visited_keys.add(key)
            try:
                result = dataset_link_refresh_service.refresh_dataset_links(
                    project_name, reserving_class, name
                )
            except Exception as exc:
                result = {
                    "ok": False,
                    "dataset_name": name,
                    "reason": "link_error",
                    "errors": [str(exc)],
                }
            for warning in result.get("warnings") or []:
                link_updates["warnings"].append({
                    "dataset_name": name,
                    "reference": str(warning.get("reference") or ""),
                    "reason": str(warning.get("reason") or ""),
                })
            if not result.get("ok"):
                link_updates["failed"].append(name)
                link_updates["errors"].append({
                    "dataset_name": name,
                    "reason": _clean_text(result.get("reason")) or "link_error",
                    "errors": [str(item) for item in result.get("errors") or []],
                })
                continue
            if result.get("changed"):
                fresh.append(name)
                link_updates["refreshed"].append(name)
                queue.append(name)
    return fresh


def _dfm_unchanged_names(dfm_updates: Mapping[str, Any] | None) -> List[str]:
    """Name every DFM output the DFM wave visited whose published values held.

    A DFM recomputed to the same publication -- its ratio basis moved, say,
    which only redraws the Ratios tab -- one only status-refreshed, or one
    skipped because none of its real inputs changed is current as it stands.
    The Result Selection wave has to hear that: it walks from the same roots,
    reaches such a DFM as their dependent, and would otherwise block it as a
    precedent still waiting for an explicit refresh, refusing every Result
    Selection that loads it.
    """
    names: List[str] = []
    if not isinstance(dfm_updates, Mapping):
        return names
    for item in dfm_updates.get("updated", []):
        if isinstance(item, Mapping) and not item.get("output_changed"):
            names.extend(
                _clean_text(value)
                for value in (item.get("dataset_name"), item.get("dataset_type"))
                if _clean_text(value)
            )
    for field in ("status_refreshed", "skipped"):
        names.extend(
            _clean_text(item.get("dataset_name"))
            for item in dfm_updates.get(field, [])
            if isinstance(item, Mapping) and _clean_text(item.get("dataset_name"))
        )
    return names


def _recalculate_dependents_impl(
    project_name: str,
    reserving_class: str,
    changed_dataset_name: str,
    changed_dataset_type_name: str = "",
    *,
    include_dfm: bool = True,
    include_result_selection: bool = True,
    include_berquist_sherman: bool = True,
    include_bornhuetter_ferguson: bool = True,
    include_cape_cod: bool = True,
    include_bootstrap: bool = True,
    finalize_method_review_status: bool = True,
    rebuild_index: bool = True,
    additional_roots: Sequence[Tuple[str, str]] | None = None,
    progress_callback: Callable[[str, int, int, str], None] | None = None,
) -> Dict[str, Any]:
    # Extra roots let one Engine-hosted walk cover several coalesced saves.
    # A progress callback exception (for example a lost reserving-class
    # lease) intentionally aborts the walk; the follow-up walk self-heals.
    changed_root_names = [changed_dataset_name, changed_dataset_type_name]
    seen_root_keys = {
        _canon_dataset_name(name)
        for name in changed_root_names
        if _canon_dataset_name(name)
    }
    for extra_root in additional_roots or []:
        for value in extra_root:
            cleaned = _clean_text(value)
            key = _canon_dataset_name(cleaned)
            if cleaned and key and key not in seen_root_keys:
                seen_root_keys.add(key)
                changed_root_names.append(cleaned)

    def _notify(stage: str, completed: int, total: int, label: str) -> None:
        if progress_callback is not None:
            progress_callback(stage, completed, total, label)

    changed = list(changed_root_names)
    dfm_updates = None
    dfm_output_names: List[str] = []
    failed_dfm_names: List[str] = []
    if include_dfm:
        _notify("dfm", 0, 0, "Refreshing DFM methods")
        try:
            from app_server.services import dfm_service

            dfm_updates = dfm_service.refresh_dependents(
                project_name,
                reserving_class,
                changed,
                finalize_method_review_status=False,
            )
            dfm_output_names = [
                _clean_text(value)
                for item in dfm_updates.get("updated", [])
                if item.get("output_changed")
                for value in (item.get("dataset_name"), item.get("dataset_type"))
                if _clean_text(value)
            ]
            failed_dfm_names = [
                _clean_text(value)
                for item in dfm_updates.get("errors", [])
                for value in (item.get("dataset_name"), item.get("dataset_type"))
                if _clean_text(value)
            ]
        except Exception as err:
            dfm_updates = {
                "ok": False,
                "errors": [{"reason": str(err)}],
                "updated": [],
            }
    changed.extend([*dfm_output_names, *failed_dfm_names])
    # Link-driven datasets (plain inputs whose cells are driven by ArcRho
    # links) refresh in the wave order their sources publish in: once here for
    # links reading the saved roots and DFM outputs, once per recalculated
    # dataset inside the loop below, and once more after the method waves for
    # links reading method outputs. ``visited_link_keys`` is the cycle guard.
    link_updates: Dict[str, List[Any]] = {
        "refreshed": [],
        "failed": [],
        "warnings": [],
        "errors": [],
    }
    visited_link_keys: Set[str] = {
        _canon_dataset_name(name) for name in changed if _canon_dataset_name(name)
    }
    _notify("linked_datasets", 0, 0, "Refreshing link-driven datasets")
    changed.extend(
        _refresh_link_driven_dependents(
            project_name,
            reserving_class,
            changed,
            visited_link_keys,
            link_updates,
        )
    )
    dataset_type_rows = _dataset_type_rows(project_name)
    targets = list(_existing_downstream_keys(project_name, reserving_class, changed, dataset_type_rows))
    rows_by_key = {
        _canon_dataset_name(item.get("name")): item
        for item in dataset_type_rows
        if item.get("calculated") and not item.get("generated") and _clean_text(item.get("formula"))
    }
    known_names = [item.get("name") for item in dataset_type_rows if _clean_text(item.get("name"))]
    dependencies_by_key = {
        key: {
            _canon_dataset_name(name)
            for name in _formula_components(row.get("formula") or "", known_names)
            if _canon_dataset_name(name)
        }
        for key, row in rows_by_key.items()
    }
    results: List[Dict[str, Any]] = []
    failed_or_blocked: Set[str] = {
        _canon_dataset_name(name) for name in failed_dfm_names if _canon_dataset_name(name)
    }
    processed_target_keys: Set[str] = set()
    for key in targets:
        if key in processed_target_keys:
            continue
        processed_target_keys.add(key)
        row = rows_by_key.get(key)
        if not row:
            continue
        _notify(
            "calculated_datasets",
            len(results),
            len([target for target in targets if target in rows_by_key]),
            str(row.get("name") or ""),
        )
        blocked_by = sorted(dependencies_by_key.get(key, set()) & failed_or_blocked)
        if blocked_by:
            result = {
                "ok": False,
                "dataset_type_name": row["name"],
                "skipped": True,
                "reason": "upstream_calculation_failed",
                "errors": [
                    "Skipped because an upstream calculated dependency did not refresh: "
                    + ", ".join(blocked_by)
                ],
            }
        else:
            try:
                result = recalculate_dataset(
                    project_name,
                    reserving_class,
                    row["name"],
                    dataset_type_rows=dataset_type_rows,
                    mark_dependents_review=False,
                )
            except Exception as exc:
                result = {
                    "ok": False,
                    "dataset_type_name": row["name"],
                    "skipped": True,
                    "reason": "calculation_error",
                    "errors": [str(exc)],
                }
        if not result.get("ok"):
            failed_or_blocked.add(key)
        step = {
            **result,
            "status": "updated" if result.get("ok") else "skipped",
        }
        results.append(step)
        if include_dfm:
            try:
                from app_server.services import dfm_service

                calculated_name = _clean_text(result.get("dataset_type_name") or row["name"])
                next_dfm = dfm_service.refresh_dependents(
                    project_name,
                    reserving_class,
                    [calculated_name],
                    blocked_precedent_names=[calculated_name] if not result.get("ok") else [],
                    finalize_method_review_status=False,
                )
            except Exception as err:
                next_dfm = {
                    "ok": False,
                    "updated": [],
                    "status_refreshed": [],
                    "skipped": [],
                    "errors": [{"reason": str(err)}],
                }
            if dfm_updates is None:
                dfm_updates = next_dfm
            else:
                dfm_updates["ok"] = bool(dfm_updates.get("ok")) and bool(next_dfm.get("ok"))
                for field in ("updated", "status_refreshed", "skipped", "errors"):
                    dfm_updates.setdefault(field, []).extend(next_dfm.get(field, []))
            next_output_roots = [
                _clean_text(value)
                for item in next_dfm.get("updated", [])
                if item.get("output_changed")
                for value in (item.get("dataset_name"), item.get("dataset_type"))
                if _clean_text(value)
            ]
            next_failed_roots = [
                _clean_text(value)
                for item in next_dfm.get("errors", [])
                for value in (item.get("dataset_name"), item.get("dataset_type"))
                if _clean_text(value)
            ]
            for name in next_output_roots:
                if _canon_dataset_name(name) not in {_canon_dataset_name(item) for item in dfm_output_names}:
                    dfm_output_names.append(name)
            for name in next_failed_roots:
                if _canon_dataset_name(name) not in {_canon_dataset_name(item) for item in failed_dfm_names}:
                    failed_dfm_names.append(name)
                if _canon_dataset_name(name):
                    failed_or_blocked.add(_canon_dataset_name(name))
            next_roots = [*next_output_roots, *next_failed_roots]
            if next_roots:
                for next_key in _existing_downstream_keys(
                    project_name,
                    reserving_class,
                    next_roots,
                    dataset_type_rows,
                ):
                    if next_key not in processed_target_keys and next_key not in targets:
                        targets.append(next_key)
        if result.get("ok"):
            link_fresh_names = _refresh_link_driven_dependents(
                project_name,
                reserving_class,
                [_clean_text(result.get("dataset_type_name") or row["name"])],
                visited_link_keys,
                link_updates,
            )
            for name in link_fresh_names:
                for next_key in _existing_downstream_keys(
                    project_name,
                    reserving_class,
                    [name],
                    dataset_type_rows,
                ):
                    if next_key not in processed_target_keys and next_key not in targets:
                        targets.append(next_key)

    failed_dataset_names = [
        _clean_text(result.get("dataset_type_name"))
        for result in results
        if not result.get("ok") and _clean_text(result.get("dataset_type_name"))
    ]
    for failed_name in failed_dataset_names:
        if failed_name:
            dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
                project_name,
                reserving_class,
                [failed_name],
            )

    result_selection_updates = None
    if include_result_selection:
        _notify("result_selection", 0, 0, "Refreshing Result Selection methods")
        try:
            from app_server.services import result_selection_service

            fresh_names = [
                *changed_root_names,
                *dfm_output_names,
                *failed_dfm_names,
                *link_updates["refreshed"],
                *link_updates["failed"],
            ]
            fresh_names.extend(
                item.get("dataset_type_name")
                for item in results
                if item.get("ok") and _clean_text(item.get("dataset_type_name"))
            )
            result_selection_updates = result_selection_service.refresh_dependents(
                project_name,
                reserving_class,
                fresh_names,
                rebuild_index=False,
                allow_status_current=True,
                blocked_precedent_names=[
                    *failed_dfm_names,
                    *failed_dataset_names,
                    *link_updates["failed"],
                ],
                unchanged_precedent_names=_dfm_unchanged_names(dfm_updates),
                finalize_method_review_status=False,
            )
        except Exception as err:
            result_selection_updates = {
                "ok": False,
                "errors": [{"reason": str(err)}],
                "updated": [],
            }

    # A Berquist Sherman output is a triangle other methods build on -- a DFM
    # over the adjusted paid claims, say -- while its own ultimate claim counts
    # usually come from a Result Selection. So the wave runs after Result
    # Selection and before the vector methods: a republished B&S triangle is
    # fed through DFM, calculated, and Result Selection inside the wave, and
    # what that nested walk refreshed joins the roots of the BF, CC, and
    # Bootstrap waves below.
    berquist_sherman_updates = None
    if include_berquist_sherman:
        _notify("berquist_sherman", 0, 0, "Refreshing Berquist Sherman methods")
        try:
            from app_server.services import berquist_sherman_service

            dfm_fresh_names = [
                _clean_text(value)
                for item in (dfm_updates or {}).get("updated", [])
                for value in (item.get("dataset_name"), item.get("dataset_type"))
                if _clean_text(value)
            ]
            dfm_fresh_names.extend(
                _clean_text(item.get("dataset_name"))
                for item in (dfm_updates or {}).get("status_refreshed", [])
                if _clean_text(item.get("dataset_name"))
            )
            calculated_fresh_names = [
                _clean_text(item.get("dataset_type_name"))
                for item in results
                if item.get("ok") and _clean_text(item.get("dataset_type_name"))
            ]
            result_selection_fresh_names = [
                _clean_text(item.get("dataset_name"))
                for field in ("updated", "status_refreshed")
                for item in (result_selection_updates or {}).get(field, [])
                if _clean_text(item.get("dataset_name"))
            ]
            result_selection_fresh_names.extend(
                _clean_text(name)
                for name in (result_selection_updates or {}).get("downstream_fresh_names", [])
                if _clean_text(name)
            )
            failed_result_selection_names = [
                _clean_text(item.get("dataset_name"))
                for item in (result_selection_updates or {}).get("errors", [])
                if _clean_text(item.get("dataset_name"))
            ]
            failed_result_selection_names.extend(
                _clean_text(name)
                for name in (result_selection_updates or {}).get("downstream_blocked_names", [])
                if _clean_text(name)
            )
            berquist_sherman_roots = [
                *changed_root_names,
                *dfm_fresh_names,
                *failed_dfm_names,
                *calculated_fresh_names,
                *failed_dataset_names,
                *link_updates["refreshed"],
                *link_updates["failed"],
                *result_selection_fresh_names,
                *failed_result_selection_names,
            ]
            berquist_sherman_updates = berquist_sherman_service.refresh_dependents(
                project_name,
                reserving_class,
                berquist_sherman_roots,
                rebuild_index=False,
                blocked_precedent_names=[
                    *failed_dfm_names,
                    *failed_dataset_names,
                    *failed_result_selection_names,
                    *link_updates["failed"],
                ],
                finalize_method_review_status=False,
            )
        except Exception as err:
            berquist_sherman_updates = {
                "ok": False,
                "errors": [{"reason": str(err)}],
                "updated": [],
            }
    berquist_sherman_fresh_names = [
        _clean_text(item.get("dataset_name"))
        for field in ("updated", "status_refreshed")
        for item in (berquist_sherman_updates or {}).get(field, [])
        if _clean_text(item.get("dataset_name"))
    ]
    berquist_sherman_fresh_names.extend(
        _clean_text(name)
        for name in (berquist_sherman_updates or {}).get("downstream_fresh_names", [])
        if _clean_text(name)
    )
    failed_berquist_sherman_names = [
        _clean_text(item.get("dataset_name"))
        for item in (berquist_sherman_updates or {}).get("errors", [])
        if _clean_text(item.get("dataset_name"))
    ]
    failed_berquist_sherman_names.extend(
        _clean_text(name)
        for name in (berquist_sherman_updates or {}).get("downstream_blocked_names", [])
        if _clean_text(name)
    )

    bornhuetter_ferguson_updates = None
    if include_bornhuetter_ferguson:
        _notify(
            "bornhuetter_ferguson", 0, 0, "Refreshing Bornhuetter Ferguson methods"
        )
        try:
            from app_server.services import bornhuetter_ferguson_service

            dfm_fresh_names = [
                _clean_text(value)
                for item in (dfm_updates or {}).get("updated", [])
                for value in (item.get("dataset_name"), item.get("dataset_type"))
                if _clean_text(value)
            ]
            dfm_fresh_names.extend(
                _clean_text(item.get("dataset_name"))
                for item in (dfm_updates or {}).get("status_refreshed", [])
                if _clean_text(item.get("dataset_name"))
            )
            calculated_fresh_names = [
                _clean_text(item.get("dataset_type_name"))
                for item in results
                if item.get("ok") and _clean_text(item.get("dataset_type_name"))
            ]
            result_selection_fresh_names = [
                _clean_text(item.get("dataset_name"))
                for field in ("updated", "status_refreshed")
                for item in (result_selection_updates or {}).get(field, [])
                if _clean_text(item.get("dataset_name"))
            ]
            result_selection_fresh_names.extend(
                _clean_text(name)
                for name in (result_selection_updates or {}).get("downstream_fresh_names", [])
                if _clean_text(name)
            )
            failed_result_selection_names = [
                _clean_text(item.get("dataset_name"))
                for item in (result_selection_updates or {}).get("errors", [])
                if _clean_text(item.get("dataset_name"))
            ]
            failed_result_selection_names.extend(
                _clean_text(name)
                for name in (result_selection_updates or {}).get("downstream_blocked_names", [])
                if _clean_text(name)
            )
            bf_roots = [
                *changed_root_names,
                *dfm_fresh_names,
                *failed_dfm_names,
                *calculated_fresh_names,
                *failed_dataset_names,
                *link_updates["refreshed"],
                *link_updates["failed"],
                *result_selection_fresh_names,
                *failed_result_selection_names,
                *berquist_sherman_fresh_names,
                *failed_berquist_sherman_names,
            ]
            bornhuetter_ferguson_updates = bornhuetter_ferguson_service.refresh_dependents(
                project_name,
                reserving_class,
                bf_roots,
                rebuild_index=False,
                blocked_precedent_names=[
                    *failed_dfm_names,
                    *failed_dataset_names,
                    *failed_result_selection_names,
                    *failed_berquist_sherman_names,
                    *link_updates["failed"],
                ],
                finalize_method_review_status=False,
            )
        except Exception as err:
            bornhuetter_ferguson_updates = {
                "ok": False,
                "errors": [{"reason": str(err)}],
                "updated": [],
            }

    cape_cod_updates = None
    if include_cape_cod:
        _notify("cape_cod", 0, 0, "Refreshing Cape Cod methods")
        try:
            from app_server.services import cape_cod_service

            dfm_fresh_names = [
                _clean_text(value)
                for item in (dfm_updates or {}).get("updated", [])
                for value in (item.get("dataset_name"), item.get("dataset_type"))
                if _clean_text(value)
            ]
            dfm_fresh_names.extend(
                _clean_text(item.get("dataset_name"))
                for item in (dfm_updates or {}).get("status_refreshed", [])
                if _clean_text(item.get("dataset_name"))
            )
            calculated_fresh_names = [
                _clean_text(item.get("dataset_type_name"))
                for item in results
                if item.get("ok") and _clean_text(item.get("dataset_type_name"))
            ]
            result_selection_fresh_names = [
                _clean_text(item.get("dataset_name"))
                for field in ("updated", "status_refreshed")
                for item in (result_selection_updates or {}).get(field, [])
                if _clean_text(item.get("dataset_name"))
            ]
            result_selection_fresh_names.extend(
                _clean_text(name)
                for name in (result_selection_updates or {}).get("downstream_fresh_names", [])
                if _clean_text(name)
            )
            failed_result_selection_names = [
                _clean_text(item.get("dataset_name"))
                for item in (result_selection_updates or {}).get("errors", [])
                if _clean_text(item.get("dataset_name"))
            ]
            failed_result_selection_names.extend(
                _clean_text(name)
                for name in (result_selection_updates or {}).get("downstream_blocked_names", [])
                if _clean_text(name)
            )
            bornhuetter_ferguson_fresh_names = [
                _clean_text(item.get("dataset_name"))
                for field in ("updated", "status_refreshed")
                for item in (bornhuetter_ferguson_updates or {}).get(field, [])
                if _clean_text(item.get("dataset_name"))
            ]
            failed_bornhuetter_ferguson_names = [
                _clean_text(item.get("dataset_name"))
                for item in (bornhuetter_ferguson_updates or {}).get("errors", [])
                if _clean_text(item.get("dataset_name"))
            ]
            cape_cod_roots = [
                *changed_root_names,
                *dfm_fresh_names,
                *failed_dfm_names,
                *calculated_fresh_names,
                *failed_dataset_names,
                *link_updates["refreshed"],
                *link_updates["failed"],
                *result_selection_fresh_names,
                *failed_result_selection_names,
                *berquist_sherman_fresh_names,
                *failed_berquist_sherman_names,
                *bornhuetter_ferguson_fresh_names,
                *failed_bornhuetter_ferguson_names,
            ]
            cape_cod_updates = cape_cod_service.refresh_dependents(
                project_name,
                reserving_class,
                cape_cod_roots,
                rebuild_index=False,
                blocked_precedent_names=[
                    *failed_dfm_names,
                    *failed_dataset_names,
                    *failed_result_selection_names,
                    *failed_berquist_sherman_names,
                    *failed_bornhuetter_ferguson_names,
                    *link_updates["failed"],
                ],
                finalize_method_review_status=False,
            )
        except Exception as err:
            cape_cod_updates = {
                "ok": False,
                "errors": [{"reason": str(err)}],
                "updated": [],
            }

    bootstrap_updates = None
    if include_bootstrap:
        _notify("bootstrap", 0, 0, "Refreshing Bootstrap methods")
        try:
            from app_server.services import bootstrap_service

            # A Bootstrap embeds its DFM's observed triangle and selected
            # ratios, so it must see every refreshed DFM — not only the ones
            # whose published ultimate vector changed.
            dfm_fresh_names = [
                _clean_text(value)
                for item in (dfm_updates or {}).get("updated", [])
                for value in (item.get("dataset_name"), item.get("dataset_type"))
                if _clean_text(value)
            ]
            dfm_fresh_names.extend(
                _clean_text(item.get("dataset_name"))
                for item in (dfm_updates or {}).get("status_refreshed", [])
                if _clean_text(item.get("dataset_name"))
            )
            calculated_fresh_names = [
                _clean_text(item.get("dataset_type_name"))
                for item in results
                if item.get("ok") and _clean_text(item.get("dataset_type_name"))
            ]
            result_selection_fresh_names = [
                _clean_text(item.get("dataset_name"))
                for field in ("updated", "status_refreshed")
                for item in (result_selection_updates or {}).get(field, [])
                if _clean_text(item.get("dataset_name"))
            ]
            result_selection_fresh_names.extend(
                _clean_text(name)
                for name in (result_selection_updates or {}).get("downstream_fresh_names", [])
                if _clean_text(name)
            )
            failed_result_selection_names = [
                _clean_text(item.get("dataset_name"))
                for item in (result_selection_updates or {}).get("errors", [])
                if _clean_text(item.get("dataset_name"))
            ]
            failed_result_selection_names.extend(
                _clean_text(name)
                for name in (result_selection_updates or {}).get("downstream_blocked_names", [])
                if _clean_text(name)
            )
            bornhuetter_ferguson_fresh_names = [
                _clean_text(item.get("dataset_name"))
                for field in ("updated", "status_refreshed")
                for item in (bornhuetter_ferguson_updates or {}).get(field, [])
                if _clean_text(item.get("dataset_name"))
            ]
            failed_bornhuetter_ferguson_names = [
                _clean_text(item.get("dataset_name"))
                for item in (bornhuetter_ferguson_updates or {}).get("errors", [])
                if _clean_text(item.get("dataset_name"))
            ]
            cape_cod_fresh_names = [
                _clean_text(item.get("dataset_name"))
                for field in ("updated", "status_refreshed")
                for item in (cape_cod_updates or {}).get(field, [])
                if _clean_text(item.get("dataset_name"))
            ]
            failed_cape_cod_names = [
                _clean_text(item.get("dataset_name"))
                for item in (cape_cod_updates or {}).get("errors", [])
                if _clean_text(item.get("dataset_name"))
            ]
            bootstrap_roots = [
                *changed_root_names,
                *dfm_fresh_names,
                *failed_dfm_names,
                *calculated_fresh_names,
                *failed_dataset_names,
                *link_updates["refreshed"],
                *link_updates["failed"],
                *result_selection_fresh_names,
                *failed_result_selection_names,
                *berquist_sherman_fresh_names,
                *failed_berquist_sherman_names,
                *bornhuetter_ferguson_fresh_names,
                *failed_bornhuetter_ferguson_names,
                *cape_cod_fresh_names,
                *failed_cape_cod_names,
            ]
            bootstrap_updates = bootstrap_service.refresh_dependents(
                project_name,
                reserving_class,
                bootstrap_roots,
                rebuild_index=False,
                blocked_precedent_names=[
                    *failed_dfm_names,
                    *failed_dataset_names,
                    *failed_result_selection_names,
                    *failed_berquist_sherman_names,
                    *failed_bornhuetter_ferguson_names,
                    *failed_cape_cod_names,
                    *link_updates["failed"],
                ],
                finalize_method_review_status=False,
            )
        except Exception as err:
            bootstrap_updates = {
                "ok": False,
                "errors": [{"reason": str(err)}],
                "updated": [],
            }

    # The late link pass covers links that read method outputs — a candidate
    # ultimate computed from the Result Selection's published indicated, say.
    # Methods reading what it refreshed cannot be re-walked inside this walk
    # (the cycle guard is what makes such loops converge), so they are marked
    # review-needed instead and the next explicit save picks them up.
    late_link_roots: List[str] = []
    for updates in (
        dfm_updates,
        result_selection_updates,
        berquist_sherman_updates,
        bornhuetter_ferguson_updates,
        cape_cod_updates,
        bootstrap_updates,
    ):
        for field in ("updated", "status_refreshed"):
            late_link_roots.extend(
                _clean_text(value)
                for item in (updates or {}).get(field, [])
                for value in (item.get("dataset_name"), item.get("dataset_type"))
                if _clean_text(value)
            )
        late_link_roots.extend(
            _clean_text(name)
            for name in (updates or {}).get("downstream_fresh_names", [])
            if _clean_text(name)
        )
    late_link_fresh = _refresh_link_driven_dependents(
        project_name,
        reserving_class,
        late_link_roots,
        visited_link_keys,
        link_updates,
    )
    if late_link_fresh or link_updates["failed"]:
        dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
            project_name,
            reserving_class,
            [*late_link_fresh, *link_updates["failed"]],
        )

    index_error = ""
    if finalize_method_review_status:
        _notify("finalize", 0, 0, "Finalizing review statuses")
        dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
            project_name,
            reserving_class,
            changed_root_names,
        )

    if rebuild_index:
        _notify("index", 0, 0, "Rebuilding the dataset index")
        try:
            dataset_instance_index_service.rebuild_index(project_name, reserving_class)
        except Exception as err:
            index_error = str(err)

    overall_ok = all(item.get("ok") for item in results)
    if dfm_updates is not None:
        overall_ok = overall_ok and bool(dfm_updates.get("ok"))
    if result_selection_updates is not None:
        overall_ok = overall_ok and bool(result_selection_updates.get("ok"))
    if berquist_sherman_updates is not None:
        overall_ok = overall_ok and bool(berquist_sherman_updates.get("ok"))
    if bornhuetter_ferguson_updates is not None:
        overall_ok = overall_ok and bool(bornhuetter_ferguson_updates.get("ok"))
    if cape_cod_updates is not None:
        overall_ok = overall_ok and bool(cape_cod_updates.get("ok"))
    if bootstrap_updates is not None:
        overall_ok = overall_ok and bool(bootstrap_updates.get("ok"))
    overall_ok = overall_ok and not link_updates["failed"]
    return {
        "ok": overall_ok,
        "project_name": project_name,
        "reserving_class": reserving_class,
        "changed_dataset_name": changed_dataset_name,
        "changed_dataset_type_name": changed_dataset_type_name,
        "targets": [
            rows_by_key[key]["name"]
            for key in targets
            if key in rows_by_key
        ],
        "steps": results,
        "updated": [item for item in results if item.get("ok")],
        "skipped": [item for item in results if not item.get("ok")],
        "dfm_updates": dfm_updates,
        "result_selection_updates": result_selection_updates,
        "berquist_sherman_updates": berquist_sherman_updates,
        "bornhuetter_ferguson_updates": bornhuetter_ferguson_updates,
        "cape_cod_updates": cape_cod_updates,
        "bootstrap_updates": bootstrap_updates,
        "link_updates": link_updates,
        "index_ok": not index_error,
        "index_error": index_error,
    }


_CASCADE_DOMAIN_FIELDS = (
    "dfm_updates",
    "result_selection_updates",
    "berquist_sherman_updates",
    "bornhuetter_ferguson_updates",
    "cape_cod_updates",
    "bootstrap_updates",
)


def cascade_failure_reasons(report: Mapping[str, Any]) -> List[str]:
    """Name every dependent that declined inside a walk report, and why.

    A method wave reports a failed downstream walk as a single line per
    output ("Downstream refresh failed after BF publication.") and tucks the
    walk itself under ``cascade``, which the flattened save message drops.
    Flattening the walk's own failures into ``"<dataset>: <reason>"`` lines
    lets that message say which dependent declined instead of only that one
    did. Nested cascades are unwound the same way so the innermost reason
    surfaces.
    """
    reasons: List[str] = []
    if not isinstance(report, Mapping):
        return reasons

    def add(name: Any, reason: Any) -> None:
        text_name = _clean_text(name)
        text_reason = _clean_text(reason)
        text = f"{text_name}: {text_reason}" if text_name and text_reason else (text_name or text_reason)
        if text and text not in reasons:
            reasons.append(text)

    for item in report.get("skipped") or []:
        if not isinstance(item, Mapping):
            continue
        details = [_clean_text(error) for error in item.get("errors") or [] if _clean_text(error)]
        add(item.get("dataset_type_name") or item.get("dataset_name"), "; ".join(details) or item.get("reason"))
    link_domain = report.get("link_updates")
    if isinstance(link_domain, Mapping):
        for error in link_domain.get("errors") or []:
            if not isinstance(error, Mapping):
                continue
            details = [_clean_text(text) for text in error.get("errors") or [] if _clean_text(text)]
            add(error.get("dataset_name"), "; ".join(details) or error.get("reason"))
    for field in _CASCADE_DOMAIN_FIELDS:
        domain = report.get(field)
        if not isinstance(domain, Mapping) or domain.get("ok", True):
            continue
        for error in domain.get("errors") or []:
            if not isinstance(error, Mapping):
                continue
            nested = cascade_failure_reasons(error.get("cascade")) if isinstance(error.get("cascade"), Mapping) else []
            if nested:
                for text in nested:
                    if text not in reasons:
                        reasons.append(text)
                continue
            add(error.get("dataset_name") or error.get("method_name"), error.get("reason"))
    return reasons


def recalculate_dependents(
    project_name: str,
    reserving_class: str,
    changed_dataset_name: str,
    changed_dataset_type_name: str = "",
    *,
    include_dfm: bool = True,
    include_result_selection: bool = True,
    include_berquist_sherman: bool = True,
    include_bornhuetter_ferguson: bool = True,
    include_cape_cod: bool = True,
    include_bootstrap: bool = True,
    finalize_method_review_status: bool = True,
    rebuild_index: bool = True,
    additional_roots: Sequence[Tuple[str, str]] | None = None,
    progress_callback: Callable[[str, int, int, str], None] | None = None,
) -> Dict[str, Any]:
    with dataset_sidecar_status_service.reserving_class_io_lock(project_name, reserving_class):
        return _recalculate_dependents_impl(
            project_name,
            reserving_class,
            changed_dataset_name,
            changed_dataset_type_name,
            include_dfm=include_dfm,
            include_result_selection=include_result_selection,
            include_berquist_sherman=include_berquist_sherman,
            include_bornhuetter_ferguson=include_bornhuetter_ferguson,
            include_cape_cod=include_cape_cod,
            include_bootstrap=include_bootstrap,
            finalize_method_review_status=finalize_method_review_status,
            rebuild_index=rebuild_index,
            additional_roots=additional_roots,
            progress_callback=progress_callback,
        )


def preview_dependents(
    project_name: str,
    reserving_class: str,
    changed_dataset_name: str,
    changed_dataset_type_name: str = "",
    values: List[List[Any]] | None = None,
    mask: List[List[bool]] | None = None,
    origin_labels: List[str] | None = None,
    development_labels: List[str] | None = None,
) -> Dict[str, Any]:
    source_arr = _array_from_preview_values(values, mask)
    if source_arr.ndim != 2 or source_arr.size == 0:
        return {"ok": False, "skipped": True, "reason": "empty_preview_values", "steps": []}

    changed = [changed_dataset_name, changed_dataset_type_name]
    targets = _downstream_keys(project_name, changed)
    rows_by_key = _calculated_rows_by_key(project_name)
    overrides: Dict[str, np.ndarray] = {
        _canon_dataset_name(name): source_arr
        for name in changed
        if _canon_dataset_name(name)
    }
    steps: List[Dict[str, Any]] = []

    all_rows = _dataset_type_rows(project_name)
    known_names = [item["name"] for item in all_rows]
    for key in targets:
        row = rows_by_key.get(key)
        if not row:
            continue

        expr, refs = _replace_formula_refs(row["formula"], known_names)
        ordered_components = [refs[var] for var in sorted(refs.keys(), key=lambda item: int(item[2:]))]
        settings = _existing_target_settings(project_name, reserving_class, row["name"])
        component_values, _precedents, errors = _load_components(
            project_name,
            reserving_class,
            ordered_components,
            settings,
            component_overrides=overrides,
        )
        if errors:
            steps.append({
                "ok": False,
                "status": "skipped",
                "dataset_type_name": row["name"],
                "reason": "dependency_error",
                "errors": errors,
            })
            continue

        eval_values: Dict[str, np.ndarray] = {}
        for var, ref_name in refs.items():
            try:
                idx = ordered_components.index(ref_name)
            except ValueError:
                continue
            if f"_d{idx}" in component_values:
                eval_values[var] = component_values[f"_d{idx}"]
        try:
            parsed = ast.parse(expr, mode="eval")
            with np.errstate(divide="ignore", invalid="ignore"):
                result = _eval_ast(parsed, eval_values)
        except Exception as exc:
            steps.append({
                "ok": False,
                "status": "skipped",
                "dataset_type_name": row["name"],
                "reason": "formula_error",
                "errors": [str(exc)],
            })
            continue

        arr = np.asarray(result, dtype="float64")
        if arr.ndim == 0:
            first = next(iter(eval_values.values()), source_arr)
            arr = np.full(first.shape, float(arr), dtype="float64")
        if arr.ndim == 1:
            arr = arr.reshape((-1, 1))
        if arr.ndim != 2:
            steps.append({
                "ok": False,
                "status": "skipped",
                "dataset_type_name": row["name"],
                "reason": "unsupported_result_shape",
            })
            continue

        overrides[_canon_dataset_name(row["name"])] = arr
        data_format = row.get("data_format") or "Triangle"
        steps.append({
            "ok": True,
            "status": "preview",
            "dataset_type_name": row["name"],
            "dataset_name": row["name"],
            "source_kind": "calculated_preview",
            "data_format": data_format,
            "values": _latest_diagonal_or_vector_values(arr, data_format),
            "matrix_values": _jsonable_matrix(arr),
            "mask": _matrix_mask(arr),
            "origin_labels": [str(item) for item in (origin_labels or [])],
            "development_labels": [str(item) for item in (development_labels or [])],
        })

    return {
        "ok": True,
        "project_name": project_name,
        "reserving_class": reserving_class,
        "changed_dataset_name": changed_dataset_name,
        "changed_dataset_type_name": changed_dataset_type_name,
        "targets": [
            rows_by_key[key]["name"]
            for key in targets
            if key in rows_by_key
        ],
        "steps": steps,
        "updated": [item for item in steps if item.get("ok")],
        "skipped": [item for item in steps if not item.get("ok")],
    }


def _rows_by_key_from_normalized_rows(rows: List[List[Any]]) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    for row in rows or []:
        if not isinstance(row, list):
            continue
        name = _clean_text(row[0] if len(row) > 0 else "")
        key = _canon_dataset_name(name)
        if not key:
            continue
        out[key] = {
            "name": name,
            "calculated": _bool_value(row[3] if len(row) > 3 else False),
            "formula": _clean_text(row[4] if len(row) > 4 else ""),
            "generated": _bool_value(row[6] if len(row) > 6 else False),
        }
    return out


def changed_formula_dataset_type_names(previous_rows: List[List[Any]], next_rows: List[List[Any]]) -> List[str]:
    previous = _rows_by_key_from_normalized_rows(previous_rows)
    current = _rows_by_key_from_normalized_rows(next_rows)
    names_by_key = _dataset_type_name_by_key_from_rows(next_rows)
    changed: List[str] = []
    for key, row in current.items():
        prev = previous.get(key)
        if (
            prev is None
            or bool(prev.get("calculated")) != bool(row.get("calculated"))
            or bool(prev.get("generated")) != bool(row.get("generated"))
            or _clean_text(prev.get("formula")) != _clean_text(row.get("formula"))
        ):
            name = names_by_key.get(key) or row.get("name")
            if _clean_text(name):
                changed.append(str(name))
    return changed


def _dataset_type_name_by_key_from_rows(rows: List[List[Any]]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for row in rows or []:
        if not isinstance(row, list):
            continue
        name = _clean_text(row[0] if len(row) > 0 else "")
        key = _canon_dataset_name(name)
        if key:
            out[key] = name
    return out


def _write_sidecar_json(path: str, payload: Dict[str, Any]) -> None:
    payload = dict(payload)
    payload.pop("instance_name", None)
    payload.pop("dataset_type_name", None)
    dataset_sidecar_status_service.write_sidecar(path, payload)


def _sidecar_method_type(project_name: str, reserving_class: str, dataset_name: str) -> str:
    payload = dataset_sidecar_status_service.read_sidecar(
        dataset_sidecar_status_service.sidecar_path(project_name, reserving_class, dataset_name)
    )
    return dataset_sidecar_status_service.normalize_method_type(
        payload.get("method_type"), payload.get("source_kind")
    )


def find_dataset_type_removal_blockers(
    project_name: str,
    planned: "dataset_types_plan_service.DatasetTypesChangePlan",
) -> List[Dict[str, Any]]:
    """Instances of disappearing dataset types that other objects still read.

    Deleting a dataset type deletes the definition its instances are described
    by, so it may only happen once nothing downstream reads those instances --
    the same rule that already governs deleting one dataset. A dependent that
    is itself an instance of a type being removed in the same change is not a
    blocker: it is leaving too.

    Only the instances the plan already found are opened: the plan knows which
    classes hold a departing type, so no other class is visited.

    Returns one entry per blocked type, each naming the instances still read
    and, for each, the datasets or methods reading them.
    """

    removed_keys = {
        key
        for key in (_canon_dataset_name(name) for name in planned.removed_types)
        if key
    }
    if not removed_keys:
        return []

    blocked_by_type: Dict[str, Dict[str, Any]] = {}
    for affected in planned.classes:
        departing = [
            instance
            for instance in affected.instances
            if _canon_dataset_name(instance.dataset_type) in removed_keys
        ]
        leaving = {_canon_dataset_name(instance.name) for instance in departing}
        for instance in departing:
            payload = dataset_sidecar_status_service.read_sidecar(
                dataset_sidecar_status_service.sidecar_path(
                    project_name, affected.reserving_class, instance.name
                )
            )
            dependents = [
                {
                    "dataset_name": dependent_name,
                    "method_type": _sidecar_method_type(
                        project_name, affected.reserving_class, dependent_name
                    ),
                }
                for dependent_name in dataset_sidecar_status_service.entry_names(
                    payload.get("dependents")
                )
                if _canon_dataset_name(dependent_name)
                and _canon_dataset_name(dependent_name) not in leaving
            ]
            if not dependents:
                continue
            entry = blocked_by_type.setdefault(
                _canon_dataset_name(instance.dataset_type),
                {"dataset_type": instance.dataset_type, "instances": []},
            )
            entry["instances"].append({
                "reserving_class": affected.reserving_class,
                "dataset_name": instance.name,
                "dependents": dependents,
            })

    return [blocked_by_type[key] for key in sorted(blocked_by_type)]


def _rename_dataset_instance(
    project_name: str,
    reserving_class: str,
    sidecar_path: str,
    payload: Dict[str, Any],
    new_name: str,
) -> str:
    """Move one plain dataset instance to a new name; return its new sidecar path.

    The CSV keeps everything after the name in its file name, the old sidecar
    file goes, and each precedent's ``dependents`` entry follows the name. The
    runtime cache provenance is dropped rather than moved: its record is bound
    to the dataset name, so the next open re-validates the cache exactly as it
    does for a cache with no record. The caller holds the class I/O lock and
    writes the returned path itself.
    """

    old_name = _clean_text(payload.get("dataset_name")) or _sidecar_instance_name(sidecar_path)
    old_stem = sanitize_dataset_file_name(old_name)
    new_stem = sanitize_dataset_file_name(new_name)

    csv_file = _clean_text(payload.get("csv_file"))
    if csv_file:
        cache_dir = config.get_project_dataset_cache_dir(project_name, reserving_class)
        old_csv = os.path.join(cache_dir, csv_file)
        if os.path.isfile(old_csv) and csv_file.lower().startswith(old_stem.lower()):
            new_csv_file = new_stem + csv_file[len(old_stem):]
            os.replace(old_csv, os.path.join(cache_dir, new_csv_file))
            runtime_cache_provenance_service.remove(old_csv)
            payload["csv_file"] = new_csv_file

    new_path = dataset_sidecar_status_service.sidecar_path(project_name, reserving_class, new_name)
    if os.path.normcase(os.path.abspath(new_path)) != os.path.normcase(os.path.abspath(sidecar_path)):
        try:
            os.remove(sidecar_path)
        except FileNotFoundError:
            pass
    payload["dataset_name"] = new_name

    precedents = dataset_sidecar_status_service.entry_names(payload.get("precedents"))
    for precedent_name in precedents:
        precedent_path = dataset_sidecar_status_service.sidecar_path(
            project_name, reserving_class, precedent_name
        )
        with dataset_sidecar_status_service.sidecar_write_lock(precedent_path):
            precedent = _read_sidecar(precedent_path)
            if not precedent:
                continue
            if dataset_sidecar_status_service._remove_dependent(precedent, old_name):
                dataset_sidecar_status_service._add_dependent(precedent, new_name)
                _write_sidecar_json(precedent_path, precedent)
    return new_path


def _sidecar_instance_name(path: str) -> str:
    """The dataset instance a sidecar file describes, decoded from its name."""

    stem = os.path.splitext(os.path.basename(path))[0]
    return dataset_instance_index_service._normalize_cached_dataset_name(stem)


def apply_planned_dataset_types_change(
    project_name: str,
    planned: "dataset_types_plan_service.DatasetTypesChangePlan",
    *,
    on_progress: Callable[[str, str, int, int], None] | None = None,
) -> Dict[str, Any]:
    """Rewrite the instances a plan named, then rebuild what changed.

    Runs after the table is written, class by class under each class's I/O
    lock. Every instance the plan listed has its sidecar re-derived from the
    new table: a renamed type's instances take the new type name, and the
    ones the plan marked for renaming move with it. The calculated datasets
    whose formula or kind changed are then recalculated and one Engine
    propagation job per class walks their dependents.

    ``on_progress`` receives ``(stage, label, completed, total)``. The unit is
    one dataset instance, then one reserving class for each class the change
    makes the walk revisit, so the count a caller shows is work done rather
    than stages passed.
    """

    def report(stage: str, label: str, completed: int, total: int) -> None:
        if on_progress is None:
            return
        try:
            on_progress(stage, label, completed, total)
        except Exception:
            # Progress is telemetry. A publisher that cannot write its status
            # must not abort a rebuild that is already under way.
            pass

    changed_keys = {
        _canon_dataset_name(name)
        for name in planned.changed_types
        if _canon_dataset_name(name)
    }
    rows_by_key = _calculated_rows_by_key(project_name)
    sidecars_updated = 0
    datasets_renamed = 0
    recalc_seeds: Set[Tuple[str, str]] = set()
    errors: List[str] = []

    datasets_total = sum(len(affected.instances) for affected in planned.classes)
    total_units = datasets_total
    completed_units = 0
    report("graphs", "Rebuilding dataset dependency graphs", 0, total_units)

    for affected in planned.classes:
        reserving_class = affected.reserving_class
        with dataset_sidecar_status_service.reserving_class_io_lock(project_name, reserving_class):
            for instance in affected.instances:
                completed_units += 1
                path = dataset_sidecar_status_service.sidecar_path(
                    project_name, reserving_class, instance.name
                )
                try:
                    with dataset_sidecar_status_service.sidecar_write_lock(path):
                        payload = _read_sidecar(path)
                        if not payload:
                            raise RuntimeError("Dataset sidecar is missing.")
                        before = json.dumps(payload, sort_keys=True, ensure_ascii=False)
                        payload.pop("instance_name", None)
                        payload.pop("dataset_type_name", None)
                        if instance.new_dataset_type != instance.dataset_type:
                            payload["dataset_type"] = instance.new_dataset_type
                        moved = False
                        if instance.rename_to:
                            path = _rename_dataset_instance(
                                project_name, reserving_class, path, payload, instance.rename_to
                            )
                            moved = True
                            datasets_renamed += 1
                        apply_sidecar_graph_fields(payload, project_name, instance.new_dataset_type)
                        after = json.dumps(payload, sort_keys=True, ensure_ascii=False)
                        if moved or before != after:
                            _write_sidecar_json(path, payload)
                            sidecars_updated += 1
                except Exception as exc:
                    errors.append(f"{reserving_class} / {instance.name}: {exc}")
                    continue
                report(
                    "graphs",
                    f"Rebuilding dependency graphs: {reserving_class}",
                    completed_units,
                    total_units,
                )
                type_key = _canon_dataset_name(instance.new_dataset_type)
                if type_key in changed_keys and type_key in rows_by_key:
                    recalc_seeds.add((reserving_class, rows_by_key[type_key]["name"]))

    # Recalculate each changed-formula dataset itself (the saved objects),
    # then enqueue one Engine propagation job per reserving class covering all
    # of that class's changed roots; the job walks the dependents and rebuilds
    # the index on the server host.
    from app_server.services import dependent_propagation_service

    seeds_by_reserving_class: Dict[str, List[str]] = {}
    for reserving_class, dataset_type in sorted(recalc_seeds):
        seeds_by_reserving_class.setdefault(reserving_class, []).append(dataset_type)

    # The classes the change makes the walk revisit are only known now, so the
    # denominator grows once here rather than being guessed up front.
    total_units = datasets_total + len(seeds_by_reserving_class)
    chains: List[Dict[str, Any]] = []
    for reserving_class, dataset_types in seeds_by_reserving_class.items():
        report(
            "recalculate",
            f"Recalculating {reserving_class}",
            completed_units,
            total_units,
        )
        steps: List[Dict[str, Any]] = []
        for dataset_type in dataset_types:
            try:
                first = recalculate_dataset(project_name, reserving_class, dataset_type)
            except Exception as exc:
                first = {
                    "ok": False,
                    "dataset_type_name": dataset_type,
                    "reason": str(exc),
                }
            steps.append(
                {**first, "status": "updated" if first.get("ok") else "skipped"}
            )
        try:
            dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
                project_name,
                reserving_class,
                dataset_types,
            )
        except Exception as exc:
            errors.append(f"{reserving_class} review marking: {exc}")
        propagation = dependent_propagation_service.enqueue_save_propagation(
            project_name,
            reserving_class,
            [
                dependent_propagation_service.changed_root(dataset_type, dataset_type)
                for dataset_type in dataset_types
            ],
        )
        chains.append({
            "ok": all(step.get("ok") for step in steps) and bool(propagation.get("ok")),
            "reserving_class": reserving_class,
            "changed_dataset_type_name": ", ".join(dataset_types),
            "steps": steps,
            "updated": [step for step in steps if step.get("ok")],
            "skipped": [step for step in steps if not step.get("ok")],
            "propagation": propagation,
        })
        completed_units += 1

    # Every class the plan named had sidecars rewritten -- a type name or a
    # rename is index content -- so each one's index is rebuilt.
    for affected in planned.classes:
        report(
            "index",
            f"Rebuilding the dataset index: {affected.reserving_class}",
            completed_units,
            total_units,
        )
        try:
            dataset_instance_index_service.rebuild_index(project_name, affected.reserving_class)
        except Exception as exc:
            errors.append(f"{affected.reserving_class} index rebuild: {exc}")

    return {
        "ok": not errors and all(chain.get("ok") for chain in chains),
        "project_name": project_name,
        "changed_dataset_types": list(planned.changed_types),
        "sidecars_updated": sidecars_updated,
        "datasets_renamed": datasets_renamed,
        "datasets_total": datasets_total,
        "classes_total": planned.class_count,
        "chains": chains,
        "errors": errors,
    }
