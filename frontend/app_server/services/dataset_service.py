"""Dataset / triangle data operations."""
from __future__ import annotations

import getpass
import hashlib
import json
import os
import re
import shutil
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from typing import Any, Dict, Iterable, Iterator, List, Tuple, cast

import numpy as np
import pandas as pd
from fastapi import HTTPException

from arcrho_api.dataset_display_contract import DEFAULT_SHOW_SUBTOTAL, normalize_show_subtotal
from arcrho_api.dataset_link_contract import link_precedent_names
from arcrho_api.io import persisted_json_text
from arcrho_api.sidecar_audit_contract import (
    AUDIT_ACTION_INSERT,
    AUDIT_ACTION_UPDATE,
    append_audit_entry,
    normalize_audit_log,
)
from arcrho_api.sidecar_core_contract import (
    SIDECAR_LINKED_DEVELOPMENT_FIELD,
    display_lengths,
    finalize_sidecar,
    linked_length_fields,
    linked_lengths,
    stored_length_fields,
    stored_lengths,
)
from arcrho_api.timestamps import utc_now_text
from arcrho_api.triangle_rollup import rollup_triangle, scatter_triangle
from app_server import config
from app_server.helpers import (
    _canon_dataset_name,
    atomic_write_csv,
    build_dataset_cache_file_name,
    sanitize_dataset_file_name,
)
from app_server.services import (
    dataset_instance_index_service,
    dataset_number_format_service,
    dataset_sidecar_status_service,
    dependent_propagation_service,
    user_identity_service,
)


_ORIGIN_YEAR_RE = re.compile(r"^(\d{4})$")
_ORIGIN_HALF_RE = re.compile(r"^(\d{4})\s*H([12])$", re.IGNORECASE)
_ORIGIN_QUARTER_RE = re.compile(r"^(\d{4})\s*Q([1-4])$", re.IGNORECASE)
_ORIGIN_MONTH_RE = re.compile(r"^(\d{4})(0[1-9]|1[0-2])$")
_ORIGIN_MONTH_NAME_RE = re.compile(
    r"^(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
    r"Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{4})$",
    re.IGNORECASE,
)
_ORIGIN_MONTH_NUMBERS = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}
_ORIGIN_KIND_BY_LENGTH = {12: "year", 6: "half", 3: "quarter", 1: "month"}
# The grid sends an empty cell as a JSON null. ``np.where`` declares both of its
# branches as array-likes, so the null travels under a plainly typed name.
_BLANK_CELL: Any = None
_METHOD_CALCULATED_TYPES = frozenset((
    dataset_sidecar_status_service.METHOD_TYPE_DFM,
    dataset_sidecar_status_service.METHOD_TYPE_RESULT_SELECTION,
    dataset_sidecar_status_service.METHOD_TYPE_BORN_HUETTER_FERGUSON,
    dataset_sidecar_status_service.METHOD_TYPE_CAPE_COD,
    dataset_sidecar_status_service.METHOD_TYPE_BOOTSTRAP,
    dataset_sidecar_status_service.METHOD_TYPE_BERQUIST_SHERMAN_SR,
    dataset_sidecar_status_service.METHOD_TYPE_BERQUIST_SHERMAN_CRA,
))
_CACHED_LOAD_HYDRATION_MAX_WORKERS = 6
_CACHED_LOAD_HYDRATION_EXECUTOR = ThreadPoolExecutor(
    max_workers=_CACHED_LOAD_HYDRATION_MAX_WORKERS,
    thread_name_prefix="arcrho-cache-hydrate",
)
def _dataset_sidecar_write_lock(path: str) -> threading.RLock:
    return dataset_sidecar_status_service.sidecar_write_lock(path)


def _parse_origin_label(value: Any) -> Tuple[str, int] | None:
    label = str(value if value is not None else "").strip()
    match = _ORIGIN_YEAR_RE.fullmatch(label)
    if match:
        year = int(match.group(1))
        return ("year", year) if year > 0 else None
    match = _ORIGIN_HALF_RE.fullmatch(label)
    if match:
        year = int(match.group(1))
        return ("half", year * 2 + int(match.group(2)) - 1) if year > 0 else None
    match = _ORIGIN_QUARTER_RE.fullmatch(label)
    if match:
        year = int(match.group(1))
        return ("quarter", year * 4 + int(match.group(2)) - 1) if year > 0 else None
    match = _ORIGIN_MONTH_RE.fullmatch(label)
    if match:
        year = int(match.group(1))
        return ("month", year * 12 + int(match.group(2)) - 1) if year > 0 else None
    match = _ORIGIN_MONTH_NAME_RE.fullmatch(label)
    if match:
        year = int(match.group(2))
        month = _ORIGIN_MONTH_NUMBERS.get(match.group(1).lower())
        return ("month", year * 12 + month - 1) if year > 0 and month else None
    return None


def _validate_origin_labels(
    value: Any,
    expected_count: int,
    origin_length: int | None = None,
) -> Tuple[List[str], str]:
    if not isinstance(value, list) or not value:
        return [], "no origin labels were returned"
    labels = [str(item if item is not None else "").strip() for item in value]
    if len(labels) != expected_count:
        return [], f"origin label count {len(labels)} does not match dataset row count {expected_count}"
    parsed = [_parse_origin_label(label) for label in labels]
    if any(item is None for item in parsed):
        return [], "one or more origin labels are blank or use an unsupported date format"
    kinds = {item[0] for item in parsed if item is not None}
    if len(kinds) != 1:
        return [], "origin labels mix incompatible date formats"
    expected_kind = (
        _ORIGIN_KIND_BY_LENGTH.get(origin_length) if origin_length is not None else None
    )
    if expected_kind and expected_kind not in kinds:
        return [], f"origin labels do not match the requested {origin_length}-month period length"
    sequence = [item[1] for item in parsed if item is not None]
    if any(current != previous + 1 for previous, current in zip(sequence, sequence[1:])):
        return [], "origin labels are not consecutive"
    return labels, ""


def _origin_labels_error(ds_id: str, project_name: str, reason: str) -> HTTPException:
    project = str(project_name or "").strip() or "(unknown)"
    detail = (
        f"Cannot load dataset '{ds_id}': valid origin labels could not be resolved for project '{project}'"
        f" ({reason}). Set a valid Origin Start Date in Project Settings, then refresh the dataset."
    )
    return HTTPException(422, detail)


def _load_project_header_labels(
    ds_id: str,
    path: str,
    project_name: str,
    period_length: int,
    *,
    period_type: int = 0,
    transposed: bool = False,
    calendar: bool = False,
) -> List[str]:
    project = str(project_name or "").strip()
    if not project:
        raise _origin_labels_error(ds_id, project, "project name is missing")
    try:
        length = int(period_length)
    except (TypeError, ValueError):
        raise _origin_labels_error(ds_id, project, "origin period length is invalid")
    if period_type == 0 and length not in _ORIGIN_KIND_BY_LENGTH:
        raise _origin_labels_error(ds_id, project, f"origin period length '{period_length}' is unsupported")

    try:
        project_data_dir = os.path.normcase(os.path.realpath(config.get_project_data_dir(project)))
    except ValueError as err:
        raise HTTPException(404, str(err))
    dataset_path = os.path.normcase(os.path.realpath(path))
    try:
        belongs_to_project = os.path.commonpath((project_data_dir, dataset_path)) == project_data_dir
    except ValueError:
        belongs_to_project = False
    if not belongs_to_project:
        raise HTTPException(
            422,
            f"Cannot load dataset '{ds_id}' for project '{project}': the registered dataset belongs to a different project.",
        )

    try:
        from app_server.services import arcrho_runtime_service

        if period_type == 0 and not transposed and not calendar:
            header_result = arcrho_runtime_service.get_project_headers(
                project,
                length,
                timeout_sec=config.ENGINE_REQUEST_TIMEOUT_SEC,
            )
        else:
            header_result = arcrho_runtime_service.get_project_headers(
                project,
                length,
                timeout_sec=config.ENGINE_REQUEST_TIMEOUT_SEC,
                period_type=period_type,
                transposed=transposed,
                calendar=calendar,
            )
    except HTTPException as err:
        detail = str(err.detail or "ArcRho project headers could not be loaded")
        raise HTTPException(err.status_code, f"Cannot load dataset '{ds_id}': {detail}")
    except OSError as err:
        raise HTTPException(500, f"Cannot load dataset '{ds_id}': failed to read ArcRho project headers: {str(err)}")

    if not header_result.get("ok"):
        status = str(header_result.get("status") or "unavailable").strip()
        message = str(header_result.get("message") or "").strip()
        if status.casefold() == "timeout":
            raise HTTPException(
                504,
                f"Cannot load dataset '{ds_id}' for project '{project}': "
                f"{message or 'timed out while loading ArcRho project headers. Try again.'}",
            )
        raise HTTPException(
            503,
            f"Cannot load dataset '{ds_id}' for project '{project}': ArcRho project headers are {status}.",
        )
    labels = header_result.get("labels")
    return [str(item if item is not None else "").strip() for item in labels] if isinstance(labels, list) else []


def _resolve_origin_labels(
    ds_id: str,
    path: str,
    project_name: str,
    origin_length: int,
    expected_count: int,
) -> List[str]:
    labels, header_reason = _validate_origin_labels(
        _load_project_header_labels(ds_id, path, project_name, origin_length),
        expected_count,
        origin_length,
    )
    if labels:
        return labels
    raise _origin_labels_error(ds_id, project_name, header_reason)


def _resolve_development_labels(
    ds_id: str,
    path: str,
    project_name: str,
    development_length: int,
    expected_count: int,
    *,
    calendar: bool = False,
    fallback_labels: List[str] | None = None,
) -> List[str]:
    """The column ages of a triangle view, as the project's own settings state them.

    Project Settings own the development axis the way they own the origin one:
    the ages count back from the Development End Date, so a project valued in
    August reads 8m, 20m, 32m and never the 12m, 24m, 36m of a December one. A
    sidecar's own list is written once, by whatever created the dataset, and
    nothing updates it afterwards, so it is a fallback for a project whose
    headers cannot be read and never the answer while they can.

    The project's headers run to the end of its own grid, which can be later
    than the valuation date a triangle stops at, so they are cut to the columns
    the dataset has.
    """

    fallback = [str(item) for item in fallback_labels] if fallback_labels else []
    usable_fallback = fallback if len(fallback) == expected_count and all(fallback) else []
    try:
        labels = _load_project_header_labels(
            ds_id,
            path,
            project_name,
            development_length,
            period_type=1,
            transposed=True,
            calendar=calendar,
        )
    except HTTPException:
        if usable_fallback:
            return usable_fallback
        raise
    if len(labels) >= expected_count and all(labels[:expected_count]):
        return labels[:expected_count]
    if usable_fallback:
        return usable_fallback
    reason = (
        f"development label count {len(labels)} does not match dataset column count {expected_count}"
        if labels
        else "no development labels were returned"
    )
    project = str(project_name or "").strip() or "(unknown)"
    raise HTTPException(
        422,
        f"Cannot load dataset '{ds_id}': valid development labels could not be resolved "
        f"for project '{project}' ({reason}). Refresh the dataset after verifying Project Settings.",
    )


def infer_shape(path: str) -> Tuple[int, int]:
    df = pd.read_csv(path, header=None)
    return int(df.shape[0]), int(df.shape[1])


def load_triangle_values(path: str) -> pd.DataFrame:
    return pd.read_csv(path, header=None, dtype="float64")


def triangle_mask(n_origin: int, n_dev: int) -> np.ndarray:
    r = np.arange(n_origin)[:, None]
    c = np.arange(n_dev)[None, :]
    return (r + c < n_dev)


def diagonal_indices(n_origin: int, n_dev: int, k: int = 0) -> List[Tuple[int, int]]:
    mask = triangle_mask(n_origin, n_dev)
    out = []
    for r in range(n_origin):
        c = n_dev - 1 - r - k
        if 0 <= c < n_dev and mask[r, c]:
            out.append((r, c))
    return out


def list_datasets() -> List[Dict[str, Any]]:
    out = []
    for ds_id, path in config.DATASETS.items():
        if not os.path.exists(path):
            continue
        n_origin, n_dev = infer_shape(path)
        st = os.stat(path)
        out.append({
            "id": ds_id,
            "path": path,
            "shape": {"n_origin": n_origin, "n_dev": n_dev},
            "mtime": st.st_mtime,
        })
    return out


def list_cached_dataset_names(project_name: str, reserving_class: str, refresh: bool = False) -> Dict[str, Any]:
    project = str(project_name if project_name is not None else "").strip()
    rc = str(reserving_class if reserving_class is not None else "").strip()
    if not project or not rc:
        raise HTTPException(400, "project_name and reserving_class are required.")
    return dataset_instance_index_service.get_index(project, rc, refresh=refresh)


def get_cached_dataset_index_signature(project_name: str, reserving_class: str) -> Dict[str, Any]:
    project = str(project_name if project_name is not None else "").strip()
    rc = str(reserving_class if reserving_class is not None else "").strip()
    if not project or not rc:
        raise HTTPException(400, "project_name and reserving_class are required.")
    return dataset_instance_index_service.get_index_signature(project, rc)


def delete_cached_datasets(project_name: str, reserving_class: str, dataset_names: List[str]) -> Dict[str, Any]:
    project = str(project_name if project_name is not None else "").strip()
    rc = str(reserving_class if reserving_class is not None else "").strip()
    if not project or not rc:
        raise HTTPException(400, "project_name and reserving_class are required.")
    return dataset_instance_index_service.delete_cached_datasets(project, rc, dataset_names)


def _normalize_number_format(value: Any) -> str:
    text = str(value or "").replace("\r", " ").replace("\n", " ").replace("\t", " ").strip()
    return (text or "0,000")[:64]


def _normalize_decimal_places(value: Any) -> int:
    try:
        n = int(value)
    except (TypeError, ValueError):
        # The same default the shared number-format settings use, so a dataset
        # with nothing recorded reads as whole numbers rather than picking up a
        # decimal place nobody asked for.
        return dataset_number_format_service.DEFAULT_DECIMAL_PLACES
    return max(0, min(6, n))


def _normalize_origin_labels(value: Any) -> List[str]:
    if not isinstance(value, list):
        return []
    return [str(item) for item in value]


def _current_user_name() -> str:
    """Display name stamped onto dataset metadata and audit entries.

    Sidecars carry the configured full name, not the Windows login, so the
    dataset table reads the same identity ResQ-imported rows already show.
    """
    display_name = user_identity_service.get_current_display_name()
    if display_name:
        return display_name
    for value in (os.environ.get("USERNAME"), os.environ.get("USER")):
        text = str(value or "").strip()
        if text:
            return text
    try:
        return str(getpass.getuser() or "").strip() or "unknown"
    except Exception:
        return "unknown"


def _normalize_dataset_audit_log(value: Any) -> List[Dict[str, str]]:
    """The one audit policy (``arcrho_api.sidecar_audit_contract``) on a stored log.

    Every action is kept -- including the Engine's ``Auto Refresh`` entries,
    which an earlier version of this normalizer silently discarded on every
    dataset save -- with consecutive automatic entries collapsed and the cap
    applied.
    """

    return normalize_audit_log(value)


def _append_dataset_audit_entry(payload: Dict[str, Any], action: str, *, event_date: str | None = None, user_name: str | None = None) -> None:
    action_value = (
        AUDIT_ACTION_INSERT
        if str(action or "").strip().lower() == "insert"
        else AUDIT_ACTION_UPDATE
    )
    payload["audit_log"] = append_audit_entry(
        payload.get("audit_log"),
        event_date=event_date or _now_utc_iso(),
        action=action_value,
        user=str(user_name or "").strip() or _current_user_name(),
    )


def _normalize_dataset_external_links(
    value: Any,
    *,
    strict: bool = False,
) -> List[Dict[str, Any]]:
    if value is None:
        return []
    if not isinstance(value, list):
        if strict:
            raise HTTPException(400, "external_links must be a list.")
        return []

    normalized: List[Dict[str, Any]] = []
    seen_links: set[Tuple[str, Tuple[Tuple[int, int, str], ...]]] = set()
    owned_targets: set[Tuple[int, int]] = set()

    def invalid(detail: str) -> bool:
        if strict:
            raise HTTPException(400, detail)
        return False

    def column_index(column: str) -> int:
        result = 0
        for character in column.upper():
            result = result * 26 + (ord(character) - 64)
        return result - 1

    def column_name(index: int) -> str:
        result = ""
        value = index + 1
        while value > 0:
            value, remainder = divmod(value - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def reference_bounds(reference: str) -> Tuple[int, int, int, int] | None:
        match = re.fullmatch(
            r"\s*=?\s*(?:'((?:[^']|'')*)'|([^!]+))!\s*"
            r"\$?([A-Z]+)\$?([1-9][0-9]*)"
            r"(?:\s*:\s*\$?([A-Z]+)\$?([1-9][0-9]*))?\s*",
            reference,
            re.IGNORECASE,
        )
        if not match:
            return None
        source = str(match.group(1) if match.group(1) is not None else match.group(2) or "")
        source = source.replace("''", "'").strip()
        open_bracket = source.find("[")
        close_bracket = source.find("]", open_bracket + 1)
        if (
            open_bracket < 0
            or close_bracket <= open_bracket + 1
            or close_bracket >= len(source) - 1
        ):
            return None
        start_column = column_index(match.group(3))
        start_row = int(match.group(4)) - 1
        end_column = column_index(match.group(5) or match.group(3))
        end_row = int(match.group(6) or match.group(4)) - 1
        return (
            min(start_row, end_row),
            max(start_row, end_row),
            min(start_column, end_column),
            max(start_column, end_column),
        )

    def normalize_source_cell(source_cell: Any) -> Tuple[str, int, int] | None:
        if not isinstance(source_cell, str):
            return None
        normalized = source_cell.strip().replace("$", "").upper()
        match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", normalized)
        if not match:
            return None
        return normalized, int(match.group(2)) - 1, column_index(match.group(1))

    for raw_link in value:
        if hasattr(raw_link, "model_dump"):
            raw_link = raw_link.model_dump()
        if not isinstance(raw_link, dict):
            invalid("Each external link must be an object.")
            continue

        raw_reference = raw_link.get("reference")
        if not isinstance(raw_reference, str):
            invalid("Each external link reference must be a string.")
            continue
        reference = raw_reference.strip()
        if not reference:
            invalid("Each external link reference must not be blank.")
            continue

        raw_targets = raw_link.get("target_cells")
        if not isinstance(raw_targets, list) or not raw_targets:
            invalid("Each external link must include at least one target cell.")
            continue

        targets: List[Dict[str, Any]] = []
        seen_targets: Dict[Tuple[int, int], str | None] = {}
        link_is_invalid = False
        for raw_target in raw_targets:
            if hasattr(raw_target, "model_dump"):
                raw_target = raw_target.model_dump()
            if not isinstance(raw_target, dict):
                invalid("Each external link target cell must be an object.")
                link_is_invalid = True
                break
            row = raw_target.get("row")
            column = raw_target.get("column")
            if (
                not isinstance(row, int)
                or not isinstance(column, int)
                or isinstance(row, bool)
                or isinstance(column, bool)
                or row < 0
                or column < 0
            ):
                invalid(
                    "External link target row and column must be nonnegative integers.",
                )
                link_is_invalid = True
                break

            raw_source_cell = raw_target.get("source_cell")
            parsed_source_cell = (
                normalize_source_cell(raw_source_cell)
                if raw_source_cell is not None
                else None
            )
            if raw_source_cell is not None and parsed_source_cell is None:
                invalid("External link source_cell must be a valid Excel cell address.")
                link_is_invalid = True
                break
            source_cell = parsed_source_cell[0] if parsed_source_cell else None
            target_key = (row, column)
            if target_key in seen_targets:
                if seen_targets[target_key] != source_cell:
                    invalid(
                        "An external link target cell cannot map to more than one source cell.",
                    )
                    link_is_invalid = True
                    break
                continue
            seen_targets[target_key] = source_cell
            targets.append({"row": row, "column": column, "source_cell": source_cell})

        if link_is_invalid:
            continue

        if not targets:
            invalid("Each external link must include at least one valid target cell.")
            continue

        bounds = reference_bounds(reference)
        if bounds is None:
            invalid("Each external link reference must be a standalone Excel workbook reference.")
            continue
        row0, row1, column0, column1 = bounds
        has_explicit_sources = [target["source_cell"] is not None for target in targets]
        if any(has_explicit_sources) and not all(has_explicit_sources):
            invalid("External link target cells must either all include source_cell or all omit it.")
            continue

        if all(has_explicit_sources):
            seen_source_cells: set[str] = set()
            explicit_sources_valid = True
            for target in targets:
                parsed_source_cell = normalize_source_cell(target["source_cell"])
                if parsed_source_cell is None:
                    explicit_sources_valid = False
                    break
                source_cell, source_row, source_column = parsed_source_cell
                if not (
                    row0 <= source_row <= row1
                    and column0 <= source_column <= column1
                    and source_cell not in seen_source_cells
                ):
                    explicit_sources_valid = False
                    break
                seen_source_cells.add(source_cell)
                target["source_cell"] = source_cell
            if not explicit_sources_valid:
                invalid(
                    "External link source_cell values must be unique cells within the reference range.",
                )
                continue
        else:
            source_cell_count = (row1 - row0 + 1) * (column1 - column0 + 1)
            if source_cell_count != len(targets):
                invalid(
                    "Legacy external links without source_cell must map the full source range.",
                )
                continue
            source_cells = (
                f"{column_name(source_column)}{source_row + 1}"
                for source_row in range(row0, row1 + 1)
                for source_column in range(column0, column1 + 1)
            )
            for target, source_cell in zip(targets, source_cells):
                target["source_cell"] = source_cell

        link_key = (
            reference,
            tuple(
                (target["row"], target["column"], target["source_cell"])
                for target in targets
            ),
        )
        if link_key in seen_links:
            continue
        target_keys = {(target["row"], target["column"]) for target in targets}
        if target_keys & owned_targets:
            invalid("An external link target cell cannot belong to more than one link.")
            continue
        seen_links.add(link_key)
        owned_targets.update(target_keys)
        normalized.append({"reference": reference, "target_cells": targets})

    return normalized


def _normalize_dataset_internal_links(
    value: Any,
    *,
    strict: bool = False,
) -> List[Dict[str, Any]]:
    """Validate the ``internal_links`` sidecar field (ArcRho dataset cell links).

    Mirrors ``_normalize_dataset_external_links``: strict on save, lenient on
    load. Target cells are zero-based untransposed coordinates of this dataset;
    source cells are zero-based coordinates of the referenced dataset.
    """

    if value is None:
        return []
    if not isinstance(value, list):
        if strict:
            raise HTTPException(400, "internal_links must be a list.")
        return []

    from app_server.services import dataset_internal_link_service

    normalized: List[Dict[str, Any]] = []
    seen_links: set[Tuple[str, Tuple[Tuple[int, int, int, int], ...]]] = set()
    owned_targets: set[Tuple[int, int]] = set()

    def invalid(detail: str) -> bool:
        if strict:
            raise HTTPException(400, detail)
        return False

    def nonnegative_int(raw: Any) -> int | None:
        if not isinstance(raw, int) or isinstance(raw, bool) or raw < 0:
            return None
        return raw

    for raw_link in value:
        if hasattr(raw_link, "model_dump"):
            raw_link = raw_link.model_dump()
        if not isinstance(raw_link, dict):
            invalid("Each internal link must be an object.")
            continue

        raw_reference = raw_link.get("reference")
        if not isinstance(raw_reference, str) or not raw_reference.strip():
            invalid("Each internal link reference must be a nonblank string.")
            continue
        try:
            reference = dataset_internal_link_service.canonical_internal_reference(raw_reference)
        except HTTPException as err:
            invalid(f"Internal link reference is invalid: {err.detail}")
            continue

        raw_targets = raw_link.get("target_cells")
        if not isinstance(raw_targets, list) or not raw_targets:
            invalid("Each internal link must include at least one target cell.")
            continue

        targets: List[Dict[str, Any]] = []
        seen_targets: set[Tuple[int, int]] = set()
        seen_sources: set[Tuple[int, int]] = set()
        link_is_invalid = False
        for raw_target in raw_targets:
            if hasattr(raw_target, "model_dump"):
                raw_target = raw_target.model_dump()
            if not isinstance(raw_target, dict):
                invalid("Each internal link target cell must be an object.")
                link_is_invalid = True
                break
            row = nonnegative_int(raw_target.get("row"))
            column = nonnegative_int(raw_target.get("column"))
            source_row = nonnegative_int(raw_target.get("source_row"))
            source_column = nonnegative_int(raw_target.get("source_column"))
            if row is None or column is None or source_row is None or source_column is None:
                invalid(
                    "Internal link target row, column, source_row, and source_column "
                    "must be nonnegative integers.",
                )
                link_is_invalid = True
                break
            target_key = (row, column)
            source_key = (source_row, source_column)
            if target_key in seen_targets or source_key in seen_sources:
                invalid("Internal link target and source cells must be unique within a link.")
                link_is_invalid = True
                break
            seen_targets.add(target_key)
            seen_sources.add(source_key)
            targets.append(
                {
                    "row": row,
                    "column": column,
                    "source_row": source_row,
                    "source_column": source_column,
                }
            )

        if link_is_invalid or not targets:
            if not link_is_invalid:
                invalid("Each internal link must include at least one valid target cell.")
            continue

        link_key = (
            reference,
            tuple(
                (target["row"], target["column"], target["source_row"], target["source_column"])
                for target in targets
            ),
        )
        if link_key in seen_links:
            continue
        target_keys = {(target["row"], target["column"]) for target in targets}
        if target_keys & owned_targets:
            invalid("An internal link target cell cannot belong to more than one link.")
            continue
        seen_links.add(link_key)
        owned_targets.update(target_keys)
        normalized.append({"reference": reference, "target_cells": targets})

    return normalized


def _normalize_dataset_formula_links(
    value: Any,
    *,
    strict: bool = False,
) -> List[Dict[str, Any]]:
    """Validate the ``formula_links`` sidecar field (calculated dataset cells).

    Mirrors ``_normalize_dataset_internal_links``: strict on save, lenient on
    load. Target cells are zero-based untransposed coordinates of this dataset;
    result cells are zero-based coordinates of the formula's result matrix.
    """

    if value is None:
        return []
    if not isinstance(value, list):
        if strict:
            raise HTTPException(400, "formula_links must be a list.")
        return []

    from app_server.services import dataset_formula_link_service

    normalized: List[Dict[str, Any]] = []
    seen_links: set[Tuple[str, Tuple[Tuple[int, int, int, int], ...]]] = set()
    owned_targets: set[Tuple[int, int]] = set()

    def invalid(detail: str) -> bool:
        if strict:
            raise HTTPException(400, detail)
        return False

    def nonnegative_int(raw: Any) -> int | None:
        if not isinstance(raw, int) or isinstance(raw, bool) or raw < 0:
            return None
        return raw

    for raw_link in value:
        if hasattr(raw_link, "model_dump"):
            raw_link = raw_link.model_dump()
        if not isinstance(raw_link, dict):
            invalid("Each formula link must be an object.")
            continue

        raw_formula = raw_link.get("formula")
        if not isinstance(raw_formula, str) or not raw_formula.strip():
            invalid("Each formula link formula must be a nonblank string.")
            continue
        try:
            formula = dataset_formula_link_service.canonical_dataset_formula(raw_formula)
        except HTTPException as err:
            invalid(f"Formula link formula is invalid: {err.detail}")
            continue

        raw_targets = raw_link.get("target_cells")
        if not isinstance(raw_targets, list) or not raw_targets:
            invalid("Each formula link must include at least one target cell.")
            continue

        targets: List[Dict[str, Any]] = []
        seen_targets: set[Tuple[int, int]] = set()
        seen_results: set[Tuple[int, int]] = set()
        link_is_invalid = False
        for raw_target in raw_targets:
            if hasattr(raw_target, "model_dump"):
                raw_target = raw_target.model_dump()
            if not isinstance(raw_target, dict):
                invalid("Each formula link target cell must be an object.")
                link_is_invalid = True
                break
            row = nonnegative_int(raw_target.get("row"))
            column = nonnegative_int(raw_target.get("column"))
            result_row = nonnegative_int(raw_target.get("result_row"))
            result_column = nonnegative_int(raw_target.get("result_column"))
            if row is None or column is None or result_row is None or result_column is None:
                invalid(
                    "Formula link target row, column, result_row, and result_column "
                    "must be nonnegative integers.",
                )
                link_is_invalid = True
                break
            target_key = (row, column)
            result_key = (result_row, result_column)
            if target_key in seen_targets or result_key in seen_results:
                invalid("Formula link target and result cells must be unique within a link.")
                link_is_invalid = True
                break
            seen_targets.add(target_key)
            seen_results.add(result_key)
            targets.append(
                {
                    "row": row,
                    "column": column,
                    "result_row": result_row,
                    "result_column": result_column,
                }
            )

        if link_is_invalid or not targets:
            if not link_is_invalid:
                invalid("Each formula link must include at least one valid target cell.")
            continue

        link_key = (
            formula,
            tuple(
                (target["row"], target["column"], target["result_row"], target["result_column"])
                for target in targets
            ),
        )
        if link_key in seen_links:
            continue
        target_keys = {(target["row"], target["column"]) for target in targets}
        if target_keys & owned_targets:
            invalid("A formula link target cell cannot belong to more than one link.")
            continue
        seen_links.add(link_key)
        owned_targets.update(target_keys)
        normalized.append({"formula": formula, "target_cells": targets})

    return normalized


def _require_disjoint_dataset_link_targets(
    external_links: List[Dict[str, Any]],
    internal_links: List[Dict[str, Any]],
    formula_links: List[Dict[str, Any]] | None = None,
) -> None:
    owned: set[Tuple[int, int]] = set()
    for links in (external_links, internal_links, formula_links or []):
        for link in links:
            for target in link.get("target_cells") or []:
                key = (target["row"], target["column"])
                if key in owned:
                    raise HTTPException(
                        400,
                        "A dataset cell can hold only one link: an Excel workbook, "
                        "another ArcRho dataset, or a formula.",
                    )
                owned.add(key)


def _write_dataset_sidecar_payload(path: str, payload: Dict[str, Any]) -> None:
    with _dataset_sidecar_write_lock(path):
        tmp_path = f"{path}.{uuid.uuid4()}.tmp"
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(tmp_path, "w", encoding="utf-8", newline="\n") as fh:
                fh.write(persisted_json_text(finalize_sidecar(payload)))
            os.replace(tmp_path, path)
        except PermissionError:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except OSError:
                pass
            raise HTTPException(423, "Dataset sidecar is locked or inaccessible.")
        except OSError as err:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except OSError:
                pass
            raise HTTPException(500, f"Failed to write dataset sidecar: {str(err)}")


def _write_dataset_csv_and_sidecar(
    dataframe: pd.DataFrame,
    csv_path: str,
    sidecar_path: str,
    payload: Dict[str, Any],
) -> None:
    csv_existed = os.path.exists(csv_path)
    rollback_path = f"{csv_path}.{uuid.uuid4()}.rollback" if csv_existed else ""
    csv_replaced = False
    preserve_rollback = False
    try:
        if rollback_path:
            shutil.copy2(csv_path, rollback_path)
        atomic_write_csv(dataframe, csv_path)
        csv_replaced = True
        _write_dataset_sidecar_payload(sidecar_path, payload)
    except Exception:
        if csv_replaced:
            try:
                if rollback_path and os.path.exists(rollback_path):
                    os.replace(rollback_path, csv_path)
                    rollback_path = ""
                elif not csv_existed and os.path.exists(csv_path):
                    os.remove(csv_path)
            except OSError as rollback_error:
                preserve_rollback = bool(rollback_path and os.path.exists(rollback_path))
                recovery_detail = (
                    f" Recovery copy retained at {rollback_path}."
                    if preserve_rollback
                    else ""
                )
                raise HTTPException(
                    500,
                    "Dataset sidecar save failed and the dataset CSV could not be restored: "
                    f"{str(rollback_error)}.{recovery_detail}",
                )
        raise
    finally:
        if rollback_path and not preserve_rollback:
            try:
                if os.path.exists(rollback_path):
                    os.remove(rollback_path)
            except OSError:
                pass


def _empty_dataset_values(
    data_format: str,
    origin_length: int,
    development_length: int,
    triangle_shape_mask: np.ndarray | None = None,
) -> pd.DataFrame:
    fmt = str(data_format or "").strip().lower()
    n_origin = max(1, int(origin_length))
    n_dev = 1 if fmt == "vector" else max(1, int(development_length))
    values = np.zeros((n_origin, n_dev), dtype="float64")
    if fmt != "vector":
        mask = triangle_shape_mask
        if not isinstance(mask, np.ndarray) or mask.shape != (n_origin, n_dev):
            mask = triangle_mask(n_origin, n_dev)
        values = np.where(mask, 0.0, np.nan)
    return pd.DataFrame(values)


def _dataset_values_to_frame(
    values: List[List[Any]],
    mask: List[List[bool]] | None = None,
) -> pd.DataFrame:
    if not isinstance(values, list) or not values:
        raise HTTPException(400, "values must include at least one row.")
    width = 0
    for row in values:
        if not isinstance(row, list):
            raise HTTPException(400, "values must be a rectangular array.")
        width = max(width, len(row))
    if width <= 0:
        raise HTTPException(400, "values must include at least one column.")

    out: List[List[float]] = []
    for r, row in enumerate(values):
        if len(row) != width:
            raise HTTPException(400, "values must be a rectangular array.")
        out_row: List[float] = []
        mask_row = mask[r] if isinstance(mask, list) and r < len(mask) and isinstance(mask[r], list) else None
        for c, raw in enumerate(row):
            has_value = True if mask_row is None or c >= len(mask_row) else bool(mask_row[c])
            if not has_value or raw is None:
                out_row.append(np.nan)
                continue
            try:
                out_row.append(float(raw))
            except (TypeError, ValueError):
                raise HTTPException(400, "values must contain only numeric or null cells.")
        out.append(out_row)
    return pd.DataFrame(out, dtype="float64")


def _ym_to_month_index(value: Any) -> int | None:
    text = str(value if value is not None else "").strip()
    if not text:
        return None
    digits = "".join(ch for ch in text if ch.isdigit())
    if len(digits) >= 6:
        year = int(digits[:4])
        month = int(digits[4:6])
    elif len(digits) == 4:
        year = int(digits)
        month = 1
    else:
        return None
    if year <= 0 or month < 1 or month > 12:
        return None
    return year * 12 + (month - 1)


def _general_settings_month_bounds(project_name: str) -> tuple[int, int, int]:
    """The Origin Start, Origin End and Development End months of a project.

    Each is a month index (``year * 12 + month - 1``), read from the project's
    General Settings; every triangle shape in the project is built from them.
    """
    try:
        path = config.get_general_settings_path(project_name)
    except ValueError as err:
        raise HTTPException(404, str(err))
    if not os.path.exists(path):
        raise HTTPException(
            422,
            f"Cannot create dataset for project '{project_name}': General Settings are missing. "
            "Set valid Origin Start Date, Origin End Date, and Development End Date values, then try again.",
        )
    try:
        with open(path, "r", encoding="utf-8") as fh:
            payload = json.load(fh)
    except PermissionError:
        raise HTTPException(423, "General Settings are locked or inaccessible.")
    except json.JSONDecodeError as err:
        raise HTTPException(422, f"Cannot create dataset: General Settings JSON is invalid: {str(err)}")
    except OSError as err:
        raise HTTPException(500, f"Cannot read General Settings: {str(err)}")
    if not isinstance(payload, dict):
        raise HTTPException(422, "Cannot create dataset: General Settings must contain a JSON object.")

    def require_boundary_month(field: str, label: str) -> int:
        raw = str(payload.get(field) or "").strip()
        if not re.fullmatch(r"\d{6}", raw):
            raise HTTPException(
                422,
                f"Cannot create dataset for project '{project_name}': {label} is missing or invalid in Project Settings.",
            )
        month_index = _ym_to_month_index(raw)
        if month_index is None:
            raise HTTPException(
                422,
                f"Cannot create dataset for project '{project_name}': {label} is missing or invalid in Project Settings.",
            )
        return month_index

    origin_start_month = require_boundary_month("origin_start_date", "Origin Start Date")
    origin_end_month = require_boundary_month("origin_end_date", "Origin End Date")
    development_end_month = require_boundary_month("development_end_date", "Development End Date")
    if origin_end_month < origin_start_month:
        raise HTTPException(422, "Cannot create dataset: Origin End Date must not be before Origin Start Date.")
    if development_end_month < origin_start_month:
        raise HTTPException(422, "Cannot create dataset: Development End Date must not be before Origin Start Date.")
    return origin_start_month, origin_end_month, development_end_month


def valuation_months(project_name: str) -> int:
    """Months from the project's Origin Start Date through its Development End Date.

    The count a roll-up of a stored triangle is anchored on: the newest cell
    of every row is valued on the Development End Date, so a coarser view
    counts its development periods back from there.
    """
    origin_start_month, _, development_end_month = _general_settings_month_bounds(project_name)
    return development_end_month - origin_start_month + 1


def _empty_dataset_geometry_from_general_settings(
    project_name: str,
    origin_period_length: int,
    development_period_length: int,
) -> tuple[int, int, np.ndarray | None]:
    origin_start_month, origin_end_month, development_end_month = _general_settings_month_bounds(project_name)
    origin_period = max(1, int(origin_period_length or 1))
    development_period = max(1, int(development_period_length or 1))
    origin_count = ((origin_end_month - origin_start_month) // origin_period) + 1
    development_count = ((development_end_month - origin_start_month) // development_period) + 1
    origin_offsets = np.arange(origin_count)[:, None] * origin_period
    development_offsets = np.arange(development_count)[None, :] * development_period
    mask = origin_start_month + origin_offsets + development_offsets <= development_end_month
    return origin_count, development_count, mask


def triangle_grid_shape(
    project_name: str,
    origin_length: int,
    development_length: int,
) -> Dict[str, Any]:
    """The rows, columns, and cells a triangle created at this shape would have.

    A hand-entered triangle is drawn in the window before there is a file to
    read it from, so the grid asks for the geometry the empty CSV would be
    written at: the cells stop on the project's calendar diagonal, where the
    Origin Start Date, the Origin End Date, and the Development End Date put
    them together. A Vector has no diagonal and needs no answer.
    """
    p = str(project_name or "").strip()
    if not p:
        raise HTTPException(400, "project_name is required.")
    origin_count, development_count, mask = _empty_dataset_geometry_from_general_settings(
        p,
        max(1, int(origin_length or 1)),
        max(1, int(development_length or 1)),
    )
    return {
        "ok": True,
        "project_name": p,
        "origin_count": int(origin_count),
        "development_count": int(development_count),
        "mask": mask.tolist(),
    }


def _triangle_shape_masked(
    mask: List[List[bool]],
    project_name: str,
    origin_length: int,
    development_length: int,
) -> List[List[bool]]:
    """Drop from ``mask`` every cell past the project's calendar diagonal.

    A hand-entered triangle has the cells :func:`triangle_grid_shape` gives it,
    whatever its file happens to hold: a figure written past the diagonal --
    by a window that laid the grid out on a rule of its own -- is neither shown
    nor editable, and the next save writes that cell blank.
    """
    try:
        _, _, shape = _empty_dataset_geometry_from_general_settings(
            project_name, origin_length, development_length
        )
    except HTTPException:
        return mask
    rows, columns = shape.shape
    return [
        [
            bool(cell) and r < rows and c < columns and bool(shape[r, c])
            for c, cell in enumerate(row)
        ]
        for r, row in enumerate(mask)
    ]


def valuation_origin_row_count(project_name: str, origin_period_length: int) -> int | None:
    """Count the origin periods that start on or before the Development End Date.

    A vector keeps every configured origin row, so its rows after the valuation
    period may hold values (full-year inputs) or stay blank. This count is the
    boundary a negative index counts back from; it is ``None`` when the project
    has no usable General Settings.
    """
    period = max(1, int(origin_period_length or 1))
    try:
        _, _, mask = _empty_dataset_geometry_from_general_settings(project_name, period, period)
    except HTTPException:
        return None
    return int(mask[:, 0].sum())


def _containing_project_name_for_dataset(path: str) -> str:
    projects_root = str(config.PROJECT_SETTINGS_DIR or "").strip()
    if not projects_root:
        return ""
    root_path = os.path.realpath(projects_root)
    dataset_path = os.path.realpath(path)
    try:
        if os.path.commonpath((os.path.normcase(root_path), os.path.normcase(dataset_path))) != os.path.normcase(root_path):
            return ""
        relative = os.path.relpath(dataset_path, root_path)
    except ValueError:
        return ""
    parts = os.path.normpath(relative).split(os.sep)
    if len(parts) < 3 or parts[1].casefold() != str(config.PROJECT_DATA_DIR).casefold():
        return ""
    return str(parts[0]).strip()


def _dataset_owning_project_name(path: str, payload: Dict[str, Any]) -> str:
    """The project a cached dataset belongs to.

    The folder holding the CSV decides it, and the name stored in the sidecar
    is only a fallback for a cache outside the projects tree. Duplicating or
    renaming a project copies that stored name verbatim, so in a duplicate it
    still names the project the data came from.
    """

    return (
        _containing_project_name_for_dataset(path)
        or str(payload.get("project_name") or "").strip()
    )


def _dataset_patch_mask(path: str, n_origin: int, n_dev: int) -> np.ndarray:
    try:
        sidecar_path = dataset_instance_index_service._dataset_sidecar_path_for_cached_csv(path)
        payload = _read_dataset_sidecar(sidecar_path)
        project_name = _dataset_owning_project_name(path, payload)
        # Stored, not displayed: the mask covers the CSV at ``path``, whose
        # blank cells are laid out at the shape that file holds.
        stored_origin, stored_development = stored_lengths(payload)
        origin_period_len = max(1, stored_origin)
        dev_period_len = max(1, stored_development)
        if project_name:
            _, _, mask = _empty_dataset_geometry_from_general_settings(
                project_name,
                origin_period_len,
                dev_period_len,
            )
            if isinstance(mask, np.ndarray) and mask.shape == (n_origin, n_dev):
                return mask
    except HTTPException:
        raise
    except Exception:
        pass
    return triangle_mask(n_origin, n_dev)


def _create_empty_cached_dataset_impl(
    project_name: str,
    reserving_class: str,
    dataset_type: str,
    *,
    instance_name: str = "",
    data_format: str = "Triangle",
    origin_length: int = 12,
    development_length: int = 12,
    cumulative: bool = True,
    calendar: bool = False,
) -> Dict[str, Any]:
    p, rc, ds_type = _require_dataset_fields(project_name, reserving_class, dataset_type)
    instance = str(instance_name or ds_type).strip()
    if not instance:
        raise HTTPException(400, "instance_name or dataset_type is required.")

    try:
        from app_server.services import calculated_dataset_service

        calc_result = calculated_dataset_service.recalculate_dataset(p, rc, ds_type)
    except Exception as err:
        calc_result = {"ok": False, "reason": "calculation_error", "errors": [str(err)]}
    if calc_result.get("ok"):
        csv_path = str(calc_result.get("path") or "")
        if csv_path:
            ds_id = "arcrhotri_" + hashlib.sha1(csv_path.encode("utf-8")).hexdigest()[:16]
            config.DATASETS[ds_id] = csv_path
            calculated_updates = dependent_propagation_service.enqueue_marked_save_propagation(
                p,
                rc,
                ds_type,
                ds_type,
            )
            index_error = ""
            try:
                dataset_instance_index_service.rebuild_index(p, rc)
            except Exception as err:
                index_error = str(err)
            try:
                n_origin, n_dev = infer_shape(csv_path)
            except Exception:
                n_origin, n_dev = 0, 0
            return {
                "ok": True,
                "project_name": p,
                "reserving_class": rc,
                "dataset_name": instance,
                "dataset_type": ds_type,
                "source_kind": "calculated",
                "data_format": "Calculated",
                "origin_length": 12,
                "development_length": 12,
                "shape": {"n_origin": n_origin, "n_development": n_dev},
                "csv_file": os.path.basename(csv_path),
                "ds_id": ds_id,
                "path": csv_path,
                "sidecar_path": str(calc_result.get("sidecar_path") or ""),
                "calculated": True,
                "calculated_updates": calculated_updates,
                "propagation_ok": bool(calculated_updates and calculated_updates.get("ok")),
                "index_ok": not index_error,
                "index_error": index_error,
            }
    if calc_result.get("reason") not in {"not_calculated"}:
        detail = "; ".join(str(item) for item in calc_result.get("errors") or [])
        if not detail:
            detail = str(calc_result.get("reason") or "Failed to calculate dataset.")
        raise HTTPException(422, detail)

    try:
        data_dir = config.get_project_dataset_cache_dir(p, rc)
        sidecar_dir = config.get_project_dataset_sidecar_dir(p, rc)
    except ValueError as err:
        raise HTTPException(404, str(err))

    origin_period_len = max(1, int(origin_length))
    dev_period_len = max(1, int(development_length))
    origin_count, dev_count, triangle_shape_mask = _empty_dataset_geometry_from_general_settings(
        p,
        origin_period_len,
        dev_period_len,
    )
    fmt = str(data_format or "Triangle").strip() or "Triangle"
    folder = data_dir
    csv_stem = build_dataset_cache_file_name(instance, fmt, origin_period_len, dev_period_len, cumulative, calendar)
    csv_path = os.path.join(folder, f"{csv_stem}.csv")
    sidecar_path = os.path.join(sidecar_dir, f"{sanitize_dataset_file_name(instance)}.json")
    now = _now_utc_iso()
    user_name = _current_user_name()

    df = _empty_dataset_values(fmt, origin_count, dev_count, triangle_shape_mask)
    payload = {
        "dataset_name": instance,
        "dataset_type": ds_type,
        "reserving_class": rc,
        "project_name": p,
        "source_kind": "input",
        "data_format": fmt,
        "show_subtotal": DEFAULT_SHOW_SUBTOTAL,
        "csv_file": os.path.basename(csv_path),
        "created": now,
        "modified_by": user_name,
        "updated_at": now,
        "method_type": dataset_sidecar_status_service.METHOD_TYPE_NONE,
        "status": dataset_sidecar_status_service.STATUS_CURRENT,
    }
    if fmt.strip().lower() == "vector":
        payload["period_length"] = origin_period_len
    else:
        payload["origin_length"] = origin_period_len
        payload["development_length"] = dev_period_len
        payload["cumulative"] = bool(cumulative)
        payload["calendar"] = bool(calendar)
    # The empty CSV is written at the requested shape, so that is the shape
    # this dataset's values are stored at.
    payload.update(stored_length_fields(fmt, origin_period_len, dev_period_len))
    _append_dataset_audit_entry(payload, "Insert", event_date=now, user_name=user_name)
    from app_server.services import calculated_dataset_service

    calculated_dataset_service.apply_sidecar_graph_fields(payload, p, ds_type)

    try:
        os.makedirs(folder, exist_ok=True)
        os.makedirs(sidecar_dir, exist_ok=True)
        _write_dataset_csv_and_sidecar(df, csv_path, sidecar_path, payload)
    except PermissionError:
        raise HTTPException(423, "Dataset cache file is locked or inaccessible.")
    except OSError as err:
        raise HTTPException(500, f"Failed to create empty dataset cache: {str(err)}")

    dataset_sidecar_status_service.refresh_method_statuses_for_dependents(p, rc, [instance])
    calculated_updates = dependent_propagation_service.enqueue_save_propagation(
        p,
        rc,
        [dependent_propagation_service.changed_root(instance, ds_type)],
    )
    index_error = ""
    try:
        dataset_instance_index_service.rebuild_index(p, rc)
    except Exception as err:
        index_error = str(err)

    ds_id = "arcrhotri_" + hashlib.sha1(csv_path.encode("utf-8")).hexdigest()[:16]
    config.DATASETS[ds_id] = csv_path
    return {
        "ok": True,
        "project_name": p,
        "reserving_class": rc,
        "dataset_name": instance,
        "dataset_type": ds_type,
        "source_kind": "input",
        "data_format": fmt,
        "origin_length": origin_period_len,
        "development_length": dev_period_len,
        "shape": {"n_origin": origin_count, "n_development": 1 if fmt.strip().lower() == "vector" else dev_count},
        "csv_file": os.path.basename(csv_path),
        "ds_id": ds_id,
        "path": csv_path,
        "sidecar_path": sidecar_path,
        "calculated_updates": calculated_updates,
        "propagation_ok": bool(calculated_updates and calculated_updates.get("ok")),
        "index_ok": not index_error,
        "index_error": index_error,
    }


def create_empty_cached_dataset(
    project_name: str,
    reserving_class: str,
    dataset_type: str,
    **kwargs: Any,
) -> Dict[str, Any]:
    # Dependent propagation runs on ArcRho Engine; block the create before any
    # write when no live Engine can pick the job up or another walk is still
    # rewriting this reserving class.
    dependent_propagation_service.require_reserving_class_writable(
        project_name, reserving_class
    )
    with dataset_sidecar_status_service.reserving_class_io_lock(project_name, reserving_class):
        return _create_empty_cached_dataset_impl(
            project_name,
            reserving_class,
            dataset_type,
            **kwargs,
        )


def register_rollup_handle(ds_id: str, recipe: Dict[str, Any]) -> None:
    """Bind a dataset id to the roll-up that builds it from a finer stored CSV.

    A hand-entered triangle is only ever stored at the shape it was typed at.
    A coarser view of it is built again on every read from that CSV, so a
    later edit to the figures can never be served from an older view.
    """

    config.DATASET_ROLLUPS[ds_id] = dict(recipe)


def _rolled_up_dataset(ds_id: str) -> Tuple[pd.DataFrame, float] | None:
    recipe = config.DATASET_ROLLUPS.get(ds_id)
    if not recipe:
        return None
    source_path = str(recipe.get("source_path") or "")
    if not source_path or not os.path.exists(source_path):
        return None
    rows = pd.read_csv(
        source_path, header=None, dtype="float64", keep_default_na=True
    ).to_numpy().tolist()
    values = rollup_triangle(
        rows,
        source_origin_length=int(recipe["source_origin_length"]),
        source_development_length=int(recipe["source_development_length"]),
        target_origin_length=int(recipe["target_origin_length"]),
        target_development_length=int(recipe["target_development_length"]),
        valuation_months=int(recipe["valuation_months"]),
        cumulative=bool(recipe.get("cumulative", True)),
        calendar=bool(recipe.get("calendar", False)),
    )
    return pd.DataFrame(values, dtype="float64"), os.stat(source_path).st_mtime


def get_dataset(ds_id: str, project_name: str, origin_length: int) -> Dict[str, Any] | None:
    path = config.DATASETS.get(ds_id)
    if not path:
        return None
    rolled_up = _rolled_up_dataset(ds_id)
    if rolled_up is None:
        if not os.path.exists(path):
            return None
        df = pd.read_csv(path, header=None, dtype="float64", keep_default_na=True)
        mtime = os.stat(path).st_mtime
    else:
        df, mtime = rolled_up
    n_origin, n_dev = df.shape

    origin_labels = _resolve_origin_labels(ds_id, path, project_name, origin_length, n_origin)
    dev_labels = [str(12 * (j + 1)) for j in range(n_dev)]

    values = df.to_numpy()
    mask = ~np.isnan(values)

    return {
        "id": ds_id,
        "origin_labels": origin_labels,
        "dev_labels": dev_labels,
        "values": np.where(np.isnan(values), _BLANK_CELL, values).tolist(),
        "mask": mask.tolist(),
        "mtime": mtime,
    }


def get_diagonal(
    ds_id: str, project_name: str, origin_length: int, k: int = 0
) -> Dict[str, Any] | None:
    path = config.DATASETS.get(ds_id)
    if not path:
        return None
    rolled_up = _rolled_up_dataset(ds_id)
    if rolled_up is None:
        if not os.path.exists(path):
            return None
        df = load_triangle_values(path)
    else:
        df = rolled_up[0]
    n_origin, n_dev = df.shape
    origin_labels = _resolve_origin_labels(ds_id, path, project_name, origin_length, n_origin)
    dev_labels = [str(12 * (j + 1)) for j in range(n_dev)]

    idx = diagonal_indices(n_origin, n_dev, k=k)
    items = []
    for r, c in idx:
        # A cell reads back as a pandas scalar, which the stubs allow to be a
        # complex number; a triangle only ever holds real numbers or a blank.
        v = cast(Any, df.iat[r, c])
        items.append({
            "r": r,
            "c": c,
            "origin": origin_labels[r],
            "dev": dev_labels[c],
            "value": None if pd.isna(v) else float(v),
        })

    return {"id": ds_id, "k": k, "items": items}


def register_dataset_handle(dataset_id: str, csv_path: str) -> None:
    """Bind a dataset id to its cached CSV for the id-addressed grid routes.

    The registry is per process. A cached-dataset load that ran on the ArcRho
    Server host must register its handle here on the client, using the CSV
    path already rebased onto this PC's workspace root, or the grid patch and
    diagonal routes for that id would find nothing.
    """

    ds_id = str(dataset_id or "").strip()
    path = str(csv_path or "").strip()
    if ds_id and path:
        config.DATASETS[ds_id] = path


def _require_dataset_fields(project_name: str, reserving_class: str, dataset_name: str) -> Tuple[str, str, str]:
    p = str(project_name if project_name is not None else "")
    rc = str(reserving_class if reserving_class is not None else "")
    ds = str(dataset_name if dataset_name is not None else "")
    if not p.strip() or not rc.strip() or not ds.strip():
        raise HTTPException(400, "project_name, reserving_class, and dataset_name are required.")
    return p, rc, ds


def _get_dataset_sidecar_path(
    project_name: str,
    reserving_class: str,
    dataset_name: str,
    csv_file: str = "",
    *,
    sidecar_dir: str = "",
) -> str:
    _ = csv_file
    if not sidecar_dir:
        try:
            sidecar_dir = config.get_project_dataset_sidecar_dir(project_name, reserving_class)
        except ValueError as err:
            raise HTTPException(404, str(err))
    ds_file = sanitize_dataset_file_name(dataset_name)
    return os.path.join(sidecar_dir, f"{ds_file}.json")


def _now_utc_iso() -> str:
    return utc_now_text()


def _int_or_default(value: Any, default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _read_dataset_sidecar(path: str) -> Dict[str, Any]:
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            payload = json.load(fh)
    except PermissionError:
        raise HTTPException(423, "Dataset sidecar is locked or inaccessible.")
    except OSError as err:
        raise HTTPException(500, f"Failed to read dataset sidecar: {str(err)}")
    except json.JSONDecodeError as err:
        raise HTTPException(500, f"Invalid dataset sidecar JSON format: {str(err)}")
    return payload if isinstance(payload, dict) else {}


def _dataset_type_calculation_map(project_name: str) -> Dict[str, tuple[bool, str]]:
    try:
        from app_server.services import calculated_dataset_service

        rows = calculated_dataset_service._dataset_type_rows(project_name)
    except Exception:
        return {}
    out: Dict[str, tuple[bool, str]] = {}
    for row in rows:
        name_key = str(row.get("name") or "").strip().lower()
        if not name_key:
            continue
        formula = str(row.get("formula") or "").strip()
        out[name_key] = (
            bool(row.get("calculated") and not row.get("generated") and formula),
            formula,
        )
    return out


def _is_app_calculated_dataset_type(
    project_name: str,
    dataset_type_name: str,
    *,
    calculation_map: Dict[str, tuple[bool, str]] | None = None,
) -> tuple[bool, str]:
    name_key = str(dataset_type_name or "").strip().lower()
    if not name_key:
        return False, ""
    resolved = calculation_map if calculation_map is not None else _dataset_type_calculation_map(project_name)
    return resolved.get(name_key, (False, ""))


def _dataset_index_entry_map(project_name: str, reserving_class: str) -> Dict[str, Dict[str, str]]:
    """Map every dataset in the reserving class to the fields a chip needs.

    ``index.json`` already records each instance's canonical name, its dataset
    type and its method type, and the reserving-class index is validated against
    one folder listing rather than a read per file. Resolving the Details graph
    from it costs one read instead of one read per precedent and dependent, which
    over a network share is the difference between a visible wait and none.
    """

    try:
        index = dataset_instance_index_service.get_index(project_name, reserving_class, refresh=False)
    except Exception:
        return {}
    out: Dict[str, Dict[str, str]] = {}
    for item in index.get("files") or []:
        if not isinstance(item, dict):
            continue
        name = str(item.get("name") or "").strip()
        if not name:
            continue
        out[name.lower()] = {
            "dataset_name": name,
            "dataset_type": str(item.get("dataset_type") or "").strip(),
            "method_type": dataset_sidecar_status_service.normalize_method_type(
                item.get("method_type"),
                item.get("source_kind"),
            ),
        }
    return out


def _sidecar_graph_entries(
    project_name: str,
    reserving_class: str,
    entries: Any,
    *,
    include_formula: bool = False,
    include_method_type: bool = False,
    calculation_map: Dict[str, tuple[bool, str]] | None = None,
    index_map: Dict[str, Dict[str, str]] | None = None,
) -> List[Dict[str, str]]:
    """Fill in the chip fields the persisted graph does not carry.

    A sidecar stores its precedents and dependents as bare names, but a chip also
    needs the neighbour's dataset type - to show the calculated formula on hover -
    and its method type, which decides what a click opens. Both are read from the
    reserving-class index rather than from each neighbour's own sidecar, so a
    graph of any width costs one index read instead of one file read per chip.
    """

    out = dataset_sidecar_status_service.name_entries(
        dataset_sidecar_status_service.entry_names(entries)
    )
    if not include_formula and not include_method_type:
        return out
    resolved_index = (
        index_map
        if index_map is not None
        else _dataset_index_entry_map(project_name, reserving_class)
    )
    for item in out:
        name = str(item.get("dataset_name") or "").strip()
        if not name:
            continue
        indexed = resolved_index.get(name.lower()) or {}
        dataset_name = str(indexed.get("dataset_name") or name).strip()
        dataset_type = str(indexed.get("dataset_type") or name).strip()
        _, type_formula = _is_app_calculated_dataset_type(
            project_name,
            dataset_type,
            calculation_map=calculation_map,
        )
        formula = str(type_formula or "").strip()
        item["dataset_name"] = dataset_name or name
        item["dataset_type"] = dataset_type or name
        if include_method_type:
            item["method_type"] = str(
                indexed.get("method_type")
                or dataset_sidecar_status_service.METHOD_TYPE_NONE
            )
        if formula:
            item["formula"] = formula
    return out


def load_dataset_sidecar(project_name: str, reserving_class: str, dataset_name: str) -> Dict[str, Any]:
    p, rc, ds = _require_dataset_fields(project_name, reserving_class, dataset_name)
    path = _get_dataset_sidecar_path(p, rc, ds)
    payload = _read_dataset_sidecar(path)
    if not payload:
        return {
            "ok": True,
            "exists": False,
            "project_name": p,
            "reserving_class": rc,
            "dataset_name": ds,
            "external_links": [],
            "internal_links": [],
            "formula_links": [],
            "show_subtotal": DEFAULT_SHOW_SUBTOTAL,
            "path": path,
        }
    dataset_type = str(payload.get("dataset_type") or ds)
    data_format = str(payload.get("data_format") or "")
    is_vector = data_format.strip().lower() == "vector"
    period_length = payload.get("period_length")
    origin_length = period_length if is_vector else payload.get("origin_length")
    development_length = period_length if is_vector else payload.get("development_length")
    stored_origin_length, stored_development_length = stored_lengths(payload)
    linked_origin_length, linked_development_length = linked_lengths(payload)
    calculation_map = _dataset_type_calculation_map(p)
    # Both chip rows resolve their neighbours from the same index read, so a
    # graph is one lookup wide however many precedents and dependents it holds.
    has_graph = bool(
        dataset_sidecar_status_service.entry_names(payload.get("precedents"))
        or dataset_sidecar_status_service.entry_names(payload.get("dependents"))
    )
    index_map = _dataset_index_entry_map(p, rc) if has_graph else {}
    app_calculated, formula = _is_app_calculated_dataset_type(
        p,
        dataset_type,
        calculation_map=calculation_map,
    )
    precedents = _sidecar_graph_entries(
        p,
        rc,
        payload.get("precedents"),
        include_method_type=True,
        calculation_map=calculation_map,
        index_map=index_map,
    )
    dependents = _sidecar_graph_entries(
        p,
        rc,
        payload.get("dependents"),
        include_formula=True,
        calculation_map=calculation_map,
        index_map=index_map,
    )
    return {
        "ok": True,
        "exists": True,
        # The folder the sidecar was read from owns it. A duplicated or renamed
        # project keeps the old name in every copied sidecar, and a caller that
        # believed it would rebuild this dataset into the wrong project.
        "project_name": p,
        "reserving_class": str(payload.get("reserving_class") or rc),
        "dataset_name": str(payload.get("dataset_name") or ds),
        "dataset_type": dataset_type,
        "instance_name": str(payload.get("dataset_name") or ds),
        "data_format": data_format,
        "period_length": period_length if is_vector else None,
        # Both shapes: the display pair is what the window reopens at, the
        # stored pair is how fine the data underneath it really is.
        "origin_length": origin_length,
        "development_length": development_length,
        "stored_period_length": stored_origin_length if is_vector else None,
        "stored_origin_length": stored_origin_length,
        "stored_development_length": stored_development_length,
        # The display the dataset's cell links were written against, which the
        # display pair above may since have left behind.
        "linked_period_length": linked_origin_length if is_vector else None,
        "linked_origin_length": linked_origin_length,
        "linked_development_length": linked_development_length,
        "origin_labels": _normalize_origin_labels(payload.get("origin_labels")),
        "cumulative": payload.get("cumulative"),
        "transposed": payload.get("transposed"),
        "calendar": payload.get("calendar"),
        "show_subtotal": normalize_show_subtotal(payload.get("show_subtotal")),
        "number_format": _normalize_number_format(payload.get("number_format") or "0,000"),
        "decimal_places": _normalize_decimal_places(payload.get("decimal_places")),
        "csv_file": str(payload.get("csv_file") or ""),
        "source_kind": str(payload.get("source_kind") or ""),
        "method_type": dataset_sidecar_status_service.normalize_method_type(
            payload.get("method_type"),
            payload.get("source_kind"),
        ),
        "status": dataset_sidecar_status_service.normalize_status(payload.get("status")),
        "notes": str(payload.get("notes") or ""),
        "calculated": True if app_calculated else payload.get("calculated"),
        "formula": str(formula or ""),
        "external_links": _normalize_dataset_external_links(payload.get("external_links")),
        "internal_links": _normalize_dataset_internal_links(payload.get("internal_links")),
        "formula_links": _normalize_dataset_formula_links(payload.get("formula_links")),
        "precedents": precedents,
        "dependents": dependents,
        "modified_by": str(payload.get("modified_by") or ""),
        "created": str(payload.get("created") or ""),
        "updated_at": str(payload.get("updated_at") or ""),
        "audit_log": _normalize_dataset_audit_log(payload.get("audit_log")),
        "path": path,
    }


def _cached_csv_candidates(
    project_name: str,
    reserving_class: str,
    dataset_name: str,
    sidecar: Dict[str, Any],
) -> Iterator[str]:
    try:
        data_dir = config.get_project_dataset_cache_dir(project_name, reserving_class)
    except ValueError as err:
        raise HTTPException(404, str(err))
    seen = set()

    def candidate_path(name: str) -> str:
        clean = os.path.basename(str(name or "").strip())
        key = clean.lower()
        if not clean or key in seen:
            return ""
        seen.add(key)
        return os.path.join(data_dir, clean)

    csv_file = str(sidecar.get("csv_file") or "").strip()
    if csv_file:
        candidate = candidate_path(csv_file)
        if candidate:
            yield candidate
    base = sanitize_dataset_file_name(dataset_name, "Dataset")
    candidate = candidate_path(f"{base}.csv")
    if candidate:
        yield candidate

    # Directory enumeration is the network-drive fallback, not the normal path.
    # The caller checks each yielded preferred path before requesting the next.
    if os.path.isdir(data_dir):
        prefix = f"{base}@".lower()
        for filename in os.listdir(data_dir):
            name_l = filename.lower()
            if not name_l.endswith(".csv"):
                continue
            if name_l == f"{base}.csv".lower() or name_l.startswith(prefix):
                candidate = candidate_path(filename)
                if candidate:
                    yield candidate


def _parse_length_scoped_cache_name(filename: str) -> Dict[str, Any]:
    stem, ext = os.path.splitext(os.path.basename(filename))
    if ext.lower() != ".csv":
        return {}
    parts = stem.split("@")
    if len(parts) >= 2 and parts[-1].strip().isdigit():
        period = int(parts[-1].strip())
        return {
            "origin_length": period,
            "development_length": period,
        }
    if len(parts) < 5:
        return {}
    origin = parts[-4].strip()
    development = parts[-3].strip()
    cumulative = parts[-2].strip().lower()
    calendar = parts[-1].strip().lower()
    if not origin.isdigit() or not development.isdigit():
        return {}
    if cumulative not in {"cum", "inc"} or calendar not in {"dev", "cal"}:
        return {}
    return {
        "origin_length": int(origin),
        "development_length": int(development),
        "cumulative": cumulative == "cum",
        "calendar": calendar == "cal",
    }


def load_cached_dataset_values(
    project_name: str,
    reserving_class: str,
    dataset_name: str,
    *,
    csv_file: str = "",
    origin_length: int | None = None,
    development_length: int | None = None,
    cumulative: bool = True,
    calendar: bool = False,
    at_display_shape: bool = False,
    at_linked_shape: bool = False,
) -> Dict[str, Any]:
    """Read a dataset's CSV with the sidecar settings a window needs beside it.

    ``origin_length`` / ``development_length`` in the response are the shape
    of the ``values`` returned, and the ``stored_*`` pair is the shape the
    dataset's own file is held at. By default the values are the file's own
    rows, which is what a method reading its inputs wants. With
    ``at_display_shape`` a hand-entered dataset shown coarser than it is
    stored is rolled up to the display shape its sidecar saved, the view the
    Dataset window opens at, built from the stored CSV the same way a run
    builds it and never written beside it. ``at_linked_shape`` rolls it up
    the same way to the display its cell links were written against, which
    is where a link refresh reads and writes its cells.
    """
    p, rc, ds = _require_dataset_fields(project_name, reserving_class, dataset_name)
    sidecar_path = _get_dataset_sidecar_path(p, rc, ds)
    sidecar = _read_dataset_sidecar(sidecar_path)
    try:
        data_dir = config.get_project_dataset_cache_dir(p, rc)
    except ValueError as err:
        raise HTTPException(404, str(err))
    exact_candidates: List[str] = []
    exact_requested = bool(csv_file or (origin_length and development_length))
    if csv_file:
        exact_candidates.append(
            os.path.join(data_dir, os.path.basename(str(csv_file).strip()))
        )
    elif sidecar.get("csv_file"):
        exact_candidates.append(
            os.path.join(data_dir, os.path.basename(str(sidecar.get("csv_file")).strip()))
        )
    if origin_length and development_length:
        cache_name = build_dataset_cache_file_name(
            ds,
            sidecar.get("data_format") or "Triangle",
            origin_length,
            development_length,
            cumulative,
            calendar,
        )
        exact_candidates.append(os.path.join(data_dir, f"{cache_name}.csv"))
    candidates: Iterable[str] = (
        exact_candidates
        if exact_requested
        else _cached_csv_candidates(p, rc, ds, sidecar)
    )

    seen = set()
    csv_path = ""
    for candidate in candidates:
        key = os.path.normcase(os.path.abspath(candidate))
        if key in seen:
            continue
        seen.add(key)
        if os.path.exists(candidate) and os.path.isfile(candidate):
            csv_path = candidate
            break
    if not csv_path:
        if exact_requested:
            raise HTTPException(404, f"Requested cached dataset CSV not found for '{ds}'.")
        raise HTTPException(404, f"Cached dataset CSV not found for '{ds}'.")
    try:
        df = pd.read_csv(csv_path, header=None)
    except PermissionError:
        raise HTTPException(423, "Dataset cache CSV is locked or inaccessible.")
    except OSError as err:
        raise HTTPException(500, f"Failed to read dataset cache CSV: {str(err)}")
    except Exception as err:
        raise HTTPException(500, f"Invalid dataset cache CSV format: {str(err)}")
    df = df.astype(object).where(pd.notnull(df), None)
    values = df.values.tolist()
    parsed_name = _parse_length_scoped_cache_name(os.path.basename(csv_path))
    data_format = str(sidecar.get("data_format") or "")
    is_vector = data_format.strip().lower() == "vector"
    # Stored, not displayed: these describe the CSV just read. The file name
    # states its own shape; the sidecar's stored pair answers for a cache
    # whose name does not.
    sidecar_origin_length, sidecar_development_length = stored_lengths(sidecar)
    resolved_origin_length = _int_or_default(parsed_name.get("origin_length") or sidecar_origin_length, max(1, len(values)))
    resolved_development_length = _int_or_default(parsed_name.get("development_length") or sidecar_development_length, max(1, len(values[0]) if values else 1))
    if not (sidecar_origin_length and sidecar_development_length):
        sidecar_origin_length, sidecar_development_length = resolved_origin_length, resolved_development_length
    linked_origin_length, linked_development_length = linked_lengths(sidecar)
    dataset_id = "arcrhotri_" + hashlib.sha1(csv_path.encode("utf-8")).hexdigest()[:16]
    handle_path = csv_path
    if at_display_shape or at_linked_shape:
        view = _display_view_of_stored_values(
            p,
            ds,
            sidecar,
            csv_path,
            values,
            linked_lengths(sidecar) if at_linked_shape else display_lengths(sidecar),
        )
        if view is not None:
            values, resolved_origin_length, resolved_development_length, dataset_id, handle_path = view
    origin_labels, _ = _validate_origin_labels(
        sidecar.get("origin_labels"),
        len(values),
        resolved_origin_length,
    )
    column_count = max((len(row) for row in values), default=0)
    development_labels = _normalize_origin_labels(sidecar.get("development_labels"))
    # The origin-header, development-header, and Dataset Type reads are
    # independent network I/O; the bounded pool prices a mapped-drive open at
    # one hydration chain of latency instead of three sequential chains.
    origin_future = None
    if not origin_labels:
        origin_future = _CACHED_LOAD_HYDRATION_EXECUTOR.submit(
            _resolve_origin_labels,
            dataset_id,
            csv_path,
            p,
            resolved_origin_length,
            len(values),
        )
    development_future = None
    if is_vector:
        if len(development_labels) != column_count and column_count == 1:
            development_labels = ["Ultimate"]
    else:
        # A triangle's column ages belong to the project, not to the file: they
        # count back from its Development End Date, and the view they are asked
        # for is the one on screen, whether or not the file is stored finer.
        # The sidecar's own list rides along only as the answer for a project
        # whose headers cannot be read.
        development_future = _CACHED_LOAD_HYDRATION_EXECUTOR.submit(
            _resolve_development_labels,
            dataset_id,
            csv_path,
            p,
            resolved_development_length,
            column_count,
            calendar=bool(sidecar.get("calendar")),
            fallback_labels=development_labels,
        )
    formula_future = _CACHED_LOAD_HYDRATION_EXECUTOR.submit(
        _is_app_calculated_dataset_type,
        p,
        str(sidecar.get("dataset_type") or ds),
    )
    # Collect in the original sequential order so an origin-label failure
    # keeps precedence over development-label and formula failures.
    if origin_future is not None:
        origin_labels = origin_future.result()
    if development_future is not None:
        development_labels = development_future.result()
    _, dataset_type_formula = formula_future.result()
    try:
        file_mtime = os.stat(csv_path).st_mtime
    except OSError:
        file_mtime = None
    register_dataset_handle(dataset_id, handle_path)
    mask = [[value is not None for value in row] for row in values]
    if not is_vector and str(sidecar.get("source_kind") or "").strip().casefold() == "input":
        mask = _triangle_shape_masked(
            mask, p, resolved_origin_length, resolved_development_length
        )
    return {
        "ok": True,
        "id": dataset_id,
        "project_name": p,
        "reserving_class": rc,
        "dataset_name": str(sidecar.get("dataset_name") or ds),
        "dataset_type": str(sidecar.get("dataset_type") or ds),
        "data_format": data_format,
        "origin_length": resolved_origin_length,
        "development_length": resolved_development_length,
        # The stored pair rides with every load, as it does with the sidecar
        # load and save, so the window knows how fine the file under the
        # shape it shows really is.
        "stored_period_length": sidecar_origin_length if is_vector else None,
        "stored_origin_length": sidecar_origin_length,
        "stored_development_length": sidecar_development_length,
        "linked_period_length": linked_origin_length if is_vector else None,
        "linked_origin_length": linked_origin_length,
        "linked_development_length": linked_development_length,
        "origin_labels": origin_labels,
        "dev_labels": development_labels,
        "mask": mask,
        "mtime": file_mtime,
        "csv_file": os.path.basename(csv_path),
        "source_kind": str(sidecar.get("source_kind") or ""),
        "method_type": dataset_sidecar_status_service.normalize_method_type(sidecar.get("method_type"), sidecar.get("source_kind")),
        "status": dataset_sidecar_status_service.normalize_status(sidecar.get("status")),
        "notes": str(sidecar.get("notes") or ""),
        "cumulative": sidecar.get("cumulative"),
        "transposed": sidecar.get("transposed"),
        "calendar": sidecar.get("calendar"),
        "show_subtotal": normalize_show_subtotal(sidecar.get("show_subtotal")),
        "number_format": _normalize_number_format(sidecar.get("number_format") or "0,000"),
        "decimal_places": _normalize_decimal_places(sidecar.get("decimal_places")),
        "formula": dataset_type_formula,
        "calculated": sidecar.get("calculated"),
        "external_links": _normalize_dataset_external_links(sidecar.get("external_links")),
        "internal_links": _normalize_dataset_internal_links(sidecar.get("internal_links")),
        "formula_links": _normalize_dataset_formula_links(sidecar.get("formula_links")),
        "precedents": sidecar.get("precedents") if isinstance(sidecar.get("precedents"), list) else [],
        "dependents": sidecar.get("dependents") if isinstance(sidecar.get("dependents"), list) else [],
        "audit_log": _normalize_dataset_audit_log(sidecar.get("audit_log")),
        "exists": bool(sidecar),
        "path": handle_path,
        "sidecar_path": sidecar_path,
        "values": values,
    }


def _display_view_of_stored_values(
    project_name: str,
    dataset_name: str,
    sidecar: Dict[str, Any],
    csv_path: str,
    values: List[List[Any]],
    target: Tuple[int, int],
) -> Tuple[List[List[Any]], int, int, str, str] | None:
    """Roll a hand-entered dataset's stored rows up to the *target* lengths.

    Returns ``(values, origin_length, development_length, dataset_id, path)``
    for the view, or ``None`` when *target* is the stored shape or the file
    cannot be rolled up to it. The view is registered under
    its own handle, the id of the file a dataset created at that shape would
    have, so the id-addressed grid routes serve the same roll-up; nothing is
    written, because the stored CSV is the only copy of the figures.
    """

    from app_server.services import precedent_cache_service

    display_origin, display_development = target
    if (display_origin, display_development) == stored_lengths(sidecar):
        return None
    if precedent_cache_service.rollup_reason(sidecar, display_origin, display_development):
        return None
    rows = precedent_cache_service.rollup_rows(
        project_name, sidecar, values, display_origin, display_development
    )
    stored_origin, stored_development = stored_lengths(sidecar)
    cumulative = bool(sidecar.get("cumulative", True))
    calendar = bool(sidecar.get("calendar", False))
    view_name = build_dataset_cache_file_name(
        dataset_name,
        str(sidecar.get("data_format") or "Triangle"),
        display_origin,
        display_development,
        cumulative,
        calendar,
    )
    view_path = os.path.join(os.path.dirname(csv_path), f"{view_name}.csv")
    view_id = "arcrhotri_" + hashlib.sha1(view_path.encode("utf-8")).hexdigest()[:16]
    register_rollup_handle(
        view_id,
        {
            "source_path": csv_path,
            "source_origin_length": stored_origin,
            "source_development_length": stored_development,
            "target_origin_length": display_origin,
            "target_development_length": display_development,
            "valuation_months": valuation_months(project_name),
            "cumulative": cumulative,
            "calendar": calendar,
        },
    )
    return rows, display_origin, display_development, view_id, view_path


def _dataset_cache_dir(project_name: str, reserving_class: str) -> str:
    try:
        return config.get_project_dataset_cache_dir(project_name, reserving_class)
    except ValueError as err:
        raise HTTPException(404, str(err))


def _frame_holds_values(frame: pd.DataFrame) -> bool:
    """Whether *frame* holds a number that is neither blank nor zero."""

    if frame.empty:
        return False
    return bool(np.any(np.nan_to_num(frame.to_numpy(dtype="float64"), nan=0.0) != 0.0))


def _stored_csv_holds_values(csv_path: str) -> bool:
    """Whether the CSV at *csv_path* holds a value that is not blank or zero.

    The stored shape of a hand-entered dataset may move only while its file
    holds nothing, which is the same "blank or zero" test the Data tab already
    applies before it lets a length be lowered.
    """

    if not csv_path or not os.path.exists(csv_path):
        return False
    return _frame_holds_values(load_triangle_values(csv_path))


_DATASET_LINK_FIELDS = (
    ("external_links", _normalize_dataset_external_links),
    ("internal_links", _normalize_dataset_internal_links),
    ("formula_links", _normalize_dataset_formula_links),
)


def _record_linked_shape(
    payload: Dict[str, Any],
    existing: Dict[str, Any],
    data_format: str,
    origin_length: int,
    development_length: int,
) -> None:
    """Stamp *payload* with the display its cell links were written against.

    A link names a cell of the grid that was on screen when it was entered,
    so the development width this save's values come at becomes the linked
    one whenever the links themselves changed, and the width the sidecar
    already records stays whenever they did not -- the display may have moved
    on without them. The origin axis is not recorded: ``linked_lengths`` reads
    it from the store, which is the only origin period a link can be entered
    at. A sidecar with no link carries no linked width.
    """

    payload.pop(SIDECAR_LINKED_DEVELOPMENT_FIELD, None)
    if not any(payload.get(field) for field, _normalize in _DATASET_LINK_FIELDS):
        return
    links_changed = any(
        normalize(payload.get(field)) != normalize(existing.get(field))
        for field, normalize in _DATASET_LINK_FIELDS
    )
    linked = (origin_length, development_length) if links_changed else linked_lengths(existing)
    payload.update(linked_length_fields(data_format, *linked))


def scatter_view_into_store(
    project_name: str,
    values_frame: pd.DataFrame,
    *,
    stored_lengths: Tuple[int, int],
    view_lengths: Tuple[int, int],
    cumulative: bool,
) -> pd.DataFrame:
    """Write a coarser development view of a triangle into its stored cells.

    Each cell of the view is the row's cumulative figure at its valuation
    date, so it goes to the stored cell valued there and the rest of the
    store goes to zero, the way ResQ rebuilds a triangle written at a coarse
    display. A coarser origin row has no single cell to write to, so the
    origin lengths must agree.
    """

    if view_lengths[0] != stored_lengths[0]:
        raise HTTPException(400, "Values can be entered only at the stored origin period.")
    return pd.DataFrame(
        scatter_triangle(
            [
                [None if pd.isna(cell) else float(cell) for cell in row]
                for row in values_frame.to_numpy()
            ],
            source_origin_length=stored_lengths[0],
            source_development_length=stored_lengths[1],
            target_origin_length=view_lengths[0],
            target_development_length=view_lengths[1],
            valuation_months=valuation_months(project_name),
            cumulative=bool(cumulative),
        ),
        dtype="float64",
    )


def _save_dataset_sidecar_impl(
    project_name: str,
    reserving_class: str,
    dataset_name: str,
    *,
    dataset_type: str = "",
    instance_name: str = "",
    source_kind: str = "",
    data_format: str = "",
    origin_length: int,
    development_length: int,
    display_at: Tuple[int, int] | None = None,
    stored_development_length: int | None = None,
    stored_values_cleared: bool = False,
    cumulative: bool = True,
    transposed: bool = False,
    calendar: bool = False,
    show_subtotal: bool | None = None,
    number_format: str = "",
    decimal_places: int | None = None,
    origin_labels: List[str] | None = None,
    csv_file: str = "",
    method_type: str = "",
    status: int | None = None,
    notes: str | None = None,
    precedents: List[str] | None = None,
    external_links: List[Any] | None = None,
    internal_links: List[Any] | None = None,
    formula_links: List[Any] | None = None,
    values: List[List[Any]] | None = None,
    mask: List[List[bool]] | None = None,
) -> Dict[str, Any]:
    p, rc, ds = _require_dataset_fields(project_name, reserving_class, dataset_name)
    if origin_length <= 0 or development_length <= 0:
        raise HTTPException(400, "origin_length and development_length must be positive.")
    # ``origin_length`` / ``development_length`` are the shape ``values`` come
    # at and, unless ``display_at`` says otherwise, the display the sidecar
    # records. A link refresh reads and writes at the linked shape while the
    # dataset is shown at another, and passes that display here.
    display_origin_months, display_development_months = display_at or (
        int(origin_length),
        int(development_length),
    )
    normalized_external_links = (
        _normalize_dataset_external_links(external_links, strict=True)
        if external_links is not None
        else None
    )
    normalized_internal_links = (
        _normalize_dataset_internal_links(internal_links, strict=True)
        if internal_links is not None
        else None
    )
    normalized_formula_links = (
        _normalize_dataset_formula_links(formula_links, strict=True)
        if formula_links is not None
        else None
    )

    path = _get_dataset_sidecar_path(p, rc, ds, csv_file=csv_file)
    existing = _read_dataset_sidecar(path)
    existing_precedents = dataset_sidecar_status_service.entry_names(existing.get("precedents"))
    created = str(existing.get("created") or "") if existing else ""
    if not created:
        created = _now_utc_iso()
    user_name = _current_user_name()
    dataset_type_value = str(dataset_type or existing.get("dataset_type") or ds)
    app_calculated, formula = _is_app_calculated_dataset_type(p, dataset_type_value)
    source_kind_value = str(source_kind or existing.get("source_kind") or ("calculated" if app_calculated else "input"))
    data_format_value = str(data_format or existing.get("data_format") or "Triangle")
    is_vector = data_format_value.strip().lower() == "vector"
    method_type_value = dataset_sidecar_status_service.normalize_method_type(method_type or existing.get("method_type"), source_kind_value)
    method_calculated = method_type_value in _METHOD_CALCULATED_TYPES
    number_format_value = _normalize_number_format(number_format or existing.get("number_format") or "0,000")
    # A caller that says nothing about decimal places is not asking for them to
    # change. Methods that republish their output dataset -- Result Selection
    # and Berquist-Sherman among them -- send only the fields they own, and
    # would otherwise reset the display the user chose on every save.
    decimal_places_value = _normalize_decimal_places(
        decimal_places if decimal_places is not None else existing.get("decimal_places")
    )
    show_subtotal_value = normalize_show_subtotal(
        show_subtotal if show_subtotal is not None else existing.get("show_subtotal")
    )
    if values is not None and app_calculated:
        raise HTTPException(400, "Calculated datasets cannot save editable grid values.")

    csv_path = ""
    csv_file_value = str(csv_file or existing.get("csv_file") or "")
    stored_origin_months = int(origin_length)
    stored_development_months = int(development_length)
    if stored_development_length is not None and not is_vector:
        # ResQ's `Stored at` spinner: a triangle that holds nothing may be
        # stored finer than it is shown, at any factor of the display length.
        # The origin store has no control of its own and follows the display
        # one, and a vector has neither, so both ignore the field.
        requested_stored_development = int(stored_development_length)
        if (
            requested_stored_development <= 0
            or int(development_length) % requested_stored_development != 0
        ):
            raise HTTPException(
                400, "The stored development length must be a factor of the development length."
            )
        stored_development_months = requested_stored_development
    superseded_csv_path = ""
    relabel_empty_input = False
    if source_kind_value.strip().casefold() == "input" and existing and not csv_file:
        # A hand-entered dataset's own CSV is its data, and the request's
        # lengths are only the shape it is displayed at, so the stored shape
        # and the file it names both stay put. The one exception is a dataset
        # whose file holds nothing: nothing is lost by relabelling it, so the
        # shape asked for becomes the stored one and the old file goes.
        existing_origin, existing_development = stored_lengths(existing)
        if is_vector:
            existing_shape = (existing_origin,)
            requested_shape = (stored_origin_months,)
            display_shape = (int(origin_length),)
        else:
            existing_shape = (existing_origin, existing_development)
            requested_shape = (stored_origin_months, stored_development_months)
            display_shape = (int(origin_length), int(development_length))
        if all(months > 0 for months in existing_shape) and (
            existing_shape != requested_shape or display_shape != existing_shape
        ):
            existing_csv_file = os.path.basename(str(existing.get("csv_file") or "").strip())
            existing_csv_path = (
                os.path.join(_dataset_cache_dir(p, rc), existing_csv_file)
                if existing_csv_file
                else ""
            )
            # A client that set every value to 0 and reshaped the grid before
            # entering the values it sends is replacing the file outright, so
            # the values the old file still holds do not fix the shape.
            if _stored_csv_holds_values(existing_csv_path) and not stored_values_cleared:
                if not is_vector and stored_development_length is not None and (
                    stored_development_months != existing_development
                ):
                    raise HTTPException(
                        400,
                        "The stored development length cannot be changed while the dataset "
                        "holds values.",
                    )
                if values is not None and display_shape != existing_shape:
                    # ResQ relaxes the development axis alone: values entered at
                    # a coarser development view are scattered into the stored
                    # cells below, while a coarser origin period has no single
                    # valuation date to write to and a vector neither.
                    development_is_a_view_of_the_store = (
                        not is_vector
                        and existing_development > 0
                        and int(development_length) % existing_development == 0
                    )
                    if not development_is_a_view_of_the_store:
                        shape_text = (
                            f"period length {existing_origin}"
                            if is_vector
                            else f"origin length {existing_origin} and development length {existing_development}"
                        )
                        raise HTTPException(
                            422,
                            f"Dataset '{ds}' stores its values at {shape_text}. Values can be entered "
                            "only at the stored period; set the lengths back to edit.",
                        )
                    if int(origin_length) != existing_origin:
                        raise HTTPException(
                            400, "Values can be entered only at the stored origin period."
                        )
                stored_origin_months, stored_development_months = existing_origin, existing_development
            elif existing_shape != requested_shape:
                relabel_empty_input = True
                superseded_csv_path = existing_csv_path
    if values is not None or relabel_empty_input:
        # The file is the dataset's data, so it is named for the shape it is
        # written at -- the stored one, which is the display shape unless this
        # save asked for a finer store.
        csv_stem = build_dataset_cache_file_name(
            ds,
            data_format_value,
            stored_origin_months,
            stored_development_months,
            cumulative,
            calendar,
        )
        csv_file_value = f"{csv_stem}.csv"
        csv_path = os.path.join(_dataset_cache_dir(p, rc), csv_file_value)

    action_value = "Update" if existing else "Insert"
    updated_at = _now_utc_iso()
    payload = {
        **existing,
        "dataset_name": ds,
        "dataset_type": dataset_type_value,
        "reserving_class": rc,
        "project_name": p,
        "source_kind": source_kind_value,
        "data_format": data_format_value,
        "transposed": bool(transposed),
        "show_subtotal": show_subtotal_value,
        "number_format": number_format_value,
        "decimal_places": decimal_places_value,
        "csv_file": csv_file_value,
        "calculated": True if app_calculated or method_calculated else (False if values is not None else existing.get("calculated")),
        "method_type": method_type_value,
        "notes": str(notes if notes is not None else existing.get("notes") or ""),
        "created": created,
        "modified_by": user_name,
        "updated_at": updated_at,
    }
    if is_vector:
        payload["period_length"] = display_origin_months
        for obsolete_key in (
            "origin_length",
            "development_length",
            "development_count",
            "cumulative",
            "calendar",
            "stored_origin_length",
            "stored_development_length",
            SIDECAR_LINKED_DEVELOPMENT_FIELD,
        ):
            payload.pop(obsolete_key, None)
    else:
        payload["origin_length"] = display_origin_months
        payload["development_length"] = display_development_months
        payload["cumulative"] = bool(cumulative)
        payload["calendar"] = bool(calendar)
        payload.pop("period_length", None)
        payload.pop("stored_period_length", None)
    if values is not None or csv_file or relabel_empty_input:
        # This save names the CSV -- it writes one from ``values``, relabels an
        # empty one, or the caller published one and passed its name -- so the
        # lengths that file is written at are the shape the values are stored
        # at. A settings-only save leaves the CSV alone, and the stored shape
        # the sidecar already records travels with it.
        payload.update(
            stored_length_fields(data_format_value, stored_origin_months, stored_development_months)
        )
    if origin_labels is not None:
        payload["origin_labels"] = _normalize_origin_labels(origin_labels)
    if normalized_external_links is not None:
        payload["external_links"] = normalized_external_links
    if normalized_internal_links is not None:
        payload["internal_links"] = normalized_internal_links
    if normalized_formula_links is not None:
        payload["formula_links"] = normalized_formula_links
    _record_linked_shape(payload, existing, data_format_value, int(origin_length), int(development_length))
    _require_disjoint_dataset_link_targets(
        _normalize_dataset_external_links(payload.get("external_links")),
        _normalize_dataset_internal_links(payload.get("internal_links")),
        _normalize_dataset_formula_links(payload.get("formula_links")),
    )
    _append_dataset_audit_entry(payload, action_value, event_date=updated_at, user_name=user_name)
    payload.pop("instance_name", None)
    payload.pop("dataset_type_name", None)
    from app_server.services import calculated_dataset_service

    calculated_dataset_service.apply_sidecar_graph_fields(payload, p, dataset_type_value)
    if precedents is not None:
        # One entry shape for every method type (``arcrho_api.sidecar_core_contract``):
        # a Result Selection's precedents are entries like everyone else's, so the
        # dependents written on the far side of the same link match them.
        payload["precedents"] = dataset_sidecar_status_service.name_entries(precedents)
    elif method_type_value == dataset_sidecar_status_service.METHOD_TYPE_RESULT_SELECTION:
        payload["precedents"] = []
    elif method_type_value != dataset_sidecar_status_service.METHOD_TYPE_NONE and existing_precedents:
        payload["precedents"] = dataset_sidecar_status_service.name_entries(existing_precedents)
    force_status = status
    if force_status is None and method_type_value != dataset_sidecar_status_service.METHOD_TYPE_NONE:
        force_status = existing.get("status")
    dataset_sidecar_status_service.apply_status_fields(
        payload,
        p,
        rc,
        ds,
        path=path,
        method_type=method_type_value,
        force_status=force_status,
    )
    if values is not None or relabel_empty_input:
        stored_shape_is_display = (stored_origin_months, stored_development_months) == (
            int(origin_length),
            int(development_length),
        )
        values_frame = _dataset_values_to_frame(values, mask) if values is not None else None
        if values_frame is not None and stored_shape_is_display:
            df = values_frame
        elif values_frame is not None and not is_vector:
            # A coarser development view of the store: each entered cell is the
            # row's cumulative figure at its valuation date, so it goes to the
            # stored cell valued there and the rest of the store goes to zero,
            # the way ResQ rebuilds a triangle written at a coarse display.
            df = scatter_view_into_store(
                p,
                values_frame,
                stored_lengths=(stored_origin_months, stored_development_months),
                view_lengths=(int(origin_length), int(development_length)),
                cumulative=bool(cumulative),
            )
        else:
            # Values arrive at the display shape, which a store finer than the
            # display cannot hold as it stands. Only a dataset that is empty
            # can be stored finer, so those values carry nothing and the file
            # is written as an empty grid at the stored shape instead.
            if values_frame is not None and _frame_holds_values(values_frame):
                raise HTTPException(
                    422,
                    f"Dataset '{ds}' is stored at development length {stored_development_months}. "
                    "Values can be entered only at the stored period; set the lengths back to edit.",
                )
            origin_count, development_count, empty_mask = _empty_dataset_geometry_from_general_settings(
                p, stored_origin_months, stored_development_months
            )
            df = _empty_dataset_values(data_format_value, origin_count, development_count, empty_mask)
        try:
            _write_dataset_csv_and_sidecar(df, csv_path, path, payload)
        except PermissionError:
            raise HTTPException(423, "Dataset cache CSV is locked or inaccessible.")
        except OSError as err:
            raise HTTPException(500, f"Failed to write dataset cache CSV: {str(err)}")
        if superseded_csv_path and os.path.normcase(superseded_csv_path) != os.path.normcase(csv_path):
            # The relabelled dataset is stored at its new shape now, so the
            # file it was stored at before is not the dataset's data any more.
            try:
                os.remove(superseded_csv_path)
            except OSError:
                pass
    else:
        _write_dataset_sidecar_payload(path, payload)
    ds_id = ""
    file_mtime = None
    if csv_path:
        ds_id = "arcrhotri_" + hashlib.sha1(csv_path.encode("utf-8")).hexdigest()[:16]
        config.DATASETS[ds_id] = csv_path
        try:
            file_mtime = os.stat(csv_path).st_mtime
        except OSError:
            file_mtime = None
    if precedents is not None:
        dataset_sidecar_status_service.update_precedent_dependents(
            p,
            rc,
            ds,
            existing_precedents,
            precedents,
            require_new_precedents=(
                method_type_value == dataset_sidecar_status_service.METHOD_TYPE_RESULT_SELECTION
            ),
        )
    elif method_type_value == dataset_sidecar_status_service.METHOD_TYPE_NONE:
        # ArcRho cell links are instance-level graph edges: the datasets this
        # save's links read gain (or lose) a dependents entry naming this
        # dataset, so the dependent-propagation walk and the delete check see
        # the link the same way they see a formula edge.
        own_key = _canon_dataset_name(ds)
        old_link_names = [
            name
            for name in link_precedent_names(
                _normalize_dataset_internal_links(existing.get("internal_links")),
                _normalize_dataset_formula_links(existing.get("formula_links")),
            )
            if _canon_dataset_name(name) != own_key
        ]
        new_link_names = [
            name
            for name in link_precedent_names(
                payload.get("internal_links"),
                payload.get("formula_links"),
            )
            if _canon_dataset_name(name) != own_key
        ]
        if old_link_names or new_link_names:
            dataset_sidecar_status_service.update_precedent_dependents(
                p,
                rc,
                ds,
                old_link_names,
                new_link_names,
            )
    unreviewed_precedents = dataset_sidecar_status_service.review_needed_precedent_names(
        p,
        rc,
        payload.get("precedents"),
    ) if method_type_value != dataset_sidecar_status_service.METHOD_TYPE_NONE else []
    status_updates = dataset_sidecar_status_service.refresh_method_statuses_for_dependents(p, rc, [ds])

    calculated_updates = dependent_propagation_service.enqueue_save_propagation(
        p,
        rc,
        [dependent_propagation_service.changed_root(ds, dataset_type_value)],
    )
    index_error = ""
    try:
        dataset_instance_index_service.rebuild_index(p, rc)
    except Exception as err:
        index_error = str(err)
    # The rebuild above just rewrote the index, so both chip rows read it once
    # rather than opening every neighbour's sidecar in turn.
    saved_calculation_map = _dataset_type_calculation_map(p)
    saved_index_map = _dataset_index_entry_map(p, rc)
    saved_stored_lengths = stored_lengths(payload)
    return {
        "ok": True,
        "project_name": p,
        "reserving_class": rc,
        "dataset_name": ds,
        "dataset_type": payload["dataset_type"],
        "instance_name": ds,
        "data_format": payload["data_format"],
        "period_length": payload.get("period_length") if is_vector else None,
        "origin_length": payload.get("period_length") if is_vector else payload["origin_length"],
        "development_length": payload.get("period_length") if is_vector else payload["development_length"],
        # The stored pair travels back with the save for the same reason the
        # load carries it: the caller has just been told the display shape and
        # needs the shape underneath it to know whether that display is a
        # roll-up. A save of a still-empty dataset moves it, so a caller cannot
        # assume the pair it sent in.
        "stored_period_length": saved_stored_lengths[0] if is_vector else None,
        "stored_origin_length": saved_stored_lengths[0],
        "stored_development_length": saved_stored_lengths[1],
        # And the linked pair, which this save may have left where it was while
        # the display moved on.
        "linked_period_length": linked_lengths(payload)[0] if is_vector else None,
        "linked_origin_length": linked_lengths(payload)[0],
        "linked_development_length": linked_lengths(payload)[1],
        "origin_labels": _normalize_origin_labels(payload.get("origin_labels")),
        "cumulative": payload.get("cumulative"),
        "transposed": payload["transposed"],
        "calendar": payload.get("calendar"),
        "show_subtotal": payload["show_subtotal"],
        "number_format": payload["number_format"],
        "decimal_places": payload["decimal_places"],
        "csv_file": payload["csv_file"],
        "source_kind": payload["source_kind"],
        "method_type": payload["method_type"],
        "status": payload["status"],
        "notes": payload["notes"],
        "external_links": _normalize_dataset_external_links(payload.get("external_links")),
        "internal_links": _normalize_dataset_internal_links(payload.get("internal_links")),
        "formula_links": _normalize_dataset_formula_links(payload.get("formula_links")),
        "precedents": _sidecar_graph_entries(
            p,
            rc,
            payload.get("precedents"),
            include_method_type=True,
            calculation_map=saved_calculation_map,
            index_map=saved_index_map,
        ),
        "dependents": _sidecar_graph_entries(
            p,
            rc,
            payload.get("dependents"),
            include_formula=True,
            calculation_map=saved_calculation_map,
            index_map=saved_index_map,
        ),
        # The Details tab renders the same three rows — formula, precedents,
        # dependents — from whichever of the two answers it holds, so a caller
        # that just saved never has to load the sidecar back to fill this in.
        # It is the value the save already derived, not a second reading of it.
        "formula": str(formula or ""),
        "updated_at": payload["updated_at"],
        "audit_log": payload["audit_log"],
        "path": path,
        "csv_path": csv_path,
        "ds_id": ds_id,
        "file_mtime": file_mtime,
        "calculated_updates": calculated_updates,
        "propagation_ok": bool(calculated_updates and calculated_updates.get("ok")),
        "status_updates": status_updates,
        "unreviewed_precedents": unreviewed_precedents,
        "unreviewed_precedent_count": len(unreviewed_precedents),
        "index_ok": not index_error,
        "index_error": index_error,
    }


def save_dataset_sidecar(
    project_name: str,
    reserving_class: str,
    dataset_name: str,
    **kwargs: Any,
) -> Dict[str, Any]:
    # Dependent propagation runs on ArcRho Engine; block the save before any
    # write when no live Engine can pick the job up or another walk is still
    # rewriting this reserving class.
    dependent_propagation_service.require_reserving_class_writable(
        project_name, reserving_class
    )
    with dataset_sidecar_status_service.reserving_class_io_lock(project_name, reserving_class):
        return _save_dataset_sidecar_impl(
            project_name,
            reserving_class,
            dataset_name,
            **kwargs,
        )


def save_propagation_roots(
    project_name: str,
    reserving_class: str,
    dataset_name: str,
    *,
    dataset_type: str = "",
    csv_file: str = "",
    **_ignored: Any,
) -> List[Tuple[str, str]]:
    """Return the changed roots ``save_dataset_sidecar`` would propagate from.

    The two-step save plans the dependent closure before anything is written,
    so this mirrors ``_save_dataset_sidecar_impl``'s root exactly, including
    its fallback to the existing sidecar's ``dataset_type`` and then to the
    instance name.
    """

    p, rc, ds = _require_dataset_fields(project_name, reserving_class, dataset_name)
    existing = _read_dataset_sidecar(_get_dataset_sidecar_path(p, rc, ds, csv_file=csv_file))
    return [(ds, str(dataset_type or existing.get("dataset_type") or ds))]


def _save_dataset_notes_impl(project_name: str, reserving_class: str, dataset_name: str, notes: str) -> Dict[str, Any]:
    """Update notes in the owning dataset sidecar; no standalone notes file exists."""
    p, rc, ds = _require_dataset_fields(project_name, reserving_class, dataset_name)
    path = _get_dataset_sidecar_path(p, rc, ds)
    with _dataset_sidecar_write_lock(path):
        payload = _read_dataset_sidecar(path)
        if not payload:
            raise HTTPException(404, f"Dataset sidecar not found for '{ds}'.")
        payload["notes"] = str(notes if notes is not None else "")
        payload["modified_by"] = _current_user_name()
        payload["updated_at"] = _now_utc_iso()
        _write_dataset_sidecar_payload(path, payload)
    return {
        "ok": True,
        "project_name": p,
        "reserving_class": rc,
        "dataset_name": ds,
        "notes": payload["notes"],
        "updated_at": payload["updated_at"],
        "path": path,
    }


def save_dataset_notes(project_name: str, reserving_class: str, dataset_name: str, notes: str) -> Dict[str, Any]:
    with dataset_sidecar_status_service.reserving_class_io_lock(project_name, reserving_class):
        return _save_dataset_notes_impl(project_name, reserving_class, dataset_name, notes)

def _patch_dataset_impl(
    ds_id: str, items: list, file_mtime: float | None = None
) -> Dict[str, Any] | None:
    path = config.DATASETS.get(ds_id)
    if not path or not os.path.exists(path):
        return None

    st = os.stat(path)
    if file_mtime is not None and abs(st.st_mtime - file_mtime) > 1e-6:
        return {"conflict": True}

    df = load_triangle_values(path)
    n_origin, n_dev = df.shape
    mask = _dataset_patch_mask(path, n_origin, n_dev)

    applied = 0
    rejected: List[Dict[str, Any]] = []

    for it in items:
        r, c = it.r, it.c
        if r >= n_origin or c >= n_dev:
            rejected.append({"r": r, "c": c, "reason": "out_of_range"})
            continue
        if not mask[r, c]:
            rejected.append({"r": r, "c": c, "reason": "outside_triangle"})
            continue

        df.iat[r, c] = np.nan if it.value is None else float(it.value)
        applied += 1

    if applied == 0:
        return {
            "ok": True,
            "applied": 0,
            "rejected": rejected,
            "mtime": st.st_mtime,
            "calculated_updates": None,
            "propagation_ok": True,
        }

    sidecar_path = ""
    sidecar_payload: Dict[str, Any] = {}
    if applied > 0:
        sidecar_path = dataset_instance_index_service._dataset_sidecar_path_for_cached_csv(path)
        sidecar_payload = _read_dataset_sidecar(sidecar_path)
        if sidecar_payload:
            # Written back below, so a sidecar carried in by a duplication stops
            # naming the project it was copied from once the grid is saved.
            owning_project = _dataset_owning_project_name(path, sidecar_payload)
            if owning_project:
                sidecar_payload["project_name"] = owning_project
            audit_at = _now_utc_iso()
            user_name = _current_user_name()
            sidecar_payload["updated_at"] = audit_at
            sidecar_payload["modified_by"] = user_name
            _append_dataset_audit_entry(sidecar_payload, "Update", event_date=audit_at, user_name=user_name)
            dataset_name = str(sidecar_payload.get("dataset_name") or sidecar_payload.get("dataset_type") or "").strip()
            if dataset_name:
                dataset_sidecar_status_service.apply_status_fields(
                    sidecar_payload,
                    str(sidecar_payload.get("project_name") or ""),
                    str(sidecar_payload.get("reserving_class") or ""),
                    dataset_name,
                    path=sidecar_path,
                )
    if sidecar_payload:
        _write_dataset_csv_and_sidecar(df, path, sidecar_path, sidecar_payload)
    else:
        atomic_write_csv(df, path)
    st2 = os.stat(path)
    if sidecar_payload:
        try:
            dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
                str(sidecar_payload.get("project_name") or ""),
                str(sidecar_payload.get("reserving_class") or ""),
                [sidecar_payload.get("dataset_name") or sidecar_payload.get("dataset_type")],
            )
        except Exception:
            pass
    calculated_updates: Dict[str, Any] | None = None
    if sidecar_payload:
        project_value = str(sidecar_payload.get("project_name") or "").strip()
        reserving_value = str(sidecar_payload.get("reserving_class") or "").strip()
        dataset_value = str(
            sidecar_payload.get("dataset_name")
            or sidecar_payload.get("dataset_type")
            or ""
        ).strip()
        if project_value and reserving_value and dataset_value:
            calculated_updates = dependent_propagation_service.enqueue_save_propagation(
                project_value,
                reserving_value,
                [
                    dependent_propagation_service.changed_root(
                        dataset_value,
                        str(sidecar_payload.get("dataset_type") or ""),
                    )
                ],
            )
        else:
            calculated_updates = {
                "ok": False,
                "skipped": True,
                "reason": "missing_sidecar_context",
            }
    else:
        calculated_updates = {
            "ok": False,
            "skipped": True,
            "reason": "missing_sidecar_context",
        }

    return {
        "ok": True,
        "applied": applied,
        "rejected": rejected,
        "mtime": st2.st_mtime,
        "calculated_updates": calculated_updates,
        "propagation_ok": bool(calculated_updates and calculated_updates.get("ok")),
    }


def patch_dataset(
    ds_id: str, items: list, file_mtime: float | None = None
) -> Dict[str, Any] | None:
    path = config.DATASETS.get(ds_id)
    if not path or not os.path.exists(path):
        return _patch_dataset_impl(ds_id, items, file_mtime)
    sidecar_path = dataset_instance_index_service._dataset_sidecar_path_for_cached_csv(path)
    sidecar = _read_dataset_sidecar(sidecar_path)
    project_name = _dataset_owning_project_name(path, sidecar)
    reserving_class = str(sidecar.get("reserving_class") or "").strip()
    if not project_name or not reserving_class:
        # Dependent propagation runs on ArcRho Engine; block the grid save
        # before any write when no live Engine instance can pick the job up.
        dependent_propagation_service.require_engine_available()
        return _patch_dataset_impl(ds_id, items, file_mtime)
    # Block the grid save before any write when no live Engine can pick the
    # job up or another walk is still rewriting this reserving class.
    dependent_propagation_service.require_reserving_class_writable(
        project_name, reserving_class
    )
    with dataset_sidecar_status_service.reserving_class_io_lock(project_name, reserving_class):
        return _patch_dataset_impl(ds_id, items, file_mtime)
