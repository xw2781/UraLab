"""ArcRho runtime request operations."""
from __future__ import annotations

import getpass
import hashlib
import json
import os
import re
import uuid
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List

import pandas as pd
from fastapi import HTTPException
from arcrho_api.dataset_index_contract import (
    INDEX_FILE_NAME as DATASET_INDEX_FILE_NAME,
    index_rebuild_reason,
    scan_folder_signature,
)
from arcrho_api.engine_dataset_sidecar_contract import build_engine_dataset_sidecar
from arcrho_api.dataset_display_contract import normalize_show_subtotal
from arcrho_api.field_mapping_contract import (
    DATE_ROLE_DEVELOPMENT,
    DATE_ROLE_ORIGIN,
)
from arcrho_api.sidecar_core_contract import stored_length_fields
from arcrho_api.timestamps import utc_now_text, format_persisted_timestamp
from arcrho_api.triangle_rollup import rollup_reason, rollup_triangle
from arcrho_engine_calculation_contract import OUTPUT_VARIANT_TEMPORARY_VIEW

from app_server import config
from app_server.helpers import (
    _canon_dataset_name,
    sanitize_dataset_file_name,
    set_data_path_like_vba,
)
from app_server.services import (
    dataset_instance_index_service,
    dataset_number_format_service,
    dataset_sidecar_status_service,
    engine_calculation_service,
    file_read_cache,
    project_settings_service,
    runtime_cache_provenance_service,
    user_identity_service,
)
from app_server.services.data_processing_rules_service import (
    get_processing_config_hash,
    get_processing_provenance,
    is_imported_snapshot_payload,
)

def _pair_value(pairs: list, key: str) -> str:
    key_l = key.strip().lower()
    for pair_key, pair_value in pairs:
        if str(pair_key or "").strip().lower() == key_l:
            return str(pair_value or "").strip()
    return ""


def _dataset_sidecar_path(data_path: str, pairs: list) -> str:
    dataset_name = _pair_value(pairs, "InstanceName") or _pair_value(pairs, "DatasetName") or _pair_value(pairs, "TriangleName")
    dataset_file = sanitize_dataset_file_name(dataset_name)
    dataset_dir = os.path.dirname(data_path)
    if os.path.basename(dataset_dir).lower() == config.DATASET_CACHE_DIR.lower():
        sidecar_dir = os.path.join(os.path.dirname(dataset_dir), config.DATASET_SIDECAR_DIR)
    else:
        sidecar_dir = os.path.join(dataset_dir, config.DATASET_SIDECAR_DIR)
    return os.path.join(sidecar_dir, f"{dataset_file}.json")


def _runtime_cache_provenance_path(data_path: str) -> str:
    return runtime_cache_provenance_service.provenance_path(data_path)


def _utc_timestamp_from_stat(value: float) -> str:
    return format_persisted_timestamp(datetime.fromtimestamp(value, timezone.utc))


def _pair_int_value(pairs: list, key: str, default: int) -> int:
    try:
        return int(_pair_value(pairs, key) or default)
    except (TypeError, ValueError):
        return default


def _pair_bool_value(pairs: list, key: str, default: bool) -> bool:
    raw = _pair_value(pairs, key)
    if not raw:
        return default
    text = raw.strip().lower()
    if text in {"true", "yes", "1"}:
        return True
    if text in {"false", "no", "0"}:
        return False
    return default


def _clean_cache_text(value: Any) -> str:
    return str(value or "").strip()


def _cache_text_matches(left: Any, right: Any) -> bool:
    return _clean_cache_text(left) == _clean_cache_text(right)


def _cache_payload_name_matches(payload: Dict[str, Any], expected_name: str) -> bool:
    if not expected_name:
        return False
    return _cache_text_matches(payload.get("dataset_name"), expected_name)


def _safe_read_json(path: str) -> Dict[str, Any]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def _request_dataset_name(pairs: list) -> str:
    return _pair_value(pairs, "InstanceName") or _pair_value(pairs, "DatasetName") or _pair_value(pairs, "TriangleName")


ProcessingHashGetter = Callable[[], str]
FileFingerprintGetter = Callable[[str], Dict[str, Any]]
CalculatedValidationMemo = Dict[str, bool]


def _processing_hash_getter(pairs: list) -> ProcessingHashGetter:
    project_name = _pair_value(pairs, "ProjectName")
    current_hash: str | None = None

    def get_current_hash() -> str:
        nonlocal current_hash
        if current_hash is not None:
            return current_hash
        if not project_name:
            raise HTTPException(422, "ProjectName is required to validate a generated dataset cache.")
        try:
            current_hash = get_processing_config_hash(project_name)
        except HTTPException:
            raise
        except Exception as error:
            raise HTTPException(
                503,
                (
                    f"Unable to validate the processing settings for '{project_name}'. "
                    "The existing dataset cache was left unchanged."
                ),
            ) from error
        return current_hash

    return get_current_hash


def _processing_hash_matches(
    processing: Any,
    pairs: list,
    processing_hash_getter: ProcessingHashGetter | None = None,
) -> bool:
    if not isinstance(processing, dict):
        return False
    stored_hash = _clean_cache_text(processing.get("config_hash"))
    project_name = _pair_value(pairs, "ProjectName")
    if not stored_hash or not project_name:
        return False
    get_current_hash = processing_hash_getter or _processing_hash_getter(pairs)
    return stored_hash == get_current_hash()


def _file_fingerprint_getter() -> FileFingerprintGetter:
    fingerprints: Dict[str, Dict[str, Any]] = {}

    def get_fingerprint(path: str) -> Dict[str, Any]:
        key = os.path.normcase(os.path.abspath(path))
        if key not in fingerprints:
            fingerprints[key] = runtime_cache_provenance_service.file_fingerprint(path)
        return fingerprints[key]

    return get_fingerprint


def _stored_file_fingerprint_matches(
    payload: Dict[str, Any],
    path: str,
    get_fingerprint: FileFingerprintGetter,
    *,
    prefix: str = "",
    verify_content: bool = True,
) -> bool:
    expected_size = payload.get(f"{prefix}size")
    expected_mtime_ns = payload.get(f"{prefix}mtime_ns")
    expected_sha256 = _clean_cache_text(payload.get(f"{prefix}sha256")).lower()
    if (
        expected_size in (None, "")
        or expected_mtime_ns in (None, "")
        or not expected_sha256
    ):
        return False
    current_stat = os.stat(path)
    try:
        if (
            int(current_stat.st_size) != int(expected_size)
            or int(current_stat.st_mtime_ns) != int(expected_mtime_ns)
        ):
            return False
    except (TypeError, ValueError):
        return False
    if not verify_content:
        return True
    current = get_fingerprint(path)
    try:
        return _clean_cache_text(current.get("sha256")).lower() == expected_sha256
    except (TypeError, ValueError):
        return False


def _calculated_precedent_sidecar(
    precedent: Dict[str, Any],
    path: str,
    pairs: list,
) -> Dict[str, Any]:
    dataset_name = _clean_cache_text(
        precedent.get("dataset_name")
        or precedent.get("dataset_type")
    )
    if not dataset_name:
        return {}
    sidecar_path = _dataset_sidecar_path(
        path,
        [("InstanceName", dataset_name)],
    )
    try:
        with open(sidecar_path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except FileNotFoundError:
        return {}
    except json.JSONDecodeError:
        return {}
    except OSError as error:
        project_name = _pair_value(pairs, "ProjectName") or "(unknown)"
        raise HTTPException(
            503,
            (
                f"Unable to validate a calculated dataset dependency for '{project_name}'. "
                "The existing calculated cache was left unchanged."
            ),
        ) from error
    return payload if isinstance(payload, dict) else {}


def _calculated_precedent_request_pairs(
    pairs: list,
    precedent: Dict[str, Any],
    sidecar: Dict[str, Any],
) -> list:
    dataset_type = _clean_cache_text(
        precedent.get("dataset_type")
        or sidecar.get("dataset_type")
        or sidecar.get("dataset_name")
    )
    instance_name = _clean_cache_text(
        precedent.get("dataset_name")
        or sidecar.get("dataset_name")
        or dataset_type
    )
    data_format = _clean_cache_text(
        precedent.get("data_format")
        or sidecar.get("data_format")
        or "Triangle"
    )
    return _dependency_request_pairs(
        pairs,
        dataset_type,
        data_format,
        instance_name=instance_name,
        settings=_dependency_cache_settings(
            precedent,
            data_format,
            _clean_cache_text(precedent.get("path")),
        ),
    )


def _reserving_class_state_trusted(
    reserving_class_dir: str,
    memo: CalculatedValidationMemo | None,
) -> bool:
    """Return True when the class's persisted index is current evidence.

    Dependent propagation is a single locked Engine-hosted job (business-logic
    contract rule 15), so a persisted ``index.json`` whose ``folder_signature``
    still matches a fresh folder listing proves no dataset, method, or sidecar
    file in the reserving class changed since the last completed rebuild. The
    check costs one index read plus three directory listings; it never raises,
    and any doubt falls back to the deep per-precedent fingerprint walk.
    """
    memo_key = "class-state-trusted::" + os.path.normcase(
        os.path.abspath(reserving_class_dir)
    )
    if memo is not None and memo_key in memo:
        return memo[memo_key]
    trusted = False
    try:
        index_path = os.path.join(reserving_class_dir, DATASET_INDEX_FILE_NAME)
        with open(index_path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
        if isinstance(data, dict):
            scan = scan_folder_signature(reserving_class_dir)
            trusted = not index_rebuild_reason(
                data,
                expected_folder_signature=scan.signature,
            )
    except (OSError, json.JSONDecodeError, UnicodeDecodeError, ValueError):
        trusted = False
    if memo is not None:
        memo[memo_key] = trusted
    return trusted


def _calculated_dependencies_match(
    payload: Dict[str, Any],
    pairs: list,
    data_path: str,
    *,
    processing_hash_getter: ProcessingHashGetter | None = None,
    file_fingerprint_getter: FileFingerprintGetter | None = None,
    validation_memo: CalculatedValidationMemo | None = None,
    validation_stack: set[str] | None = None,
    verify_content: bool = True,
) -> bool:
    cache_key = "::".join([
        os.path.normcase(os.path.abspath(data_path)),
        _canon_dataset_name(_pair_value(pairs, "ProjectName")),
        _canon_dataset_name(_pair_value(pairs, "Path")),
        _canon_dataset_name(
            _pair_value(pairs, "DatasetName")
            or _pair_value(pairs, "TriangleName")
        ),
        _canon_dataset_name(_pair_value(pairs, "InstanceName")),
        _canon_dataset_name(_pair_value(pairs, "Function")),
        "content" if verify_content else "metadata",
    ])
    memo = validation_memo if validation_memo is not None else {}
    if cache_key in memo:
        return memo[cache_key]
    active_stack = set(validation_stack or set())
    if cache_key in active_stack:
        memo[cache_key] = False
        return False
    active_stack.add(cache_key)

    precedents = payload.get("precedents")
    project_name = _pair_value(pairs, "ProjectName") or "(unknown)"
    dataset_type = _clean_cache_text(
        payload.get("dataset_type")
        or _pair_value(pairs, "DatasetName")
        or _pair_value(pairs, "TriangleName")
    )
    from app_server.services import calculated_dataset_service

    contract = calculated_dataset_service.calculated_dataset_contract(
        project_name,
        dataset_type,
    )
    if contract is None:
        memo[cache_key] = False
        return False
    current_formula = _clean_cache_text(contract.get("formula"))
    stored_names = [
        _canon_dataset_name(item.get("dataset_name"))
        for item in precedents or []
        if isinstance(item, dict) and _canon_dataset_name(item.get("dataset_name"))
    ]
    current_names = [
        _canon_dataset_name(name)
        for name in contract.get("precedents") or []
        if _canon_dataset_name(name)
    ]
    if stored_names != current_names:
        memo[cache_key] = False
        return False
    # The sidecar names the precedents and nothing else: it is shared,
    # location-independent data. What this cache was built from -- the formula
    # and each input's file and fingerprint -- is the technical record beside
    # the CSV, and a cache without one has no evidence to offer.
    record = _calculated_cache_record(data_path, pairs)
    if record is None or _clean_cache_text(record.get("formula")) != current_formula:
        memo[cache_key] = False
        return False
    if not current_names:
        memo[cache_key] = True
        return True

    recorded_dependencies = {
        _canon_dataset_name(item.get("dataset_type") or item.get("dataset_name")): item
        for item in record.get("dependencies") or []
        if isinstance(item, dict)
    }
    dataset_folder = os.path.dirname(data_path)
    if (
        os.path.basename(dataset_folder).lower()
        == config.TEMPORARY_VIEW_DATASET_CACHE_DIR.lower()
    ):
        dataset_folder = os.path.dirname(dataset_folder)
    method_folder = os.path.join(
        os.path.dirname(dataset_folder),
        config.METHOD_DATA_DIR,
    )
    if not _path_is_within_folder(data_path, dataset_folder):
        memo[cache_key] = False
        return False
    if (
        dataset_sidecar_status_service.normalize_status(payload.get("status"))
        == dataset_sidecar_status_service.STATUS_CURRENT
        and _reserving_class_state_trusted(os.path.dirname(dataset_folder), memo)
    ):
        # Open fast path: a current sidecar status inside a class whose folder
        # signature still matches the persisted index needs no ancestry walk.
        memo[cache_key] = True
        return True
    current_precedent_contracts = (
        contract.get("precedent_contracts")
        if isinstance(contract.get("precedent_contracts"), dict)
        else {}
    )
    get_fingerprint = file_fingerprint_getter or _file_fingerprint_getter()
    for name in current_names:
        precedent = recorded_dependencies.get(name)
        if not isinstance(precedent, dict):
            memo[cache_key] = False
            return False
        path = _clean_cache_text(precedent.get("path"))
        if not path or not (
            _path_is_within_folder(path, dataset_folder)
            or _path_is_within_folder(path, method_folder)
        ):
            memo[cache_key] = False
            return False
        try:
            if not _stored_file_fingerprint_matches(
                precedent,
                path,
                get_fingerprint,
                verify_content=verify_content,
            ):
                memo[cache_key] = False
                return False
            input_path = _clean_cache_text(precedent.get("input_path"))
            if input_path:
                if not _path_is_within_folder(input_path, dataset_folder):
                    memo[cache_key] = False
                    return False
                if not _stored_file_fingerprint_matches(
                    precedent,
                    input_path,
                    get_fingerprint,
                    prefix="input_",
                    verify_content=verify_content,
                ):
                    memo[cache_key] = False
                    return False

            stored_source_kind = _clean_cache_text(precedent.get("source_kind")).lower()
            dependency_sidecar: Dict[str, Any] = {}
            if stored_source_kind in {"", "engine", "calculated"}:
                dependency_sidecar = _calculated_precedent_sidecar(
                    precedent,
                    path,
                    pairs,
                )
            source_kind = (
                stored_source_kind
                or _clean_cache_text(dependency_sidecar.get("source_kind")).lower()
            )
            dependency_definition = current_precedent_contracts.get(name)
            if isinstance(dependency_definition, dict):
                current_data_format = _clean_cache_text(
                    dependency_definition.get("data_format")
                ).lower()
                stored_data_format = _clean_cache_text(
                    precedent.get("data_format")
                    or dependency_sidecar.get("data_format")
                    or _cache_path_data_format(path)
                ).lower()
                if (
                    current_data_format
                    and stored_data_format
                    and stored_data_format != current_data_format
                ):
                    memo[cache_key] = False
                    return False
            if (
                isinstance(dependency_definition, dict)
                and stored_source_kind in {"", "input", "engine", "calculated"}
            ):
                if dependency_definition.get("generated"):
                    source_kind = "engine"
                elif (
                    dependency_definition.get("calculated")
                    and _clean_cache_text(dependency_definition.get("formula"))
                ):
                    source_kind = "calculated"
                else:
                    source_kind = "input"
            if (
                not source_kind
                and runtime_cache_provenance_service.provenance_exists(path)
            ):
                source_kind = "engine"
            if source_kind in {"engine", "calculated"}:
                validation_precedent = dict(precedent)
                if isinstance(dependency_definition, dict):
                    current_data_format = _clean_cache_text(
                        dependency_definition.get("data_format")
                    )
                    if current_data_format:
                        validation_precedent["data_format"] = current_data_format
                dependency_pairs = _calculated_precedent_request_pairs(
                    pairs,
                    validation_precedent,
                    dependency_sidecar,
                )
                if not arcrho_tri_cache_matches(
                    path,
                    dependency_pairs,
                    allow_runtime_cache_provenance=source_kind == "engine",
                    processing_hash_getter=processing_hash_getter,
                    file_fingerprint_getter=get_fingerprint,
                    calculated_validation_memo=memo,
                    calculated_validation_stack=active_stack,
                    verify_content=verify_content,
                ):
                    memo[cache_key] = False
                    return False
        except FileNotFoundError:
            memo[cache_key] = False
            return False
        except OSError as error:
            raise HTTPException(
                503,
                (
                    f"Unable to validate calculated dataset inputs for '{project_name}'. "
                    "The existing calculated cache was left unchanged."
                ),
            ) from error
    memo[cache_key] = True
    return True


def _processing_config_matches(
    payload: Dict[str, Any],
    pairs: list,
    data_path: str = "",
    processing_hash_getter: ProcessingHashGetter | None = None,
    file_fingerprint_getter: FileFingerprintGetter | None = None,
    calculated_validation_memo: CalculatedValidationMemo | None = None,
    calculated_validation_stack: set[str] | None = None,
    verify_content: bool = True,
) -> bool:
    source_kind = _clean_cache_text(payload.get("source_kind")).lower()
    if source_kind == "calculated":
        return _calculated_dependencies_match(
            payload,
            pairs,
            data_path,
            processing_hash_getter=processing_hash_getter,
            file_fingerprint_getter=file_fingerprint_getter,
            validation_memo=calculated_validation_memo,
            validation_stack=calculated_validation_stack,
            verify_content=verify_content,
        )
    if source_kind != "engine" or is_imported_snapshot_payload(payload):
        return True
    processing: Any = None
    filename = os.path.basename(str(data_path or "").strip())
    if filename:
        csv_file = _clean_cache_text(payload.get("csv_file"))
        if not csv_file or os.path.basename(csv_file) != filename:
            return False
        processing = payload.get("processing")
    else:
        processing = payload.get("processing")
    return _processing_hash_matches(processing, pairs, processing_hash_getter)


def _runtime_cache_identity(data_path: str, pairs: list) -> Dict[str, str]:
    return {
        "csv_file": os.path.basename(data_path),
        "dataset_name": _request_dataset_name(pairs),
        "dataset_type": _pair_value(pairs, "DatasetName") or _pair_value(pairs, "TriangleName"),
        "reserving_class": _pair_value(pairs, "Path"),
        "project_name": _pair_value(pairs, "ProjectName"),
        "function": _pair_value(pairs, "Function"),
    }


def _calculated_cache_identity(data_path: str, pairs: list) -> Dict[str, str]:
    return runtime_cache_provenance_service.calculated_cache_identity(
        data_path,
        project_name=_pair_value(pairs, "ProjectName"),
        reserving_class=_pair_value(pairs, "Path"),
        dataset_name=_request_dataset_name(pairs),
        dataset_type=_pair_value(pairs, "DatasetName") or _pair_value(pairs, "TriangleName"),
    )


def _calculated_cache_record(
    data_path: str,
    pairs: list,
    *,
    bind_to_csv: bool = True,
) -> Dict[str, Any] | None:
    """The technical record of what built this calculated CSV, or None.

    A read that fails for any reason other than absence is a retryable error:
    the existing cache is never called stale because the share hiccuped.
    """
    try:
        return runtime_cache_provenance_service.calculated_record(
            data_path,
            expected_identity=_calculated_cache_identity(data_path, pairs),
            bind_to_csv=bind_to_csv,
        )
    except OSError as error:
        project_name = _pair_value(pairs, "ProjectName") or "(unknown)"
        raise HTTPException(
            503,
            (
                f"Unable to validate calculated dataset inputs for '{project_name}'. "
                "The existing calculated cache was left unchanged."
            ),
        ) from error


def _runtime_cache_provenance_matches(
    data_path: str,
    pairs: list,
    processing_hash_getter: ProcessingHashGetter | None = None,
    file_fingerprint_getter: FileFingerprintGetter | None = None,
    *,
    verify_content: bool = True,
) -> bool:
    identity = _runtime_cache_identity(data_path, pairs)
    if not all(identity.values()):
        return False
    get_current_hash = processing_hash_getter or _processing_hash_getter(pairs)
    try:
        if not runtime_cache_provenance_service.provenance_exists(data_path):
            return False
        return runtime_cache_provenance_service.matches(
            data_path,
            expected_identity=identity,
            processing_config_hash=get_current_hash(),
            fingerprint_getter=file_fingerprint_getter,
            verify_content=verify_content,
        )
    except OSError as error:
        project_name = _pair_value(pairs, "ProjectName") or "(unknown)"
        raise HTTPException(
            503,
            (
                f"Unable to read technical cache provenance for '{project_name}'. "
                "The existing CSV was left unchanged."
            ),
        ) from error


def _write_runtime_cache_provenance(
    data_path: str,
    pairs: list,
    processing_hash_getter: ProcessingHashGetter | None = None,
) -> bool:
    """Record only the processing identity of a background-generated cache.

    ``WriteSidecar: false`` callers must not overwrite the user-editable dataset
    sidecar, but their generated CSV still needs durable provenance before it can
    be safely reused after a restart.
    """
    if not os.path.isfile(data_path):
        return False
    identity = _runtime_cache_identity(data_path, pairs)
    project_name = _pair_value(pairs, "ProjectName")
    if not all(identity.values()):
        return False
    try:
        get_current_hash = processing_hash_getter or _processing_hash_getter(pairs)
        return runtime_cache_provenance_service.record(
            data_path,
            identity=identity,
            processing=get_processing_provenance(
                project_name,
                config_hash=get_current_hash(),
            ),
        )
    except Exception as error:
        print(f"Unable to record ArcRho runtime cache provenance: {error}")
        return False


def _remove_runtime_cache_provenance(data_path: str) -> None:
    try:
        runtime_cache_provenance_service.remove(data_path)
    except OSError as error:
        print(f"Unable to remove ArcRho runtime cache provenance: {error}")


def _require_runtime_cache_provenance(
    data_path: str,
    pairs: list,
    processing_hash_getter: ProcessingHashGetter,
) -> bool:
    if _write_runtime_cache_provenance(data_path, pairs, processing_hash_getter):
        return True
    raise HTTPException(
        503,
        (
            "The generated dataset CSV is available, but ArcRho could not record "
            "its technical cache provenance. The CSV was left unchanged; check "
            "write access to the reserving-class data folder and retry."
        ),
    )


def _triangle_sidecar_payload(
    data_path: str,
    pairs: list,
    *,
    local_only: bool = False,
    validate_processing_for_path: bool = True,
    processing_hash_getter: ProcessingHashGetter | None = None,
    file_fingerprint_getter: FileFingerprintGetter | None = None,
    calculated_validation_memo: CalculatedValidationMemo | None = None,
    calculated_validation_stack: set[str] | None = None,
    verify_content: bool = True,
) -> Dict[str, Any]:
    sidecar_path = _dataset_sidecar_path(data_path, pairs)
    payload = _safe_read_json(sidecar_path)
    if not payload:
        return {}
    expected_name = _request_dataset_name(pairs)
    if not _cache_payload_name_matches(payload, expected_name):
        return {}
    if not _cache_text_matches(payload.get("reserving_class"), _pair_value(pairs, "Path")):
        return {}
    # The project is settled by where the sidecar was found, not by the name it
    # stores: this path was built from the requested project's folders, so a
    # sidecar read from it belongs to that project. Duplicating or renaming a
    # project copies the stored name verbatim, and comparing it would reject
    # every cache in the new project and force a full rebuild of each one.
    source_kind = _clean_cache_text(payload.get("source_kind")).lower()
    if not local_only and source_kind != "input":
        return {}
    if validate_processing_for_path and not _processing_config_matches(
        payload,
        pairs,
        data_path,
        processing_hash_getter,
        file_fingerprint_getter,
        calculated_validation_memo,
        calculated_validation_stack,
        verify_content=verify_content,
    ):
        return {}
    data_format = _clean_cache_text(payload.get("data_format") or "Triangle").lower()
    if data_format and data_format != "triangle":
        return {}
    return payload


def _manual_input_sidecar_payload(data_path: str, pairs: list) -> Dict[str, Any]:
    return _triangle_sidecar_payload(data_path, pairs, local_only=False)


def _is_stale_input_variant(data_path: str, pairs: list) -> bool:
    """Is this a coarser copy of a hand-entered dataset left behind on disk?

    A hand-entered dataset is current only in the CSV its sidecar names. Any
    other copy beside it is a view an older release wrote down, and nothing
    updates it when the figures are edited, so it is never served.
    """

    payload = _safe_read_json(_dataset_sidecar_path(data_path, pairs))
    if not isinstance(payload, dict):
        return False
    if _clean_cache_text(payload.get("source_kind")).lower() != "input":
        return False
    csv_file = _clean_cache_text(payload.get("csv_file"))
    return bool(csv_file) and os.path.basename(csv_file) != os.path.basename(data_path)


def _is_generated_triangle_payload(payload: Dict[str, Any]) -> bool:
    source_kind = _clean_cache_text(payload.get("source_kind")).lower()
    return source_kind == "engine"


def _parse_cache_variant(filename: str) -> Dict[str, Any]:
    stem, ext = os.path.splitext(os.path.basename(filename))
    if ext.lower() != ".csv":
        return {}
    parts = stem.split("@")
    if len(parts) < 5:
        return {}
    origin = parts[-4].strip()
    dev = parts[-3].strip()
    cum = parts[-2].strip().lower()
    cal = parts[-1].strip().lower()
    if not origin.isdigit() or not dev.isdigit():
        return {}
    if cum not in {"cum", "inc"} or cal not in {"dev", "cal"}:
        return {}
    return {
        "base": "@".join(parts[:-4]),
        "origin_length": int(origin),
        "development_length": int(dev),
        "cumulative": cum == "cum",
        "calendar": cal == "cal",
    }


def _local_cache_candidates(
    data_path: str,
    pairs: list,
    *,
    local_only: bool = False,
    processing_hash_getter: ProcessingHashGetter | None = None,
    file_fingerprint_getter: FileFingerprintGetter | None = None,
    calculated_validation_memo: CalculatedValidationMemo | None = None,
    calculated_validation_stack: set[str] | None = None,
    verify_content: bool = True,
) -> list[Dict[str, Any]]:
    payload = _triangle_sidecar_payload(
        data_path,
        pairs,
        local_only=local_only,
        validate_processing_for_path=False,
        processing_hash_getter=processing_hash_getter,
        file_fingerprint_getter=file_fingerprint_getter,
        calculated_validation_memo=calculated_validation_memo,
        calculated_validation_stack=calculated_validation_stack,
        verify_content=verify_content,
    )
    if not payload:
        return []
    dataset_dir = os.path.dirname(data_path)
    expected_base = sanitize_dataset_file_name(_request_dataset_name(pairs))
    if not os.path.isdir(dataset_dir):
        return []
    out: list[Dict[str, Any]] = []
    seen_paths: set[str] = set()

    def add_candidate(filename: str) -> None:
        parsed = _parse_cache_variant(filename)
        if not parsed or parsed["base"] != expected_base:
            return
        path = os.path.join(dataset_dir, os.path.basename(filename))
        path_key = os.path.normcase(os.path.abspath(path))
        if path_key in seen_paths or not os.path.isfile(path):
            return
        if not _processing_config_matches(
            payload,
            pairs,
            path,
            processing_hash_getter,
            file_fingerprint_getter,
            calculated_validation_memo,
            calculated_validation_stack,
            verify_content=verify_content,
        ):
            return
        seen_paths.add(path_key)
        out.append({
            **parsed,
            "path": path,
            "payload": payload,
        })

    preferred_csv = _clean_cache_text(payload.get("csv_file"))
    if preferred_csv:
        add_candidate(preferred_csv)
        if out and _can_derive_cache(out[0], pairs, data_path)[0]:
            return out

    for filename in os.listdir(dataset_dir):
        add_candidate(filename)
    return out


def _can_derive_cache(candidate: Dict[str, Any], pairs: list, target_path: str) -> tuple[bool, str]:
    if candidate.get("path") == target_path:
        return False, "exact target already handled"
    target_origin = _pair_int_value(pairs, "OriginLength", 12)
    target_dev = _pair_int_value(pairs, "DevelopmentLength", 12)
    source_origin = int(candidate.get("origin_length") or 0)
    source_dev = int(candidate.get("development_length") or 0)
    calendar = bool(candidate.get("calendar"))
    if calendar != _pair_bool_value(pairs, "Calendar", False):
        return False, "calendar mode differs"
    if bool(candidate.get("cumulative")) != _pair_bool_value(pairs, "Cumulative", True):
        return False, "cumulative mode differs"
    reason = rollup_reason(source_origin, source_dev, target_origin, target_dev, calendar=calendar)
    return (not reason), reason


def _derive_triangle_cache(candidate: Dict[str, Any], pairs: list, target_path: str) -> Dict[str, Any]:
    source_path = str(candidate["path"])
    source_origin = int(candidate["origin_length"])
    source_dev = int(candidate["development_length"])
    target_origin = _pair_int_value(pairs, "OriginLength", 12)
    target_dev = _pair_int_value(pairs, "DevelopmentLength", 12)
    origin_factor = target_origin // source_origin
    dev_factor = target_dev // source_dev
    from app_server.services import dataset_service

    # The view is valued on the project's Development End Date, like every
    # triangle created at the requested shape would be.
    valuation = dataset_service.valuation_months(_pair_value(pairs, "ProjectName"))
    df = pd.read_csv(source_path, header=None, dtype="float64", keep_default_na=True)
    values = rollup_triangle(
        df.to_numpy().tolist(),
        source_origin_length=source_origin,
        source_development_length=source_dev,
        target_origin_length=target_origin,
        target_development_length=target_dev,
        valuation_months=valuation,
        cumulative=_pair_bool_value(pairs, "Cumulative", True),
        calendar=_pair_bool_value(pairs, "Calendar", False),
    )
    target_rows = len(values)
    target_cols = len(values[0])
    derived = {
        "source_path": source_path,
        "source_origin_length": source_origin,
        "source_development_length": source_dev,
        "origin_factor": origin_factor,
        "development_factor": dev_factor,
        "target_rows": target_rows,
        "target_cols": target_cols,
    }

    payload = candidate.get("payload")
    if isinstance(payload, dict) and _clean_cache_text(payload.get("source_kind")).lower() == "input":
        # A hand-entered triangle is the only copy of figures nobody can
        # produce again, so a coarser view of it is never written beside it:
        # the grid is handed the roll-up recipe and builds the view from the
        # stored CSV on every read.
        dataset_service.register_rollup_handle(
            _arcrho_dataset_id(target_path, pairs),
            {
                "source_path": source_path,
                "source_origin_length": source_origin,
                "source_development_length": source_dev,
                "target_origin_length": target_origin,
                "target_development_length": target_dev,
                "valuation_months": valuation,
                "cumulative": _pair_bool_value(pairs, "Cumulative", True),
                "calendar": _pair_bool_value(pairs, "Calendar", False),
            },
        )
        derived["in_memory"] = True
        return derived

    os.makedirs(os.path.dirname(target_path), exist_ok=True)
    tmp_path = f"{target_path}.{uuid.uuid4()}.tmp"
    pd.DataFrame(values).to_csv(tmp_path, header=False, index=False)
    os.replace(tmp_path, target_path)
    return derived


def _arcrho_dataset_id(data_path: str, pairs: list | None = None) -> str:
    function_name = _pair_value(pairs or [], "Function").strip().lower()
    prefix = "arcrhovec_" if function_name == "arcrhovec" else "arcrhotri_"
    return prefix + hashlib.sha1(data_path.encode("utf-8")).hexdigest()[:16]


def _register_arcrho_dataset(data_path: str, pairs: list | None = None) -> str:
    ds_id = _arcrho_dataset_id(data_path, pairs)
    config.DATASETS[ds_id] = data_path
    return ds_id


def resolve_local_triangle_cache(
    data_path: str,
    pairs: list,
    allow_derived: bool = True,
    materialize: bool = True,
    local_only: bool = False,
    materialize_path: str | None = None,
    refresh_index_on_materialize: bool = True,
    allow_runtime_cache_provenance: bool = False,
    processing_hash_getter: ProcessingHashGetter | None = None,
    file_fingerprint_getter: FileFingerprintGetter | None = None,
    calculated_validation_memo: CalculatedValidationMemo | None = None,
    calculated_validation_stack: set[str] | None = None,
    verify_content: bool = True,
) -> Dict[str, Any]:
    target_path = materialize_path or data_path
    get_processing_hash = processing_hash_getter or _processing_hash_getter(pairs)
    get_file_fingerprint = file_fingerprint_getter or _file_fingerprint_getter()
    validation_memo = (
        calculated_validation_memo
        if calculated_validation_memo is not None
        else {}
    )
    if not _is_stale_input_variant(data_path, pairs) and arcrho_tri_cache_matches(
        data_path,
        pairs,
        allow_runtime_cache_provenance=allow_runtime_cache_provenance,
        processing_hash_getter=get_processing_hash,
        file_fingerprint_getter=get_file_fingerprint,
        calculated_validation_memo=validation_memo,
        calculated_validation_stack=calculated_validation_stack,
        verify_content=verify_content,
    ):
        payload = _triangle_sidecar_payload(
            data_path,
            pairs,
            local_only=True,
            processing_hash_getter=get_processing_hash,
            file_fingerprint_getter=get_file_fingerprint,
            calculated_validation_memo=validation_memo,
            calculated_validation_stack=calculated_validation_stack,
            verify_content=verify_content,
        )
        return {
            "ok": True,
            "status": "cache_exact",
            "data_path": data_path,
            "manual_source_found": bool(_manual_input_sidecar_payload(data_path, pairs)),
            "generated_source_found": bool(payload and _is_generated_triangle_payload(payload)),
            "local_source_found": True,
        }

    payload = _triangle_sidecar_payload(
        data_path,
        pairs,
        local_only=local_only,
        validate_processing_for_path=False,
        processing_hash_getter=get_processing_hash,
        file_fingerprint_getter=get_file_fingerprint,
        calculated_validation_memo=validation_memo,
        calculated_validation_stack=calculated_validation_stack,
    )
    if not payload:
        return {
            "ok": False,
            "status": "missing_sidecar",
            "manual_source_found": False,
            "generated_source_found": False,
            "local_source_found": False,
            "message": f"Input triangle cache sidecar was not found for '{_request_dataset_name(pairs)}'.",
            "data_path": data_path,
        }
    generated_source_found = _is_generated_triangle_payload(payload)
    manual_source_found = bool(
        _triangle_sidecar_payload(
            data_path,
            pairs,
            local_only=False,
            processing_hash_getter=get_processing_hash,
            file_fingerprint_getter=get_file_fingerprint,
            calculated_validation_memo=validation_memo,
            calculated_validation_stack=calculated_validation_stack,
            verify_content=verify_content,
        )
    )
    if not allow_derived:
        return {
            "ok": False,
            "status": "cache_missing",
            "manual_source_found": manual_source_found,
            "generated_source_found": generated_source_found,
            "local_source_found": True,
            "message": f"Input triangle cache was not found for '{_request_dataset_name(pairs)}'.",
            "data_path": data_path,
        }

    candidates = _local_cache_candidates(
        data_path,
        pairs,
        local_only=local_only,
        processing_hash_getter=get_processing_hash,
        file_fingerprint_getter=get_file_fingerprint,
        calculated_validation_memo=validation_memo,
        calculated_validation_stack=calculated_validation_stack,
        verify_content=verify_content,
    )
    if not candidates:
        return {
            "ok": False,
            "status": "cache_missing",
            "manual_source_found": manual_source_found,
            "generated_source_found": generated_source_found,
            "local_source_found": True,
            "message": f"Input triangle cache was not found for '{_request_dataset_name(pairs)}'.",
            "data_path": data_path,
        }

    rejected: list[str] = []
    candidates.sort(key=lambda item: (int(item.get("origin_length") or 999999), int(item.get("development_length") or 999999)))
    for candidate in candidates:
        can_derive, reason = _can_derive_cache(candidate, pairs, target_path)
        if not can_derive:
            if reason:
                rejected.append(reason)
            continue
        if not materialize:
            return {
                "ok": True,
                "status": "cache_derivable",
                "data_path": target_path,
                "manual_source_found": manual_source_found,
                "generated_source_found": generated_source_found,
                "local_source_found": True,
                "derived": {
                    "source_path": str(candidate["path"]),
                    "source_origin_length": int(candidate.get("origin_length") or 0),
                    "source_development_length": int(candidate.get("development_length") or 0),
                },
            }
        try:
            derived = _derive_triangle_cache(candidate, pairs, target_path)
        except Exception as err:
            rejected.append(str(err))
            continue
        if refresh_index_on_materialize and not derived.get("in_memory"):
            try:
                dataset_instance_index_service.rebuild_index(_pair_value(pairs, "ProjectName"), _pair_value(pairs, "Path"))
            except Exception:
                pass
        return {
            "ok": True,
            "status": "cache_derived",
            "data_path": target_path,
            "manual_source_found": manual_source_found,
            "generated_source_found": generated_source_found,
            "local_source_found": True,
            "derived": derived,
        }

    detail = rejected[0] if rejected else "no compatible finer cache was found"
    return {
        "ok": False,
        "status": "cache_missing" if generated_source_found else "cache_not_derivable",
        "manual_source_found": manual_source_found,
        "generated_source_found": generated_source_found,
        "local_source_found": True,
        "message": (
            f"Input triangle '{_request_dataset_name(pairs)}' exists as a local cache that cannot derive "
            f"{_pair_int_value(pairs, 'OriginLength', 12)}x{_pair_int_value(pairs, 'DevelopmentLength', 12)} periods: {detail}."
        ),
        "data_path": data_path,
    }


def _engine_stored_lengths(project_name: str, origin: int, development: int) -> tuple:
    """The months per period an Engine regeneration of this dataset can reach.

    A generated dataset is rebuilt from the project's source table, so the
    finest shape it can be produced at is that table's own date granularity,
    however coarse the period this request asked to see it at. A project whose
    mapping records no granularity keeps the requested shape.
    """
    from app_server.services import field_mapping_service

    months = field_mapping_service.load_source_period_months(project_name)
    return (
        months.get(DATE_ROLE_ORIGIN, origin),
        months.get(DATE_ROLE_DEVELOPMENT, development),
    )


def _apply_dataset_sidecar_shape_fields(
    payload: Dict[str, Any],
    pairs: list,
    *,
    is_vector: bool,
    stored: tuple,
) -> None:
    origin = _pair_int_value(pairs, "OriginLength", 12)
    development = _pair_int_value(pairs, "DevelopmentLength", 12)
    stored_origin, stored_development = stored
    if is_vector:
        payload["period_length"] = origin
        for obsolete_key in (
            "origin_length",
            "development_length",
            "development_count",
            "cumulative",
            "calendar",
            "stored_origin_length",
            "stored_development_length",
        ):
            payload.pop(obsolete_key, None)
        payload.update(stored_length_fields("Vector", stored_origin))
        return
    payload["origin_length"] = origin
    payload["development_length"] = development
    payload["cumulative"] = _pair_bool_value(pairs, "Cumulative", True)
    payload["calendar"] = _pair_bool_value(pairs, "Calendar", False)
    payload.pop("period_length", None)
    payload.pop("stored_period_length", None)
    payload.update(stored_length_fields("Triangle", stored_origin, stored_development))


def _set_processing_provenance(
    payload: Dict[str, Any],
    project_name: str,
    data_path: str,
) -> None:
    payload["processing"] = get_processing_provenance(project_name)


def _write_dataset_sidecar_impl(data_path: str, pairs: list) -> None:
    dataset_type = _pair_value(pairs, "DatasetName") or _pair_value(pairs, "TriangleName")
    instance_name = _pair_value(pairs, "InstanceName") or dataset_type
    if not instance_name:
        return
    sidecar_path = _dataset_sidecar_path(data_path, pairs)
    project_name = _pair_value(pairs, "ProjectName")
    reserving_class = _pair_value(pairs, "Path")
    is_vector = _pair_value(pairs, "Function").strip().lower() == "arcrhovec"
    data_format = "Vector" if is_vector else "Triangle"
    user_name = user_identity_service.get_current_display_name() or getpass.getuser()
    updated_at = utc_now_text()
    requested_origin = _pair_int_value(pairs, "OriginLength", 12)
    requested_development = _pair_int_value(pairs, "DevelopmentLength", 12)
    if os.path.exists(sidecar_path):
        payload = dataset_sidecar_status_service.read_sidecar(sidecar_path)
        if not payload:
            return
        payload["method_type"] = dataset_sidecar_status_service.METHOD_TYPE_NONE
        payload["status"] = dataset_sidecar_status_service.STATUS_CURRENT
        payload["updated_at"] = updated_at
        # The engine just regenerated this dataset from source tables, so a
        # ResQ-import source_modified stamp no longer describes its content.
        payload.pop("source_modified", None)
        payload["modified_by"] = user_name
        payload["data_format"] = data_format
        # A sidecar copied in by a project duplication still names the project
        # it came from. Restamping it here retires that name as each dataset is
        # regenerated, so nothing downstream can read it and act on the wrong
        # project.
        payload["project_name"] = project_name
        payload["show_subtotal"] = normalize_show_subtotal(payload.get("show_subtotal"))
        generated = _clean_cache_text(payload.get("source_kind")).lower() == "engine"
        if generated:
            _set_processing_provenance(payload, project_name, data_path)
        _apply_dataset_sidecar_shape_fields(
            payload,
            pairs,
            is_vector=is_vector,
            # Only a dataset the Engine rebuilds from the source table can
            # claim the source's granularity; anything else holds exactly the
            # shape this cache was written at.
            stored=(
                _engine_stored_lengths(project_name, requested_origin, requested_development)
                if generated
                else (requested_origin, requested_development)
            ),
        )
        payload["csv_file"] = os.path.basename(data_path)
        from app_server.services.dataset_service import _append_dataset_audit_entry

        _append_dataset_audit_entry(payload, "Update", event_date=updated_at, user_name=user_name)
        dataset_sidecar_status_service.write_sidecar(sidecar_path, payload)
        dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
            project_name,
            reserving_class,
            [instance_name, dataset_type],
        )
        return
    created = utc_now_text()
    try:
        created = _utc_timestamp_from_stat(os.stat(data_path).st_ctime)
    except OSError:
        pass
    display_settings = dataset_number_format_service.dataset_type_number_format_settings(dataset_type)
    stored_origin, stored_development = _engine_stored_lengths(
        project_name, requested_origin, requested_development
    )
    payload = build_engine_dataset_sidecar(
        project_name=project_name,
        reserving_class=reserving_class,
        dataset_name=instance_name,
        dataset_type=dataset_type,
        data_format=data_format,
        csv_file=os.path.basename(data_path),
        user=user_name,
        created=created,
        updated_at=updated_at,
        number_format=display_settings["number_format"],
        decimal_places=display_settings["decimal_places"],
        origin_length=requested_origin,
        development_length=requested_development,
        period_length=requested_origin if is_vector else None,
        stored_origin_length=stored_origin,
        stored_development_length=stored_development,
        stored_period_length=stored_origin if is_vector else None,
        cumulative=_pair_bool_value(pairs, "Cumulative", True),
        calendar=_pair_bool_value(pairs, "Calendar", False),
        processing=get_processing_provenance(project_name),
    )
    from app_server.services import calculated_dataset_service

    calculated_dataset_service.apply_sidecar_graph_fields(
        payload,
        _pair_value(pairs, "ProjectName"),
        dataset_type,
    )
    dataset_sidecar_status_service.write_sidecar(sidecar_path, payload)
    dataset_sidecar_status_service.refresh_method_statuses_for_dependents(
        project_name,
        reserving_class,
        [instance_name, dataset_type],
    )


def _write_dataset_sidecar(data_path: str, pairs: list) -> None:
    project_name = _pair_value(pairs, "ProjectName")
    reserving_class = _pair_value(pairs, "Path")
    sidecar_path = _dataset_sidecar_path(data_path, pairs)
    with (
        dataset_sidecar_status_service.reserving_class_io_lock(project_name, reserving_class),
        dataset_sidecar_status_service.sidecar_write_lock(sidecar_path),
    ):
        _write_dataset_sidecar_impl(data_path, pairs)


def _refresh_dataset_instance_index_after_cache_write(pairs: list) -> None:
    project_name = _pair_value(pairs, "ProjectName")
    reserving_class = _pair_value(pairs, "Path")
    if not project_name or not reserving_class:
        return
    try:
        dataset_instance_index_service.rebuild_index(project_name, reserving_class)
    except Exception:
        return


def arcrho_tri_cache_matches(
    data_path: str,
    pairs: list,
    *,
    allow_runtime_cache_provenance: bool = False,
    processing_hash_getter: ProcessingHashGetter | None = None,
    file_fingerprint_getter: FileFingerprintGetter | None = None,
    calculated_validation_memo: CalculatedValidationMemo | None = None,
    calculated_validation_stack: set[str] | None = None,
    verify_content: bool = True,
) -> bool:
    if not os.path.isfile(data_path):
        return False

    def runtime_provenance_matches() -> bool:
        return allow_runtime_cache_provenance and _runtime_cache_provenance_matches(
            data_path,
            pairs,
            processing_hash_getter,
            file_fingerprint_getter,
            verify_content=verify_content,
        )

    sidecar_path = _dataset_sidecar_path(data_path, pairs)
    try:
        with open(sidecar_path, "r", encoding="utf-8") as f:
            payload = json.load(f)
    except FileNotFoundError:
        return runtime_provenance_matches()
    except json.JSONDecodeError:
        return runtime_provenance_matches()
    except OSError as error:
        project_name = _pair_value(pairs, "ProjectName") or "(unknown)"
        raise HTTPException(
            503,
            (
                f"Unable to validate the dataset cache for '{project_name}'. "
                "The existing CSV was left unchanged."
            ),
        ) from error
    if not isinstance(payload, dict):
        return runtime_provenance_matches()
    sidecar_allows_runtime_provenance = (
        _clean_cache_text(payload.get("source_kind")).lower() == "engine"
    )

    def sidecar_mismatch() -> bool:
        return (
            runtime_provenance_matches()
            if sidecar_allows_runtime_provenance
            else False
        )

    expected_name = _pair_value(pairs, "InstanceName")
    dataset_type = _pair_value(pairs, "DatasetName") or _pair_value(pairs, "TriangleName")
    if not expected_name:
        expected_name = dataset_type
    if not _cache_payload_name_matches(payload, expected_name):
        return sidecar_mismatch()
    if not _cache_text_matches(payload.get("reserving_class"), _pair_value(pairs, "Path")):
        return sidecar_mismatch()
    # The sidecar's own folder settles which project it belongs to; the stored
    # project name is copied verbatim by a duplicate or a rename, so comparing
    # it would treat every cache in a duplicated project as foreign.
    if not _processing_config_matches(
        payload,
        pairs,
        data_path,
        processing_hash_getter,
        file_fingerprint_getter,
        calculated_validation_memo,
        calculated_validation_stack,
        verify_content=verify_content,
    ):
        return sidecar_mismatch()
    if not _pair_value(pairs, "InstanceName") and dataset_type:
        payload_type = payload.get("dataset_type")
        if payload_type and not _cache_text_matches(payload_type, dataset_type):
            return sidecar_mismatch()
    return True


def _require_valid_header_project_settings(pairs: list) -> Dict[str, Any]:
    project_name = _pair_value(pairs, "ProjectName")
    settings = project_settings_service.get_general_settings(project_name)
    data = settings.get("data") if isinstance(settings.get("data"), dict) else {}
    origin_start = str(data.get("origin_start_date") or "").strip()
    match = re.fullmatch(r"(\d{4})(0[1-9]|1[0-2])", origin_start)
    if not settings.get("exists") or not match or int(match.group(1)) <= 0:
        raise HTTPException(
            422,
            f"Cannot load ArcRho project headers for '{project_name}': Origin Start Date is missing or invalid. "
            "Set a valid Origin Start Date in Project Settings, then try again.",
        )
    return settings


def arcrho_headers(pairs: list, timeout_sec: float) -> Dict[str, Any]:
    settings = _require_valid_header_project_settings(pairs)
    data_path = set_data_path_like_vba(pairs)
    request_file = None

    settings_path = str(settings.get("path") or "").strip()
    if os.path.exists(data_path) and settings_path and os.path.exists(settings_path):
        try:
            if os.path.getmtime(data_path) < os.path.getmtime(settings_path):
                os.remove(data_path)
        except PermissionError:
            raise HTTPException(423, "ArcRho project headers cache is locked or inaccessible.")
        except OSError as err:
            raise HTTPException(500, f"Failed to refresh ArcRho project headers cache: {str(err)}")

    if not os.path.exists(data_path):
        try:
            os.makedirs(os.path.dirname(data_path), exist_ok=True)
        except OSError as err:
            raise HTTPException(500, f"Failed to create ArcRho headers data folder: {str(err)}")
        outcome = engine_calculation_service.run_engine_calculation(
            pairs, data_path, max(0.1, float(timeout_sec))
        )
        request_file = outcome.get("request_file")
        if not outcome["ok"]:
            return {
                "ok": False,
                "status": outcome["status"],
                "message": outcome.get("message")
                or "Timed out while loading ArcRho project headers. Verify the data engine is running, then try again.",
                "request_file": request_file,
                "data_path": data_path,
            }

    raw = file_read_cache.read_text_file_cached(data_path).strip()

    parts = [x.strip() for x in raw.replace("\n", ",").split(",") if x.strip()]

    return {
        "ok": True,
        "labels": parts,
        "request_file": request_file,
        "data_path": data_path,
    }


def _header_cache_pairs(project_name: str, period_type: int, transposed: bool, period_length: int, calendar: bool = False) -> list:
    return [
        ("Function", "ArcRhoHeaders"),
        ("periodType", str(period_type)),
        ("Transposed", str(transposed)),
        ("Calendar", str(calendar)),
        ("PeriodLength", str(period_length)),
        ("ProjectName", project_name),
        ("StoredPeriodLength", str(-1)),
    ]


def get_project_headers(
    project_name: str,
    period_length: int,
    timeout_sec: float,
    *,
    period_type: int = 0,
    transposed: bool = False,
    calendar: bool = False,
) -> Dict[str, Any]:
    """Load ArcRho headers for a project without exposing request-pair details."""
    project = str(project_name or "").strip()
    if not project:
        raise HTTPException(400, "ProjectName is required")
    try:
        length = int(period_length)
    except (TypeError, ValueError):
        raise HTTPException(400, "PeriodLength must be a positive integer")
    if length <= 0:
        raise HTTPException(400, "PeriodLength must be a positive integer")
    pairs = _header_cache_pairs(project, int(period_type), bool(transposed), length, bool(calendar))
    return arcrho_headers(pairs, timeout_sec=max(0.1, float(timeout_sec)))


def _target_header_cache_paths(project_name: str, origin_length: Any, development_length: Any) -> list[str]:
    targets: list[str] = []
    seen = set()
    specs = (
        (0, False, origin_length, False),
        (1, True, development_length, False),
        (1, True, development_length, True),
    )
    for period_type, transposed, length, calendar in specs:
        try:
            period_length = int(length)
        except (TypeError, ValueError):
            continue
        if period_length <= 0:
            continue
        path = set_data_path_like_vba(_header_cache_pairs(project_name, period_type, transposed, period_length, calendar))
        normalized = os.path.normcase(os.path.abspath(path))
        if normalized in seen:
            continue
        seen.add(normalized)
        targets.append(path)
    return targets


def clear_arcrho_headers_cache(project_name: str, origin_length: Any = None, development_length: Any = None) -> Dict[str, Any]:
    project_name_clean = str(project_name or "").strip()
    if not project_name_clean:
        raise HTTPException(400, "ProjectName is required")

    try:
        data_dir = config.get_project_data_dir(project_name_clean)
    except ValueError as e:
        raise HTTPException(404, str(e))

    cleared_files = []
    target_paths = _target_header_cache_paths(project_name_clean, origin_length, development_length)
    if target_paths:
        try:
            for path in target_paths:
                if not os.path.exists(path):
                    continue
                os.remove(path)
                cleared_files.append(os.path.basename(path))
        except PermissionError:
            raise HTTPException(423, "Cannot clear ArcRhoHeaders cache files because the project data folder is locked.")
        except OSError as e:
            raise HTTPException(500, f"Failed to clear ArcRhoHeaders cache files: {str(e)}")

        return {
            "ok": True,
            "project_name": project_name_clean,
            "data_dir": data_dir,
            "cleared_count": len(cleared_files),
            "cleared_files": cleared_files,
            "targeted": True,
        }

    if not os.path.isdir(data_dir):
        return {
            "ok": True,
            "project_name": project_name_clean,
            "data_dir": data_dir,
            "cleared_count": 0,
            "cleared_files": [],
            "targeted": False,
        }

    try:
        with os.scandir(data_dir) as it:
            for entry in it:
                if not entry.is_file():
                    continue
                name_l = entry.name.strip().lower()
                if not name_l.endswith(".csv"):
                    continue
                if not name_l.startswith("arcrhoheaders"):
                    continue
                os.remove(entry.path)
                cleared_files.append(entry.name)
    except PermissionError:
        raise HTTPException(423, "Cannot clear ArcRhoHeaders cache files because the project data folder is locked.")
    except OSError as e:
        raise HTTPException(500, f"Failed to clear ArcRhoHeaders cache files: {str(e)}")

    return {
        "ok": True,
        "project_name": project_name_clean,
        "data_dir": data_dir,
        "cleared_count": len(cleared_files),
        "cleared_files": cleared_files,
        "targeted": False,
    }


def arcrho_projects() -> Dict[str, Any]:
    seen = set()
    out = []
    index_data = project_settings_service._read_project_index()
    for item in index_data.get("projects", []):
        name = str(item.get("name", "") or "").strip()
        if name and name not in seen:
            out.append(name)
            seen.add(name)

    return {"sheet": "Virtual Projects", "projects": out, "folders": index_data.get("folders", [])}


def _local_cache_response(local_result: Dict[str, Any], data_path: str, pairs: list | None = None) -> Dict[str, Any]:
    ds_id = _register_arcrho_dataset(data_path, pairs)
    return {
        "ok": True,
        "need_request": False,
        "ds_id": ds_id,
        "request_file": None,
        "data_path": data_path,
        "local_cache_status": local_result.get("status"),
        "derived": local_result.get("derived"),
        "calculated_updates": None,
    }


def _recalculate_dependents_after_cache_write(pairs: list) -> Dict[str, Any] | None:
    project_name = _pair_value(pairs, "ProjectName")
    reserving_class = _pair_value(pairs, "Path")
    dataset_type = _pair_value(pairs, "DatasetName") or _pair_value(pairs, "TriangleName")
    dataset_name = _pair_value(pairs, "InstanceName") or dataset_type
    if not project_name or not reserving_class or not dataset_name:
        return None

    try:
        from app_server.services import dependent_propagation_service

        return dependent_propagation_service.enqueue_marked_save_propagation(
            project_name,
            reserving_class,
            dataset_name,
            dataset_type,
        )
    except Exception as err:
        return {"ok": False, "skipped": True, "reason": str(err)}


def _dependency_request_pairs(
    pairs: list,
    dataset_name: str,
    data_format: str,
    *,
    instance_name: str = "",
    settings: Dict[str, Any] | None = None,
) -> list:
    excluded = {
        "function",
        "datasetname",
        "trianglename",
        "vectorname",
        "instancename",
        "originlength",
        "developmentlength",
        "periodlength",
    }
    resolved_settings = settings or {}
    dependency_pairs = [
        (
            key,
            resolved_settings.get(str(key or "").strip().lower(), value),
        )
        for key, value in pairs
        if str(key or "").strip().lower() not in excluded
    ]
    is_vector = str(data_format or "").strip().lower() == "vector"
    function_name = "ArcRhoVec" if is_vector else "ArcRhoTri"
    vector_period = resolved_settings.get(
        "periodlength",
        resolved_settings.get(
            "originlength",
            _pair_value(pairs, "PeriodLength")
            or _pair_value(pairs, "OriginLength")
            or "12",
        ),
    )
    dimensions = (
        [
            ("OriginLength", vector_period),
            ("DevelopmentLength", vector_period),
        ]
        if is_vector
        else [
            (
                "OriginLength",
                resolved_settings.get(
                    "originlength",
                    _pair_value(pairs, "OriginLength") or "12",
                ),
            ),
            (
                "DevelopmentLength",
                resolved_settings.get(
                    "developmentlength",
                    _pair_value(pairs, "DevelopmentLength") or "12",
                ),
            ),
        ]
    )
    output = [
        ("Function", function_name),
        ("DatasetName", dataset_name),
        ("InstanceName", instance_name or dataset_name),
        *dimensions,
        *dependency_pairs,
    ]
    return [
        (key, "" if value is None else str(value))
        for key, value in output
    ]


def _dependency_cache_settings(
    descriptor: Dict[str, Any],
    data_format: str,
    path: str,
) -> Dict[str, Any]:
    settings: Dict[str, Any] = {}
    field_map = {
        "origin_length": "originlength",
        "development_length": "developmentlength",
        "period_length": "periodlength",
        "cumulative": "cumulative",
        "calendar": "calendar",
    }
    for source_key, pair_key in field_map.items():
        if descriptor.get(source_key) not in (None, ""):
            settings[pair_key] = descriptor[source_key]

    if str(data_format or "").strip().lower() == "vector":
        stem = os.path.splitext(os.path.basename(path))[0]
        parts = stem.rsplit("@", 1)
        if len(parts) == 2 and parts[1].isdigit():
            settings["periodlength"] = int(parts[1])
        return settings

    parsed = _parse_cache_variant(os.path.basename(path))
    if parsed:
        settings.update({
            "originlength": parsed["origin_length"],
            "developmentlength": parsed["development_length"],
            "cumulative": parsed["cumulative"],
            "calendar": parsed["calendar"],
        })
    return settings


def _cache_path_data_format(path: str) -> str:
    if _parse_cache_variant(os.path.basename(path)):
        return "Triangle"
    stem = os.path.splitext(os.path.basename(path))[0]
    parts = stem.rsplit("@", 1)
    return "Vector" if len(parts) == 2 and parts[1].isdigit() else ""


def _calculation_dataset_key(pairs: list) -> str:
    parts = [
        _canon_dataset_name(_pair_value(pairs, "ProjectName")),
        _canon_dataset_name(_pair_value(pairs, "Path")),
        _canon_dataset_name(
            _pair_value(pairs, "DatasetName")
            or _pair_value(pairs, "TriangleName")
        ),
    ]
    return "::".join(parts) if all(parts) else ""


def _materialize_calculated_dependencies(
    pairs: list,
    dependencies: list[Any],
    timeout_sec: float,
    *,
    local_only: bool,
    allow_derived: bool,
    calculation_stack: set[str],
    dataset_folder: str,
) -> List[Dict[str, Any]]:
    project_name = _pair_value(pairs, "ProjectName")
    if not project_name:
        return []

    from app_server.services import calculated_dataset_service

    rows_by_key = {
        _canon_dataset_name(row.get("name")): row
        for row in calculated_dataset_service._dataset_type_rows(project_name)
        if _canon_dataset_name(row.get("name"))
    }
    results: List[Dict[str, Any]] = []
    for dependency in dependencies:
        descriptor = dependency if isinstance(dependency, dict) else {}
        dependency_name = _clean_cache_text(
            descriptor.get("dataset_type_name")
            or dependency
        )
        row = rows_by_key.get(_canon_dataset_name(dependency_name))
        if not row:
            continue
        can_materialize = bool(
            row.get("generated")
            or (row.get("calculated") and str(row.get("formula") or "").strip())
        )
        if not can_materialize:
            continue
        instance_name = _clean_cache_text(
            descriptor.get("dataset_name")
            or row.get("name")
            or dependency_name
        )
        data_format = _clean_cache_text(
            row.get("data_format")
            or descriptor.get("data_format")
            or "Triangle"
        )
        stored_path = _clean_cache_text(descriptor.get("path"))
        stored_data_format = _clean_cache_text(
            descriptor.get("data_format")
            or _cache_path_data_format(stored_path)
        )
        if (
            stored_path
            and stored_data_format
            and stored_data_format.lower() != data_format.lower()
        ):
            stored_path = ""
        dependency_pairs = _dependency_request_pairs(
            pairs,
            str(row.get("name") or dependency_name),
            data_format,
            instance_name=instance_name,
            settings=_dependency_cache_settings(
                descriptor,
                data_format,
                stored_path,
            ),
        )
        canonical_dependency_path = set_data_path_like_vba(dependency_pairs)
        dependency_path = (
            stored_path
            if (
                stored_path
                and _path_is_within_folder(stored_path, dataset_folder)
                and _same_resolved_path(stored_path, canonical_dependency_path)
            )
            else canonical_dependency_path
        )
        result = run_arcrho_tri(
            dependency_pairs,
            dependency_path,
            timeout_sec=timeout_sec,
            force_refresh=False,
            local_only=local_only,
            allow_derived=allow_derived,
            write_sidecar=False,
            recalculate_dependents_on_cache_write=False,
            calculation_stack=calculation_stack,
        )
        results.append({"dataset_type_name": dependency_name, **result})
    return results


def _recalculate_requested_app_dataset(
    pairs: list,
    requested_data_path: str,
    timeout_sec: float,
    *,
    local_only: bool,
    allow_derived: bool,
    calculation_stack: set[str] | None = None,
    recalculate_dependents: bool = True,
    refresh_index: bool = True,
) -> Dict[str, Any] | None:
    project_name = _pair_value(pairs, "ProjectName")
    reserving_class = _pair_value(pairs, "Path")
    dataset_type = _pair_value(pairs, "DatasetName") or _pair_value(pairs, "TriangleName")
    if not project_name or not reserving_class or not dataset_type:
        return None

    dataset_key = _calculation_dataset_key(pairs)
    active_stack = set(calculation_stack or set())
    if dataset_key in active_stack:
        return {
            "ok": False,
            "status": "calculation_failed",
            "need_request": False,
            "data_path": requested_data_path,
            "message": f"Calculated dataset dependency cycle detected: {dataset_type}",
        }
    active_stack.add(dataset_key)

    try:
        from app_server.services import calculated_dataset_service

        contract = calculated_dataset_service.calculated_dataset_contract(
            project_name,
            dataset_type,
        )
        if contract is None:
            return None
        dependency_names = [
            str(item).strip()
            for item in contract.get("precedents") or []
            if str(item).strip()
        ]
        # Which files built the previous output -- the exact DFM method and its
        # input in particular -- is read from the technical record beside the
        # CSV. The cache is being rebuilt, so the record is read unbound.
        record = _calculated_cache_record(
            requested_data_path,
            pairs,
            bind_to_csv=False,
        ) or {}
        stored_precedents = {
            _canon_dataset_name(item.get("dataset_type") or item.get("dataset_name")): item
            for item in record.get("dependencies") or []
            if isinstance(item, dict)
            and _canon_dataset_name(item.get("dataset_type") or item.get("dataset_name"))
        }
        dependency_descriptors = [
            {
                **dict(stored_precedents.get(_canon_dataset_name(name)) or {}),
                "dataset_type_name": name,
            }
            for name in dependency_names
        ]
        dataset_folder = os.path.dirname(requested_data_path)
        current_dependency_contracts = (
            contract.get("precedent_contracts")
            if isinstance(contract.get("precedent_contracts"), dict)
            else {}
        )
        component_paths = {
            _canon_dataset_name(item.get("dataset_type_name")): _clean_cache_text(
                item.get("path")
            )
            for item in dependency_descriptors
            if _canon_dataset_name(item.get("dataset_type_name"))
            and _clean_cache_text(item.get("path")).lower().endswith(".csv")
            and _path_is_within_folder(
                _clean_cache_text(item.get("path")),
                dataset_folder,
            )
            and (
                not _clean_cache_text(
                    (
                        current_dependency_contracts.get(
                            _canon_dataset_name(item.get("dataset_type_name"))
                        )
                        or {}
                    ).get("data_format")
                )
                or _clean_cache_text(
                    item.get("data_format")
                    or _cache_path_data_format(_clean_cache_text(item.get("path")))
                ).lower()
                == _clean_cache_text(
                    (
                        current_dependency_contracts.get(
                            _canon_dataset_name(item.get("dataset_type_name"))
                        )
                        or {}
                    ).get("data_format")
                ).lower()
            )
        }
        component_method_sources: Dict[str, Dict[str, str]] = {}
        for item in dependency_descriptors:
            dependency_key = _canon_dataset_name(
                item.get("dataset_type_name")
            )
            dependency_definition = (
                current_dependency_contracts.get(dependency_key) or {}
            )
            currently_runtime_materialized = bool(
                dependency_definition.get("generated")
                or (
                    dependency_definition.get("calculated")
                    and _clean_cache_text(dependency_definition.get("formula"))
                )
            )
            if (
                dependency_key
                and not currently_runtime_materialized
                and _clean_cache_text(item.get("source_kind")).lower()
                == "dfm_method"
                and _clean_cache_text(item.get("path"))
            ):
                component_method_sources[dependency_key] = {
                    "path": _clean_cache_text(item.get("path")),
                    "input_path": _clean_cache_text(item.get("input_path")),
                }
        dependency_results = _materialize_calculated_dependencies(
            pairs,
            dependency_descriptors,
            timeout_sec,
            local_only=local_only,
            allow_derived=allow_derived,
            calculation_stack=active_stack,
            dataset_folder=dataset_folder,
        )
        failed_dependencies = [
            item for item in dependency_results if not item.get("ok")
        ]
        if failed_dependencies:
            messages = [
                _clean_cache_text(
                    item.get("message")
                    or item.get("reason")
                    or item.get("status")
                )
                for item in failed_dependencies
            ]
            message = "; ".join(item for item in messages if item)
            return {
                "ok": False,
                "status": "calculation_failed",
                "need_request": False,
                "data_path": requested_data_path,
                "message": message or "Failed to refresh a calculated dataset dependency.",
                "dependency_results": dependency_results,
            }

        for item in dependency_results:
            dependency_name = _canon_dataset_name(item.get("dataset_type_name"))
            dependency_path = _clean_cache_text(
                item.get("data_path")
                or item.get("path")
            )
            if (
                dependency_name
                and dependency_path.lower().endswith(".csv")
                and _path_is_within_folder(dependency_path, dataset_folder)
            ):
                component_paths[dependency_name] = dependency_path
        recalculate_kwargs: Dict[str, Any] = {
            "component_paths": component_paths,
        }
        if component_method_sources:
            recalculate_kwargs["component_method_sources"] = (
                component_method_sources
            )
        result = calculated_dataset_service.recalculate_dataset(
            project_name,
            reserving_class,
            dataset_type,
            **recalculate_kwargs,
        )
    except HTTPException:
        raise
    except Exception as err:
        return {
            "ok": False,
            "status": "calculation_error",
            "need_request": False,
            "data_path": requested_data_path,
            "message": str(err),
        }

    if not result.get("ok"):
        errors = [str(item).strip() for item in result.get("errors") or [] if str(item).strip()]
        message = "; ".join(errors) or str(result.get("reason") or "Failed to calculate dataset.")
        return {
            "ok": False,
            "status": "calculation_failed",
            "need_request": False,
            "data_path": requested_data_path,
            "message": message,
            "calculation": result,
            "dependency_results": dependency_results,
        }

    data_path = str(result.get("path") or requested_data_path)
    ds_id = _register_arcrho_dataset(data_path, pairs)
    calculated_updates = (
        _recalculate_dependents_after_cache_write(pairs)
        if recalculate_dependents
        else None
    )
    if refresh_index and not recalculate_dependents:
        try:
            dataset_instance_index_service.rebuild_index(project_name, reserving_class)
        except Exception:
            pass
    return {
        "ok": True,
        "need_request": False,
        "ds_id": ds_id,
        "request_file": None,
        "data_path": data_path,
        "local_cache_status": "calculated",
        "calculated": True,
        "calculated_updates": calculated_updates,
        "sidecar_written": True,
    }


def _normalize_temporary_session_id(value: Any) -> str:
    try:
        return str(uuid.UUID(str(value or "").strip()))
    except (AttributeError, TypeError, ValueError) as err:
        raise HTTPException(422, "TemporarySessionId must be a valid UUID.") from err


def _path_is_within_folder(path: str, folder: str) -> bool:
    child = os.path.normcase(os.path.realpath(os.path.abspath(path)))
    parent = os.path.normcase(os.path.realpath(os.path.abspath(folder)))
    try:
        return os.path.commonpath([child, parent]) == parent
    except ValueError:
        return False


def _same_resolved_path(left: str, right: str) -> bool:
    return (
        os.path.normcase(os.path.realpath(os.path.abspath(left)))
        == os.path.normcase(os.path.realpath(os.path.abspath(right)))
    )


def temporary_dataset_path(data_path: str, pairs: list) -> str:
    """Return the Temporary view cache path beside the canonical ``data_path``."""
    project_name = _pair_value(pairs, "ProjectName")
    reserving_class = _pair_value(pairs, "Path")
    if not project_name:
        raise HTTPException(400, "ProjectName is required for a temporary dataset request.")
    if not reserving_class:
        raise HTTPException(400, "Path is required for a temporary dataset request.")

    temporary_cache_dir = config.get_project_temporary_view_dataset_cache_dir(
        project_name,
        reserving_class,
    )

    dataset_filename = os.path.basename(data_path)
    if not dataset_filename:
        raise HTTPException(400, "Temporary dataset cache file name is invalid.")
    temporary_data_path = os.path.join(temporary_cache_dir, dataset_filename)
    if not _path_is_within_folder(temporary_data_path, temporary_cache_dir):
        raise HTTPException(400, "Temporary dataset path is outside its cache folder.")
    return temporary_data_path


def _temporary_dataset_response(
    data_path: str,
    pairs: list,
    temporary_session_id: str,
    *,
    need_request: bool,
    request_file: str | None,
    local_result: Dict[str, Any] | None = None,
    force_refresh: bool = False,
    cache_cleared: bool = False,
) -> Dict[str, Any]:
    ds_id = _register_arcrho_dataset(data_path, pairs)
    out: Dict[str, Any] = {
        "ok": True,
        "need_request": need_request,
        "ds_id": ds_id,
        "request_file": request_file,
        "data_path": data_path,
        "local_cache_status": (local_result or {}).get("status") or "temporary_cache",
        "derived": (local_result or {}).get("derived"),
        "calculated_updates": None,
        "sidecar_written": False,
        "temporary_cache": True,
        "temporary_session_id": temporary_session_id,
    }
    if force_refresh:
        out["cache_cleared"] = cache_cleared
    return out


def arcrho_precheck(
    data_path: str,
    pairs: list,
    *,
    local_only: bool = False,
    allow_derived: bool = True,
    temporary_session_id: str | None = None,
    allow_runtime_cache_provenance: bool = False,
) -> Dict[str, Any]:
    session_id = _normalize_temporary_session_id(temporary_session_id) if temporary_session_id else None
    temporary_data_path = temporary_dataset_path(data_path, pairs) if session_id else None
    local_result = resolve_local_triangle_cache(
        data_path,
        pairs,
        allow_derived=allow_derived,
        materialize=False,
        local_only=local_only,
        allow_runtime_cache_provenance=allow_runtime_cache_provenance and not session_id,
        # Precheck is advisory. Execute requests keep the authoritative
        # SHA-256 comparison before they reuse any matching cache.
        verify_content=False,
    )
    local_available = bool(local_result.get("ok"))
    canonical_cache_exact = local_result.get("status") == "cache_exact"
    temporary_cache_exists = bool(temporary_data_path and os.path.isfile(temporary_data_path))
    use_temporary_cache = temporary_cache_exists and not canonical_cache_exact
    cache_exists = local_available or temporary_cache_exists
    manual_source_found = bool(local_result.get("manual_source_found"))
    generated_source_found = bool(local_result.get("generated_source_found"))
    need_request = not cache_exists and not manual_source_found and (not local_only or generated_source_found)
    resolved_data_path = (
        temporary_data_path
        if temporary_data_path and not canonical_cache_exact
        else data_path
    )
    result = {
        "ok": True,
        "need_request": need_request,
        "cache_exists": cache_exists,
        "data_path": resolved_data_path,
        "ds_id": _arcrho_dataset_id(resolved_data_path, pairs),
        "local_cache_status": "temporary_cache" if use_temporary_cache else local_result.get("status"),
        "local_cache_message": None if use_temporary_cache else local_result.get("message"),
        "manual_source_found": manual_source_found,
        "generated_source_found": generated_source_found,
    }
    if session_id:
        result["temporary_session_id"] = session_id
        result["temporary_cache"] = use_temporary_cache
    return result


def _run_temporary_arcrho_tri(
    pairs: list,
    data_path: str,
    timeout_sec: float,
    *,
    temporary_session_id: str,
    force_refresh: bool,
    local_only: bool,
    allow_derived: bool,
) -> Dict[str, Any]:
    session_id = _normalize_temporary_session_id(temporary_session_id)
    temporary_data_path = temporary_dataset_path(data_path, pairs)
    get_processing_hash = _processing_hash_getter(pairs)

    local_result = resolve_local_triangle_cache(
        data_path,
        pairs,
        allow_derived=allow_derived,
        materialize=False,
        local_only=local_only,
        processing_hash_getter=get_processing_hash,
    )
    if local_result.get("ok") and local_result.get("status") == "cache_exact" and not force_refresh:
        out = _local_cache_response(local_result, data_path, pairs)
        out["temporary_cache"] = False
        out["temporary_session_id"] = session_id
        return out

    temporary_cache_exists = os.path.isfile(temporary_data_path)
    if temporary_cache_exists and not force_refresh:
        return _temporary_dataset_response(
            temporary_data_path,
            pairs,
            session_id,
            need_request=False,
            request_file=None,
        )

    if local_result.get("ok") and local_result.get("status") == "cache_derivable" and not force_refresh:
        derived_result = resolve_local_triangle_cache(
            data_path,
            pairs,
            allow_derived=allow_derived,
            materialize=True,
            local_only=local_only,
            materialize_path=temporary_data_path,
            refresh_index_on_materialize=False,
            processing_hash_getter=get_processing_hash,
        )
        if derived_result.get("ok"):
            derived_data_path = str(derived_result.get("data_path") or temporary_data_path)
            if os.path.normcase(os.path.abspath(derived_data_path)) == os.path.normcase(os.path.abspath(data_path)):
                out = _local_cache_response(derived_result, data_path, pairs)
                out["temporary_cache"] = False
                out["temporary_session_id"] = session_id
                return out
            return _temporary_dataset_response(
                derived_data_path,
                pairs,
                session_id,
                need_request=False,
                request_file=None,
                local_result=derived_result,
            )
        local_result = derived_result

    manual_source_found = bool(local_result.get("manual_source_found"))
    generated_source_found = bool(local_result.get("generated_source_found"))
    if (local_only and not generated_source_found) or manual_source_found:
        message = str(local_result.get("message") or "Input triangle cache is not available.")
        if force_refresh and manual_source_found:
            message = "Manual input triangle caches cannot be refreshed from the DFM/Dataset loader."
        return {
            "ok": False,
            "status": local_result.get("status") or "local_cache_unavailable",
            "need_request": False,
            "data_path": temporary_data_path,
            "message": message,
            "local_only": bool(local_only),
            "manual_source_found": manual_source_found,
            "temporary_session_id": session_id,
        }

    cache_cleared = False
    if (force_refresh or not temporary_cache_exists) and os.path.exists(temporary_data_path):
        try:
            os.remove(temporary_data_path)
            cache_cleared = True
        except OSError as err:
            raise HTTPException(423, f"Cannot clear temporary ArcRho tri file: {str(err)}") from err

    need_request = force_refresh or not temporary_cache_exists
    request_file = None
    if need_request:
        try:
            os.makedirs(os.path.dirname(temporary_data_path), exist_ok=True)
        except OSError as err:
            raise HTTPException(500, f"Failed to create temporary ArcRho tri data folder: {str(err)}") from err
        outcome = engine_calculation_service.run_engine_calculation(
            pairs,
            temporary_data_path,
            max(0.1, float(timeout_sec)),
            output_variant=OUTPUT_VARIANT_TEMPORARY_VIEW,
        )
        request_file = outcome.get("request_file")
        if not outcome["ok"]:
            timeout_out: Dict[str, Any] = {
                "ok": False,
                "status": outcome["status"],
                "need_request": True,
                "request_file": request_file,
                "data_path": temporary_data_path,
                "temporary_cache": True,
                "temporary_session_id": session_id,
            }
            if outcome.get("message"):
                timeout_out["message"] = outcome["message"]
            if force_refresh:
                timeout_out["cache_cleared"] = cache_cleared
            return timeout_out

    return _temporary_dataset_response(
        temporary_data_path,
        pairs,
        session_id,
        need_request=need_request,
        request_file=request_file,
        force_refresh=force_refresh,
        cache_cleared=cache_cleared,
    )


def run_arcrho_tri(
    pairs: list,
    data_path: str,
    timeout_sec: float,
    force_refresh: bool = False,
    local_only: bool = False,
    allow_derived: bool = True,
    write_sidecar: bool = True,
    temporary_session_id: str | None = None,
    recalculate_dependents_on_cache_write: bool = True,
    # A batch that regenerates every dataset in one reserving class rebuilds
    # the class index once at the end instead of after each dataset: the
    # rebuild reads every sidecar and method payload in the folder, so paying
    # it per dataset is quadratic in the size of the class. A reader that
    # arrives in between still gets a current index, because the persisted one
    # is checked against the folder signature before it is served.
    refresh_index: bool = True,
    calculation_stack: set[str] | None = None,
) -> Dict[str, Any]:
    calculation_key = _calculation_dataset_key(pairs)
    if calculation_key and calculation_key in set(calculation_stack or set()):
        dataset_type = (
            _pair_value(pairs, "DatasetName")
            or _pair_value(pairs, "TriangleName")
        )
        return {
            "ok": False,
            "status": "calculation_failed",
            "need_request": False,
            "data_path": data_path,
            "message": f"Calculated dataset dependency cycle detected: {dataset_type}",
        }
    if temporary_session_id:
        return _run_temporary_arcrho_tri(
            pairs,
            data_path,
            timeout_sec,
            temporary_session_id=temporary_session_id,
            force_refresh=force_refresh,
            local_only=local_only,
            allow_derived=allow_derived,
        )

    request_file = None
    cache_cleared = False
    get_processing_hash = _processing_hash_getter(pairs)

    local_result = resolve_local_triangle_cache(
        data_path,
        pairs,
        allow_derived=allow_derived,
        local_only=local_only,
        refresh_index_on_materialize=(
            write_sidecar and refresh_index and not recalculate_dependents_on_cache_write
        ),
        allow_runtime_cache_provenance=not write_sidecar,
        processing_hash_getter=get_processing_hash,
    )
    if local_result.get("ok") and not force_refresh:
        out = _local_cache_response(local_result, data_path, pairs)
        derived_in_memory = bool((local_result.get("derived") or {}).get("in_memory"))
        if local_result.get("status") == "cache_derived" and not derived_in_memory:
            if write_sidecar:
                _write_dataset_sidecar(data_path, pairs)
                if refresh_index and not recalculate_dependents_on_cache_write:
                    _refresh_dataset_instance_index_after_cache_write(pairs)
            else:
                out["cache_provenance_recorded"] = _require_runtime_cache_provenance(
                    data_path,
                    pairs,
                    get_processing_hash,
                )
        if (
            write_sidecar
            and recalculate_dependents_on_cache_write
            and not derived_in_memory
            and local_result.get("status") != "cache_exact"
        ):
            out["calculated_updates"] = _recalculate_dependents_after_cache_write(pairs)
        return out

    calculated_result = _recalculate_requested_app_dataset(
        pairs,
        data_path,
        timeout_sec,
        local_only=local_only,
        allow_derived=allow_derived,
        calculation_stack=calculation_stack,
        recalculate_dependents=bool(
            write_sidecar and recalculate_dependents_on_cache_write
        ),
        refresh_index=bool(write_sidecar and refresh_index),
    )
    if calculated_result is not None:
        return calculated_result
    manual_source_found = bool(local_result.get("manual_source_found"))
    generated_source_found = bool(local_result.get("generated_source_found"))
    if (local_only and not generated_source_found) or manual_source_found:
        message = str(local_result.get("message") or "Input triangle cache is not available.")
        if force_refresh and manual_source_found:
            message = "Manual input triangle caches cannot be refreshed from the DFM/Dataset loader."
        return {
            "ok": False,
            "status": local_result.get("status") or "local_cache_unavailable",
            "need_request": False,
            "data_path": data_path,
            "message": message,
            "local_only": bool(local_only),
            "manual_source_found": manual_source_found,
        }

    cache_matches = local_result.get("status") == "cache_exact"
    if (force_refresh or not cache_matches) and os.path.exists(data_path):
        try:
            os.remove(data_path)
            _remove_runtime_cache_provenance(data_path)
            cache_cleared = True
        except OSError as e:
            raise HTTPException(423, f"Cannot clear cached ArcRho tri file: {str(e)}")

    need_request = force_refresh or (not cache_matches)
    if need_request:
        try:
            os.makedirs(os.path.dirname(data_path), exist_ok=True)
        except OSError as err:
            raise HTTPException(500, f"Failed to create ArcRho tri data folder: {str(err)}")
        outcome = engine_calculation_service.run_engine_calculation(
            pairs, data_path, max(0.1, float(timeout_sec))
        )
        request_file = outcome.get("request_file")
        if not outcome["ok"]:
            timeout_out: Dict[str, Any] = {
                "ok": False,
                "status": outcome["status"],
                "need_request": True,
                "request_file": request_file,
                "data_path": data_path,
            }
            if outcome.get("message"):
                timeout_out["message"] = outcome["message"]
            if force_refresh:
                timeout_out["cache_cleared"] = cache_cleared
            return timeout_out

    if not write_sidecar:
        cache_provenance_recorded = (
            _require_runtime_cache_provenance(data_path, pairs, get_processing_hash)
            if need_request
            else None
        )
        ds_id = _register_arcrho_dataset(data_path, pairs)
        out: Dict[str, Any] = {
            "ok": True,
            "need_request": need_request,
            "ds_id": ds_id,
            "request_file": request_file,
            "data_path": data_path,
            "calculated_updates": None,
            "sidecar_written": False,
        }
        if cache_provenance_recorded is not None:
            out["cache_provenance_recorded"] = cache_provenance_recorded
        if force_refresh:
            out["cache_cleared"] = cache_cleared
        return out

    calculated_updates = None
    try:
        _write_dataset_sidecar(data_path, pairs)
        calculated_updates = (
            _recalculate_dependents_after_cache_write(pairs)
            if recalculate_dependents_on_cache_write
            else None
        )
        if refresh_index and not recalculate_dependents_on_cache_write:
            _refresh_dataset_instance_index_after_cache_write(pairs)
    except OSError as err:
        raise HTTPException(500, f"Failed to write ArcRho tri dataset metadata: {str(err)}")

    ds_id = _register_arcrho_dataset(data_path, pairs)

    out: Dict[str, Any] = {
        "ok": True,
        "need_request": need_request,
        "ds_id": ds_id,
        "request_file": request_file,
        "data_path": data_path,
        "calculated_updates": calculated_updates,
    }
    if force_refresh:
        out["cache_cleared"] = cache_cleared
    return out
