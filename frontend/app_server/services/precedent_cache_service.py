"""Locate the cache CSV a method reads for one precedent at the method's own period.

A method can run at a finer or coarser period than the cache its precedent's
sidecar names. An Engine-generated dataset is rebuilt at the requested lengths
through the same ``/arcrho/tri`` runtime the browser uses, leaving the sidecar
and its primary cache untouched; a hand-entered dataset is rolled up in memory
from its own CSV, so no coarser copy of it is ever written or trusted; a method
output already publishes its coarser vector variants, so those are looked up on
disk. DFM and Result Selection share this one resolver rather than each
deciding on its own.
"""

from __future__ import annotations

import os
from typing import Any, List, Mapping, Sequence

from arcrho_api import triangle_rollup
from arcrho_api.sidecar_core_contract import is_vector_format, stored_lengths

from app_server import config
from app_server.helpers import build_dataset_cache_file_name, set_data_path_like_vba


def _clean(value: Any) -> str:
    return str(value if value is not None else "").strip()


def _positive_int(value: Any) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        return 0
    return number if number > 0 else 0


def source_period(sidecar: Mapping[str, Any]) -> int:
    """Return the origin period length a sidecar's cache is stored at, or 0.

    Stored, not displayed: this decides which file on disk holds the values,
    so it must describe the CSV the sidecar names rather than the coarser
    shape a window may be showing it at.
    """

    return stored_lengths(sidecar)[0]


def sidecar_csv_path(project_name: str, reserving_class: str, sidecar: Mapping[str, Any]) -> str:
    """The CSV a sidecar names as its own, or an empty string when it names none."""

    csv_file = os.path.basename(_clean(sidecar.get("csv_file")))
    if not csv_file:
        return ""
    return os.path.join(config.get_project_dataset_cache_dir(project_name, reserving_class), csv_file)


def _rollup_arguments(
    sidecar: Mapping[str, Any],
    origin_length: Any,
    development_length: Any = None,
) -> dict:
    stored_origin, stored_development = stored_lengths(sidecar)
    if is_vector_format(sidecar.get("data_format")):
        # A vector holds one column, so only its rows are aggregated: a plain
        # block sum, which is the calendar form of the roll-up.
        return {
            "source_origin_length": stored_origin,
            "source_development_length": stored_development,
            "target_origin_length": _positive_int(origin_length),
            "target_development_length": stored_development,
            "cumulative": True,
            "calendar": True,
        }
    return {
        "source_origin_length": stored_origin,
        "source_development_length": stored_development,
        "target_origin_length": _positive_int(origin_length),
        "target_development_length": _positive_int(development_length) or stored_development,
        "cumulative": bool(sidecar.get("cumulative", True)),
        "calendar": bool(sidecar.get("calendar", False)),
    }


def rollup_reason(
    sidecar: Mapping[str, Any],
    origin_length: Any,
    development_length: Any = None,
) -> str:
    """Why the sidecar's own CSV cannot be rolled up to those lengths, or ``""``.

    Only a hand-entered dataset is rolled up. Its CSV is the finest copy of
    figures that add, and it is the only kind of dataset that cannot simply be
    produced again at the period a method wants it at.
    """

    if _clean(sidecar.get("source_kind")).lower() != "input":
        return "only a hand-entered dataset is rolled up in memory"
    arguments = _rollup_arguments(sidecar, origin_length, development_length)
    return triangle_rollup.rollup_reason(
        arguments["source_origin_length"],
        arguments["source_development_length"],
        arguments["target_origin_length"],
        arguments["target_development_length"],
        calendar=arguments["calendar"],
    )


def rollup_rows(
    project_name: str,
    sidecar: Mapping[str, Any],
    rows: Sequence[Sequence[Any]],
    origin_length: Any,
    development_length: Any = None,
) -> List[List[float | None]]:
    """Aggregate a hand-entered dataset's own rows to the coarser lengths.

    The view is valued on the project's Development End Date, like every
    dataset of the project is.
    """

    from app_server.services import dataset_service

    return triangle_rollup.rollup_triangle(
        rows,
        valuation_months=dataset_service.valuation_months(project_name),
        **_rollup_arguments(sidecar, origin_length, development_length),
    )


def precedent_source(
    project_name: str,
    reserving_class: str,
    dataset_name: str,
    sidecar: Mapping[str, Any],
    origin_length: Any,
) -> tuple[str, bool]:
    """The CSV a method reads for this precedent at ``origin_length`` months.

    Returns ``(csv_path, needs_rollup)``. A precedent stored at the method's
    own period is read from the CSV its sidecar names. A finer one is served
    the way the DFM already serves it: an Engine-generated dataset is rebuilt
    at the method's period, and a hand-entered one is read from its own CSV
    and rolled up in memory, so a monthly exposure vector feeds a yearly
    method. ``csv_path`` is empty when the sidecar names no CSV. Raises
    ``RuntimeError`` with the reason when the precedent cannot be brought to
    the method's period, which is the case for a coarser one.
    """

    stored = source_period(sidecar)
    target = _positive_int(origin_length)
    if not stored or not target or stored == target:
        return sidecar_csv_path(project_name, reserving_class, sidecar), False
    if _clean(sidecar.get("source_kind")).lower() == "engine":
        try:
            path = materialize_engine_source(
                project_name, reserving_class, dataset_name, sidecar, target
            )
        except RuntimeError as exc:
            raise RuntimeError(
                f"could not be generated at {target} months: {exc}"
            ) from exc
        return path, False
    reason = rollup_reason(sidecar, target)
    if reason:
        raise RuntimeError(
            f"uses {stored}-month origins; expected {target} ({reason})"
        )
    return sidecar_csv_path(project_name, reserving_class, sidecar), True


def materialize_engine_source(
    project_name: str,
    reserving_class: str,
    dataset_name: str,
    sidecar: Mapping[str, Any],
    origin_length: int,
    development_length: int | None = None,
) -> str:
    """Rebuild an Engine-generated dataset at the requested lengths and return its cache path."""

    data_format = _clean(sidecar.get("data_format")) or "Triangle"
    function = "ArcRhoVec" if data_format.lower() == "vector" else "ArcRhoTri"
    dataset_type = _clean(sidecar.get("dataset_type")) or dataset_name
    cumulative = bool(sidecar.get("cumulative", True))
    transposed = bool(sidecar.get("transposed", False))
    calendar = bool(sidecar.get("calendar", False))
    pairs = [
        ("Function", function),
        ("Path", reserving_class),
        ("DatasetName", dataset_type),
        ("InstanceName", dataset_name),
        ("Cumulative", str(cumulative)),
        ("Transposed", str(transposed)),
        ("Calendar", str(calendar)),
        ("ProjectName", project_name),
        ("OriginLength", str(origin_length)),
        ("DevelopmentLength", str(development_length or origin_length)),
    ]
    path = set_data_path_like_vba(pairs)
    from app_server.services import arcrho_runtime_service

    result = arcrho_runtime_service.run_arcrho_tri(
        pairs,
        path,
        timeout_sec=config.ENGINE_REQUEST_TIMEOUT_SEC,
        local_only=False,
        allow_derived=True,
        write_sidecar=False,
    )
    if not result.get("ok") or not os.path.isfile(path):
        raise RuntimeError(result.get("message") or f"Unable to materialize '{dataset_name}' at {origin_length} months.")
    return path


def precedent_csv_path(
    project_name: str,
    reserving_class: str,
    dataset_name: str,
    sidecar: Mapping[str, Any],
    origin_length: int,
    *,
    exact: bool,
) -> str:
    """Return the cache CSV holding ``dataset_name`` at ``origin_length`` months.

    ``exact`` demands a cache at that length even when the sidecar's own cache
    already has it; otherwise the sidecar's cache is used whenever its period
    matches.
    """

    data_dir = config.get_project_dataset_cache_dir(project_name, reserving_class)
    data_format = _clean(sidecar.get("data_format")) or "Triangle"
    source_kind = _clean(sidecar.get("source_kind")).lower()
    native_period = source_period(sidecar)
    sidecar_path = os.path.join(data_dir, os.path.basename(_clean(sidecar.get("csv_file")))) \
        if _clean(sidecar.get("csv_file")) else ""
    needs_exact = exact or bool(native_period and native_period != origin_length)
    if needs_exact:
        if source_kind == "engine":
            return materialize_engine_source(project_name, reserving_class, dataset_name, sidecar, origin_length)
        if native_period and origin_length < native_period:
            raise RuntimeError(
                f"Exact {origin_length}-month cache cannot be derived from the current "
                f"{native_period}-month output for '{dataset_name}'."
            )
        filename = build_dataset_cache_file_name(
            dataset_name,
            data_format,
            origin_length,
            origin_length,
            True,
            False,
        ) + ".csv"
        target = os.path.join(data_dir, filename)
        if os.path.isfile(target):
            return target
        if native_period == origin_length and sidecar_path and os.path.isfile(sidecar_path):
            return sidecar_path
        raise RuntimeError(f"Exact {origin_length}-month cache is missing for '{dataset_name}'.")
    if sidecar_path and os.path.isfile(sidecar_path):
        return sidecar_path
    raise RuntimeError(f"Cached dataset CSV is missing for '{dataset_name}'.")
