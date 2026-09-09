from __future__ import annotations

import calendar
import csv
import getpass
import io
import math
import os
import re
import uuid
from contextlib import contextmanager
from contextvars import ContextVar
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from functools import wraps
from pathlib import Path

from arcrho_api.bornhuetter_ferguson_contract import (
    BF_JSON_FORMAT,
    BF_METHOD_TYPE,
    BF_SOURCE_KIND,
    bornhuetter_ferguson_precedent_names,
    build_bornhuetter_ferguson_output_sidecar,
    recalculate_bornhuetter_ferguson_method,
)
from arcrho_api.cape_cod_contract import (
    CC_JSON_FORMAT,
    CC_METHOD_TYPE,
    CC_PRIOR_ULTIMATE_MODES,
    CC_SCALING_TYPES,
    CC_SOURCE_KIND,
    build_cape_cod_output_sidecar,
    cape_cod_precedent_names,
    recalculate_cape_cod_method,
)
from arcrho_api.dfm_contract import build_dfm_output_sidecar, dfm_output_variants
from arcrho_api.dataset_display_contract import normalize_show_subtotal
from arcrho_api.dataset_link_contract import DatasetLinkError, canonical_dataset_formula
from arcrho_api.engine_dataset_sidecar_contract import build_engine_dataset_sidecar
from arcrho_api.sidecar_audit_contract import AUDIT_ACTION_INSERT, AUDIT_ACTION_UPDATE
from arcrho_api.sidecar_core_contract import (
    dependency_entries,
    stored_length_fields,
)
from arcrho_api.timestamps import format_persisted_timestamp, utc_now_text

from .catalog import (
    _apply_sidecar_graph_meta,
    _canon_dataset_name,
    _is_calculated_dataset_type,
    _is_generated_dataset_type,
    _triangle_source_kind,
)
from .engine import import_user_identity_service
from .core import (
    BS_CRA_FILE_PREFIX,
    BS_CRA_JSON_FORMAT,
    BS_CRA_METHOD_TYPE,
    BS_CRA_SOURCE_KIND,
    BS_SR_FILE_PREFIX,
    BS_SR_JSON_FORMAT,
    BS_SR_METHOD_TYPE,
    BS_SR_SOURCE_KIND,
    DATASET_CACHE_DIR,
    DATASET_SIDECAR_DIR,
    DEFAULT_CALENDAR,
    DEFAULT_CUMULATIVE,
    METHOD_TYPE_BF_CODE,
    METHOD_TYPE_BS_CRA_CODE,
    METHOD_TYPE_BS_SR_CODE,
    METHOD_TYPE_CAPE_COD_CODE,
    METHOD_TYPE_NONE_CODE,
    METHOD_TYPE_DFM_CODE,
    METHOD_TYPE_RESULT_SELECTION_CODE,
    _bool_value,
    _call_member,
    _clean_name,
    _dataset_cache_csv_file_name,
    _encode_name_part,
    _is_result_selection_method_type,
    _iso_or_text,
    _json_sidecar_name,
    _method_type_code,
    _method_type_name,
    _normalize_import_name,
    normalize_method_status,
    persisted_json_text,
    _safe_attr,
    _safe_int_attr,
    _safe_read_json,
    _try_call_member,
    _vector_cache_csv_file_name,
    _write_csv_matrix,
    _write_json,
    _write_sidecar_json,
)
from .number_formats import (
    dataset_type_decimal_places,
    dataset_type_number_format,
    number_format_entry,
)


PROJECT_NAME = "NJ_Annual_Prod_202605_Fake"
RS_JSON_FORMAT = "arcrho-result-selection-v4"
METHOD_DATA_DIR = "methods"
RS_JSON_VALUE_DECIMAL_PLACES = 6

BS_SR_ADJUSTMENT_TYPES = {
    0: "unadjusted",
    1: "pairs",
    2: "all",
    3: "loess",
}
# Mirrors normalizeLoessSpan in
# frontend/ui/method_pages/berquist_sherman/calculation_helpers.js so both
# producers persist the same loess_span for the same logical inputs. Both B&S
# variants carry one, for the Settlement Rate adjusted paid claims and for the
# Case Reserve Adequacy current average case reserves.
BS_DEFAULT_LOESS_SPAN = 7
BS_MIN_LOESS_SPAN = 2
BS_MAX_LOESS_SPAN = 99
BS_CRA_INFLATION_TYPES = {
    0: "case_column",
    1: "case_all",
    2: "paid_column",
    3: "paid_all",
    4: "user",
}
BS_CRA_AVERAGE_CASE_RESERVE_TYPES = {
    0: "latest",
    1: "monotone",
    2: "loess",
    3: "user",
}
# Mirrors ROLE_DEFINITIONS in
# frontend/ui/method_pages/berquist_sherman/berquist_sherman_main.js: both
# producers record one number format per source role, in the same order.
BS_SOURCE_ROLES = {
    "sr": ("paid_claims", "closed_claim_numbers", "ultimate_claim_numbers"),
    "cra": ("paid_claims", "incurred_claims", "reported_claim_numbers", "closed_claim_numbers"),
}

_DEFER_GRAPH_ENRICHMENT_DEPTH: ContextVar[int] = ContextVar(
    "resq_migration_defer_graph_enrichment_depth",
    default=0,
)

_STRICT_RESQ_EXTRACTION: ContextVar[bool] = ContextVar(
    "resq_migration_strict_extraction",
    default=False,
)


class StrictResQExtractionError(RuntimeError):
    """A required ResQ value could not be read during selective extraction."""


def _strict_extractor(function):
    """Scope an extractor's opt-in strict flag across its nested helpers."""

    @wraps(function)
    def wrapped(*args, **kwargs):
        token = _STRICT_RESQ_EXTRACTION.set(bool(kwargs.get("strict", False)))
        try:
            return function(*args, **kwargs)
        finally:
            _STRICT_RESQ_EXTRACTION.reset(token)

    return wrapped


def _strict_failure(message: str, exc: Exception | None = None):
    error = StrictResQExtractionError(message)
    if exc is None:
        raise error
    raise error from exc


def _extract_attr(source, member_name: str, default=None, *, context: str = "ResQ object"):
    """Read a persisted COM property, preserving tolerant bulk-import behavior."""

    if not _STRICT_RESQ_EXTRACTION.get():
        return _safe_attr(source, member_name, default)
    try:
        return getattr(source, member_name)
    except Exception as exc:
        _strict_failure(f"Could not read {context}.{member_name} during strict ResQ extraction.", exc)


def _extract_int_attr(source, member_name: str, default: int = 0, *, context: str = "ResQ object") -> int:
    if not _STRICT_RESQ_EXTRACTION.get():
        return _safe_int_attr(source, member_name, default)
    value = _extract_attr(source, member_name, default, context=context)
    try:
        return int(value)
    except Exception as exc:
        _strict_failure(
            f"Could not convert {context}.{member_name}={value!r} to an integer during strict ResQ extraction.",
            exc,
        )


def resq_stored_lengths(
    item,
    *,
    is_vector: bool,
    origin_length: int,
    development_length: int | None = None,
) -> tuple[int, int]:
    """The months per period ResQ holds *item*'s data at, per axis.

    ResQ keeps a displayed length (``OriginLength`` / ``DevelopmentLength``,
    ``PeriodLength`` on a vector) and a stored one (``StoredOriginLength`` /
    ``StoredDevelopmentLength``, ``StoredPeriodLength``), and the displayed
    length is always a whole multiple of the stored one. The stored length is
    the only record of how fine a dataset's data is: a generated dataset is
    stored at the source data's granularity whatever period it is shown at,
    and a hand-entered one at the shape it was typed at. A stored length ResQ
    does not answer, or one the displayed length is not a multiple of, reads
    as the displayed length, which is the shape the values can be read at
    either way.
    """
    display_origin = int(origin_length or 0)
    display_development = int(
        display_origin if development_length is None else (development_length or 0)
    )

    def stored(member: str, display: int) -> int:
        value = _safe_int_attr(item, member, 0)
        if value <= 0 or display <= 0 or display % value != 0:
            return display
        return value

    if is_vector:
        period = stored("StoredPeriodLength", display_origin)
        return period, period
    return (
        stored("StoredOriginLength", display_origin),
        stored("StoredDevelopmentLength", display_development),
    )


def _restore_displayed_lengths(item, previous: list[tuple[str, object]]) -> None:
    for member, value in reversed(previous):
        try:
            setattr(item, member, value)
        except Exception:
            pass


@contextmanager
def _displayed_at(item, lengths: dict[str, int]):
    """Show *item* at the displayed lengths in *lengths* for the block.

    ResQ hands out a dataset's values at its displayed length, so reading the
    data at the shape it is stored at means showing it at that shape first.
    The same switch the export macro makes before it writes values. Yields
    ``True`` when every member was set and ``False`` when ResQ refused one, in
    which case the values are read at the shape the dataset already shows. The
    displayed lengths are put back on the way out either way, and nothing is
    saved, so the ResQ project is left exactly as it was found.
    """
    previous: list[tuple[str, object]] = []
    switched = True
    try:
        for member, value in lengths.items():
            current = _safe_int_attr(item, member, 0)
            if current <= 0:
                switched = False
                break
            if current == int(value):
                continue
            try:
                setattr(item, member, int(value))
            except Exception:
                switched = False
                break
            previous.append((member, current))
        if not switched:
            _restore_displayed_lengths(item, previous)
            previous = []
        yield switched
    finally:
        _restore_displayed_lengths(item, previous)


@contextmanager
def defer_sidecar_graph_enrichment():
    """Defer per-write graph work until the caller performs a bulk graph refresh."""
    token = _DEFER_GRAPH_ENRICHMENT_DEPTH.set(_DEFER_GRAPH_ENRICHMENT_DEPTH.get() + 1)
    try:
        yield
    finally:
        _DEFER_GRAPH_ENRICHMENT_DEPTH.reset(token)


def _apply_graph_meta_best_effort(meta: dict, dataset_type: str, rc_dir: Path, **kwargs) -> None:
    if _DEFER_GRAPH_ENRICHMENT_DEPTH.get() > 0:
        return
    try:
        _apply_sidecar_graph_meta(meta, dataset_type, rc_dir, **kwargs)
    except Exception as exc:
        meta.setdefault("precedents", [])
        meta.setdefault("dependents", [])
        meta["graph_metadata_error"] = str(exc)


def configure_extractors(
    *,
    project_name: str,
    rs_json_format: str,
    method_data_dir: str,
    bf_json_format: str | None = None,
    cc_json_format: str | None = None,
) -> None:
    global PROJECT_NAME, RS_JSON_FORMAT, METHOD_DATA_DIR

    PROJECT_NAME = str(project_name)
    RS_JSON_FORMAT = str(rs_json_format)
    if bf_json_format and str(bf_json_format) != BF_JSON_FORMAT:
        raise ValueError(
            f"The ResQ producer only supports canonical BF format {BF_JSON_FORMAT!r}."
        )
    if cc_json_format and str(cc_json_format) != CC_JSON_FORMAT:
        raise ValueError(
            f"The ResQ producer only supports canonical Cape Cod format {CC_JSON_FORMAT!r}."
        )
    METHOD_DATA_DIR = str(method_data_dir)


_MONTH_NAME_NUMBERS = {
    name.lower(): number
    for number, name in enumerate(calendar.month_abbr)
    if number
}


def _period_end_date(year: int, month: int) -> datetime:
    return datetime(year, month, calendar.monthrange(year, month)[1])


def _origin_date_from_label(label: str) -> datetime | None:
    """Return a date that falls inside the origin period a ResQ label names.

    ResQ resolves an ``OriginDate`` to the origin period containing that date, so
    a sub-annual label must map inside its own period.  Returning the period's
    last day keeps an annual label on 31 December exactly as before, while
    ``2025 Q1`` now resolves to 31 March 2025 instead of collapsing onto Q4.
    """
    text = _normalize_import_name(label)
    year_match = re.search(r"(?<!\d)(\d{4})(?!\d)", text)
    if not year_match:
        # Bare indices such as "1" are ResQ label fallbacks, not calendar years.
        return None
    year = int(year_match.group(1))
    remainder = f"{text[: year_match.start()]} {text[year_match.end():]}".strip().lower()
    if not remainder:
        return _period_end_date(year, 12)
    quarter = re.search(r"(?<![a-z0-9])q\s*([1-4])(?![0-9])", remainder)
    if quarter:
        return _period_end_date(year, int(quarter.group(1)) * 3)
    half = re.search(r"(?<![a-z0-9])h\s*([1-2])(?![0-9])", remainder)
    if half:
        return _period_end_date(year, int(half.group(1)) * 6)
    for name, number in _MONTH_NAME_NUMBERS.items():
        if re.search(rf"(?<![a-z]){name}(?![a-z])", remainder):
            return _period_end_date(year, number)
    month = re.search(r"(?<![a-z0-9])m\s*(1[0-2]|[1-9])(?![0-9])", remainder)
    if month:
        return _period_end_date(year, int(month.group(1)))
    return _period_end_date(year, 12)


def _resolve_origin_date(source, origin_index: int, label: str) -> datetime | None:
    """Prefer ResQ's own origin date for a row; fall back to parsing its label."""
    try:
        value = _try_call_member(
            source,
            "GetOriginDate",
            [((origin_index,), {}), ((), {"OriginIndex": origin_index})],
        )
    except Exception:
        value = None
    if value is not None:
        try:
            return datetime(int(value.year), int(value.month), int(value.day))
        except Exception:
            pass
    return _origin_date_from_label(label)

def _triangle_development_count(triangle, origin_index: int) -> int | None:
    """Return how many development periods one origin row has populated.

    ``None`` means ResQ could not answer at all.  Callers must keep that distinct
    from ``0``: a zero count is ResQ stating that an origin period beyond the
    valuation date holds no data yet, and its row must stay empty.
    """
    origin_date = _resolve_origin_date(
        triangle,
        origin_index,
        _triangle_origin_label(triangle, origin_index),
    )
    call_shapes = [
        ((), {"OriginDate": origin_date}) if origin_date is not None else None,
        ((origin_index,), {}),
        ((), {"OriginIndex": origin_index}),
        ((), {"arg0": origin_index}),
        ((), {}),
    ]
    call_shapes = [shape for shape in call_shapes if shape is not None]
    errors: list[Exception] = []
    for name in ("DevelopmentCount", "DevCount"):
        try:
            return int(_try_call_member(triangle, name, call_shapes))
        except Exception as exc:
            errors.append(exc)
            continue
    try:
        return int(getattr(triangle, "DevelopmentCount"))
    except Exception as exc:
        errors.append(exc)
        if _STRICT_RESQ_EXTRACTION.get():
            _strict_failure(
                f"Could not read triangle DevelopmentCount for origin index {origin_index}.",
                errors[0],
            )
        return None

def _triangle_origin_label(triangle, origin_index: int) -> str:
    errors: list[Exception] = []
    for name in ("OriginLabel", "OriginLabels"):
        try:
            return _normalize_import_name(_try_call_member(triangle, name, [((origin_index,), {}), ((), {"OriginIndex": origin_index})]))
        except Exception as exc:
            errors.append(exc)
            continue
    if _STRICT_RESQ_EXTRACTION.get():
        _strict_failure(
            f"Could not read triangle origin label for origin index {origin_index}.",
            errors[0] if errors else None,
        )
    return str(origin_index)

def _triangle_development_label(triangle, dev_index: int) -> str:
    errors: list[Exception] = []
    for name in ("DevelopmentLabel", "DevelopmentLabels", "DevLabel"):
        try:
            return _normalize_import_name(_try_call_member(triangle, name, [((dev_index,), {}), ((), {"DevIndex": dev_index})]))
        except Exception as exc:
            errors.append(exc)
            continue
    if _STRICT_RESQ_EXTRACTION.get():
        _strict_failure(
            f"Could not read triangle development label for development index {dev_index}.",
            errors[0] if errors else None,
        )
    return str(dev_index)

def _triangle_value(triangle, origin_index: int, dev_index: int):
    call_shapes = [
        ((origin_index, dev_index), {}),
        ((), {"OriginIndex": origin_index, "DevIndex": dev_index}),
        ((), {"OriginIndex": origin_index, "DevelopmentIndex": dev_index}),
    ]
    for name in ("ValuesByIndex", "Values", "Value", "Data", "TriangleValues"):
        try:
            return _try_call_member(triangle, name, call_shapes)
        except Exception:
            continue
    raise AttributeError(
        "Could not read triangle values. Tried ValuesByIndex, Values, Value, Data, and TriangleValues "
        f"for cell ({origin_index}, {dev_index})."
    )

def _vector_origin_count(vector) -> int:
    errors: list[Exception] = []
    for name in ("OriginCount", "Count", "Length"):
        try:
            value = int(getattr(vector, name))
            if value > 0:
                return value
        except Exception as exc:
            errors.append(exc)
        try:
            value = int(_call_member(vector, name))
            if value > 0:
                return value
        except Exception as exc:
            errors.append(exc)
            continue
    if _STRICT_RESQ_EXTRACTION.get():
        _strict_failure(
            "Could not read a positive vector OriginCount/Count/Length.",
            errors[0] if errors else None,
        )
    return 0

def _vector_origin_label(vector, origin_index: int) -> str:
    errors: list[Exception] = []
    for name in ("OriginLabel", "OriginLabels", "Label", "Labels"):
        try:
            return _normalize_import_name(_try_call_member(vector, name, [((origin_index,), {}), ((), {"OriginIndex": origin_index})]))
        except Exception as exc:
            errors.append(exc)
            continue
    if _STRICT_RESQ_EXTRACTION.get():
        _strict_failure(
            f"Could not read vector origin label for origin index {origin_index}.",
            errors[0] if errors else None,
        )
    return str(origin_index)

def _vector_value(vector, origin_index: int):
    origin_date = _resolve_origin_date(
        vector,
        origin_index,
        _vector_origin_label(vector, origin_index),
    )
    call_shapes = [
        ((origin_index,), {}),
        ((), {"OriginIndex": origin_index}),
        ((), {"Index": origin_index}),
        ((), {"arg0": origin_index}),
        ((), {"OriginDate": origin_date}) if origin_date is not None else None,
    ]
    call_shapes = [shape for shape in call_shapes if shape is not None]
    for name in ("ValuesByIndex", "Values", "Value", "Data", "VectorValues"):
        try:
            return _try_call_member(vector, name, call_shapes)
        except Exception:
            continue
    raise AttributeError(
        "Could not read vector values. Tried ValuesByIndex, Values, Value, Data, and VectorValues "
        f"for origin index {origin_index}."
    )

def _read_triangle_body(triangle, name: str) -> dict:
    """Read a triangle's rows, values, and labels at the shape it is showing."""
    origin_count = _extract_int_attr(triangle, "OriginCount", 0, context=f"triangle {name!r}")
    if origin_count <= 0:
        try:
            origin_count = int(_call_member(triangle, "OriginCount"))
        except Exception:
            origin_count = 0
    if origin_count <= 0:
        raise ValueError(f"Triangle {name!r} does not expose a positive OriginCount.")

    first_row_dev_count = _triangle_development_count(triangle, 1)
    max_dev_count = first_row_dev_count if first_row_dev_count is not None else 0
    if max_dev_count <= 0:
        max_dev_count = _extract_int_attr(
            triangle,
            "DevelopmentCount",
            0,
            context=f"triangle {name!r}",
        )
    if max_dev_count <= 0:
        raise ValueError(f"Triangle {name!r} does not expose a positive DevelopmentCount.")
    row_dev_counts = [_triangle_development_count(triangle, i) for i in range(1, origin_count + 1)]
    known_row_dev_counts = [count for count in row_dev_counts if count is not None]
    if known_row_dev_counts:
        max_dev_count = max(max_dev_count, max(known_row_dev_counts))

    values: list[list] = []
    attempted_cells = 0
    value_errors: list[Exception] = []
    for i, row_dev_count in enumerate(row_dev_counts, start=1):
        row: list = []
        for j in range(1, max_dev_count + 1):
            # A known count of 0 means the origin period holds no data yet, so the
            # row stays empty.  Only an unknown (None) count falls back to reading
            # every column, because ResQ pads past the diagonal with 0.0.
            if row_dev_count is not None and j > row_dev_count:
                row.append(None)
                continue
            attempted_cells += 1
            try:
                row.append(_triangle_value(triangle, i, j))
            except Exception as exc:
                if _STRICT_RESQ_EXTRACTION.get():
                    _strict_failure(
                        f"Could not read triangle {name!r} value at cell ({i}, {j}).",
                        exc,
                    )
                value_errors.append(exc)
                row.append(None)
        values.append(row)
    if attempted_cells > 0 and len(value_errors) == attempted_cells:
        raise ValueError(f"Failed to read any values for triangle {name!r}: {value_errors[0]}")

    return {
        "origin_count": origin_count,
        "development_count": max_dev_count,
        "origin_labels": [_triangle_origin_label(triangle, i) for i in range(1, origin_count + 1)],
        "development_labels": [
            _triangle_development_label(triangle, j) for j in range(1, max_dev_count + 1)
        ],
        "values": values,
    }


@_strict_extractor
def export_triangle(
    triangle,
    *,
    method_type_code: int | None = None,
    strict: bool = False,
    at_stored_shape: bool = False,
) -> dict:
    """Extract a ResQ Triangle COM object into ArcRho CSV values and metadata.

    ``origin_length`` / ``development_length`` in the payload are the shape ResQ
    displays the triangle at; ``stored_origin_length`` /
    ``stored_development_length`` are the shape its ``values`` were read at.
    By default that is the displayed shape. With ``at_stored_shape`` the
    triangle is shown at ResQ's stored lengths while it is read, so a monthly
    triangle displayed yearly is copied month by month, the only way its
    finer figures ever leave ResQ. Should ResQ refuse the switch, the values
    are read at the displayed shape and the payload says so.
    """
    del strict
    name = _normalize_import_name(_extract_attr(triangle, "Name", "", context="triangle"))
    dataset_type_obj = _extract_attr(triangle, "DatasetType", None, context=f"triangle {name!r}")
    dataset_type = _normalize_import_name(
        _extract_attr(dataset_type_obj, "Name", "", context=f"triangle {name!r} DatasetType")
    )
    category_obj = _extract_attr(
        dataset_type_obj,
        "Category",
        None,
        context=f"triangle {name!r} DatasetType",
    )
    category = _normalize_import_name(
        _extract_attr(category_obj, "Name", "", context=f"triangle {name!r} Category")
    )
    data_format = _extract_int_attr(
        dataset_type_obj,
        "DataFormat",
        0,
        context=f"triangle {name!r} DatasetType",
    )
    if method_type_code is None:
        method_type_code = _extract_int_attr(
            triangle,
            "MethodType",
            METHOD_TYPE_NONE_CODE,
            context=f"triangle {name!r}",
        )
    method_type = _method_type_name(method_type_code)
    origin_length = _extract_int_attr(triangle, "OriginLength", 12, context=f"triangle {name!r}")
    dev_length = _extract_int_attr(triangle, "DevelopmentLength", 12, context=f"triangle {name!r}")
    stored_origin, stored_dev = resq_stored_lengths(
        triangle,
        is_vector=False,
        origin_length=origin_length,
        development_length=dev_length,
    )
    read_origin, read_dev = origin_length, dev_length
    switch: dict[str, int] = {}
    if at_stored_shape:
        if stored_origin != origin_length:
            switch["OriginLength"] = stored_origin
        if stored_dev != dev_length:
            switch["DevelopmentLength"] = stored_dev
    with _displayed_at(triangle, switch) as switched:
        if switch and switched:
            read_origin, read_dev = stored_origin, stored_dev
        body = _read_triangle_body(triangle, name)

    user = _normalize_import_name(_extract_attr(triangle, "User", "", context=f"triangle {name!r}"))
    created = _iso_or_text(_extract_attr(triangle, "Created", "", context=f"triangle {name!r}"))
    modified = _iso_or_text(_extract_attr(triangle, "Modified", "", context=f"triangle {name!r}"))
    notes = str(_extract_attr(triangle, "Notes", "", context=f"triangle {name!r}") or "")

    return {
        "name": name,
        "dataset_type": dataset_type,
        "category": category,
        "data_format": data_format,
        "method_type": method_type,
        "method_type_code": method_type_code,
        "origin_length": origin_length,
        "development_length": dev_length,
        "stored_origin_length": read_origin,
        "stored_development_length": read_dev,
        **body,
        "user": user,
        "created": created,
        "modified": modified,
        "notes": notes,
        "status": normalize_method_status(
            _extract_attr(triangle, "Status", 0, context=f"triangle {name!r}")
        ),
    }

def write_triangle_export(payload: dict, rc_path: str, rc_dir: Path) -> Path:
    name = _normalize_import_name(payload["name"])
    dataset_type = _normalize_import_name(payload.get("dataset_type")) or name
    origin_length = int(payload["origin_length"])
    dev_length = int(payload["development_length"])
    # The CSV is written at the shape the values were read at, which is
    # ResQ's stored shape for a hand-entered dataset; the display shape is the
    # one ResQ showed it at, and the app rolls the file up to it on read.
    stored_origin, stored_dev = _triangle_payload_stored_lengths(payload)
    csv_name = _dataset_cache_csv_file_name(
        name,
        stored_origin,
        stored_dev,
        cumulative=DEFAULT_CUMULATIVE,
        calendar=DEFAULT_CALENDAR,
    )
    csv_path = rc_dir / DATASET_CACHE_DIR / csv_name
    _write_csv_matrix(csv_path, payload["values"])

    updated_at = payload.get("modified") or utc_now_text()
    method_source_kind = _clean_name(payload.get("source_kind"))
    is_berquist_sherman = method_source_kind in {BS_SR_SOURCE_KIND, BS_CRA_SOURCE_KIND}
    source_kind = method_source_kind if is_berquist_sherman else _triangle_source_kind(name, dataset_type)
    is_app_calculated = (not is_berquist_sherman) and source_kind == "calculated"
    meta_path = rc_dir / DATASET_SIDECAR_DIR / _json_sidecar_name(name)
    existing = _safe_read_json(meta_path)
    meta = {
        "dataset_name": name,
        "dataset_type": dataset_type,
        "dataset_category": _normalize_import_name(payload.get("category")),
        "reserving_class": rc_path,
        "project_name": PROJECT_NAME,
        "source_kind": source_kind,
        "calculated": is_berquist_sherman or is_app_calculated,
        "source": (
            "resq_berquist_sherman_sr_triangle"
            if method_source_kind == BS_SR_SOURCE_KIND
            else "resq_berquist_sherman_cra_triangle"
            if method_source_kind == BS_CRA_SOURCE_KIND
            else "resq_triangle"
        ),
        "data_format": "Triangle",
        "origin_length": origin_length,
        "development_length": dev_length,
        # Display and stored shapes both follow ResQ's own: the display pair
        # is what ResQ showed, the stored pair is the shape the CSV holds.
        **stored_length_fields("Triangle", stored_origin, stored_dev),
        "development_count": payload.get("development_count", 0),
        "origin_labels": payload.get("origin_labels", []),
        "development_labels": payload.get("development_labels", []),
        "cumulative": DEFAULT_CUMULATIVE,
        "calendar": DEFAULT_CALENDAR,
        "show_subtotal": normalize_show_subtotal(existing.get("show_subtotal")),
        "number_format": dataset_type_number_format(rc_path, dataset_type),
        "decimal_places": dataset_type_decimal_places(rc_path, dataset_type),
        "csv_file": csv_name,
        "created": payload.get("created", ""),
        "modified_by": payload.get("user", ""),
        "notes": str(payload.get("notes") or ""),
        "updated_at": updated_at,
    }
    if is_berquist_sherman:
        meta["method_name"] = _normalize_import_name(payload.get("method_name")) or name
        meta["method_type"] = _method_type_name(payload.get("method_type"))
        meta["precedents"] = dependency_entries([
            _normalize_import_name(item)
            for item in payload.get("precedents", [])
            if _normalize_import_name(item)
        ])
        meta["dependents"] = []
        meta["status"] = normalize_method_status(payload.get("status"))
        _apply_graph_meta_best_effort(meta, dataset_type, rc_dir, preserve_precedents=True)
    else:
        _apply_graph_meta_best_effort(meta, dataset_type, rc_dir)
    _write_sidecar_json(meta_path, meta)
    return csv_path


def _bs_indexed_value(method, member_name: str, *indices: int):
    keyword_names = (
        ("DevIndex",)
        if len(indices) == 1
        else ("OriginIndex", "DevIndex")
    )
    keyword_args = {
        name: value
        for name, value in zip(keyword_names, indices)
    }
    try:
        value = _try_call_member(
            method,
            member_name,
            [
                (tuple(indices), {}),
                ((), keyword_args),
            ],
        )
    except Exception as exc:
        if _STRICT_RESQ_EXTRACTION.get():
            index_text = ", ".join(str(index) for index in indices)
            _strict_failure(
                f"Could not read Berquist Sherman {member_name} at index ({index_text}).",
                exc,
            )
        raise
    if value is None and _STRICT_RESQ_EXTRACTION.get():
        index_text = ", ".join(str(index) for index in indices)
        _strict_failure(
            f"Berquist Sherman {member_name} returned no value at index ({index_text})."
        )
    return value


def _bs_source_name(method, attr_name: str) -> str:
    context = f"Berquist Sherman {attr_name}"
    source = _extract_attr(method, attr_name, None, context="Berquist Sherman method")
    name = _normalize_import_name(_extract_attr(source, "Name", "", context=context))
    if not name and _STRICT_RESQ_EXTRACTION.get():
        _strict_failure(f"{context} does not expose a source dataset name.")
    return name


def _bs_loess_span(method) -> int:
    raw = _extract_attr(method, "LoessSpan", None, context="Berquist Sherman method")
    try:
        span = int(raw)
    except (TypeError, ValueError) as exc:
        if _STRICT_RESQ_EXTRACTION.get():
            _strict_failure(
                f"Could not convert Berquist Sherman method.LoessSpan={raw!r} to an integer.",
                exc,
            )
        return BS_DEFAULT_LOESS_SPAN
    return min(BS_MAX_LOESS_SPAN, max(BS_MIN_LOESS_SPAN, span))


def _bs_selection_label(value: object, labels: dict[int, str], field_name: str) -> str:
    try:
        code = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Invalid ResQ {field_name} value: {value!r}.") from exc
    if code not in labels:
        raise ValueError(f"Unsupported ResQ {field_name} code: {code}.")
    return labels[code]


def _bs_precedents(method_tab: dict, variant: str) -> list[str]:
    keys = (
        ("paid_claims", "closed_claim_numbers", "ultimate_claim_numbers")
        if variant == "sr"
        else ("reported_claim_numbers", "closed_claim_numbers", "incurred_claims", "paid_claims")
    )
    names: list[str] = []
    seen: set[str] = set()
    for key in keys:
        name = _normalize_import_name(method_tab.get(key))
        name_key = name.casefold()
        if name and name_key not in seen:
            seen.add(name_key)
            names.append(name)
    return names


def _bs_variant_from_payload(payload: dict) -> str:
    json_format = _clean_name(payload.get("json_format")).casefold()
    if json_format == BS_SR_JSON_FORMAT:
        return "sr"
    if json_format == BS_CRA_JSON_FORMAT:
        return "cra"
    raise ValueError(f"Unsupported Berquist Sherman JSON format: {json_format!r}.")


@_strict_extractor
def export_berquist_sherman(
    method,
    variant: str,
    output_payload: dict,
    *,
    strict: bool = False,
) -> dict:
    """Extract the annual B&S configuration needed to reproduce a ResQ output."""
    del strict
    clean_variant = _clean_name(variant).casefold()
    if clean_variant not in {"sr", "cra"}:
        raise ValueError(f"Unsupported Berquist Sherman variant: {variant!r}.")

    origin_length = int(
        output_payload.get("origin_length")
        or _extract_int_attr(method, "OriginLength", 0, context="Berquist Sherman method")
    )
    development_length = int(
        output_payload.get("development_length")
        or _extract_int_attr(method, "DevelopmentLength", 0, context="Berquist Sherman method")
    )
    if origin_length != 12 or development_length != 12:
        raise ValueError(
            "ArcRho's Berquist Sherman MVP supports annual triangles only "
            f"(got origin_length={origin_length}, development_length={development_length})."
        )

    name = _normalize_import_name(output_payload.get("name")) or _normalize_import_name(
        _extract_attr(method, "Name", "", context="Berquist Sherman method")
    )
    if not name:
        raise ValueError("The ResQ Berquist Sherman method does not expose an output name.")
    output_type = _normalize_import_name(output_payload.get("dataset_type")) or name
    origin_labels = [
        _normalize_import_name(label)
        for label in output_payload.get("origin_labels", [])
    ]
    development_labels = [
        _normalize_import_name(label)
        for label in output_payload.get("development_labels", [])
    ]
    if _STRICT_RESQ_EXTRACTION.get() and (
        any(not label for label in origin_labels)
        or any(not label for label in development_labels)
    ):
        _strict_failure(f"Berquist Sherman method {name!r} has incomplete annual triangle labels.")
    origin_count = len(origin_labels)
    development_count = len(development_labels)
    if origin_count <= 0 or development_count <= 0:
        raise ValueError(f"Berquist Sherman method {name!r} does not expose annual triangle labels.")

    if clean_variant == "sr":
        method_type = BS_SR_METHOD_TYPE
        source_kind = BS_SR_SOURCE_KIND
        method_tab = {
            "paid_claims": _bs_source_name(method, "PaidClaims"),
            "closed_claim_numbers": _bs_source_name(method, "ClosedClaimNos"),
            "ultimate_claim_numbers": _bs_source_name(method, "UltimateClaimNos"),
            "origin_labels": origin_labels,
            "development_labels": development_labels,
            "selected_proportion_settled": [
                float(_bs_indexed_value(method, "SelectedProportionSettled", dev_index))
                for dev_index in range(1, development_count + 1)
            ],
            "selected_proportion_is_default": [
                _bool_value(_bs_indexed_value(method, "IsDefaultProportionSettled", dev_index))
                for dev_index in range(1, development_count + 1)
            ],
            "selected_adjustment": [],
        }
        selected_adjustment: list[list[str | None]] = []
        for origin_index in range(1, origin_count + 1):
            row_count = min(development_count, origin_count - origin_index + 1)
            row: list[str | None] = []
            for dev_index in range(1, development_count + 1):
                if dev_index > row_count:
                    row.append(None)
                    continue
                raw = _bs_indexed_value(method, "SelectedAdjustment", origin_index, dev_index)
                row.append(_bs_selection_label(raw, BS_SR_ADJUSTMENT_TYPES, "SelectedAdjustment"))
            selected_adjustment.append(row)
        method_tab["selected_adjustment"] = selected_adjustment
        method_tab["loess_span"] = _bs_loess_span(method)
        json_format = BS_SR_JSON_FORMAT
    else:
        method_type = BS_CRA_METHOD_TYPE
        source_kind = BS_CRA_SOURCE_KIND
        method_tab = {
            "reported_claim_numbers": _bs_source_name(method, "ReportedClaimNos"),
            "closed_claim_numbers": _bs_source_name(method, "ClosedClaimNos"),
            "incurred_claims": _bs_source_name(method, "IncurredClaims"),
            "paid_claims": _bs_source_name(method, "PaidClaims"),
            "origin_labels": origin_labels,
            "development_labels": development_labels,
            "inflation_selection": [
                _bs_selection_label(
                    _bs_indexed_value(method, "SelectedAvgInflation", dev_index),
                    BS_CRA_INFLATION_TYPES,
                    "SelectedAvgInflation",
                )
                for dev_index in range(1, development_count + 1)
            ],
            "user_inflation": [
                float(_bs_indexed_value(method, "UserAvgInflation", dev_index))
                for dev_index in range(1, development_count + 1)
            ],
            "average_case_reserve_selection": [
                _bs_selection_label(
                    _bs_indexed_value(method, "SelectedAvgCaseReserves", dev_index),
                    BS_CRA_AVERAGE_CASE_RESERVE_TYPES,
                    "SelectedAvgCaseReserves",
                )
                for dev_index in range(1, development_count + 1)
            ],
            "user_average_case_reserves": [
                float(_bs_indexed_value(method, "UserAvgCaseReserves", dev_index))
                for dev_index in range(1, development_count + 1)
            ],
            "loess_span": _bs_loess_span(method),
        }
        json_format = BS_CRA_JSON_FORMAT

    notes = _clean_name(
        _extract_attr(method, "Notes", "", context=f"Berquist Sherman method {name!r}")
    )
    modified = output_payload.get("modified") or utc_now_text()
    output_triangle = _extract_attr(
        method,
        "OutputTriangle",
        None,
        context=f"Berquist Sherman method {name!r}",
    )
    return {
        "json_format": json_format,
        "details_tab": {
            "name": name,
            "method_type": method_type,
            "output_type": output_type,
            "origin_length": origin_length,
            "development_length": development_length,
        },
        "method_tab": method_tab,
        "_sidecar_notes": notes,
        "_sidecar_status": normalize_method_status(
            _extract_int_attr(
                output_triangle,
                "Status",
                0,
                context=f"Berquist Sherman method {name!r} OutputTriangle",
            )
        ),
        "method_metadata": {
            "method_type": method_type,
            "source_kind": source_kind,
            "last_modified": modified,
        },
    }


def _apply_berquist_sherman_triangle_metadata(payload: dict, method_payload: dict) -> None:
    payload["notes"] = str(method_payload.pop("_sidecar_notes", "") or "")
    payload["status"] = normalize_method_status(
        method_payload.pop("_sidecar_status", payload.get("status"))
    )
    variant = _bs_variant_from_payload(method_payload)
    if variant == "sr":
        payload["source_kind"] = BS_SR_SOURCE_KIND
        payload["method_type"] = BS_SR_METHOD_TYPE
        payload["method_type_code"] = METHOD_TYPE_BS_SR_CODE
    elif variant == "cra":
        payload["source_kind"] = BS_CRA_SOURCE_KIND
        payload["method_type"] = BS_CRA_METHOD_TYPE
        payload["method_type_code"] = METHOD_TYPE_BS_CRA_CODE
    else:
        raise ValueError(f"Unsupported Berquist Sherman variant: {variant!r}.")
    details_tab = method_payload.get("details_tab") if isinstance(method_payload.get("details_tab"), dict) else {}
    method_tab = method_payload.get("method_tab") if isinstance(method_payload.get("method_tab"), dict) else {}
    payload["method_name"] = _normalize_import_name(details_tab.get("name")) or _normalize_import_name(
        payload.get("name")
    )
    payload["precedents"] = _bs_precedents(method_tab, variant)


def _backfill_berquist_sherman_precedent_origin_labels(
    payload: dict,
    variant: str,
    rc_dir: Path,
) -> None:
    method_tab = payload.get("method_tab") if isinstance(payload.get("method_tab"), dict) else {}
    origin_labels = method_tab.get("origin_labels")
    if not isinstance(origin_labels, list) or not origin_labels:
        return

    canonical_labels = [str(label) for label in origin_labels]
    for precedent in _bs_precedents(method_tab, variant):
        sidecar_path = rc_dir / DATASET_SIDECAR_DIR / _json_sidecar_name(precedent)
        if not sidecar_path.is_file():
            continue
        sidecar = _safe_read_json(sidecar_path)
        if not sidecar:
            continue
        existing_labels = sidecar.get("origin_labels")
        if existing_labels not in (None, []):
            continue
        sidecar["origin_labels"] = canonical_labels
        _write_sidecar_json(sidecar_path, sidecar)


def _dataset_number_format_entry(rc_dir: Path, dataset_name: object) -> dict:
    """The format a dataset instance displays with, as the frontend reads it.

    The sidecar is the source of truth because the user can restyle a dataset
    after import; the shared dataset-type preference only seeds a new one.
    """
    name = _normalize_import_name(dataset_name)
    if not name:
        return number_format_entry(None)
    sidecar = _safe_read_json(rc_dir / DATASET_SIDECAR_DIR / _json_sidecar_name(name))
    if isinstance(sidecar, dict) and str(sidecar.get("number_format") or "").strip():
        return number_format_entry(sidecar.get("number_format"), sidecar.get("decimal_places"))
    dataset_type = _normalize_import_name(
        sidecar.get("dataset_type") if isinstance(sidecar, dict) else ""
    ) or name
    return number_format_entry(dataset_type_number_format(rc_dir, dataset_type))


def _berquist_sherman_number_formats(payload: dict, variant: str, rc_dir: Path) -> dict:
    """The recorded formats a B&S method page renders its grids with.

    ``derived`` is the output dataset's own format, used for every calculated
    triangle; each source entry mirrors that input dataset instance.
    """
    method_tab = payload.get("method_tab") if isinstance(payload.get("method_tab"), dict) else {}
    details_tab = payload.get("details_tab") if isinstance(payload.get("details_tab"), dict) else {}
    output_type = _normalize_import_name(details_tab.get("output_type")) or _normalize_import_name(
        details_tab.get("name")
    )
    role_keys = BS_SOURCE_ROLES[variant]
    return {
        "derived": number_format_entry(dataset_type_number_format(rc_dir, output_type)),
        "sources": {
            role: _dataset_number_format_entry(rc_dir, method_tab.get(role))
            for role in role_keys
        },
    }


def write_berquist_sherman_export(payload: dict, rc_path: str, rc_dir: Path) -> Path:
    del rc_path
    variant = _bs_variant_from_payload(payload)
    details_tab = payload.get("details_tab") if isinstance(payload.get("details_tab"), dict) else {}
    name = _normalize_import_name(details_tab.get("name"))
    if variant == "sr":
        prefix = BS_SR_FILE_PREFIX
    elif variant == "cra":
        prefix = BS_CRA_FILE_PREFIX
    else:
        raise ValueError(f"Unsupported Berquist Sherman variant: {variant!r}.")
    if not name:
        raise ValueError("Berquist Sherman method JSON is missing details_tab.name.")
    out_path = rc_dir / METHOD_DATA_DIR / f"{prefix}{_encode_name_part(name)}.json"
    method_payload = dict(payload)
    method_payload.pop("_sidecar_notes", None)
    method_payload.pop("_sidecar_status", None)
    # The recorded formats need the reserving class on disk, so they are filled
    # in here rather than during the COM extraction.
    method_tab = method_payload.get("method_tab")
    if isinstance(method_tab, dict):
        method_payload["method_tab"] = {
            **method_tab,
            "number_formats": _berquist_sherman_number_formats(method_payload, variant, rc_dir),
        }
    _write_json(out_path, method_payload)
    _backfill_berquist_sherman_precedent_origin_labels(method_payload, variant, rc_dir)
    return out_path


def _find_unique_method_by_output(
    collection,
    direct_candidates,
    output_name: str,
    output_member: str,
    method_label: str,
):
    """Return the one method that actually owns the requested output.

    Some ResQ collection getters resolve a method name before an output name.
    A direct lookup is therefore only a compatibility fallback when the COM
    collection cannot be enumerated, and every candidate is still validated
    against its output object.  A complete enumeration also lets selective
    migration fail closed instead of choosing arbitrarily when corrupt or
    unusual ResQ data exposes duplicate output owners.
    """

    target = _normalize_import_name(output_name).casefold()
    if not target:
        return None

    enumerated = None
    if collection is not None:
        try:
            enumerated = list(collection)
        except Exception:
            # Older COM collection wrappers are not always iterable. Keep the
            # canonical bulk migration tolerant, but never trust their direct
            # lookup without checking the returned method's output identity.
            pass

    candidates = enumerated if enumerated is not None else list(direct_candidates)
    matches = []
    direct_identities: set[str] = set()
    for method in candidates:
        output = _safe_attr(method, output_member, None)
        candidate_output = _normalize_import_name(
            _safe_attr(output, "Name", "")
        ).casefold()
        if candidate_output != target:
            continue
        if enumerated is None:
            # Getter and Item() can return separate Python wrappers for the
            # same COM method. Method names are unique in ResQ and provide the
            # stable identity needed to avoid treating those aliases as two
            # different output owners.
            method_name = _normalize_import_name(
                _safe_attr(method, "Name", "")
            ).casefold()
            identity = method_name or f"object:{id(method)}"
            if identity in direct_identities:
                continue
            direct_identities.add(identity)
        matches.append(method)

    if len(matches) > 1:
        raise ValueError(
            f"Multiple ResQ {method_label} methods produce output {output_name!r}."
        )
    return matches[0] if matches else None


def _find_berquist_sherman_for_triangle(
    reserving_class,
    triangle_name: str,
    method_type_code: int,
) -> tuple[str, object] | None:
    if method_type_code == METHOD_TYPE_BS_SR_CODE:
        variants = (("sr", "GetBerquistShermanSR", "BerquistShermanSRs"),)
    elif method_type_code == METHOD_TYPE_BS_CRA_CODE:
        variants = (("cra", "GetBerquistShermanCRA", "BerquistShermanCRAs"),)
    else:
        return None

    target = _normalize_import_name(triangle_name).casefold()
    for variant, getter_name, collection_name in variants:
        direct_candidates = []
        try:
            method = _call_member(reserving_class, getter_name, triangle_name)
            if method is not None:
                direct_candidates.append(method)
        except Exception:
            pass
        try:
            collection = _call_member(reserving_class, collection_name)
        except Exception:
            collection = None
        if collection is not None:
            try:
                method = collection.Item(triangle_name)
                if method is not None:
                    direct_candidates.append(method)
            except Exception:
                pass
        method = _find_unique_method_by_output(
            collection,
            direct_candidates,
            target,
            "OutputTriangle",
            "Berquist Sherman",
        )
        if method is not None:
            return variant, method
    return None


def _read_vector_body(vector, name: str) -> dict:
    """Read a vector's rows, values, and labels at the period it is showing."""
    origin_count = _vector_origin_count(vector)
    if origin_count <= 0:
        raise ValueError(f"Vector {name!r} does not expose a positive OriginCount/Count.")

    values: list[list] = []
    attempted_cells = 0
    value_errors: list[Exception] = []
    for i in range(1, origin_count + 1):
        attempted_cells += 1
        try:
            values.append([_vector_value(vector, i)])
        except Exception as exc:
            if _STRICT_RESQ_EXTRACTION.get():
                _strict_failure(
                    f"Could not read vector {name!r} value at origin index {i}.",
                    exc,
                )
            value_errors.append(exc)
            values.append([None])
    if attempted_cells > 0 and len(value_errors) == attempted_cells:
        raise ValueError(f"Failed to read any values for vector {name!r}: {value_errors[0]}")

    return {
        "origin_count": origin_count,
        "origin_labels": [_vector_origin_label(vector, i) for i in range(1, origin_count + 1)],
        "values": values,
    }


@_strict_extractor
def export_vector(vector, *, strict: bool = False, at_stored_shape: bool = False) -> dict:
    """Extract a ResQ Vector COM object into ArcRho CSV values and metadata.

    ``origin_length`` in the payload is the period ResQ displays the vector
    at; ``stored_period_length`` is the period its ``values`` were read at.
    By default that is the displayed period. With ``at_stored_shape`` the
    vector is shown at ResQ's stored period while it is read, so a monthly
    vector displayed yearly is copied month by month. Should ResQ refuse the
    switch, the values are read at the displayed period and the payload says
    so.
    """
    del strict
    name = _normalize_import_name(_extract_attr(vector, "Name", "", context="vector"))
    dataset_type_obj = _extract_attr(vector, "DatasetType", None, context=f"vector {name!r}")
    dataset_type = _normalize_import_name(
        _extract_attr(dataset_type_obj, "Name", "", context=f"vector {name!r} DatasetType")
    ) or name
    category_obj = _extract_attr(
        dataset_type_obj,
        "Category",
        None,
        context=f"vector {name!r} DatasetType",
    )
    category = _normalize_import_name(
        _extract_attr(category_obj, "Name", "", context=f"vector {name!r} Category")
    )
    data_format = _extract_int_attr(
        dataset_type_obj,
        "DataFormat",
        1,
        context=f"vector {name!r} DatasetType",
    )
    method_type_code = _extract_int_attr(vector, "MethodType", -1, context=f"vector {name!r}")
    method_type = _method_type_name(method_type_code)
    # ResQ vectors expose their period granularity (in months) as PeriodLength. They do
    # not have the triangle/method-style OriginLength member, so reading "OriginLength"
    # here always missed and fell back to the default 12. Use PeriodLength, falling back
    # to OriginLength then 12 only if PeriodLength is unavailable. A vector is 1-D, so the
    # same period length applies to both the origin and (nominal) development axis.
    period_length = _extract_int_attr(vector, "PeriodLength", 0, context=f"vector {name!r}")
    if period_length <= 0:
        period_length = _extract_int_attr(vector, "OriginLength", 12, context=f"vector {name!r}")
    origin_length = period_length
    dev_length = period_length
    stored_period, _ = resq_stored_lengths(vector, is_vector=True, origin_length=period_length)
    read_period = period_length
    switch: dict[str, int] = {}
    if at_stored_shape and stored_period != period_length:
        switch["PeriodLength"] = stored_period
    with _displayed_at(vector, switch) as switched:
        if switch and switched:
            read_period = stored_period
        body = _read_vector_body(vector, name)

    user = _normalize_import_name(_extract_attr(vector, "User", "", context=f"vector {name!r}"))
    created = _iso_or_text(_extract_attr(vector, "Created", "", context=f"vector {name!r}"))
    modified = _iso_or_text(_extract_attr(vector, "Modified", "", context=f"vector {name!r}"))
    notes = str(_extract_attr(vector, "Notes", "", context=f"vector {name!r}") or "")
    formula = _clean_name(_extract_attr(vector, "Formula", "", context=f"vector {name!r}"))

    return {
        "name": name,
        "dataset_type": dataset_type,
        "category": category,
        "data_format": data_format,
        "method_type": method_type,
        "method_type_code": method_type_code,
        "origin_length": origin_length,
        "development_length": dev_length,
        "stored_period_length": read_period,
        "origin_count": body["origin_count"],
        "development_count": 1,
        "origin_labels": body["origin_labels"],
        "development_labels": ["Value"],
        "values": body["values"],
        "formula": formula,
        "user": user,
        "created": created,
        "modified": modified,
        "notes": notes,
        "status": normalize_method_status(
            _extract_attr(vector, "Status", 0, context=f"vector {name!r}")
        ),
    }

def _vector_payload_period_length(payload: dict) -> int:
    return int(payload.get("period_length") or payload.get("origin_length") or 0)


def _vector_payload_stored_period_length(payload: dict) -> int:
    """The period the payload's ``values`` were read at: its CSV's own shape."""
    return int(payload.get("stored_period_length") or _vector_payload_period_length(payload))


def _triangle_payload_stored_lengths(payload: dict) -> tuple[int, int]:
    """The shape the payload's ``values`` were read at: its CSV's own shape."""
    origin_length = int(payload.get("origin_length") or 0)
    dev_length = int(payload.get("development_length") or 0)
    return (
        int(payload.get("stored_origin_length") or origin_length),
        int(payload.get("stored_development_length") or dev_length),
    )


def _vector_payload_row_count(payload: dict) -> int:
    """The vector's actual cell/row count (ResQ ``OriginCount``).

    Not ``origin_length``/``period_length``: those hold the number of months in
    one origin/development period (12 for annual data), never how many rows the
    vector has.
    """
    count = int(payload.get("origin_count") or 0)
    if count <= 0:
        labels = payload.get("origin_labels")
        count = len(labels) if isinstance(labels, list) else 0
    if count <= 0:
        values = payload.get("values")
        count = len(values) if isinstance(values, list) else 0
    return count


# One ResQ instance-formula token: a double-quoted dataset name, a number, or
# an operator/parenthesis, each after optional whitespace. Anything else in the
# text makes the formula untranslatable.
_RESQ_INSTANCE_FORMULA_TOKEN_RE = re.compile(
    r'\s*(?:"(?P<name>[^"]+)"'
    r"|(?P<number>(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)"
    r"|(?P<op>[+\-*/^()]))"
)


def _translated_instance_formula_links(
    payload: dict,
    dataset_name: str,
    row_count: int,
    known_instance_names: object,
) -> list[dict] | None:
    """Translate a ResQ instance formula into one ArcRho in-cell formula link.

    ResQ keeps per-instance formulas on vectors whose dataset type ArcRho
    treats as a plain input — quoted dataset names combined with arithmetic,
    ``"C 91 - Current Qtr Indicated" * "H 01 - ..." / 1000``. Each quoted name
    becomes a whole-vector dataset reference (``[C 91 - ...][1:N]`` where N is
    the vector's row count, never its period length in months), the text
    is canonicalized through ``arcrho_api.dataset_link_contract`` so a later
    save round-trips it byte for byte, and the link owns every cell of the
    vector. The translation is all-or-nothing: a name not among
    ``known_instance_names`` (a frozen prior-quarter snapshot ArcRho never
    imports, say), a self-reference, or any text outside the token grammar
    falls back to the hardcoded imported values by returning ``None``.
    """

    formula = _clean_name(payload.get("formula"))
    if not formula or row_count <= 0 or known_instance_names is None:
        return None
    known_keys = {
        _canon_dataset_name(name)
        for name in known_instance_names
        if _canon_dataset_name(name)
    }
    own_key = _canon_dataset_name(dataset_name)
    pieces: list[str] = []
    referenced = False
    cursor = 0
    while cursor < len(formula) and formula[cursor:].strip():
        match = _RESQ_INSTANCE_FORMULA_TOKEN_RE.match(formula, cursor)
        if not match or match.end() == cursor:
            return None
        if match.group("name") is not None:
            referenced_name = _normalize_import_name(match.group("name"))
            key = _canon_dataset_name(referenced_name)
            if not key or key == own_key or key not in known_keys:
                return None
            pieces.append(f"[{referenced_name}][1:{row_count}]")
            referenced = True
        elif match.group("number") is not None:
            pieces.append(match.group("number"))
        else:
            pieces.append(match.group("op"))
        cursor = match.end()
    if not referenced:
        return None
    try:
        canonical = canonical_dataset_formula("=" + " ".join(pieces))
    except DatasetLinkError:
        return None
    return [{
        "formula": canonical,
        "target_cells": [
            {"row": row, "column": 0, "result_row": row, "result_column": 0}
            for row in range(row_count)
        ],
    }]


def write_vector_export(
    payload: dict,
    rc_path: str,
    rc_dir: Path,
    *,
    bf_method_payload: dict | None = None,
    cc_method_payload: dict | None = None,
    known_instance_names: object = None,
) -> Path:
    name = _normalize_import_name(payload["name"])
    dataset_type = _normalize_import_name(payload.get("dataset_type")) or name
    period_length = _vector_payload_period_length(payload)
    # The CSV is written at the period the values were read at, which is
    # ResQ's stored period for a hand-entered vector; the display period is
    # the one ResQ showed it at, served from the coarser copies written below.
    stored_period = _vector_payload_stored_period_length(payload)
    csv_name = _vector_cache_csv_file_name(name, stored_period)
    csv_path = rc_dir / DATASET_CACHE_DIR / csv_name
    _write_csv_matrix(csv_path, payload["values"])
    _write_aggregated_vector_cache_exports(payload, rc_dir)

    method_type = _method_type_name(payload.get("method_type"))
    is_result_selection = _is_result_selection_method_type(method_type)
    raw_method_type_code = _method_type_code(method_type, -1)
    is_bornhuetter_ferguson = _clean_name(payload.get("source_kind")) == BF_SOURCE_KIND
    is_cape_cod = _clean_name(payload.get("source_kind")) == CC_SOURCE_KIND
    if is_bornhuetter_ferguson:
        meta_method_type = BF_METHOD_TYPE
        meta_method_type_code = METHOD_TYPE_BF_CODE
    elif is_cape_cod:
        meta_method_type = CC_METHOD_TYPE
        meta_method_type_code = METHOD_TYPE_CAPE_COD_CODE
    elif raw_method_type_code in {METHOD_TYPE_BF_CODE, METHOD_TYPE_CAPE_COD_CODE}:
        # A method-coded vector without its exported method imports as a plain dataset.
        meta_method_type = "None"
        meta_method_type_code = METHOD_TYPE_NONE_CODE
    else:
        meta_method_type = method_type
        meta_method_type_code = payload.get("method_type_code", _method_type_code(method_type, 0))
    is_method_output = is_bornhuetter_ferguson or is_cape_cod
    is_engine_generated = (not is_result_selection) and (not is_method_output) and _is_generated_dataset_type(dataset_type)
    # ArcRho's own dataset-types library decides whether the dataset is
    # calculated. The ResQ Formula on the vector is deliberately ignored: a
    # formula only ResQ knows (a prior-quarter lookup, say) is one ArcRho can
    # neither store nor evaluate, and stamping the sidecar "calculated" for it
    # left the dataset read-only with nothing able to regenerate it.
    is_app_calculated = (
        (not is_result_selection)
        and (not is_method_output)
        and (not is_engine_generated)
        and _is_calculated_dataset_type(dataset_type)
    )
    updated_at = payload.get("modified") or utc_now_text()
    if is_result_selection:
        source_kind = "result_selection"
    elif is_bornhuetter_ferguson:
        source_kind = BF_SOURCE_KIND
    elif is_cape_cod:
        source_kind = CC_SOURCE_KIND
    elif is_engine_generated:
        source_kind = "engine"
    elif is_app_calculated:
        source_kind = "calculated"
    else:
        source_kind = "input"
    meta_path = rc_dir / DATASET_SIDECAR_DIR / _json_sidecar_name(name)
    existing = _safe_read_json(meta_path)
    meta = {
        "dataset_name": name,
        "dataset_type": dataset_type,
        "dataset_category": _normalize_import_name(payload.get("category")),
        "reserving_class": rc_path,
        "project_name": PROJECT_NAME,
        "source_kind": source_kind,
        "calculated": bool(is_app_calculated or is_result_selection or is_method_output),
        "source": (
            "resq_result_selection_vector"
            if is_result_selection
            else "resq_bornhuetter_ferguson_vector"
            if is_bornhuetter_ferguson
            else "resq_cape_cod_vector"
            if is_cape_cod
            else "resq_vector"
        ),
        "method_type": meta_method_type,
        "data_format": "Vector",
        "period_length": period_length,
        # Display and stored periods both follow ResQ's own: the display
        # period is what ResQ showed, the stored period is what the CSV holds.
        **stored_length_fields("Vector", stored_period),
        "show_subtotal": normalize_show_subtotal(existing.get("show_subtotal")),
        "origin_labels": payload.get("origin_labels", []),
        "development_labels": payload.get("development_labels", []),
        "number_format": dataset_type_number_format(rc_path, dataset_type),
        "decimal_places": dataset_type_decimal_places(rc_path, dataset_type),
        "csv_file": csv_name,
        "created": payload.get("created", ""),
        "modified_by": payload.get("user", ""),
        "notes": str(payload.get("notes") or ""),
        "updated_at": updated_at,
    }
    if is_bornhuetter_ferguson and isinstance(bf_method_payload, dict):
        publication_revision = _clean_name(
            bf_method_payload.get("method_metadata", {}).get("publication_revision")
            if isinstance(bf_method_payload.get("method_metadata"), dict)
            else ""
        )
        output_changed = _clean_name(existing.get("publication_revision")) != publication_revision
        meta = build_bornhuetter_ferguson_output_sidecar(
            bf_method_payload,
            project_name=PROJECT_NAME,
            reserving_class=rc_path,
            csv_file=csv_name,
            existing=existing,
            notes=str(payload.get("notes") or ""),
            timestamp=updated_at,
            user=payload.get("user", ""),
            output_changed=output_changed,
            append_audit=not existing or output_changed,
            status=normalize_method_status(payload.get("status")),
        )
    elif is_cape_cod and isinstance(cc_method_payload, dict):
        publication_revision = _clean_name(
            cc_method_payload.get("method_metadata", {}).get("publication_revision")
            if isinstance(cc_method_payload.get("method_metadata"), dict)
            else ""
        )
        output_changed = _clean_name(existing.get("publication_revision")) != publication_revision
        meta = build_cape_cod_output_sidecar(
            cc_method_payload,
            project_name=PROJECT_NAME,
            reserving_class=rc_path,
            csv_file=csv_name,
            existing=existing,
            notes=str(payload.get("notes") or ""),
            timestamp=updated_at,
            user=payload.get("user", ""),
            output_changed=output_changed,
            append_audit=not existing or output_changed,
            status=normalize_method_status(payload.get("status")),
        )
    elif is_method_output:
        # A BF/Cape Cod-coded vector without a matching exported method is an
        # ordinary imported dataset, not a method publication. Preserve the
        # legacy fallback rather than manufacturing an incomplete canonical
        # method sidecar.
        meta["status"] = normalize_method_status(payload.get("status"))
        meta["precedents"] = dependency_entries([
            _normalize_import_name(item)
            for item in payload.get("precedents", [])
            if _normalize_import_name(item)
        ])
        meta["dependents"] = []
    elif is_result_selection:
        meta["status"] = normalize_method_status(payload.get("status"))
        meta["precedents"] = dependency_entries([
            _normalize_import_name(item)
            for item in payload.get("precedents", [])
            if _normalize_import_name(item)
        ])
        meta["dependents"] = []
    else:
        if source_kind == "input":
            # A ResQ instance formula on a plain-input vector imports as an
            # in-cell formula link, so ArcRho re-evaluates it through the
            # dependent-propagation walk instead of freezing the copied
            # values; an untranslatable formula keeps today's hardcoded
            # values.
            translated = _translated_instance_formula_links(
                payload,
                name,
                _vector_payload_row_count(payload),
                known_instance_names,
            )
            if translated:
                meta["formula_links"] = translated
        _apply_graph_meta_best_effort(meta, dataset_type, rc_dir)
    _write_sidecar_json(meta_path, meta)
    return csv_path


def _engine_cache_created_at(csv_path: Path, fallback: str) -> str:
    """Match the app's engine sidecar `created` timestamp (CSV file ctime, UTC)."""
    try:
        ctime = csv_path.stat().st_ctime
    except OSError:
        return fallback
    return format_persisted_timestamp(datetime.fromtimestamp(ctime, timezone.utc))


def write_engine_generated_export(
    payload: dict,
    rc_path: str,
    rc_dir: Path,
    *,
    is_vector: bool,
    provenance: dict,
    csv_name: str,
    csv_path: Path,
) -> Path:
    """Write the canonical sidecar for a data-engine-generated dataset.

    The CSV at ``csv_path`` must already have been produced by the data-engine
    (see ``resq_migration.engine.generate_engine_csv``); this function only writes
    the JSON sidecar. Unlike the ResQ-copied writers, the sidecar is marked as a
    live engine cache (``source_kind='engine'`` with no ``resq_*`` source marker)
    and carries the authoritative processing provenance so the app treats the
    migrated cache as fresh rather than stale.
    """
    name = _normalize_import_name(payload["name"])
    dataset_type = _normalize_import_name(payload.get("dataset_type")) or name
    user = import_user_identity_service().get_current_display_name() or getpass.getuser()
    updated_at = utc_now_text()
    created = _engine_cache_created_at(csv_path, "")
    meta_path = rc_dir / DATASET_SIDECAR_DIR / _json_sidecar_name(name)
    existing = _safe_read_json(meta_path)
    meta = build_engine_dataset_sidecar(
        project_name=PROJECT_NAME,
        reserving_class=rc_path,
        dataset_name=name,
        dataset_type=dataset_type,
        data_format="Vector" if is_vector else "Triangle",
        csv_file=csv_name,
        user=user,
        created=created,
        updated_at=updated_at,
        number_format=dataset_type_number_format(rc_path, dataset_type),
        decimal_places=dataset_type_decimal_places(rc_path, dataset_type),
        origin_length=int(payload.get("origin_length") or 0),
        development_length=int(payload.get("development_length") or 0),
        period_length=_vector_payload_period_length(payload) if is_vector else None,
        # ResQ's own stored lengths: a generated dataset is stored at the
        # source data's granularity however coarsely ResQ displayed it, and
        # the Engine rebuilds it at any period from that same source table.
        stored_origin_length=int(payload.get("stored_origin_length") or 0) or None,
        stored_development_length=int(payload.get("stored_development_length") or 0) or None,
        stored_period_length=(
            int(payload.get("stored_period_length") or 0) or None if is_vector else None
        ),
        cumulative=DEFAULT_CUMULATIVE,
        calendar=DEFAULT_CALENDAR,
        show_subtotal=normalize_show_subtotal(existing.get("show_subtotal")),
        processing=provenance,
        source_modified=str(payload.get("modified") or "").strip(),
        audit_log=existing.get("audit_log") or (),
        audit_action=AUDIT_ACTION_UPDATE if existing else AUDIT_ACTION_INSERT,
    )

    _apply_graph_meta_best_effort(meta, dataset_type, rc_dir)
    _write_sidecar_json(meta_path, meta)
    return csv_path

def _result_selection_dataset_count(result_selection) -> int:
    errors: list[Exception] = []
    try:
        value = int(getattr(result_selection, "DatasetCount"))
        if value >= 0:
            return value
    except Exception as exc:
        errors.append(exc)
    try:
        value = int(_call_member(result_selection, "DatasetCount"))
        if value >= 0:
            return value
    except Exception as exc:
        errors.append(exc)
    if _STRICT_RESQ_EXTRACTION.get():
        _strict_failure(
            "Could not read Result Selection DatasetCount.",
            errors[0] if errors else None,
        )
    return 0


def _result_selection_origin_count(result_selection) -> int:
    errors: list[Exception] = []
    try:
        value = int(getattr(result_selection, "OriginCount"))
        if value > 0:
            return value
    except Exception as exc:
        errors.append(exc)
    try:
        value = int(_call_member(result_selection, "OriginCount"))
        if value > 0:
            return value
    except Exception as exc:
        errors.append(exc)
    if _STRICT_RESQ_EXTRACTION.get():
        _strict_failure(
            "Could not read a positive Result Selection OriginCount.",
            errors[0] if errors else None,
        )
    return 0

def _result_selection_origin_label(result_selection, origin_index: int) -> str:
    try:
        return _normalize_import_name(result_selection.OriginLabel(origin_index))
    except Exception as exc:
        if _STRICT_RESQ_EXTRACTION.get():
            _strict_failure(
                f"Could not read Result Selection origin label for origin index {origin_index}.",
                exc,
            )
        return str(origin_index)

def _result_selection_dataset(result_selection, dataset_index: int):
    return _call_member(result_selection, "Dataset", dataset_index)

def _result_selection_dataset_value(result_selection, dataset_index: int, origin_index: int, origin_length: int):
    call_shapes = [
        ((dataset_index, origin_index, origin_length), {}),
        ((), {"DatasetIndex": dataset_index, "OriginIndex": origin_index, "OriginLength": origin_length}),
    ]
    return _try_call_member(result_selection, "DatasetValues", call_shapes)

def _result_selection_weight(result_selection, dataset_index: int, origin_index: int):
    call_shapes = [
        ((dataset_index, origin_index), {}),
        ((), {"DatasetIndex": dataset_index, "OriginIndex": origin_index}),
    ]
    return _try_call_member(result_selection, "Weights", call_shapes)

def _result_selection_ultimate(result_selection, origin_index: int, origin_length: int):
    call_shapes = [
        ((origin_index, origin_length), {}),
        ((origin_index,), {}),
        ((), {"OriginIndex": origin_index, "OriginLength": origin_length}),
    ]
    return _try_call_member(result_selection, "Ultimates", call_shapes)

def _result_selection_ultimate_overridden(result_selection, origin_index: int) -> bool:
    call_shapes = [
        ((origin_index,), {}),
        ((), {"OriginIndex": origin_index}),
    ]
    return _bool_value(_try_call_member(result_selection, "UltimateOverridden", call_shapes))

def _result_selection_ratio_basis_dataset_name(result_selection) -> str:
    call_shapes = [
        ((1,), {}),
        ((), {"DatasetIndex": 1}),
        ((), {"arg0": 1}),
    ]
    try:
        dataset = _try_call_member(result_selection, "RatioBasisDataset", call_shapes)
    except Exception as exc:
        if _STRICT_RESQ_EXTRACTION.get():
            _strict_failure("Could not read Result Selection RatioBasisDataset.", exc)
        return ""
    return _normalize_import_name(
        _extract_attr(dataset, "Name", "", context="Result Selection RatioBasisDataset")
    )

def _result_selection_ratio_basis_value(result_selection, origin_index: int, origin_length: int):
    call_shapes = [
        ((origin_index, origin_length), {}),
        ((origin_index,), {}),
        ((), {"OriginIndex": origin_index, "OriginLength": origin_length}),
    ]
    return _try_call_member(result_selection, "RatioBasisValues", call_shapes)

def _rs_json_number(value):
    if value is None or isinstance(value, bool):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(number):
        return None
    if isinstance(value, int):
        return value
    # Carried whole, matching result_selection_service._round_number: the value
    # is a ResQ ultimate, and rounding the copy is what made ArcRho's weighted
    # average disagree with ResQ's.
    return 0.0 if number == 0 else number

def _result_selection_source_kind(name: str, dataset_type: str, data_format: str, method_type_code: int) -> str:
    if method_type_code == METHOD_TYPE_DFM_CODE:
        return "dfm"
    if method_type_code == METHOD_TYPE_RESULT_SELECTION_CODE:
        return "result_selection"
    if _clean_name(data_format).lower() == "triangle":
        return _triangle_source_kind(name, dataset_type)
    return "input"

def _result_selection_source_payload(result_selection, dataset_index: int, origin_count: int, origin_length: int) -> dict:
    dataset = _result_selection_dataset(result_selection, dataset_index)
    context = f"Result Selection source {dataset_index}"
    dataset_type_obj = _extract_attr(dataset, "DatasetType", None, context=context)
    name = _normalize_import_name(_extract_attr(dataset, "Name", "", context=context)) or f"Source {dataset_index}"
    dataset_type = _normalize_import_name(
        _extract_attr(dataset_type_obj, "Name", "", context=f"{context} DatasetType")
    )
    data_format_code = _extract_int_attr(
        dataset_type_obj,
        "DataFormat",
        -1,
        context=f"{context} DatasetType",
    )
    data_format = "Triangle" if data_format_code == 0 else "Vector"
    method_type_code = _extract_int_attr(
        dataset,
        "MethodType",
        METHOD_TYPE_NONE_CODE,
        context=context,
    )
    method_type = _method_type_name(method_type_code)
    values: list = []
    weights: list = []
    for origin_index in range(1, origin_count + 1):
        try:
            values.append(_rs_json_number(_result_selection_dataset_value(result_selection, dataset_index, origin_index, origin_length)))
        except Exception as exc:
            if _STRICT_RESQ_EXTRACTION.get():
                _strict_failure(
                    f"Could not read {context} value for origin index {origin_index}.",
                    exc,
                )
            values.append(None)
        try:
            weights.append(max(0.0, _rs_json_number(
                _result_selection_weight(result_selection, dataset_index, origin_index)
            ) or 0.0))
        except Exception as exc:
            if _STRICT_RESQ_EXTRACTION.get():
                _strict_failure(
                    f"Could not read {context} weight for origin index {origin_index}.",
                    exc,
                )
            weights.append(0)
    category_obj = _extract_attr(
        dataset_type_obj,
        "Category",
        None,
        context=f"{context} DatasetType",
    )
    return {
        "name": name,
        "dataset_type": dataset_type,
        "data_format": data_format,
        "method_type": method_type,
        "category": _normalize_import_name(
            _extract_attr(category_obj, "Name", "", context=f"{context} Category")
        ),
        "source_kind": _result_selection_source_kind(name, dataset_type, data_format, method_type_code),
        # Vectors do not consistently expose OriginLength; the enclosing
        # Result Selection's OriginLength is the canonical fallback, not a
        # failed cell/property read.
        "origin_length": max(1, _safe_int_attr(dataset, "OriginLength", origin_length)),
        "values": values,
        "weights": weights,
    }


def _result_selection_calculated_ultimate(loaded_datasets: list[dict], origin_count: int) -> list:
    ultimate: list = []
    for row_index in range(origin_count):
        numerator = 0.0
        denominator = 0.0
        for dataset in loaded_datasets:
            values = dataset.get("values") if isinstance(dataset.get("values"), list) else []
            weights = dataset.get("weights") if isinstance(dataset.get("weights"), list) else []
            try:
                value = float(values[row_index])
                weight = max(0.0, float(weights[row_index]))
            except (IndexError, TypeError, ValueError):
                continue
            if not math.isfinite(value) or not math.isfinite(weight) or weight <= 0:
                continue
            numerator += value * weight
            denominator += weight
        ultimate.append(_rs_json_number(numerator / denominator) if denominator > 0 else None)
    return ultimate


def _result_selection_selected_ultimate(calculated_ultimate: list, ultimate_overrides: list, origin_count: int) -> list:
    selected: list = []
    for row_index in range(origin_count):
        override = ultimate_overrides[row_index] if row_index < len(ultimate_overrides) else None
        selected.append(override if override is not None else calculated_ultimate[row_index])
    return selected


@_strict_extractor
def export_result_selection(result_selection, *, strict: bool = False) -> dict:
    """Extract a ResQ Result Selection method into ArcRho's method JSON shape."""
    del strict
    output_vector = _extract_attr(result_selection, "OutputVector", None, context="Result Selection")
    name = _normalize_import_name(
        _extract_attr(output_vector, "Name", "", context="Result Selection OutputVector")
    ) or _normalize_import_name(
        _extract_attr(result_selection, "Name", "", context="Result Selection")
    )
    dataset_type_obj = _extract_attr(
        output_vector,
        "DatasetType",
        None,
        context=f"Result Selection {name!r} OutputVector",
    )
    output_type = _normalize_import_name(
        _extract_attr(
            dataset_type_obj,
            "Name",
            "",
            context=f"Result Selection {name!r} output DatasetType",
        )
    ) or name
    origin_length = _extract_int_attr(
        result_selection,
        "OriginLength",
        12,
        context=f"Result Selection {name!r}",
    )
    origin_count = _result_selection_origin_count(result_selection)
    if origin_count <= 0:
        raise ValueError(f"Result Selection {name!r} does not expose a positive OriginCount.")
    dataset_count = _result_selection_dataset_count(result_selection)
    origin_labels = [_result_selection_origin_label(result_selection, i) for i in range(1, origin_count + 1)]
    loaded_datasets = [
        _result_selection_source_payload(result_selection, dataset_index, origin_count, origin_length)
        for dataset_index in range(1, dataset_count + 1)
    ]
    ratio_basis_dataset = _result_selection_ratio_basis_dataset_name(result_selection)
    ratio_basis_datasets = [ratio_basis_dataset] if ratio_basis_dataset else []
    ratio_basis_values = []
    if ratio_basis_dataset:
        values = []
        for origin_index in range(1, origin_count + 1):
            try:
                values.append(_rs_json_number(
                    _result_selection_ratio_basis_value(result_selection, origin_index, origin_length)
                ))
            except Exception as exc:
                if _STRICT_RESQ_EXTRACTION.get():
                    _strict_failure(
                        f"Could not read Result Selection {name!r} ratio-basis value for origin index {origin_index}.",
                        exc,
                    )
                values.append(None)
        ratio_basis_values.append({"name": ratio_basis_dataset, "values": values})
    ultimate_overrides: list = []
    for origin_index in range(1, origin_count + 1):
        try:
            overridden = _result_selection_ultimate_overridden(result_selection, origin_index)
        except Exception as exc:
            if _STRICT_RESQ_EXTRACTION.get():
                _strict_failure(
                    f"Could not read Result Selection {name!r} UltimateOverridden for origin index {origin_index}.",
                    exc,
                )
            overridden = False
        if not overridden:
            ultimate_overrides.append(None)
            continue
        try:
            ultimate_overrides.append(_rs_json_number(_result_selection_ultimate(result_selection, origin_index, origin_length)))
        except Exception as exc:
            if _STRICT_RESQ_EXTRACTION.get():
                _strict_failure(
                    f"Could not read Result Selection {name!r} ultimate for origin index {origin_index}.",
                    exc,
                )
            ultimate_overrides.append(None)
    calculated_ultimate = _result_selection_calculated_ultimate(loaded_datasets, origin_count)
    selected_ultimate = _result_selection_selected_ultimate(calculated_ultimate, ultimate_overrides, origin_count)

    try:
        notes = _clean_name(result_selection.Notes)
    except Exception as exc:
        if _STRICT_RESQ_EXTRACTION.get():
            _strict_failure(f"Could not read Result Selection {name!r}.Notes.", exc)
        notes = ""
    try:
        modified = _iso_or_text(output_vector.Modified)
    except Exception as exc:
        if _STRICT_RESQ_EXTRACTION.get():
            _strict_failure(f"Could not read Result Selection {name!r} OutputVector.Modified.", exc)
        modified = utc_now_text()

    return {
        "json_format": RS_JSON_FORMAT,
        "details_tab": {
            "name": name,
            "output_type": output_type,
            "origin_length": origin_length,
            "ratio_basis_datasets": ratio_basis_datasets,
            "active_ratio_basis_dataset": ratio_basis_dataset,
            "show_ratios_as_percentages": True,
            "statistic_decimal_places": 1,
        },
        "method_tab": {
            "origin_labels": origin_labels,
            "show_weights": True,
            "loaded_datasets": loaded_datasets,
            "ratio_basis_values": ratio_basis_values,
            "calculated_ultimate": calculated_ultimate,
            "selected_ultimate": selected_ultimate,
            "ultimate_overrides": ultimate_overrides,
        },
        "_sidecar_notes": notes,
        "_sidecar_status": normalize_method_status(
            _extract_attr(
                output_vector,
                "Status",
                0,
                context=f"Result Selection {name!r} OutputVector",
            )
        ),
        "method_metadata": {
            "last_modified": modified,
        },
    }

def _result_selection_source_names(payload: dict) -> list[str]:
    method_tab = payload.get("method_tab") if isinstance(payload.get("method_tab"), dict) else {}
    loaded_datasets = method_tab.get("loaded_datasets") if isinstance(method_tab.get("loaded_datasets"), list) else []
    names: list[str] = []
    seen: set[str] = set()
    for dataset in loaded_datasets:
        name = _normalize_import_name(dataset.get("name") if isinstance(dataset, dict) else "")
        key = name.lower()
        if name and key not in seen:
            seen.add(key)
            names.append(name)
    return names

def _result_selection_precedent_names(payload: dict) -> list[str]:
    names = _result_selection_source_names(payload)
    seen = {name.lower() for name in names}
    details_tab = payload.get("details_tab") if isinstance(payload.get("details_tab"), dict) else {}
    for raw_name in details_tab.get("ratio_basis_datasets", []) if isinstance(details_tab.get("ratio_basis_datasets"), list) else []:
        name = _normalize_import_name(raw_name)
        key = name.lower()
        if name and key not in seen:
            seen.add(key)
            names.append(name)
    return names

def _result_selection_origin_labels_from_payload(payload: dict) -> list[str]:
    method_tab = payload.get("method_tab") if isinstance(payload.get("method_tab"), dict) else {}
    labels = method_tab.get("origin_labels") if isinstance(method_tab.get("origin_labels"), list) else []
    return [_normalize_import_name(label) for label in labels if _normalize_import_name(label)]

def _apply_result_selection_vector_metadata(payload: dict, result_selection_payload: dict) -> None:
    payload["notes"] = str(result_selection_payload.pop("_sidecar_notes", "") or "")
    payload["status"] = normalize_method_status(
        result_selection_payload.pop("_sidecar_status", payload.get("status"))
    )
    payload["precedents"] = _result_selection_precedent_names(result_selection_payload)
    origin_labels = _result_selection_origin_labels_from_payload(result_selection_payload)
    if origin_labels:
        payload["origin_labels"] = origin_labels
        payload["origin_count"] = len(origin_labels)

def _bf_origin_count(method, output_vector) -> int:
    origin_count = _extract_int_attr(method, "OriginCount", 0, context="method")
    if origin_count <= 0:
        try:
            origin_count = int(_call_member(method, "OriginCount"))
        except Exception as exc:
            if _STRICT_RESQ_EXTRACTION.get():
                _strict_failure("Could not read method OriginCount.", exc)
            origin_count = 0
    if origin_count <= 0:
        origin_count = _vector_origin_count(output_vector)
    return max(0, origin_count)


def _bf_origin_label(method, origin_index: int) -> str:
    errors: list[Exception] = []
    for name in ("OriginLabel", "OriginLabels"):
        try:
            return _normalize_import_name(_try_call_member(method, name, [((), {"OriginIndex": origin_index}), ((origin_index,), {})]))
        except Exception as exc:
            errors.append(exc)
            continue
    if _STRICT_RESQ_EXTRACTION.get():
        _strict_failure(
            f"Could not read method origin label for origin index {origin_index}.",
            errors[0] if errors else None,
        )
    return ""


def _bf_origin_labels(method, output_vector, fallback_count: int = 0) -> list[str]:
    origin_count = _bf_origin_count(method, output_vector)
    if origin_count <= 0:
        origin_count = max(0, int(fallback_count or 0))
    labels: list[str] = []
    for i in range(1, origin_count + 1):
        labels.append(_bf_origin_label(method, i) or _vector_origin_label(output_vector, i))
    return labels


def _resq_percentage_developed(
    method,
    origin_labels: list[str],
    *,
    context: str,
    fallback: list,
) -> list:
    """Copy the percentage developed ResQ shows for each origin of a method.

    ResQ exposes ``PercentageDevelopedValues(i)`` for origin rows ``1..N``.
    Under its DFM development-factor settings that figure is the pattern of
    the DFM's cumulative factors at each origin's own age -- the same thing
    ArcRho reads from the DFM behind the precedent -- so the import copies it
    as-is rather than deriving a ratio: it is defined for an origin whose
    latest is zero and never drifts against a Latest the DFM was not built on.
    ``fallback`` is what a method that will not expose the values lands
    instead; strict extraction refuses such a method.
    """

    percentages: list = []
    for origin_index in range(1, len(origin_labels) + 1):
        try:
            value = _try_call_member(
                method,
                "PercentageDevelopedValues",
                [((origin_index,), {}), ((), {"OriginIndex": origin_index})],
            )
        except Exception as exc:
            if _STRICT_RESQ_EXTRACTION.get():
                _strict_failure(
                    f"Could not read {context} PercentageDevelopedValues for origin index "
                    f"{origin_index}.",
                    exc,
                )
            return list(fallback)
        percentages.append(float(value) if _rs_json_number(value) is not None else None)
    return percentages


def _latest_ultimate_ratio(latest_snapshot: dict, dfm_snapshot: dict) -> list:
    """Latest / imported ultimate, the fallback for a method without ResQ percentages."""

    latest = latest_snapshot.get("values") or []
    ultimates = dfm_snapshot.get("values") or []
    percentages: list = []
    for index in range(len(ultimates)):
        latest_value = latest[index] if index < len(latest) else None
        ultimate = ultimates[index]
        if latest_value is None or ultimate in (None, 0):
            percentages.append(None)
            continue
        percentages.append(float(latest_value) / float(ultimate))
    return percentages


def _bf_source_snapshot(
    source,
    origin_labels: list[str],
    *,
    latest: bool,
    context: str = "BF",
    role: str = "source",
) -> dict:
    """Extract the exact source vector a method consumes, without filesystem I/O.

    ``role`` is the ResQ method-dialog field the precedent came from (Latest,
    Exposure, Perc Developed, Prior), so a blank selection is reported by the
    field the operator has to fix.
    """

    name = _normalize_import_name(
        _extract_attr(source, "Name", "", context=f"{context} {role} precedent")
    )
    if not name:
        raise ValueError(
            f"The {role} input does not name a ResQ dataset; its selection is blank or broken."
        )
    values: list = []
    successful_reads = 0
    errors: list[Exception] = []
    for origin_index in range(1, len(origin_labels) + 1):
        try:
            if not latest:
                value = _vector_value(source, origin_index)
                successful_reads += 1
                values.append(value)
                continue
            development_count = _triangle_development_count(source, origin_index)
            if development_count is None or development_count <= 0:
                value = _vector_value(source, origin_index)
                successful_reads += 1
                values.append(value)
                continue
            value = None
            row_read = False
            for development_index in range(development_count, 0, -1):
                try:
                    candidate = _triangle_value(source, origin_index, development_index)
                    row_read = True
                except Exception as exc:
                    if _STRICT_RESQ_EXTRACTION.get():
                        _strict_failure(
                            f"Could not read {context} source {name!r} value at cell "
                            f"({origin_index}, {development_index}).",
                            exc,
                        )
                    errors.append(exc)
                    continue
                if _rs_json_number(candidate) is not None:
                    value = candidate
                    break
            if row_read:
                successful_reads += 1
            values.append(value)
        except Exception as exc:
            if _STRICT_RESQ_EXTRACTION.get():
                _strict_failure(
                    f"Could not read {context} source {name!r} for origin index {origin_index}.",
                    exc,
                )
            errors.append(exc)
            values.append(None)
    if origin_labels and successful_reads <= 0:
        detail = f": {errors[0]}" if errors else ""
        raise ValueError(f"Failed to read {context} source {name!r}{detail}")
    return {
        "name": name,
        "origin_labels": list(origin_labels),
        "values": values,
    }


@_strict_extractor
def export_bornhuetter_ferguson(method, *, strict: bool = False) -> dict:
    """Extract a complete, self-contained canonical BF v3 payload from ResQ."""

    del strict
    output_vector = _extract_attr(method, "OutputVector", None, context="Bornhuetter Ferguson method")
    name = _normalize_import_name(
        _extract_attr(output_vector, "Name", "", context="Bornhuetter Ferguson OutputVector")
    ) or _normalize_import_name(
        _extract_attr(method, "Name", "", context="Bornhuetter Ferguson method")
    )
    dataset_type_obj = _extract_attr(
        output_vector,
        "DatasetType",
        None,
        context=f"Bornhuetter Ferguson {name!r} OutputVector",
    )
    output_type = _normalize_import_name(
        _extract_attr(
            dataset_type_obj,
            "Name",
            "",
            context=f"Bornhuetter Ferguson {name!r} output DatasetType",
        )
    ) or name
    category_obj = _extract_attr(
        dataset_type_obj,
        "Category",
        None,
        context=f"Bornhuetter Ferguson {name!r} output DatasetType",
    )
    dataset_category = _normalize_import_name(
        _extract_attr(
            category_obj,
            "Name",
            "",
            context=f"Bornhuetter Ferguson {name!r} output Category",
        )
    )
    output_period_length = _extract_int_attr(
        output_vector,
        "PeriodLength",
        12,
        context=f"Bornhuetter Ferguson {name!r} OutputVector",
    )
    origin_length = _extract_int_attr(
        method,
        "OriginLength",
        output_period_length,
        context=f"Bornhuetter Ferguson {name!r}",
    )
    origin_labels = _bf_origin_labels(method, output_vector)
    if not origin_labels or any(not label for label in origin_labels):
        raise ValueError(f"Bornhuetter Ferguson method {name!r} does not expose complete origin labels.")
    latest_source = _extract_attr(method, "Latest", None, context=f"Bornhuetter Ferguson {name!r}")
    dfm_source = _extract_attr(
        method,
        "PercentageDeveloped",
        None,
        context=f"Bornhuetter Ferguson {name!r}",
    )
    prior_source = _extract_attr(method, "Prior", None, context=f"Bornhuetter Ferguson {name!r}")
    latest_snapshot = _bf_source_snapshot(latest_source, origin_labels, latest=True, role="Latest")
    dfm_snapshot = _bf_source_snapshot(dfm_source, origin_labels, latest=False, role="Perc Developed")
    dfm_snapshot["percentage_developed"] = _resq_percentage_developed(
        method,
        origin_labels,
        context=f"Bornhuetter Ferguson {name!r}",
        fallback=_latest_ultimate_ratio(latest_snapshot, dfm_snapshot),
    )
    prior_snapshot = _bf_source_snapshot(prior_source, origin_labels, latest=False, role="Prior")
    try:
        notes = _clean_name(method.Notes)
    except Exception as exc:
        if _STRICT_RESQ_EXTRACTION.get():
            _strict_failure(f"Could not read Bornhuetter Ferguson {name!r}.Notes.", exc)
        notes = ""
    try:
        modified = _iso_or_text(output_vector.Modified)
    except Exception as exc:
        if _STRICT_RESQ_EXTRACTION.get():
            _strict_failure(
                f"Could not read Bornhuetter Ferguson {name!r} OutputVector.Modified.",
                exc,
            )
        modified = utc_now_text()

    owned = {
        "json_format": BF_JSON_FORMAT,
        "details_tab": {
            "name": name,
            "method_type": BF_METHOD_TYPE,
            "output_type": output_type,
            "dataset_category": dataset_category,
            "origin_length": origin_length,
            "statistic_decimal_places": 1,
        },
        "method_tab": {
            "latest_dataset": latest_snapshot["name"],
            "dfm_dataset": dfm_snapshot["name"],
            "show_weights": True,
            "show_effective_weights": False,
            "prior_datasets": [
                {
                    "name": prior_snapshot["name"],
                    "values": [],
                    "weights": [1.0 for _ in origin_labels],
                }
            ],
            "origin_labels": origin_labels,
            "latest_values": [],
            "dfm_ultimate_values": [],
            "percentage_developed": [],
            "selected_prior_values": [],
            "new_ultimate": [],
        },
        "_sidecar_notes": notes,
        "method_metadata": {
            "method_type": BF_METHOD_TYPE,
            "source_kind": BF_SOURCE_KIND,
            "last_modified": modified,
            "data_refreshed": modified,
        },
    }
    payload = recalculate_bornhuetter_ferguson_method(
        owned,
        source_snapshots={
            "latest": latest_snapshot,
            "dfm": dfm_snapshot,
            "priors": [prior_snapshot],
        },
        timestamp=modified,
    )
    payload["_sidecar_notes"] = notes
    payload["_sidecar_status"] = normalize_method_status(
        _extract_attr(
            output_vector,
            "Status",
            0,
            context=f"Bornhuetter Ferguson {name!r} OutputVector",
        )
    )
    return payload


def _apply_bornhuetter_ferguson_vector_metadata(payload: dict, bf_payload: dict) -> None:
    payload["notes"] = str(bf_payload.pop("_sidecar_notes", "") or "")
    payload["status"] = normalize_method_status(
        bf_payload.pop("_sidecar_status", payload.get("status"))
    )
    payload["source_kind"] = BF_SOURCE_KIND
    payload["method_type"] = BF_METHOD_TYPE
    payload["method_type_code"] = METHOD_TYPE_BF_CODE
    payload["precedents"] = bornhuetter_ferguson_precedent_names(bf_payload)
    details_tab = bf_payload.get("details_tab") if isinstance(bf_payload.get("details_tab"), dict) else {}
    metadata = bf_payload.get("method_metadata") if isinstance(bf_payload.get("method_metadata"), dict) else {}
    payload["method_name"] = _normalize_import_name(details_tab.get("name"))
    payload["publication_revision"] = _clean_name(metadata.get("publication_revision"))
    method_tab = bf_payload.get("method_tab") if isinstance(bf_payload.get("method_tab"), dict) else {}
    origin_labels = method_tab.get("origin_labels") if isinstance(method_tab.get("origin_labels"), list) else []
    if origin_labels:
        payload["origin_labels"] = [_normalize_import_name(label) for label in origin_labels]
        payload["origin_count"] = len(origin_labels)
        payload["values"] = [[value] for value in method_tab.get("new_ultimate", [])]
    payload["origin_length"] = int(
        details_tab.get("origin_length") or payload.get("origin_length") or 12
    )
    payload["period_length"] = payload["origin_length"]


def write_bornhuetter_ferguson_export(payload: dict, rc_path: str, rc_dir: Path) -> Path:
    del rc_path
    details_tab = payload.get("details_tab") if isinstance(payload.get("details_tab"), dict) else {}
    name = _normalize_import_name(details_tab.get("name")) or BF_METHOD_TYPE
    file_name = f"BF@{_encode_name_part(name)}.json"
    out_path = rc_dir / METHOD_DATA_DIR / file_name
    method_payload = dict(payload)
    method_payload.pop("_sidecar_notes", None)
    method_payload.pop("_sidecar_status", None)
    _write_json(out_path, method_payload)
    return out_path


def _find_bornhuetter_ferguson_for_vector(reserving_class, vector_name: str):
    try:
        collection = _call_member(reserving_class, "BFMethods")
    except Exception:
        return None
    direct_candidates = []
    try:
        item = collection.Item(vector_name)
        if item is not None:
            direct_candidates.append(item)
    except Exception:
        pass
    return _find_unique_method_by_output(
        collection,
        direct_candidates,
        vector_name,
        "OutputVector",
        "Bornhuetter Ferguson",
    )


def _cc_indexed_value(method, member_name: str, origin_index: int):
    return _try_call_member(
        method,
        member_name,
        [((origin_index,), {}), ((), {"OriginIndex": origin_index})],
    )


def _cc_code_label(value: object, labels: tuple[str, ...], field_name: str) -> str:
    try:
        code = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Invalid ResQ {field_name} value: {value!r}.") from exc
    if not 0 <= code < len(labels):
        raise ValueError(f"Unsupported ResQ {field_name} code: {code}.")
    return labels[code]


# ResQ PercentageDevelopedType codes, indexed: 0 pdInternal (Latest/Ultimates),
# 1 pdPattern, 2 pdCumDevFactors, 3 pdCumDevFactorsAdjusted. For both DFM-factor
# codes the referenced vector stores per-origin DFM ultimates whose
# latest/ultimate ratio equals ResQ's PercentageDevelopedValues exactly
# (verified against the live book, max deviation 3e-16), so they import as
# latest_ultimates and stay correct when the input datasets are refreshed.
CC_PERCENTAGE_DEVELOPED_TYPE_MODES = (
    CC_PRIOR_ULTIMATE_MODES[0],
    CC_PRIOR_ULTIMATE_MODES[1],
    CC_PRIOR_ULTIMATE_MODES[0],
    CC_PRIOR_ULTIMATE_MODES[0],
)
CC_DFM_FACTOR_TYPE_CODES = (2, 3)


def _apply_dfm_factor_prior_ultimates(method, latest_snapshot: dict, prior_snapshot: dict) -> None:
    """Rewrite the prior-ultimate snapshot as latest / ResQ percentage developed.

    For the DFM-factor codes ResQ derives percentage developed from the DFM's
    cumulative development factors, not from the referenced vector's stored
    values. latest / PercentageDevelopedValues is therefore the effective prior
    ultimate: it equals the stored vector whenever the DFM output is populated,
    and when that output is empty (ResQ then reports a flat 1.0 pattern) the
    stored zeros would otherwise blank every origin of the imported method.
    """
    values = prior_snapshot["values"]
    latest_values = latest_snapshot["values"]
    for index in range(min(len(values), len(latest_values))):
        latest_value = latest_values[index]
        if not latest_value:
            continue
        try:
            percentage = float(_cc_indexed_value(method, "PercentageDevelopedValues", index + 1))
        except Exception as exc:
            if _STRICT_RESQ_EXTRACTION.get():
                _strict_failure(
                    "Could not read Cape Cod PercentageDevelopedValues for origin index "
                    f"{index + 1}.",
                    exc,
                )
            continue
        if percentage:
            values[index] = float(latest_value) / percentage


@_strict_extractor
def export_cape_cod(method, *, strict: bool = False) -> dict:
    """Extract a complete, self-contained canonical Cape Cod v1 payload from ResQ."""

    del strict
    output_vector = _extract_attr(method, "OutputVector", None, context="Cape Cod method")
    name = _normalize_import_name(
        _extract_attr(output_vector, "Name", "", context="Cape Cod OutputVector")
    ) or _normalize_import_name(_extract_attr(method, "Name", "", context="Cape Cod method"))
    dataset_type_obj = _extract_attr(
        output_vector,
        "DatasetType",
        None,
        context=f"Cape Cod {name!r} OutputVector",
    )
    output_type = _normalize_import_name(
        _extract_attr(
            dataset_type_obj,
            "Name",
            "",
            context=f"Cape Cod {name!r} output DatasetType",
        )
    ) or name
    category_obj = _extract_attr(
        dataset_type_obj,
        "Category",
        None,
        context=f"Cape Cod {name!r} output DatasetType",
    )
    dataset_category = _normalize_import_name(
        _extract_attr(
            category_obj,
            "Name",
            "",
            context=f"Cape Cod {name!r} output Category",
        )
    )
    output_period_length = _extract_int_attr(
        output_vector,
        "PeriodLength",
        12,
        context=f"Cape Cod {name!r} OutputVector",
    )
    origin_length = _extract_int_attr(
        method,
        "OriginLength",
        output_period_length,
        context=f"Cape Cod {name!r}",
    )
    origin_labels = _bf_origin_labels(method, output_vector)
    if not origin_labels or any(not label for label in origin_labels):
        raise ValueError(f"Cape Cod method {name!r} does not expose complete origin labels.")
    latest_source = _extract_attr(method, "Latest", None, context=f"Cape Cod {name!r}")
    exposure_source = _extract_attr(method, "Exposure", None, context=f"Cape Cod {name!r}")
    prior_source = _extract_attr(
        method,
        "PercentageDeveloped",
        None,
        context=f"Cape Cod {name!r}",
    )
    latest_snapshot = _bf_source_snapshot(
        latest_source, origin_labels, latest=True, context="Cape Cod", role="Latest"
    )
    exposure_snapshot = _bf_source_snapshot(
        exposure_source, origin_labels, latest=False, context="Cape Cod", role="Exposure"
    )
    prior_snapshot = _bf_source_snapshot(
        prior_source, origin_labels, latest=False, context="Cape Cod", role="Perc Developed"
    )
    pd_type_code = _extract_attr(
        method,
        "PercentageDevelopedType",
        0,
        context=f"Cape Cod {name!r}",
    )
    prior_ultimate_mode = _cc_code_label(
        pd_type_code,
        CC_PERCENTAGE_DEVELOPED_TYPE_MODES,
        "PercentageDevelopedType",
    )
    if int(pd_type_code) in CC_DFM_FACTOR_TYPE_CODES:
        _apply_dfm_factor_prior_ultimates(method, latest_snapshot, prior_snapshot)
    # The pattern behind the prior ultimate, copied from ResQ; a method that
    # exposes none leaves the snapshot without one and Cape Cod falls back to
    # Latest / Prior Ultimate, the way it treats a prior with no DFM behind it.
    prior_snapshot["percentage_developed"] = _resq_percentage_developed(
        method,
        origin_labels,
        context=f"Cape Cod {name!r}",
        fallback=[],
    )
    scaling_type = _cc_code_label(
        _extract_attr(method, "ScalingType", 0, context=f"Cape Cod {name!r}"),
        CC_SCALING_TYPES,
        "ScalingType",
    )
    trend_factor_overrides: list = []
    for origin_index in range(1, len(origin_labels) + 1):
        if _bool_value(_cc_indexed_value(method, "ManualTrendFactor", origin_index)):
            trend_factor_overrides.append(float(_cc_indexed_value(method, "TrendFactorValues", origin_index)))
        else:
            trend_factor_overrides.append(None)
    try:
        notes = _clean_name(method.Notes)
    except Exception as exc:
        if _STRICT_RESQ_EXTRACTION.get():
            _strict_failure(f"Could not read Cape Cod {name!r}.Notes.", exc)
        notes = ""
    try:
        modified = _iso_or_text(output_vector.Modified)
    except Exception as exc:
        if _STRICT_RESQ_EXTRACTION.get():
            _strict_failure(f"Could not read Cape Cod {name!r} OutputVector.Modified.", exc)
        modified = utc_now_text()

    owned = {
        "json_format": CC_JSON_FORMAT,
        "details_tab": {
            "name": name,
            "method_type": CC_METHOD_TYPE,
            "output_type": output_type,
            "dataset_category": dataset_category,
            "origin_length": origin_length,
            "statistic_decimal_places": _extract_int_attr(
                method,
                "DecimalPlaces",
                2,
                context=f"Cape Cod {name!r}",
            ),
        },
        "method_tab": {
            "latest_dataset": latest_snapshot["name"],
            "exposure_dataset": exposure_snapshot["name"],
            "prior_ultimate_dataset": prior_snapshot["name"],
            "prior_ultimate_mode": prior_ultimate_mode,
            "trend_rate": _extract_attr(method, "TrendRate", 0, context=f"Cape Cod {name!r}"),
            "auto_trend_fit": _bool_value(
                _extract_attr(method, "AutoTrendFit", False, context=f"Cape Cod {name!r}")
            ),
            "decay_factor": _extract_attr(
                method,
                "DecayFactor",
                0,
                context=f"Cape Cod {name!r}",
            ),
            "scaling_type": scaling_type,
            "alternative_ultimate_calculation": _bool_value(
                _extract_attr(
                    method,
                    "AltUltimateCalc",
                    False,
                    context=f"Cape Cod {name!r}",
                )
            ),
            "trend_factor_overrides": trend_factor_overrides,
            "origin_labels": origin_labels,
        },
        "method_metadata": {
            "method_type": CC_METHOD_TYPE,
            "source_kind": CC_SOURCE_KIND,
            "last_modified": modified,
            "data_refreshed": modified,
        },
    }
    payload = recalculate_cape_cod_method(
        owned,
        source_snapshots={
            "latest": latest_snapshot,
            "exposure": exposure_snapshot,
            "prior_ultimate": prior_snapshot,
        },
        timestamp=modified,
    )
    payload["_sidecar_notes"] = notes
    payload["_sidecar_status"] = normalize_method_status(
        _extract_attr(
            output_vector,
            "Status",
            0,
            context=f"Cape Cod {name!r} OutputVector",
        )
    )
    return payload


def _apply_cape_cod_vector_metadata(payload: dict, cc_payload: dict) -> None:
    payload["notes"] = str(cc_payload.pop("_sidecar_notes", "") or "")
    payload["status"] = normalize_method_status(
        cc_payload.pop("_sidecar_status", payload.get("status"))
    )
    payload["source_kind"] = CC_SOURCE_KIND
    payload["method_type"] = CC_METHOD_TYPE
    payload["method_type_code"] = METHOD_TYPE_CAPE_COD_CODE
    payload["precedents"] = cape_cod_precedent_names(cc_payload)
    details_tab = cc_payload.get("details_tab") if isinstance(cc_payload.get("details_tab"), dict) else {}
    metadata = cc_payload.get("method_metadata") if isinstance(cc_payload.get("method_metadata"), dict) else {}
    payload["method_name"] = _normalize_import_name(details_tab.get("name"))
    payload["publication_revision"] = _clean_name(metadata.get("publication_revision"))
    method_tab = cc_payload.get("method_tab") if isinstance(cc_payload.get("method_tab"), dict) else {}
    origin_labels = method_tab.get("origin_labels") if isinstance(method_tab.get("origin_labels"), list) else []
    if origin_labels:
        payload["origin_labels"] = [_normalize_import_name(label) for label in origin_labels]
        payload["origin_count"] = len(origin_labels)
        payload["values"] = [[value] for value in method_tab.get("cape_cod_ultimate", [])]
    payload["origin_length"] = int(
        details_tab.get("origin_length") or payload.get("origin_length") or 12
    )
    payload["period_length"] = payload["origin_length"]


def write_cape_cod_export(payload: dict, rc_path: str, rc_dir: Path) -> Path:
    del rc_path
    details_tab = payload.get("details_tab") if isinstance(payload.get("details_tab"), dict) else {}
    name = _normalize_import_name(details_tab.get("name")) or CC_METHOD_TYPE
    file_name = f"CC@{_encode_name_part(name)}.json"
    out_path = rc_dir / METHOD_DATA_DIR / file_name
    method_payload = dict(payload)
    method_payload.pop("_sidecar_notes", None)
    method_payload.pop("_sidecar_status", None)
    _write_json(out_path, method_payload)
    return out_path


def _find_cape_cod_for_vector(reserving_class, vector_name: str):
    try:
        collection = _call_member(reserving_class, "CapeCodMethods")
    except Exception:
        return None
    direct_candidates = []
    try:
        item = collection.Item(vector_name)
        if item is not None:
            direct_candidates.append(item)
    except Exception:
        pass
    return _find_unique_method_by_output(
        collection,
        direct_candidates,
        vector_name,
        "OutputVector",
        "Cape Cod",
    )

def _parse_origin_start_month(label: object, base_len: int) -> tuple[int, int] | None:
    text = _clean_name(label)
    if not text:
        return None

    if base_len == 1:
        match = re.match(r"^(\d{4})(\d{2})$", text)
        if match:
            month = int(match.group(2))
            if 1 <= month <= 12:
                return int(match.group(1)), month
        return None

    if base_len == 3:
        for pattern in (r"^(\d{4})\s*Q([1-4])$", r"^Q([1-4])\s*(\d{4})$"):
            match = re.match(pattern, text, re.I)
            if not match:
                continue
            if pattern.startswith("^(\\d"):
                year, quarter = int(match.group(1)), int(match.group(2))
            else:
                quarter, year = int(match.group(1)), int(match.group(2))
            return year, (quarter - 1) * 3 + 1
        return None

    if base_len == 6:
        for pattern in (r"^(\d{4})\s*H([1-2])$", r"^H([1-2])\s*(\d{4})$"):
            match = re.match(pattern, text, re.I)
            if not match:
                continue
            if pattern.startswith("^(\\d"):
                year, half = int(match.group(1)), int(match.group(2))
            else:
                half, year = int(match.group(1)), int(match.group(2))
            return year, (half - 1) * 6 + 1
        return None

    if base_len == 12 and re.match(r"^\d{4}$", text):
        return int(text), 1
    return None

def _numeric_or_none(value: object) -> float | None:
    if value is None:
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if number == number else None

def _vector_row_value(row: object) -> object:
    if isinstance(row, list):
        return row[0] if row else None
    return row

def _aggregate_vector_values_by_length(values: list, origin_labels: list, base_len: int, target_len: int) -> list[list]:
    if not values:
        return []
    if base_len <= 0 or target_len <= base_len or target_len % base_len != 0:
        return []
    factor = target_len // base_len
    vector = [_vector_row_value(row) for row in values]
    labels = [str(label) for label in origin_labels] if isinstance(origin_labels, list) else []

    if len(labels) == len(vector) and base_len in {1, 3, 6, 12}:
        ordered_keys: list[tuple[int, int]] = []
        buckets: dict[tuple[int, int], dict[str, object]] = {}
        parse_failed = False
        for label, raw in zip(labels, vector):
            parsed = _parse_origin_start_month(label, base_len)
            if parsed is None:
                parse_failed = True
                break
            year, month = parsed
            bucket_month = ((month - 1) // target_len) * target_len + 1
            key = (year, bucket_month)
            if key not in buckets:
                buckets[key] = {"sum": 0.0, "has_value": False}
                ordered_keys.append(key)
            number = _numeric_or_none(raw)
            if number is not None:
                buckets[key]["sum"] = float(buckets[key]["sum"]) + number
                buckets[key]["has_value"] = True
        if not parse_failed:
            return [[buckets[key]["sum"] if buckets[key]["has_value"] else None] for key in ordered_keys]

    out: list[list] = []
    for start in range(0, len(vector), factor):
        total = 0.0
        has_value = False
        for raw in vector[start:start + factor]:
            number = _numeric_or_none(raw)
            if number is None:
                continue
            total += number
            has_value = True
        out.append([total if has_value else None])
    return out

def _write_aggregated_vector_cache_exports(payload: dict, rc_dir: Path) -> list[Path]:
    # The values are at the payload's stored period, so that is the base every
    # coarser copy is summed from.
    try:
        base_len = _vector_payload_stored_period_length(payload)
    except (TypeError, ValueError):
        return []
    if base_len <= 0:
        return []
    paths: list[Path] = []
    for target_len in (3, 6, 12):
        if target_len <= base_len or target_len % base_len != 0:
            continue
        rows = _aggregate_vector_values_by_length(
            payload.get("values") if isinstance(payload.get("values"), list) else [],
            payload.get("origin_labels") if isinstance(payload.get("origin_labels"), list) else [],
            base_len,
            target_len,
        )
        if not rows:
            continue
        csv_name = _vector_cache_csv_file_name(_normalize_import_name(payload["name"]), target_len)
        csv_path = rc_dir / DATASET_CACHE_DIR / csv_name
        _write_csv_matrix(csv_path, rows)
        paths.append(csv_path)
    return paths

def write_result_selection_export(payload: dict, rc_path: str, rc_dir: Path) -> Path:
    details_tab = payload.get("details_tab") if isinstance(payload.get("details_tab"), dict) else {}
    name = _normalize_import_name(details_tab.get("name")) or "Result Selection"
    file_name = f"RS@{_encode_name_part(name)}.json"
    out_path = rc_dir / METHOD_DATA_DIR / file_name
    method_payload = dict(payload)
    method_payload.pop("_sidecar_notes", None)
    method_payload.pop("_sidecar_status", None)
    _write_json(out_path, method_payload)
    return out_path

def _find_result_selection_for_vector(reserving_class, vector_name: str):
    direct_candidates = []
    try:
        method = _call_member(reserving_class, "GetResultSelection", vector_name)
        if method is not None:
            direct_candidates.append(method)
    except Exception:
        pass
    try:
        collection = _call_member(reserving_class, "ResultSelections")
    except Exception:
        collection = None
    if collection is not None:
        try:
            method = collection.Item(vector_name)
            if method is not None:
                direct_candidates.append(method)
        except Exception:
            pass
    return _find_unique_method_by_output(
        collection,
        direct_candidates,
        vector_name,
        "OutputVector",
        "Result Selection",
    )

def _dfm_ultimate_value(dfm, origin_index: int):
    call_shapes = [
        ((origin_index,), {}),
        ((), {"OriginIndex": origin_index}),
        ((), {"arg0": origin_index}),
    ]
    try:
        return _try_call_member(dfm, "Ultimates", call_shapes)
    except Exception as exc:
        raise AttributeError(f"Could not read DFM ultimate value for origin index {origin_index}.") from exc

def export_dfm_ultimate_vector(
    dfm,
    origin_labels: list[str],
    origin_length: int,
    dev_length: int,
) -> dict:
    """Extract the DFM output vector from ResQ DFM.Ultimates into ArcRho CSV payload shape."""
    output_vector = dfm.OutputVector
    name = _normalize_import_name(output_vector.Name)
    dataset_type_obj = _safe_attr(output_vector, "DatasetType", None)
    dataset_type = _normalize_import_name(_safe_attr(dataset_type_obj, "Name", "")) or name
    category = _normalize_import_name(_safe_attr(_safe_attr(dataset_type_obj, "Category", None), "Name", ""))
    data_format = _safe_int_attr(dataset_type_obj, "DataFormat", 1)
    method_type_code = _safe_int_attr(output_vector, "MethodType", -1)
    method_type = _method_type_name(method_type_code)
    origin_count = len(origin_labels)
    if origin_count <= 0:
        raise ValueError(f"DFM output vector {name!r} does not have origin labels.")

    values: list[list] = []
    attempted_cells = 0
    value_errors: list[Exception] = []
    for i in range(1, origin_count + 1):
        attempted_cells += 1
        try:
            values.append([_dfm_ultimate_value(dfm, i)])
        except Exception as exc:
            value_errors.append(exc)
            values.append([None])
    if attempted_cells > 0 and len(value_errors) == attempted_cells:
        raise ValueError(f"Failed to read any DFM ultimate values for {name!r}: {value_errors[0]}")

    user = _normalize_import_name(_safe_attr(output_vector, "User", ""))
    created = _iso_or_text(_safe_attr(output_vector, "Created", ""))
    modified = _iso_or_text(_safe_attr(output_vector, "Modified", ""))

    return {
        "name": name,
        "dataset_type": dataset_type,
        "category": category,
        "data_format": data_format,
        "method_type": method_type,
        "method_type_code": method_type_code,
        "origin_length": origin_length,
        "development_length": dev_length,
        "origin_count": origin_count,
        "development_count": 1,
        "origin_labels": origin_labels,
        "development_labels": ["Ultimate"],
        "values": values,
        "method_name": _normalize_import_name(dfm.Name),
        "notes": str(_safe_attr(dfm, "Notes", "") or ""),
        "user": user,
        "created": created,
        "modified": modified,
        "status": normalize_method_status(_safe_attr(output_vector, "Status", 0)),
    }


def _csv_matrix_bytes(rows: list[list]) -> bytes:
    stream = io.StringIO(newline="")
    writer = csv.writer(stream, lineterminator="\n")
    writer.writerows([
        ["" if cell is None or str(cell).strip().lower() in {"none", "nan"} else str(cell).strip() for cell in row]
        for row in rows
    ])
    return stream.getvalue().encode("utf-8")


def build_dfm_ultimate_publication(
    payload: dict,
    method_payload: dict,
    rc_path: str,
    rc_dir: Path,
) -> tuple[Path, dict[Path, bytes], Path]:
    """Build every DFM output artifact without mutating disk."""

    name = _normalize_import_name(payload["name"])
    period_length = _vector_payload_period_length(payload)
    files: dict[Path, bytes] = {}
    primary_path = rc_dir / DATASET_CACHE_DIR / _vector_cache_csv_file_name(name, period_length)
    for target_length, values in dfm_output_variants(method_payload).items():
        path = rc_dir / DATASET_CACHE_DIR / _vector_cache_csv_file_name(name, target_length)
        files[path] = _csv_matrix_bytes([[value] for value in values])

    meta_path = rc_dir / DATASET_SIDECAR_DIR / _json_sidecar_name(name)
    existing = _safe_read_json(meta_path)
    publication_revision = _clean_name(
        method_payload.get("method_metadata", {}).get("publication_revision")
        if isinstance(method_payload.get("method_metadata"), dict)
        else ""
    )
    output_changed = _clean_name(existing.get("publication_revision")) != publication_revision
    updated_at = payload.get("modified") or utc_now_text()
    sidecar = build_dfm_output_sidecar(
        method_payload,
        project_name=PROJECT_NAME,
        reserving_class=rc_path,
        csv_file=primary_path.name,
        existing=existing,
        notes=None if existing else str(payload.get("notes") or ""),
        timestamp=updated_at,
        user=payload.get("user", ""),
        output_changed=output_changed,
        append_audit=not existing or output_changed,
        status=normalize_method_status(payload.get("status")),
    )
    files[meta_path] = persisted_json_text(sidecar).encode("utf-8")
    return primary_path, files, meta_path

def publish_dfm_artifacts(files: dict[Path, bytes], *, sidecar_path: Path) -> list[Path]:
    """Publish staged DFM artifacts with rollback and sidecar-last replacement."""

    ordered = sorted(files, key=lambda path: (path == sidecar_path, str(path).casefold()))
    staged: dict[Path, Path] = {}
    backups: dict[Path, bytes | None] = {}
    replaced: list[Path] = []
    try:
        for path in ordered:
            path.parent.mkdir(parents=True, exist_ok=True)
            current = path.read_bytes() if path.is_file() else None
            if current == files[path]:
                continue
            backups[path] = current
            temporary = path.with_name(f"{path.name}.{uuid.uuid4().hex}.tmp")
            temporary.write_bytes(files[path])
            staged[path] = temporary
        for path in ordered:
            temporary = staged.pop(path, None)
            if temporary is None:
                continue
            os.replace(temporary, path)
            replaced.append(path)
    except OSError as exc:
        rollback_errors: list[str] = []
        for path in reversed(replaced):
            try:
                original = backups[path]
                if original is None:
                    path.unlink(missing_ok=True)
                else:
                    temporary = path.with_name(f"{path.name}.{uuid.uuid4().hex}.rollback")
                    temporary.write_bytes(original)
                    os.replace(temporary, path)
            except OSError as rollback_exc:
                rollback_errors.append(f"{path.name}: {rollback_exc}")
        detail = f"; rollback failed: {'; '.join(rollback_errors)}" if rollback_errors else ""
        raise RuntimeError(f"Failed to publish DFM {sidecar_path.stem}: {exc}{detail}") from exc
    finally:
        for temporary in staged.values():
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass
    return replaced


def write_dfm_ultimate_vector_export(
    payload: dict,
    rc_path: str,
    rc_dir: Path,
    *,
    method_payload: dict | None = None,
) -> Path:
    """Compatibility publisher; a canonical v2 method payload is mandatory."""

    if not isinstance(method_payload, dict):
        raise ValueError("A canonical DFM v2 method payload is required to publish its output.")
    csv_path, files, sidecar_path = build_dfm_ultimate_publication(
        payload,
        method_payload,
        rc_path,
        rc_dir,
    )
    publish_dfm_artifacts(files, sidecar_path=sidecar_path)
    return csv_path
