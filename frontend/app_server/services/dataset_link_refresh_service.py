"""Re-evaluate a dataset's ArcRho cell links where the workspace is local disk.

A manual-input dataset whose cells are driven by ArcRho links — standalone
dataset references (``internal_links``) and formulas (``formula_links``) — is
an instance-level node of the dependency graph: its sidecar names the datasets
its links read as precedents, and the dependent-propagation walk calls
:func:`refresh_dataset_links` when one of them was refreshed, so the linked
values follow their sources instead of staying the snapshot of the last manual
refresh. The browser's Links-tab refresh stays the interactive twin; both
resolve references through the same app-server services and evaluate through
``arcrho_api.dataset_link_contract``, so a walk and a hand refresh can never
disagree about a value.

Failure semantics are asymmetric on purpose:

- An ArcRho-side problem — a referenced dataset missing, a non-numeric cell, a
  result that no longer covers the linked cells — fails the refresh. The
  dataset keeps its last values, the walk records the error, and downstream
  methods are flagged for review, exactly as a calculated dataset's
  dependency failure reports.
- An Excel-side problem fails only the links that read that workbook: their
  cells keep their last values and the refresh reports a warning naming the
  reference, because a workbook on another machine's local disk is routinely
  unreachable from the server host and must never block the chain.

Cells no link owns are never touched: the refresh rewrites the linked cells
inside the current CSV grid rather than regenerating the grid.
"""

from __future__ import annotations

import os
import re
from typing import Any, Dict, List, Mapping, Tuple

from fastapi import HTTPException

from arcrho_api.dataset_link_contract import (
    DatasetLinkError,
    evaluate_dataset_formula,
    parse_dataset_formula_tree,
    parse_internal_reference,
    tokenize_dataset_formula,
)
from arcrho_api.sidecar_audit_contract import AUDIT_ACTION_AUTO_REFRESH, append_audit_entry
from arcrho_api.sidecar_core_contract import stored_lengths
from arcrho_api.timestamps import utc_now_text

from app_server.services import (
    dataset_sidecar_status_service,
    excel_service,
    user_identity_service,
)

_EXCEL_CANONICAL_RE = re.compile(
    r"^'((?:[^']|'')*)'!([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?$",
    re.IGNORECASE,
)


def _clean_text(value: Any) -> str:
    return str(value if value is not None else "").strip()


def _column_index(letters: str) -> int:
    column = 0
    for character in letters.upper():
        column = column * 26 + (ord(character) - 64)
    return column - 1


def _column_name(index: int) -> str:
    name = ""
    value = index + 1
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _excel_range(token_text: str) -> Dict[str, Any] | None:
    """Book path, sheet, and row-major cell addresses of one canonical Excel token."""

    match = _EXCEL_CANONICAL_RE.fullmatch(_clean_text(token_text))
    if not match:
        return None
    source = match.group(1).replace("''", "'")
    open_bracket = source.find("[")
    close_bracket = source.find("]", open_bracket + 1)
    if open_bracket < 0 or close_bracket <= open_bracket + 1 or close_bracket >= len(source) - 1:
        return None
    book_path = source[:open_bracket] + source[open_bracket + 1 : close_bracket]
    sheet = source[close_bracket + 1 :]
    start_row = int(match.group(3)) - 1
    start_col = _column_index(match.group(2))
    end_row = int(match.group(5)) - 1 if match.group(4) else start_row
    end_col = _column_index(match.group(4)) if match.group(4) else start_col
    row_first, row_last = sorted((start_row, end_row))
    col_first, col_last = sorted((start_col, end_col))
    cells = [
        f"{_column_name(col)}{row + 1}"
        for row in range(row_first, row_last + 1)
        for col in range(col_first, col_last + 1)
    ]
    return {
        "book_path": book_path,
        "sheet": sheet,
        "rows": row_last - row_first + 1,
        "cols": col_last - col_first + 1,
        "cells": cells,
    }


def _matrix_from_flat(rows: int, cols: int, flat: List[Any]) -> Dict[str, Any]:
    return {
        "rows": rows,
        "cols": cols,
        "values": [flat[row * cols : (row + 1) * cols] for row in range(rows)],
    }


def _values_equal(a: Any, b: Any) -> bool:
    if a is None or (isinstance(a, str) and not str(a).strip()):
        return b is None or (isinstance(b, str) and not str(b).strip())
    if b is None or (isinstance(b, str) and not str(b).strip()):
        return False
    try:
        return float(a) == float(b)
    except (TypeError, ValueError):
        return str(a) == str(b)


class _LinkRefreshHardError(Exception):
    """An ArcRho-side link failure: the dataset's refresh must report an error."""


def _load_source_datasets(
    project_name: str,
    reserving_class: str,
    names: List[str],
) -> Dict[str, Mapping[str, Any]]:
    from app_server.services import dataset_service
    from app_server.services.dfm_service import _key

    datasets: Dict[str, Mapping[str, Any]] = {}
    for name in names:
        key = _key(name)
        if key in datasets:
            continue
        try:
            datasets[key] = dataset_service.load_cached_dataset_values(
                project_name,
                reserving_class,
                name,
            )
        except HTTPException as err:
            raise _LinkRefreshHardError(f"Missing dependency: {name} ({err.detail})")
    return datasets


def _resolve_internal_matrix(
    reference_text: str,
    datasets: Mapping[str, Mapping[str, Any]],
) -> Dict[str, Any]:
    """The matrix one internal reference stands for, resolved like the Links tab."""

    from app_server.services.dataset_internal_link_service import _resolved_internal_reference
    from app_server.services.dfm_service import _key

    try:
        parsed = parse_internal_reference(reference_text)
    except DatasetLinkError as err:
        raise _LinkRefreshHardError(str(err))
    dataset = datasets.get(_key(parsed["dataset_name"]))
    if dataset is None:
        raise _LinkRefreshHardError(f"Missing dependency: {parsed['dataset_name']}")
    try:
        resolved = _resolved_internal_reference(reference_text, parsed, dataset)
    except HTTPException as err:
        raise _LinkRefreshHardError(str(err.detail))
    rows = int(resolved["row_count"])
    cols = int(resolved["column_count"])
    resolved["matrix"] = _matrix_from_flat(
        rows,
        cols,
        [cell["value"] for cell in resolved["cells"]],
    )
    return resolved


def _read_excel_matrices(
    formula_links: List[Mapping[str, Any]],
) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, str]]:
    """One batched workbook read for every Excel token across all formulas.

    Returns matrices and failures both keyed by the token's canonical text; a
    failed workbook or cell fails every token that reads it, nothing else.
    """

    ranges: Dict[str, Dict[str, Any]] = {}
    failures: Dict[str, str] = {}
    for link in formula_links:
        try:
            tokens = tokenize_dataset_formula(link.get("formula"))
        except DatasetLinkError:
            continue
        for token in tokens:
            if token["type"] != "reference" or token["kind"] != "excel":
                continue
            key = token["canonical"]
            if key in ranges or key in failures:
                continue
            span = _excel_range(key)
            if span is None:
                failures[key] = "The Excel reference could not be parsed."
                continue
            ranges[key] = span

    items: List[Dict[str, str]] = []
    spans: List[Tuple[str, Dict[str, Any], int]] = []
    for key, span in ranges.items():
        spans.append((key, span, len(items)))
        items.extend(
            {"book_path": span["book_path"], "sheet": span["sheet"], "cell": cell}
            for cell in span["cells"]
        )
    matrices: Dict[str, Dict[str, Any]] = {}
    if items:
        try:
            results = excel_service.excel_read_cells_batch(items).get("results") or []
        except Exception as err:  # A workbook reader crash is every token's soft failure.
            for key, _span, _start in spans:
                failures[key] = str(err)
            return matrices, failures
        for key, span, start in spans:
            flat: List[Any] = []
            error = ""
            for offset, cell in enumerate(span["cells"]):
                result = results[start + offset] if start + offset < len(results) else None
                if not isinstance(result, dict) or not result.get("ok"):
                    error = f"{cell}: {(result or {}).get('error') or 'The workbook cell could not be read.'}"
                    break
                flat.append(result.get("value"))
            if error:
                failures[key] = error
            else:
                matrices[key] = _matrix_from_flat(span["rows"], span["cols"], flat)
    return matrices, failures


def refresh_dataset_links(
    project_name: str,
    reserving_class: str,
    dataset_name: str,
) -> Dict[str, Any]:
    """Re-evaluate one dataset's ArcRho cell links and rewrite the linked cells.

    Returns ``{ok, dataset_name, refreshed, changed, warnings}`` on success —
    ``refreshed`` is false when the dataset has no links to evaluate, and each
    warning is ``{reference, reason}`` for an Excel-read failure whose cells
    kept their last values — or ``{ok: False, dataset_name, reason, errors}``
    when an ArcRho-side reference failed and nothing was written.
    """

    from app_server.services import dataset_service

    result: Dict[str, Any] = {
        "ok": True,
        "dataset_name": _clean_text(dataset_name),
        "refreshed": False,
        "changed": False,
        "warnings": [],
    }
    sidecar_path = dataset_sidecar_status_service.sidecar_path(
        project_name, reserving_class, dataset_name
    )
    sidecar = dataset_sidecar_status_service.read_sidecar(sidecar_path)
    if not sidecar:
        return {**result, "ok": False, "reason": "missing_sidecar", "errors": ["Dataset sidecar is missing."]}
    if dataset_sidecar_status_service.normalize_method_type(
        sidecar.get("method_type"), sidecar.get("source_kind")
    ) != dataset_sidecar_status_service.METHOD_TYPE_NONE:
        return result
    if _clean_text(sidecar.get("source_kind")).casefold() != "input":
        return result
    internal_links = [link for link in sidecar.get("internal_links") or [] if isinstance(link, dict)]
    formula_links = [link for link in sidecar.get("formula_links") or [] if isinstance(link, dict)]
    if not internal_links and not formula_links:
        return result

    try:
        # A link names a cell of the grid the dataset was displayed at when the
        # link was written, so the cells are read at that shape, which is the
        # file's own only while the two agree.
        target = dataset_service.load_cached_dataset_values(
            project_name, reserving_class, dataset_name, at_linked_shape=True
        )
    except HTTPException as err:
        return {**result, "ok": False, "reason": "missing_values", "errors": [str(err.detail)]}
    values = [list(row) for row in target.get("values") or []]

    source_names: List[str] = []
    for link in internal_links:
        try:
            source_names.append(parse_internal_reference(link.get("reference"))["dataset_name"])
        except DatasetLinkError as err:
            return {**result, "ok": False, "reason": "link_error", "errors": [str(err)]}
    for link in formula_links:
        try:
            tokens = tokenize_dataset_formula(link.get("formula"))
        except DatasetLinkError as err:
            return {**result, "ok": False, "reason": "link_error", "errors": [str(err)]}
        for token in tokens:
            if token["type"] == "reference" and token["kind"] == "internal":
                try:
                    source_names.append(parse_internal_reference(token["text"])["dataset_name"])
                except DatasetLinkError as err:
                    return {**result, "ok": False, "reason": "link_error", "errors": [str(err)]}

    errors: List[str] = []
    changed = False
    try:
        datasets = _load_source_datasets(project_name, reserving_class, source_names)

        for link in internal_links:
            resolved = _resolve_internal_matrix(str(link.get("reference") or ""), datasets)
            row_start = int(resolved["row_start"])
            column_start = int(resolved["column_start"])
            column_count = int(resolved["column_count"])
            cells = resolved["cells"]
            for target_cell in link.get("target_cells") or []:
                row_offset = int(target_cell.get("source_row", -1)) - row_start
                column_offset = int(target_cell.get("source_column", -1)) - column_start
                index = row_offset * column_count + column_offset
                cell = (
                    cells[index]
                    if 0 <= row_offset and 0 <= column_offset < column_count and index < len(cells)
                    else None
                )
                if (
                    not cell
                    or int(cell["row"]) != int(target_cell.get("source_row", -1))
                    or int(cell["column"]) != int(target_cell.get("source_column", -1))
                ):
                    raise _LinkRefreshHardError(
                        f"{link.get('reference')}: The referenced cells are no longer part of the source dataset."
                    )
                row_index = int(target_cell.get("row", -1))
                column_index = int(target_cell.get("column", -1))
                if not (0 <= row_index < len(values)) or not (0 <= column_index < len(values[row_index])):
                    raise _LinkRefreshHardError(
                        f"{link.get('reference')}: The linked dataset cell is no longer part of this dataset."
                    )
                if not _values_equal(values[row_index][column_index], cell["value"]):
                    values[row_index][column_index] = cell["value"]
                    changed = True

        excel_matrices, excel_failures = _read_excel_matrices(formula_links)
        for link in formula_links:
            formula = str(link.get("formula") or "")
            tokens = tokenize_dataset_formula(formula)
            failed_reference = next(
                (
                    token["canonical"]
                    for token in tokens
                    if token["type"] == "reference"
                    and token["kind"] == "excel"
                    and token["canonical"] in excel_failures
                ),
                "",
            )
            if failed_reference:
                result["warnings"].append({
                    "reference": formula,
                    "reason": (
                        "Excel value could not be read; the linked cells keep their last values. "
                        f"({excel_failures[failed_reference]})"
                    ),
                })
                continue

            def lookup(token: Mapping[str, Any]) -> Dict[str, Any] | None:
                if token["kind"] == "excel":
                    return excel_matrices.get(token["canonical"])
                return _resolve_internal_matrix(token["text"], datasets)["matrix"]

            try:
                tree = parse_dataset_formula_tree(tokens)
                matrix = evaluate_dataset_formula(tree, lookup)
            except DatasetLinkError as err:
                raise _LinkRefreshHardError(f"{formula}: {err}")
            for target_cell in link.get("target_cells") or []:
                result_row = int(target_cell.get("result_row", -1))
                result_column = int(target_cell.get("result_column", -1))
                if not (0 <= result_row) or not (0 <= result_column):
                    raise _LinkRefreshHardError(
                        f"{formula}: The formula result no longer covers the linked cells."
                    )
                value = (
                    matrix["values"][0 if matrix["rows"] == 1 else result_row][
                        0 if matrix["cols"] == 1 else result_column
                    ]
                    if (matrix["rows"] == 1 or result_row < matrix["rows"])
                    and (matrix["cols"] == 1 or result_column < matrix["cols"])
                    else None
                )
                if value is None:
                    raise _LinkRefreshHardError(
                        f"{formula}: The formula result no longer covers the linked cells."
                    )
                row_index = int(target_cell.get("row", -1))
                column_index = int(target_cell.get("column", -1))
                if not (0 <= row_index < len(values)) or not (0 <= column_index < len(values[row_index])):
                    raise _LinkRefreshHardError(
                        f"{formula}: The linked dataset cell is no longer part of this dataset."
                    )
                if not _values_equal(values[row_index][column_index], value):
                    values[row_index][column_index] = value
                    changed = True
    except _LinkRefreshHardError as err:
        errors.append(str(err))
    if errors:
        return {**result, "ok": False, "reason": "link_error", "errors": errors}

    result["refreshed"] = True
    result["changed"] = changed
    if not changed:
        return result

    import pandas as pd

    csv_path = str(target.get("path") or "")
    now = utc_now_text()
    sidecar["updated_at"] = now
    sidecar["modified_by"] = (
        user_identity_service.get_current_display_name()
        or _clean_text(os.environ.get("USERNAME"))
        or "unknown"
    )
    sidecar["audit_log"] = append_audit_entry(
        sidecar.get("audit_log"),
        event_date=now,
        action=AUDIT_ACTION_AUTO_REFRESH,
        user=sidecar["modified_by"],
    )
    from app_server import config
    from app_server.services.dataset_service import _write_dataset_csv_and_sidecar

    frame = pd.DataFrame(values).astype(object)
    view = (int(target.get("origin_length") or 0), int(target.get("development_length") or 0))
    if all(view) and view != stored_lengths(sidecar):
        # The cells were read at a coarser view of the store, so they go back
        # into it the way a save from that view does, and to the file itself
        # rather than the view's handle.
        frame = dataset_service.scatter_view_into_store(
            project_name,
            frame,
            stored_lengths=stored_lengths(sidecar),
            view_lengths=view,
            cumulative=bool(sidecar.get("cumulative", True)),
        )
        csv_path = os.path.join(
            config.get_project_dataset_cache_dir(project_name, reserving_class),
            str(sidecar.get("csv_file") or ""),
        )
    _write_dataset_csv_and_sidecar(frame, csv_path, sidecar_path, sidecar)
    return result
