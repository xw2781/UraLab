"""Compare persisted ArcRho DFM ratio triangles against live ResQ ratio triangles.

Scope is fixed to one project and the 17 reserving-class paths listed below.
For every DFM method present in a reserving class, this reads:

  - the ratio triangle already persisted in the ArcRho method JSON
    (``ratios_tab.ratio_triangle.ratio_values``), and
  - the live ResQ ratio triangle via ``dfm.Ratios(OriginIndex, DevIndex)``,

then reports the largest absolute difference between the two, rounded to six
decimal places. Nothing is written back to ArcRho or ResQ.

Run with Python 3.10 from the repository root:

    py -3.10 python-api/migration/validation/compare_dfm_ratio_triangles.py
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path

_VALIDATION_DIR = Path(__file__).resolve().parent
_MIGRATION_DIR = _VALIDATION_DIR.parent
if str(_MIGRATION_DIR) not in sys.path:
    sys.path.insert(0, str(_MIGRATION_DIR))

import resq_data_migration as migration  # noqa: E402
from resq_migration.core import _clean_name, _encode_rc_folder, _safe_attr  # noqa: E402


TARGET_PROJECT_NAME = "NJ_Annual_Prod_2026 Q3-Aug"
RC_PATHS = [
    r"PRNJ - PA\PA\NY\Direct Group\BI Total",
    r"PRNJ - PA\PA\NY\Direct Group\MP+PIP",
    r"PRNJ - PA\PA\Penn+CT\Direct Group\BI Total",
    r"PRNJ - PA\PA\Penn+CT\Direct Group\MP+PIP",
    r"PRNJ - PA\PA\All States\Direct Group\PD+UMPD",
    r"PRNJ - PA\PA\All States\Direct Group\COL",
    r"PRNJ - PA\PA\All States\Direct Group\CMPxCAT",
    r"PRNJ - PA\PA\NJ\Direct Group\MP+PIP",
    r"PRNJ - PA\PA\NJ\Direct Group\BIR51+UMBIR51",
    r"PRNJ - PA\PA\NJ\Direct Group\BIx51+UMBIx51",
    r"HPPREF\HO+DF\NJ\Legacy\HOL",
    r"HPPREF\HO+DF\NJ\Legacy\HOPxCAT",
    r"Rider\MC\All States\Direct Group\BI+PIP",
    r"Rider\MC\All States\Direct Group\PD+UMPD",
    r"Rider\MC\All States\Direct Group\PhysDxCat",
    r"PRNJ - PA\PA\MA\Direct Group\BI Total",
    r"PRNJ - PA\PA\MA\Direct Group\MP+PIP",
]

DECIMAL_PLACES = 6
TOLERANCE = 0.5 * 10 ** (-DECIMAL_PLACES)
OUTPUT_PATH = _VALIDATION_DIR / "results" / f"dfm_ratio_diffs_{TARGET_PROJECT_NAME}.xlsx"


def _read_arcrho_dfm_methods(rc_dir: Path) -> dict[str, dict]:
    """Map DFM name -> persisted ArcRho method payload for one reserving class."""

    methods_dir = rc_dir / "methods"
    out: dict[str, dict] = {}
    if not methods_dir.is_dir():
        return out
    for path in methods_dir.glob("DFM@*.json"):
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        name = _clean_name(payload.get("details_tab", {}).get("name")) or path.stem
        out[name] = payload
    return out


def _arcrho_ratio_matrix(payload: dict) -> list[list[float | None]]:
    ratio_triangle = payload.get("ratios_tab", {}).get("ratio_triangle", {})
    values = ratio_triangle.get("ratio_values")
    return values if isinstance(values, list) else []


def _resq_ratio_matrix(dfm) -> list[list[float | None]]:
    origin_count = int(_safe_attr(dfm, "OriginCount", 0) or 0)
    matrix: list[list[float | None]] = []
    for i in range(1, origin_count + 1):
        try:
            row_dev = int(dfm.DevelopmentCount(i))
        except Exception:
            row_dev = 0
        row: list[float | None] = []
        for j in range(1, row_dev + 1):
            try:
                v = dfm.Ratios(OriginIndex=i, DevIndex=j)
                row.append(float(v) if v is not None else None)
            except Exception:
                row.append(None)
        matrix.append(row)
    return matrix


def _compare_ratio_matrices(
    arcrho_matrix: list[list[float | None]],
    resq_matrix: list[list[float | None]],
) -> tuple[float | None, str]:
    """Return (max abs diff, note). Note flags a shape mismatch, if any."""

    arcrho_shape = (len(arcrho_matrix), max((len(r) for r in arcrho_matrix), default=0))
    resq_shape = (len(resq_matrix), max((len(r) for r in resq_matrix), default=0))
    note = "" if arcrho_shape == resq_shape else f"shape mismatch: ArcRho {arcrho_shape} vs ResQ {resq_shape}"

    max_diff: float | None = None
    row_limit = min(len(arcrho_matrix), len(resq_matrix))
    for i in range(row_limit):
        a_row, r_row = arcrho_matrix[i], resq_matrix[i]
        col_limit = min(len(a_row), len(r_row))
        for j in range(col_limit):
            a_val, r_val = a_row[j], r_row[j]
            if a_val is None or r_val is None:
                if a_val is not r_val:
                    note = note or f"missing cell at origin {i + 1}, dev {j + 1}"
                continue
            diff = abs(round(float(a_val), DECIMAL_PLACES) - round(float(r_val), DECIMAL_PLACES))
            if max_diff is None or diff > max_diff:
                max_diff = diff
    return max_diff, note


def run_comparison(app_factory=None, progress=print) -> list[dict]:
    """Compare every DFM in scope and return rows describing differences found."""

    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32 is required: pip install pywin32") from exc

    previous_scope = migration._apply_runtime_scope(TARGET_PROJECT_NAME, migration.SERVER_ROOT)
    app = app_factory() if app_factory is not None else win32com.client.Dispatch("ResQ3Automation.ResQApplication")
    rows: list[dict] = []
    try:
        app.ConnectByName(migration.CONNECTION_NAME, migration.USER_NAME, migration.PASSWORD)
        project = app.Projects().Item(TARGET_PROJECT_NAME)

        for rc_index, rc_path in enumerate(RC_PATHS, start=1):
            progress(f"RC {rc_index}/{len(RC_PATHS)}: {rc_path}")
            rc_dir = migration.PROJECT_DATA_DIR / _encode_rc_folder(rc_path)
            arcrho_methods = _read_arcrho_dfm_methods(rc_dir)

            try:
                reserving_class = project.ReservingClasses().Item(rc_path)
                dfm_collection = list(reserving_class.DFMMethods())
            except Exception as exc:
                rows.append(
                    {
                        "rc_path": rc_path,
                        "dfm_name": "",
                        "max_diff": None,
                        "note": f"could not read ResQ reserving class: {type(exc).__name__}: {exc}",
                    }
                )
                continue

            seen_names: set[str] = set()
            for dfm in dfm_collection:
                name = _clean_name(_safe_attr(dfm, "Name", ""))
                if not name:
                    continue
                seen_names.add(name)
                arcrho_payload = arcrho_methods.get(name)
                if arcrho_payload is None:
                    rows.append(
                        {
                            "rc_path": rc_path,
                            "dfm_name": name,
                            "max_diff": None,
                            "note": "DFM exists in ResQ but no persisted ArcRho method JSON was found",
                        }
                    )
                    continue

                resq_matrix = _resq_ratio_matrix(dfm)
                arcrho_matrix = _arcrho_ratio_matrix(arcrho_payload)
                max_diff, note = _compare_ratio_matrices(arcrho_matrix, resq_matrix)
                if (max_diff is not None and max_diff > TOLERANCE) or note:
                    rows.append(
                        {
                            "rc_path": rc_path,
                            "dfm_name": name,
                            "max_diff": max_diff,
                            "note": note,
                        }
                    )

            for name in arcrho_methods.keys() - seen_names:
                rows.append(
                    {
                        "rc_path": rc_path,
                        "dfm_name": name,
                        "max_diff": None,
                        "note": "DFM has a persisted ArcRho method JSON but was not found in ResQ",
                    }
                )
    finally:
        try:
            app.Disconnect()
        except Exception:
            pass
        migration._restore_runtime_scope(previous_scope)
    return rows


def write_workbook(path: Path, rows: list[dict]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "DFM Ratio Diffs"
    headers = ["RC Path", "DFM Name", f"Max Diff (>{DECIMAL_PLACES} dp)", "Note"]
    sheet.append(headers)
    for row in rows:
        sheet.append([row["rc_path"], row["dfm_name"], row["max_diff"], row["note"]])
    if not rows:
        sheet.append(["No differences found.", "", "", ""])

    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.freeze_panes = "A2"
    for column_cells in sheet.columns:
        width = max(len(str(cell.value) if cell.value is not None else "") for cell in column_cells)
        sheet.column_dimensions[column_cells[0].column_letter].width = min(max(width + 2, 12), 80)

    descriptor, temporary_name = tempfile.mkstemp(prefix=f".{path.stem}-", suffix=".xlsx", dir=path.parent)
    os.close(descriptor)
    temporary_path = Path(temporary_name)
    try:
        workbook.save(temporary_path)
        os.replace(temporary_path, path)
    finally:
        workbook.close()
        if temporary_path.exists():
            temporary_path.unlink()


def main() -> int:
    rows = run_comparison()
    write_workbook(OUTPUT_PATH, rows)
    print(f"Compared {len(RC_PATHS)} reserving classes. Found {len(rows)} problematic DFM instance(s).")
    print(f"Excel report: {OUTPUT_PATH}")
    return 0 if not rows else 1


if __name__ == "__main__":
    raise SystemExit(main())
