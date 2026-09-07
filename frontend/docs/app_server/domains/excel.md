# App Server Domain: excel

## Purpose
<!-- MANUAL:BEGIN -->
Excel integration domain (workbook value reads, lightweight file metadata checks, and workbook operations).
<!-- MANUAL:END -->

## Entry Points
<!-- AUTO-GEN:BEGIN app_server.excel.entry_points -->
| Method | Path | Handler | Request Model | Schema | Service Calls |
| --- | --- | --- | --- | --- | --- |
| `POST` | `/excel/file_mtimes_batch` | `excel_file_mtimes_batch` | `ExcelFileMtimeBatchRequest` | [`app_server/schemas/excel.py`](../../../app_server/schemas/excel.py) | `excel_service.excel_file_mtimes_batch` |
| `POST` | `/excel/open_workbook` | `excel_open_workbook` | `ExcelOpenRequest` | [`app_server/schemas/excel.py`](../../../app_server/schemas/excel.py) | `excel_service.excel_open_workbook` |
| `POST` | `/excel/read_cell` | `excel_read_cell` | `ExcelCellReadRequest` | [`app_server/schemas/excel.py`](../../../app_server/schemas/excel.py) | `excel_service.excel_read_cell` |
| `POST` | `/excel/read_cells_batch` | `excel_read_cells_batch` | `ExcelBatchReadRequest` | [`app_server/schemas/excel.py`](../../../app_server/schemas/excel.py) | `excel_service.excel_read_cells_batch` |
| `POST` | `/excel/validate_links` | `excel_validate_links` | `ExcelBatchReadRequest` | [`app_server/schemas/excel.py`](../../../app_server/schemas/excel.py) | `excel_service.excel_validate_links` |
<!-- AUTO-GEN:END -->

## Key Files
<!-- AUTO-GEN:BEGIN app_server.excel.key_files -->
- [`app_server/api/excel_router.py`](../../../app_server/api/excel_router.py) - Excel COM automation routes.
- [`app_server/services/excel_service.py`](../../../app_server/services/excel_service.py) - Excel process interaction logic.
- [`app_server/schemas/excel.py`](../../../app_server/schemas/excel.py) - Excel request payload schemas.
<!-- AUTO-GEN:END -->

## External Interfaces
<!-- MANUAL:BEGIN -->
- Called by interactive Excel-based workflows.
- `/excel/file_mtimes_batch` resolves and deduplicates workbook paths, then reads file metadata with bounded concurrency while preserving request order. It does not open Excel or load workbook contents.
- `/excel/validate_links` is the check an opening Dataset or DFM method runs against its saved Excel links. It takes the same item list as `/excel/read_cells_batch` and returns the same per-cell `results`, plus a `workbooks` entry per distinct workbook carrying that file's `ok`/`mtime`. Both halves come from one grouped pass — the worker that opens a workbook also stats it — so an opening window can tell a broken reference from a merely newer workbook without a second round trip over a network share. Every cell is answered on its own: a missing sheet, an address the sheet cannot resolve, or a non-numeric value (a `#REF!` left by a deleted row) is that cell's error and leaves the other cells of the same workbook with their real values. `workbook_cell_value` is the single rule for what one workbook cell means and is shared by `excel_read_cell` and both batch reads: a cell that reads as empty — never filled, outside the used range, or a formula whose cached result is an empty or whitespace-only string — is the blank ArcRho stores as `null`, not an error. `/excel/read_cells_batch` keeps its narrower result shape, so the commit and refresh paths never pay for the workbook stats.
<!-- MANUAL:END -->

## Data/State/Caches
<!-- MANUAL:BEGIN -->
- A batch read groups its cells by workbook and then by sheet, and answers each sheet from one walk of the rectangle its requested cells span. A read-only worksheet re-reads the sheet from its first row every time it is asked for a single address, so reading a linked range address by address costs one pass per cell — a 120x120 range took roughly a quarter of an hour and now takes well under a second. Every walk is run to its end rather than abandoned once the last requested cell is answered, because a worksheet only releases the sheet's XML stream when its walk finishes, and a stream left open keeps the workbook file locked against the person editing it in Excel. `excel_read_cell` is the same reader asked for one cell, so a single read and a batch read can never disagree.
- Cell reads and the readability probe (`excel_workbook_readable`) are plain openpyxl file reads that need no Excel installation, so they run wherever the workbook is reachable — on ArcRho Server for the Excel Link Manager; only opening a workbook in Excel needs local automation. Timestamp checks use filesystem metadata only.
<!-- MANUAL:END -->

## Common Change Tasks
<!-- MANUAL:BEGIN -->
1. Add automation method: schema + router + service must stay aligned.
<!-- MANUAL:END -->

## Known Risks
<!-- MANUAL:BEGIN -->
- Excel COM timing and environment dependencies are fragile.
<!-- MANUAL:END -->
