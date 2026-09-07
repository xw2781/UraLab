// Excel Link Manager page.
//
// This runs inside a Project Instance nested window (pi-window). The host frame
// in project_instance_windows.js owns the titlebar, dragging, resizing,
// minimize/maximize/close and the dock; this page owns the inventory and the
// row actions reached from a right-click menu: open the workbook, open it
// read-only, change what it points at, and - on a Folder cell - open the
// containing folder. The window is pinned to the
// reserving class it was opened on, which arrives in the query string, so
// selecting another class in the tree leaves it alone exactly like a Dataset or
// DFM window.
//
// The table shows one row per usage - a workbook read by two datasets is two
// rows - and lives in excel_links_table.js, which owns the column widths,
// filters, and rendering. Clicking a Dataset Name cell asks the Project
// Instance page to open that dataset in DSV or that method in DFM, through the
// same arcrho:project-instance-open-dependent-dataset message a method page
// uses for a precedent.
//
// Everything about a workbook is answered by ArcRho Server: the listing's
// found/missing verdict is whether the server can open the file, and a change
// is an Engine-hosted job that opens the picked workbook there, refreshes every
// affected dataset and DFM, and flags them and their dependents Needs Review.
// Opening a workbook is the exception and runs on the client machine, through
// the desktop host, because that is where Excel is.
//
// Two messages go back to the Project Instance page, because the retarget
// writes files the host is watching:
//   arcrho:excel-links-retarget-begin  - suppress the host's index-change prompt
//   arcrho:excel-links-retarget-end    - restore it, report status, and reload
//                                        the cached dataset table when files changed
import { openContextMenu } from "/ui/shared/components/context_menu/context_menu.js?v=20260811b";
import { openPathThroughDesktopHost } from "/ui/shared/integrations/open_path.js?v=20260907b";
import { createExcelLinksTable, excelLinkDetailRows } from "/ui/project_instance/excel_links_table.js?v=20260818b";
import "/ui/shared/integrations/zoom_bridge.js?v=20260521a";

const LIST_ENDPOINT = "/excel_links/list";
const RETARGET_ENDPOINT = "/excel_links/retarget";
const EXCEL_FILE_FILTERS = [
  { name: "Excel Workbooks", extensions: ["xlsx", "xlsm", "xlsb", "xls"] },
  { name: "All Files", extensions: ["*"] },
];

function text(value) {
  return String(value ?? "").trim();
}

function count(value) {
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric > 0 ? Math.floor(numeric) : 0;
}

function detailMessage(payload, fallback) {
  const detail = payload?.detail;
  if (typeof detail === "string" && detail.trim()) return detail.trim();
  if (Array.isArray(detail) && detail.length) return detail.map((item) => item?.msg || String(item)).join("; ");
  return fallback;
}

export function normalizeExcelLinkWorkbooks(value) {
  const source = Array.isArray(value) ? value : [];
  return source
    .map((item) => ({
      workbookPath: text(item?.workbook_path),
      workbookName: text(item?.workbook_name) || text(item?.workbook_path),
      folder: text(item?.folder),
      exists: item?.exists === true,
      // The workbook's own Created/Modified/Last saved by, the workbook-side
      // answer to the dataset table's Created, Last Modified, and User. A
      // workbook that carries none - a legacy .xls, an encrypted package -
      // leaves them blank.
      created: text(item?.created),
      modified: text(item?.modified),
      lastModifiedBy: text(item?.last_modified_by),
      datasetCount: count(item?.dataset_count),
      methodCount: count(item?.method_count),
      linkCount: count(item?.link_count),
      cellCount: count(item?.cell_count),
      usages: (Array.isArray(item?.usages) ? item.usages : [])
        .map((usage) => ({
          kind: usage?.kind === "dfm" ? "dfm" : "dataset",
          name: text(usage?.name),
          datasetType: text(usage?.dataset_type),
          methodType: text(usage?.method_type),
          linkCount: count(usage?.link_count),
          cellCount: count(usage?.cell_count),
        }))
        .filter((usage) => usage.name),
    }))
    .filter((item) => item.workbookPath);
}

export function excelLinkInventorySummary({ workbookCount, visibleRows, totalRows, scanErrorCount }) {
  const books = count(workbookCount);
  const total = count(totalRows);
  const visible = count(visibleRows);
  if (!books) return "";
  const workbooks = `${books} linked workbook${books === 1 ? "" : "s"}`;
  const references = visible === total
    ? `${total} reference${total === 1 ? "" : "s"}`
    : `${visible} of ${total} references shown`;
  const errors = count(scanErrorCount);
  const skipped = errors ? ` ${errors} file${errors === 1 ? "" : "s"} could not be read.` : "";
  return `${workbooks}, ${references}.${skipped}`;
}

export function excelLinkRetargetSummary(payload) {
  const results = Array.isArray(payload?.results) ? payload.results : [];
  const failures = results.filter((item) => item?.ok === false);
  const changedFiles = count(payload?.changed_file_count);
  const changedLinks = count(payload?.changed_link_count);
  if (failures.length) {
    const first = failures[0];
    const name = text(first?.name) || "a file";
    const error = text(first?.error) || "The file could not be updated.";
    const others = failures.length > 1 ? ` (+${failures.length - 1} more)` : "";
    return {
      ok: false,
      message: `Updated ${changedFiles} of ${changedFiles + failures.length} files; ${name}: ${error}${others}`,
    };
  }
  if (!changedFiles) {
    return { ok: true, message: text(payload?.message) || "No saved links needed a change." };
  }
  const relinked = `Updated ${changedLinks} link${changedLinks === 1 ? "" : "s"} in ${changedFiles} file${changedFiles === 1 ? "" : "s"}`;
  const refreshedCells = count(payload?.refreshed_cell_count);
  const changedValueFiles = count(payload?.value_changed_file_count);
  const values = changedValueFiles
    ? `values changed in ${changedValueFiles} file${changedValueFiles === 1 ? "" : "s"}`
    : "stored values already matched";
  const refreshed = `recalculated ${refreshedCells} linked cell${refreshedCells === 1 ? "" : "s"} (${values}).`;
  const failedRefresh = count(payload?.failed_refresh_count);
  const failed = failedRefresh
    ? ` ${failedRefresh} linked cell${failedRefresh === 1 ? "" : "s"} could not be recalculated and kept the stored values.`
    : "";
  let propagation = " Affected objects and their dependents are marked Needs Review.";
  if (payload?.propagation_ok === false) {
    propagation = " Dependent recalculation reported a problem; check the affected pages.";
  } else if (text(payload?.propagation?.status) === "queued") {
    propagation = " Dependent recalculation has started; affected objects are marked Needs Review.";
  }
  return {
    ok: !failedRefresh && payload?.propagation_ok !== false,
    message: `${relinked}; ${refreshed}${failed}${propagation}`,
  };
}

const params = new URLSearchParams(window.location.search);
const inst = text(params.get("inst"));
const projectName = text(params.get("project"));
const reservingClass = text(params.get("class"));

const els = {
  refresh: document.getElementById("excelLinksRefresh"),
  table: document.getElementById("excelLinksTable"),
  wrap: document.getElementById("excelLinksTableWrap"),
  state: document.getElementById("excelLinksState"),
  status: document.getElementById("excelLinksStatus"),
  menu: document.getElementById("excelLinksMenu"),
  filterPopover: document.getElementById("excelLinksFilterPopover"),
};

const manager = {
  loading: false,
  busy: false,
  workbooks: [],
  rows: [],
  visibleRows: 0,
  requestSeq: 0,
  scanErrorCount: 0,
  // The detail row the context menu is open for, and its <tr>.
  menuRow: null,
  menuRowEl: null,
};

function postToParent(type, payload = {}) {
  try {
    window.parent?.postMessage({ type, inst, ...payload }, "*");
  } catch {}
}

function hostApi() {
  try {
    return window.ADAHost || window.parent?.ADAHost || window.top?.ADAHost || null;
  } catch {
    return null;
  }
}

function setManagerStatus(message, tone = "") {
  if (!els.status) return;
  els.status.textContent = text(message);
  els.status.className = `pi-excel-links-status${tone ? ` ${tone}` : ""}`;
}

function syncControls() {
  const blocked = manager.busy || manager.loading;
  if (els.refresh) els.refresh.disabled = blocked;
  if (blocked) {
    closeMenu();
    table?.closeFilterPopover();
  }
  document.body.setAttribute("aria-busy", blocked ? "true" : "false");
}

function setBusy(busy) {
  manager.busy = !!busy;
  syncControls();
}

function showState(message) {
  if (!els.state) return;
  els.state.textContent = text(message);
  els.state.hidden = !els.state.textContent;
}

function syncInventoryStatus() {
  setManagerStatus(excelLinkInventorySummary({
    workbookCount: manager.workbooks.length,
    visibleRows: manager.visibleRows,
    totalRows: manager.rows.length,
    scanErrorCount: manager.scanErrorCount,
  }), manager.scanErrorCount ? "error" : "");
}

const table = createExcelLinksTable({
  table: els.table,
  wrap: els.wrap,
  popover: els.filterPopover,
  onOpenUsage: (row) => openUsage(row),
  onRowMenu: (row, rowEl, event, columnKey) => openMenu(row, rowEl, event, columnKey),
  onViewChange: ({ visible, total, filtered }) => {
    manager.visibleRows = visible;
    if (manager.loading) return;
    if (!total) showState(manager.workbooks.length ? "" : "No Excel links are saved in this reserving class.");
    else showState(visible ? "" : "No rows match the current column filters.");
    if (filtered || total) syncInventoryStatus();
  },
});

function setRows(workbooks) {
  manager.workbooks = workbooks;
  manager.rows = excelLinkDetailRows(workbooks);
  closeMenu();
  table.setRows(manager.rows);
  syncControls();
}

async function loadExcelLinks() {
  const seq = ++manager.requestSeq;
  setRows([]);
  if (!projectName || !reservingClass) {
    manager.loading = false;
    showState("This window is missing its project or reserving class.");
    setManagerStatus("");
    syncControls();
    return;
  }
  manager.loading = true;
  showState("Loading Excel links...");
  setManagerStatus("");
  syncControls();
  try {
    const response = await fetch(LIST_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ project_name: projectName, reserving_class: reservingClass }),
    });
    const payload = await response.json().catch(() => ({}));
    if (seq !== manager.requestSeq) return;
    if (!response.ok || payload?.ok === false) {
      throw new Error(detailMessage(payload, `HTTP ${response.status}`));
    }
    manager.loading = false;
    manager.scanErrorCount = Array.isArray(payload?.errors) ? payload.errors.length : 0;
    setRows(normalizeExcelLinkWorkbooks(payload?.workbooks));
    if (!manager.workbooks.length) {
      showState("No Excel links are saved in this reserving class.");
      setManagerStatus("");
    }
  } catch (error) {
    if (seq !== manager.requestSeq) return;
    manager.loading = false;
    showState("Excel links could not be loaded.");
    setManagerStatus(`Could not load Excel links: ${error.message}`, "error");
  } finally {
    if (seq === manager.requestSeq) syncControls();
  }
}

// ---------------------------------------------------------------------------
// Open the dataset or DFM method a row names
// ---------------------------------------------------------------------------

// The Project Instance page owns every dataset and method window, so the same
// message a method page sends for a precedent opens the row's object here: a
// dataset row lands in DSV, a DFM row in the DFM page, both pinned to this
// window's reserving class.
function openUsage(row) {
  const name = text(row?.name);
  if (!name || manager.busy || manager.loading) return;
  const isDfm = row?.kind === "dfm";
  postToParent("arcrho:project-instance-open-dependent-dataset", {
    datasetName: name,
    reservingClass,
    projectName,
    openMethod: isDfm,
    ...(isDfm
      ? { methodType: "DFM", methodName: name }
      // The listing names the instance's Dataset Type and Method Type, so an
      // instance whose name differs from its type opens without the host
      // guessing from the reserving class the tree happens to be showing.
      : {
        datasetTypeName: text(row?.datasetType) || name,
        ...(text(row?.methodType) ? { methodType: text(row.methodType) } : {}),
      }),
  });
  setManagerStatus(isDfm ? `Opening DFM method ${name}...` : `Opening dataset ${name}...`);
}

// ---------------------------------------------------------------------------
// Row context menu
// ---------------------------------------------------------------------------

function closeMenu() {
  if (!manager.menuRow || !els.menu) return;
  // openContextMenu shows the menu with an inline display; clearing it hands
  // the menu back to the stylesheet's hidden default.
  els.menu.style.display = "";
  manager.menuRowEl?.classList.remove("context-target");
  manager.menuRow = null;
  manager.menuRowEl = null;
}

function openMenu(row, rowEl, event, columnKey = "") {
  closeMenu();
  if (manager.busy || manager.loading || !els.menu) return;
  table.closeFilterPopover();
  manager.menuRow = row;
  manager.menuRowEl = rowEl;
  rowEl.classList.add("context-target");
  // Opening the folder belongs to the Folder cell, so it appears only there
  // rather than adding a fourth item to every row's menu.
  const folderItem = els.menu.querySelector('[data-action="open-folder"]');
  if (folderItem) folderItem.hidden = columnKey !== "folder" || !text(row?.folder);
  openContextMenu(els.menu, {
    anchorEl: rowEl,
    clientX: Number(event?.clientX),
    clientY: Number(event?.clientY),
    offset: 8,
    align: "top-left",
  });
  els.menu.querySelector(".ctx-item")?.focus();
}

function wireMenu() {
  if (!els.menu) return;
  els.menu.addEventListener("click", (event) => {
    const item = event.target.closest?.(".ctx-item");
    if (!item) return;
    const row = manager.menuRow;
    closeMenu();
    if (!row) return;
    const action = item.dataset.action;
    if (action === "open-workbook") void openWorkbook(row, false);
    else if (action === "open-workbook-read-only") void openWorkbook(row, true);
    else if (action === "open-folder") void openFolder(row);
    else if (action === "change-link") void changeWorkbook(row);
  });
  document.addEventListener("mousedown", (event) => {
    if (!els.menu.contains(event.target)) closeMenu();
  }, true);
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeMenu();
  }, true);
  els.wrap?.addEventListener("scroll", closeMenu);
  window.addEventListener("resize", closeMenu);
  window.addEventListener("blur", closeMenu);
}

// ---------------------------------------------------------------------------
// Open workbook and folder
// ---------------------------------------------------------------------------

// Opening runs on the client machine through the desktop host, not on ArcRho
// Server: the workbook opens for this user in their own Excel, so a workbook
// the server cannot reach can still open here and the reverse. The same route
// hands a folder to File Explorer.
async function openThroughDesktopHost(path, messages, readOnly = false) {
  if (!path || manager.busy || manager.loading) return;
  setBusy(true);
  setManagerStatus(messages.opening);
  try {
    const result = await openPathThroughDesktopHost(path, { readOnly: !!readOnly });
    if (result?.ok) {
      setManagerStatus(messages.opened, "success");
    } else {
      setManagerStatus(`Could not open ${path}: ${text(result?.error) || messages.failed}`, "error");
    }
  } catch (error) {
    setManagerStatus(`Could not open ${path}: ${error.message}`, "error");
  } finally {
    setBusy(false);
  }
}

function openWorkbook(row, readOnly) {
  const name = text(row?.workbookName) || text(row?.workbookPath);
  return openThroughDesktopHost(text(row?.workbookPath), {
    opening: readOnly ? `Opening ${name} read-only...` : `Opening ${name}...`,
    opened: readOnly ? "Workbook opened read-only." : "Workbook opened.",
    failed: readOnly ? "The workbook could not be opened read-only." : "The workbook could not be opened.",
  }, readOnly);
}

function openFolder(row) {
  const folder = text(row?.folder);
  return openThroughDesktopHost(folder, {
    opening: `Opening ${folder}...`,
    opened: "Folder opened in File Explorer.",
    failed: "The folder could not be opened.",
  });
}

// ---------------------------------------------------------------------------
// Change link
// ---------------------------------------------------------------------------

async function changeWorkbook(row) {
  if (manager.busy || manager.loading) return;
  if (!reservingClass) return;
  const host = hostApi();
  if (!host?.pickOpenFile) {
    setManagerStatus("Changing links is available in the desktop app only.", "error");
    return;
  }
  let picked = "";
  try {
    picked = text(await host.pickOpenFile({
      startDir: row.folder,
      filters: EXCEL_FILE_FILTERS,
    }));
  } catch {
    picked = "";
  }
  if (!picked) return;

  const seq = ++manager.requestSeq;
  setBusy(true);
  setManagerStatus(`Relinking ${row.workbookName} to ${picked} on ArcRho Server and recalculating affected datasets and DFM methods...`);
  // The retarget rewrites files and rebuilds index.json on the server; the
  // host suppresses its own disk-watch prompt for this window's change.
  postToParent("arcrho:excel-links-retarget-begin");
  let summary = { ok: false, message: "" };
  let payload = null;
  try {
    const response = await fetch(RETARGET_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        project_name: projectName,
        reserving_class: reservingClass,
        old_workbook_path: row.workbookPath,
        new_workbook_path: picked,
      }),
    });
    payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(detailMessage(payload, `HTTP ${response.status}`));
    summary = excelLinkRetargetSummary(payload);
    if (seq === manager.requestSeq) {
      setRows(normalizeExcelLinkWorkbooks(payload?.workbooks));
      if (!manager.workbooks.length) showState("No Excel links are saved in this reserving class.");
    }
    setManagerStatus(summary.message, summary.ok ? "success" : "error");
  } catch (error) {
    // A refused workbook arrives as the server's own verdict ("ArcRho Server
    // cannot open the selected workbook: ..."); the picked path is named here
    // because the server redacts paths from its messages.
    setManagerStatus(`Could not change the link to ${picked}: ${error.message}`, "error");
  } finally {
    postToParent("arcrho:excel-links-retarget-end", {
      ok: !!summary.ok,
      workbookPath: picked,
      changedFileCount: count(payload?.changed_file_count),
    });
    setBusy(false);
  }
}

// The host posts arcrho:set-zoom to every nested window on load, so this page
// scales with the app exactly like a Dataset or DFM window.
window.ArcRhoZoomBridge?.wirePageZoomBridge();

els.refresh?.addEventListener("click", () => void loadExcelLinks());
wireMenu();
void loadExcelLinks();
