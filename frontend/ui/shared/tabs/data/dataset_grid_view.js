// Rendering only: read state.model + state.showBlanks and produce DOM.

import { state } from "/ui/shared/dataset/dataset_state.js";
import { $ } from "/ui/shared/tabs/data/data_tab_dom.js";
import { openContextMenu } from "/ui/shared/components/context_menu/context_menu.js";
import {
  clampDatasetDecimalPlaces,
  formatDatasetNumberValue,
  normalizeDatasetNumberFormat,
} from "/ui/shared/dataset/dataset_number_format.js";
import {
  formatDatasetOriginLabel,
  getDatasetOriginLabelText,
} from "/ui/shared/dataset/dataset_origin_labels.js";
import { renderDatasetGridPlaceholder } from "/ui/shared/tabs/data/dataset_grid_placeholder.js?v=20260809a";
import { renderDataTabChart } from "/ui/shared/tabs/data/data_tab_chart_port.js";
import { isDfmDataTabHost } from "/ui/shared/tabs/data/data_tab_context.js";
import {
  getDatasetGridTotalLayout,
  shouldShowDatasetGridTotals,
  sumDatasetGridColumn,
  sumDatasetGridRow,
} from "/ui/shared/tabs/data/dataset_grid_totals.js?v=20260830a";

let ctxMenuWired = false;
let renderNumberFormatSettings = null;
let renderVectorColumnLabel = "";
let gridEditConfig = null;

function normalizeRenderNumberFormatSettings(settings = null) {
  if (!settings || typeof settings !== "object") return null;
  return {
    numberFormat: normalizeDatasetNumberFormat(
      settings.number_format ?? settings.numberFormat ?? settings.num_format,
    ),
    decimalPlaces: clampDatasetDecimalPlaces(settings.decimal_places ?? settings.decimalPlaces),
  };
}

export function setDatasetRenderNumberFormatSettings(settings = null) {
  renderNumberFormatSettings = normalizeRenderNumberFormatSettings(settings);
}

export function setDatasetRenderVectorColumnLabel(label = "") {
  renderVectorColumnLabel = String(label || "").trim();
}

export function setDatasetGridEditConfig(config = null) {
  gridEditConfig = config && typeof config === "object" ? config : null;
}

// --- keyboard focus sink: make sure this document receives keydown after clicking a cell ---
function ensureKeySink() {
  let el = document.getElementById("keySink");
  if (el) return el;

  el = document.createElement("div");
  el.id = "keySink";
  el.tabIndex = 0;                 // make it focusable
  el.setAttribute("aria-hidden", "true");
  el.style.position = "fixed";
  el.style.left = "-9999px";
  el.style.top = "0";
  el.style.width = "1px";
  el.style.height = "1px";
  el.style.opacity = "0";
  document.body.appendChild(el);
  return el;
}

function claimDatasetFocus() {
  try { window.focus(); } catch {}
  const sink = ensureKeySink();
  try { sink.focus({ preventScroll: true }); } catch { try { sink.focus(); } catch {} }
}

function configureSelectableDatasetCell(cell, rowIndex, columnIndex, options = {}) {
  cell.classList.add("cell");
  cell.dataset.r = String(rowIndex);
  cell.dataset.c = String(columnIndex);
  cell.setAttribute("aria-selected", "false");
  if (options.readOnly) cell.setAttribute("aria-readonly", "true");
  if (Object.prototype.hasOwnProperty.call(options, "copyValue")) {
    cell.dataset.copyValue = options.copyValue == null ? "" : String(options.copyValue);
  }

  cell.addEventListener("click", (event) => {
    if (event.target?.closest?.(".dsCellInput")) return;
    claimDatasetFocus();
  });

  cell.addEventListener("contextmenu", (event) => {
    event.preventDefault();
    gridEditConfig?.onCellContextMenu?.(rowIndex, columnIndex);
    showCtxMenu(cell, event.clientX, event.clientY);
  });
}

const fmt0 = new Intl.NumberFormat("en-US", {
  minimumFractionDigits: 0,
  maximumFractionDigits: 0,
});
const DFM_PERCENT_DECIMAL_PLACES = 1;

function isPercentTriangle() {
  const triInput = document.getElementById("triInput");
  return triInput && triInput.value.includes("%");
}

function normalizeDatasetTypeKey(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function getCurrentDatasetTypeFormula() {
  const modelFormula = String(state.model?.formula || "").trim();
  if (modelFormula) return modelFormula;
  const tri = String(document.getElementById("triInput")?.value || "").trim();
  if (!tri) return "";
  const key = normalizeDatasetTypeKey(tri);
  if (!key) return "";
  const formulaMap = state.datasetTypeFormulaByKey instanceof Map ? state.datasetTypeFormulaByKey : null;
  if (!formulaMap) return "";
  return String(formulaMap.get(key) || "").trim();
}

function ensureCtxMenuWired() {
  if (ctxMenuWired) return;
  ctxMenuWired = true;

  const menu = document.getElementById("ctxMenu");
  if (!menu) return;

  menu.addEventListener("click", async (e) => {
    const btn = e.target.closest(".ctx-item");
    if (!btn) return;
    const action = btn.dataset.action || "";
    if (action === "copy_value" && typeof window.__arcRhoCopyActiveGridSelection === "function") {
      await window.__arcRhoCopyActiveGridSelection();
    } else if (gridEditConfig?.onContextAction) {
      try {
        await gridEditConfig.onContextAction(action);
      } catch (error) {
        console.error("Dataset grid context action failed", error);
      }
    }
    hideCtxMenu();
    claimDatasetFocus();
  });

  // Click anywhere else -> hide
  document.addEventListener("mousedown", (e) => {
    if (!menu.contains(e.target)) hideCtxMenu();
  });

  // ESC -> hide
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") hideCtxMenu();
  });

  // Scroll/resize -> hide (prevents "floating" menu)
  window.addEventListener("scroll", hideCtxMenu, true);
  window.addEventListener("resize", hideCtxMenu);
}

function showCtxMenu(anchorEl, clientX, clientY) {
  const menu = document.getElementById("ctxMenu");
  if (!menu) return;
  const pasteButton = menu.querySelector('[data-action="paste"]');
  if (pasteButton) pasteButton.hidden = !gridEditConfig?.canPasteSelection?.();
  const clearButton = menu.querySelector('[data-action="clear_data"]');
  if (clearButton) clearButton.hidden = !gridEditConfig?.canClearData?.();
  openContextMenu(menu, {
    anchorEl,
    clientX,
    clientY,
    offset: 8,
    align: "top-left",
  });
}

function hideCtxMenu() {
  const menu = document.getElementById("ctxMenu");
  if (!menu) return;
  menu.style.display = "none";
}

function getDecimalPlaces() {
  if (renderNumberFormatSettings) return renderNumberFormatSettings.decimalPlaces;
  if (!document.getElementById("numberFormatSelect")) return 1;
  const el = document.getElementById("decimalPlaces");
  const n = parseInt(el?.value, 10);
  if (!Number.isFinite(n)) return 1;
  return Math.max(0, Math.min(6, n)); // clamp 0..6
}

function getNumberFormatPattern() {
  if (renderNumberFormatSettings) return renderNumberFormatSettings.numberFormat;
  const input = document.getElementById("numberFormatSelect");
  return input ? normalizeDatasetNumberFormat(input.value) : "";
}

function getPercentDecimalPlaces() {
  return isDfmDataTabHost() ? DFM_PERCENT_DECIMAL_PLACES : getDecimalPlaces();
}

function detectNumberMode() {
  // 1) name contains % => percent
  if (isPercentTriangle()) return "percent";

  const model = state.model;
  if (!model || !Array.isArray(model.values) || !Array.isArray(model.mask)) {
    return "int";
  }

  const vals = model.values;
  const mask = model.mask;

  // 2) scan dataset: all non-zero numeric values in (0,1) => decimal
  let sawNonZero = false;

  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < (vals[r] || []).length; c++) {
      if (!mask[r] || !mask[r][c]) continue;

      const v = vals[r][c];
      if (v === null || v === undefined || v === "") continue;

      const n = (typeof v === "number") ? v : Number(v);
      if (!Number.isFinite(n)) continue;

      if (n === 0) continue; // exclude 0 from the check (allowed to exist)

      sawNonZero = true;

      // if ANY non-zero value is outside (0,1), it's not a ratio-like dataset
      const abs = Math.abs(n);
      if (!(abs > 0 && abs < 1)) return "int";
    }
  }

  return sawNonZero ? "decimal" : "int";
}

export function formatCellValue(v) {
  if (v === null || v === undefined || v === "") return "";

  const n = (typeof v === "number") ? v : Number(v);
  if (!Number.isFinite(n)) return "";

  const pattern = getNumberFormatPattern();
  if (pattern) return formatDatasetNumberValue(n, pattern, getDecimalPlaces());

  const mode = detectNumberMode();
  const dp = getDecimalPlaces();

  if (mode === "percent") {
    return (n * 100).toFixed(getPercentDecimalPlaces()) + "%";
  }

  if (mode === "decimal") {
    return n.toFixed(dp); // 0.000 style (no comma)
  }

  // default: 0,000
  return fmt0.format(n);
}

export function formatDatasetChartValue(value) {
  if (!Number.isFinite(value)) return "";
  const pattern = getNumberFormatPattern();
  if (pattern) return formatDatasetNumberValue(value, pattern, getDecimalPlaces());
  if (isPercentTriangle()) {
    return `${(value * 100).toFixed(getPercentDecimalPlaces())}%`;
  }
  return fmt0.format(value);
}

function getEffectiveDevLabels(model) {
  const devs = Array.isArray(model?.dev_labels) ? model.dev_labels : [];
  const vals = Array.isArray(model?.values) ? model.values : [];
  let maxCols = 0;
  for (const row of vals) {
    if (Array.isArray(row)) maxCols = Math.max(maxCols, row.length);
  }
  if (!maxCols) return devs;
  if (devs.length >= maxCols) return devs.slice(0, maxCols);
  return devs.concat(Array(maxCols - devs.length).fill(""));
}

function isTransposedView() {
  return document.getElementById("transposedChk")?.checked === true;
}

function transposeMatrix(matrix) {
  const rows = Array.isArray(matrix) ? matrix : [];
  let maxCols = 0;
  for (const row of rows) {
    if (Array.isArray(row)) maxCols = Math.max(maxCols, row.length);
  }
  const out = [];
  for (let c = 0; c < maxCols; c++) {
    const next = [];
    for (let r = 0; r < rows.length; r++) {
      next.push(rows[r]?.[c]);
    }
    out.push(next);
  }
  return out;
}

export function getDisplayDatasetModel() {
  const model = state.model;
  if (!model || !isTransposedView()) return model;

  return {
    ...model,
    origin_labels: getEffectiveDevLabels(model),
    dev_labels: Array.isArray(model.origin_labels) ? model.origin_labels.map(String) : [],
    values: transposeMatrix(model.values),
    mask: transposeMatrix(model.mask).map((row) => row.map(Boolean)),
  };
}

export function getDatasetGridSelectionLayout(model = getDisplayDatasetModel()) {
  const transposed = isTransposedView();
  const showTotals = shouldShowDatasetGridTotals({
    isDfmHost: isDfmDataTabHost(),
    formula: getCurrentDatasetTypeFormula(),
    showSubtotal: state.showSubtotal !== false,
  });
  return {
    transposed,
    showTotals,
    ...getDatasetGridTotalLayout({
      rowCount: model?.origin_labels?.length || 0,
      columnCount: getEffectiveDevLabels(model).length,
      showTotals,
      transposed,
    }),
  };
}

function captureRenderedColumnWidths(wrap) {
  const headers = Array.from(wrap?.querySelectorAll?.("thead tr:first-child > th") || []);
  if (!headers.length) return [];
  return headers.map((cell) => {
    const width = cell.getBoundingClientRect?.().width;
    return Number.isFinite(width) && width > 0 ? Math.ceil(width) : 0;
  });
}

function applyColumnWidthLock(table, widths, expectedCount) {
  if (!table || !Array.isArray(widths) || widths.length !== expectedCount) return;
  if (widths.some((width) => !Number.isFinite(width) || width <= 0)) return;
  const colgroup = document.createElement("colgroup");
  let totalWidth = 0;
  for (const width of widths) {
    totalWidth += width;
    const col = document.createElement("col");
    col.style.width = `${width}px`;
    colgroup.appendChild(col);
  }
  table.appendChild(colgroup);
  table.style.width = `${totalWidth}px`;
}

function fitRowLabelColumn(table, cornerCell) {
  if (!table || !cornerCell) return;

  const cornerStyle = getComputedStyle(cornerCell);
  const labelMeasure = document.createElement("canvas").getContext("2d");
  if (!labelMeasure) return;

  labelMeasure.font = `${cornerStyle.fontWeight} ${cornerStyle.fontSize} ${cornerStyle.fontFamily}`;
  const horizontalChrome = [
    cornerStyle.paddingLeft,
    cornerStyle.paddingRight,
    cornerStyle.borderLeftWidth,
    cornerStyle.borderRightWidth,
  ].reduce((total, value) => total + (Number.parseFloat(value) || 0), 0);
  const sharedWidth = Number.parseFloat(
    cornerStyle.getPropertyValue("--ar-spreadsheet-cell-width"),
  ) || Math.ceil(cornerCell.getBoundingClientRect().width);
  const labelWidth = Math.max(
    sharedWidth,
    Math.ceil(labelMeasure.measureText(cornerCell.textContent || "").width + horizontalChrome + 1),
  );

  table.style.setProperty("--data-tab-row-label-column-width", `${labelWidth}px`);

  const lockedFirstColumn = table.querySelector("colgroup col:first-child");
  if (!lockedFirstColumn) return;
  const previousWidth = Number.parseFloat(lockedFirstColumn.style.width) || sharedWidth;
  lockedFirstColumn.style.width = `${labelWidth}px`;
  const lockedTableWidth = Number.parseFloat(table.style.width);
  if (Number.isFinite(lockedTableWidth)) {
    table.style.width = `${lockedTableWidth + labelWidth - previousWidth}px`;
  }
}

export function renderTable() {

  const wrap = $("tableWrap");
  const lockedColumnWidths = state.editingCell ? captureRenderedColumnWidths(wrap) : [];
  wrap.innerHTML = "";
  ensureCtxMenuWired();

  const model = getDisplayDatasetModel();
  if (!model) {
    // The grid has nothing to paint yet. Which of "still arriving", "nothing
    // selected", and "load failed" that means is owned by the placeholder.
    renderDatasetGridPlaceholder(wrap);
    return;
  }

  const origins = model.origin_labels;
  const devs = getEffectiveDevLabels(model);
  const vals = model.values;
  const mask = model.mask; // True=has value, False=blank/missing
  const layout = getDatasetGridSelectionLayout(model);
  const {
    maxRow,
    maxCol,
    showTotals: showTotalRow,
    totalColumnIndex,
    totalRowIndex,
    transposed,
  } = layout;
  const showRightSideTotal = totalColumnIndex !== null;

  if (!Array.isArray(mask)) {
    wrap.innerHTML = `<div style="color:#b00;"><b>UI Error:</b> mask is missing. Update get_dataset to return mask.</div>`;
    return;
  }

  if (state.activeCell) {
    if (maxRow < 0 || maxCol < 0) {
      state.activeCell = null;
    } else {
      const r = Math.max(0, Math.min(state.activeCell.r, maxRow));
      const c = Math.max(0, Math.min(state.activeCell.c, maxCol));
      state.activeCell = { r, c };
    }
  }

  const tbl = document.createElement("table");
  tbl.classList.add("arSpreadsheetTable");
  applyColumnWidthLock(tbl, lockedColumnWidths, devs.length + 1 + (showRightSideTotal ? 1 : 0));

  // header
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");

  const th0 = document.createElement("th");
  const originLen = document.getElementById("originLenSelect")?.value || 12;
  const calendar = document.querySelector('input[name="timeMode"][value="calendar"]')?.checked === true;
  th0.textContent = transposed ? (calendar ? "Calendar Period" : "Development Period") : getDatasetOriginLabelText(originLen);
  trh.appendChild(th0);

  devs.forEach((d, c) => {
    const th = document.createElement("th");
    th.textContent = !transposed && devs.length === 1 && renderVectorColumnLabel
      ? renderVectorColumnLabel
      : (transposed ? formatDatasetOriginLabel(d, originLen) : d);

    th.classList.add("colhdr");
    th.dataset.c = String(c);

    trh.appendChild(th);
  });

  if (showRightSideTotal) {
    const th = document.createElement("th");
    th.textContent = "Total";
    th.classList.add("totalColHdr", "colhdr");
    th.dataset.c = String(totalColumnIndex);
    trh.appendChild(th);
  }

  thead.appendChild(trh);
  tbl.appendChild(thead);

  // body
  const tbody = document.createElement("tbody");

  for (let r = 0; r < origins.length; r++) {
    const tr = document.createElement("tr");

    const th = document.createElement("th");
    th.textContent = formatDatasetOriginLabel(origins[r], originLen);

    th.classList.add("rowhdr");
    th.dataset.r = String(r);

    tr.appendChild(th);

    for (let c = 0; c < devs.length; c++) {
      const td = document.createElement("td");
      const hasValue = !!(mask[r] && mask[r][c]);
      configureSelectableDatasetCell(td, r, c);

      if (!hasValue) {
        td.textContent = "";
        if (!state.showBlanks) {
          td.classList.add("na");        // visually hidden
        }
      } else {
        const v = vals[r][c];
        const isEditable = gridEditConfig?.isEditableCell?.(r, c) === true;
        if (isEditable && gridEditConfig?.isEditingCell?.(r, c)) {
          const input = document.createElement("input");
          input.className = "dsCellInput";
          input.type = "text";
          input.inputMode = "decimal";
          input.value = formatCellValue(v);
          input.classList.toggle("dsCellInputBlank", v == null);
          input.dataset.r = String(r);
          input.dataset.c = String(c);
          input.addEventListener("focus", () => gridEditConfig?.onCellFocus?.(r, c));
          input.addEventListener("mousedown", (event) => event.stopPropagation());
          input.addEventListener("keydown", (event) => gridEditConfig?.onCellKeyDown?.(r, c, event));
          input.addEventListener("input", () => gridEditConfig?.onCellInput?.(r, c, input.value, input, td));
          input.addEventListener("paste", (event) => gridEditConfig?.onCellPaste?.(r, c, event));
          input.addEventListener("change", () => gridEditConfig?.onCellCommit?.(r, c, input.value, input, td));
          input.addEventListener("blur", () => gridEditConfig?.onCellCommit?.(r, c, input.value, input, td));
          td.appendChild(input);
        } else {
          const displayNullAsZero = isEditable && v == null;
          td.textContent = formatCellValue(displayNullAsZero ? 0 : v);
          td.classList.toggle("dsNullValue", displayNullAsZero);
        }
      }

      gridEditConfig?.decorateCell?.(td, r, c);

      tr.appendChild(td);
    }

    if (showRightSideTotal) {
      const td = document.createElement("td");
      td.classList.add("totalCell");
      const sum = sumDatasetGridRow(vals, mask, r, devs.length);
      td.textContent = sum == null ? "" : formatCellValue(sum);
      configureSelectableDatasetCell(td, r, totalColumnIndex, {
        copyValue: sum,
        readOnly: true,
      });
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  tbl.appendChild(tbody);

  tbl.classList.toggle("has-total-row", showTotalRow && !showRightSideTotal);
  tbl.classList.toggle("has-total-column", showRightSideTotal);

  if (showTotalRow && !showRightSideTotal) {
    // Footer totals: sum each development column across all origin rows.
    const tfoot = document.createElement("tfoot");
    const trf = document.createElement("tr");
    const totalLabel = document.createElement("th");
    totalLabel.textContent = "Total";
    totalLabel.classList.add("rowhdr");
    totalLabel.dataset.r = String(totalRowIndex);
    trf.appendChild(totalLabel);

    for (let c = 0; c < devs.length; c++) {
      const td = document.createElement("td");
      const sum = sumDatasetGridColumn(vals, mask, c, origins.length);
      td.textContent = sum == null ? "" : formatCellValue(sum);
      td.classList.add("totalCell");
      configureSelectableDatasetCell(td, totalRowIndex, c, {
        copyValue: sum,
        readOnly: true,
      });
      trf.appendChild(td);
    }
    tfoot.appendChild(trf);
    tbl.appendChild(tfoot);
  }

  wrap.appendChild(tbl);
  fitRowLabelColumn(tbl, th0);

  if (gridEditConfig?.onTableRendered) gridEditConfig.onTableRendered();
  renderDataTabChart();
}
