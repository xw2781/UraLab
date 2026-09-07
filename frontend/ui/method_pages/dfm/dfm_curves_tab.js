/*
===============================================================================
DFM Curves Tab
The ArcRho counterpart of ResQ's Curves | Data tab: the Initial Selection from
the Ratios tab, four fitted curves, user value columns, the Include flags, the
Selected Estimate Number per development period and for the tail, and the
chain the ultimates use. The numbers come from dfm_curve_fit.js, the mirror of
arcrho_api/dfm_curves.py; this module owns only the grid and its clicks.
===============================================================================
*/
import {
  state,
  formatRatio,
  roundHalfUp,
  getCurvesTab,
  setCurvesTab,
  getCurvesTable,
  getNormalizedCurvesTab,
  getSelectedRatioValues,
  getEffectiveDevLabelsForModel,
  getRatioHeaderLabels,
  getDfmDecimalPlaces,
  isCurvesTabVisible,
  isResultsTabVisible,
  markDfmDirty,
  isDfmApplyingProgrammatically,
} from "/ui/method_pages/dfm/dfm_state.js";
import {
  FIXED_COLUMN_COUNT,
  FIT_OK,
  FIT_LIMIT,
  FIT_FAIL,
  FIT_WARNING,
  FIT_UNFITTED,
  DEFAULT_USER_COLUMN_LABEL,
  MAX_FUTURE_DEVELOPMENT_PERIODS,
} from "/ui/method_pages/dfm/dfm_curve_fit.js?v=20260903a";
import {
  renderResultsTable,
  invalidatePersistedResultsDerivations,
} from "/ui/method_pages/dfm/dfm_results_tab.js?v=20260907a";
import { openContextMenu } from "/ui/shared/components/context_menu/context_menu.js?v=20260811b";

const FIT_LABELS = Object.freeze({
  [FIT_OK]: "OK",
  [FIT_LIMIT]: "Limit",
  [FIT_FAIL]: "Fail",
  [FIT_WARNING]: "Warning",
  [FIT_UNFITTED]: "",
});
const LEAST_SQUARES_NOTICE = "This method was fitted by least squares in ResQ. ArcRho fits by log regression, so the curves shown are its log-regression fits.";

let wired = false;
let activeEditor = null;

function periodContext() {
  const model = state.model;
  if (!model || !Array.isArray(model.values) || !Array.isArray(model.mask)) return null;
  const devs = getEffectiveDevLabelsForModel(model);
  const ratioLabels = getRatioHeaderLabels(devs);
  if (!ratioLabels.length) return null;
  return { model, devs, ratioLabels, periodCount: Math.max(0, ratioLabels.length - 1) };
}

// Every mutation goes through here: the tab is rewritten as a whole, the
// cached table dropped, the Results tab told its persisted ultimates are
// stale, and the page marked dirty so the preview and Save see the change.
function commitCurvesTab(nextTab) {
  setCurvesTab(nextTab);
  invalidatePersistedResultsDerivations();
  renderDfmCurvesTab();
  if (isResultsTabVisible()) renderResultsTable();
  if (isDfmApplyingProgrammatically()) return;
  markDfmDirty();
  window.dispatchEvent(new CustomEvent("arcrho:dfm-owned-state-mutated"));
}

function currentTab() {
  const context = periodContext();
  if (!context) return null;
  return getNormalizedCurvesTab(context.model, context.devs);
}

function updateTab(mutate) {
  const tab = currentTab();
  if (!tab) return;
  mutate(tab);
  commitCurvesTab(tab);
}

// ---------------------------------------------------------------------------
// Labels and formatting
// ---------------------------------------------------------------------------

function periodRowLabel(ratioLabels, index) {
  const text = String(ratioLabels[index] ?? "");
  return text ? `(${index + 1}) ${text}` : `(${index + 1})`;
}

function tailRowLabel(ratioLabels) {
  return String(ratioLabels[ratioLabels.length - 1] ?? "Ult") || "Ult";
}

// Future period labels continue the observed ages by the last observed step:
// "(9) 101-113" is followed by "(10) 113-125".
function futureRowLabel(devs, ratioLabels, period) {
  const numbers = devs
    .map((label) => Number.parseInt(String(label ?? "").replace(/[^\d-]/g, ""), 10))
    .filter(Number.isFinite);
  if (numbers.length >= 2) {
    const step = numbers[numbers.length - 1] - numbers[numbers.length - 2];
    const start = numbers[numbers.length - 1] + step * (period - devs.length);
    if (step > 0) return `(${period}) ${start}-${start + step}`;
  }
  return `(${period})`;
}

function formatFactor(value) {
  return Number.isFinite(value) ? formatRatio(value, getDfmDecimalPlaces()) : "";
}

function formatPercent(value) {
  if (!Number.isFinite(value)) return "";
  const rounded = roundHalfUp(value * 100, 2);
  return `${(rounded ?? value * 100).toFixed(2)}%`;
}

function formatParameter(value) {
  if (!Number.isFinite(value)) return "";
  const rounded = roundHalfUp(value, 4);
  return (rounded ?? value).toFixed(4);
}

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------

function cell(text, classes = [], attrs = {}) {
  const td = document.createElement("td");
  td.textContent = text;
  classes.filter(Boolean).forEach((name) => td.classList.add(name));
  Object.entries(attrs).forEach(([key, value]) => {
    if (value !== null && value !== undefined) td.dataset[key] = String(value);
  });
  return td;
}

function rowHeader(text, classes = []) {
  const th = document.createElement("th");
  th.textContent = text;
  classes.filter(Boolean).forEach((name) => th.classList.add(name));
  return th;
}

// The spacer row's label cell is a row header like any other, so the frozen
// label column stays whole when the grid is scrolled sideways.
function blankRow(columnCount) {
  const tr = document.createElement("tr");
  tr.classList.add("dfmCurvesBlankRow");
  tr.appendChild(rowHeader("", ["dfmCurvesRowHeader"]));
  for (let i = 0; i < columnCount; i++) tr.appendChild(cell("", ["dfmCurvesBlank"]));
  return tr;
}

function syncControls(tab) {
  const fitting = document.getElementById("dfmCurvesFittingMethod");
  const notice = document.getElementById("dfmCurvesNotice");
  if (fitting) {
    const leastSquares = tab.fitting_method === "least_squares";
    let option = fitting.querySelector('option[value="least_squares"]');
    if (leastSquares && !option) {
      option = document.createElement("option");
      option.value = "least_squares";
      option.textContent = "Least Squares";
      option.disabled = true;
      fitting.appendChild(option);
    } else if (!leastSquares && option) {
      option.remove();
    }
    fitting.value = tab.fitting_method;
  }
  if (notice) notice.textContent = tab.fitting_method === "least_squares" ? LEAST_SQUARES_NOTICE : "";
  const future = document.getElementById("dfmCurvesFuturePeriodsInput");
  if (future && document.activeElement !== future) future.value = String(tab.future_development_periods);
  const freeFit = document.getElementById("dfmCurvesFreeFitC");
  if (freeFit) freeFit.checked = !!tab.free_fit_c;
}

export function renderDfmCurvesTab() {
  const wrap = document.getElementById("dfmCurvesWrap");
  if (!wrap) return;
  const context = periodContext();
  if (!context) {
    wrap.replaceChildren();
    return;
  }
  const { model, devs, ratioLabels, periodCount } = context;
  const tab = getNormalizedCurvesTab(model, devs);
  syncControls(tab);
  const table = getCurvesTable(model, devs);
  const columns = table.columns;
  const derivedCount = 5; // Selected Estimate Number, Selected Value, Cumulative Value, Cumulative %, Incremental %
  const totalColumns = 1 + columns.length + derivedCount; // + Include after the Initial Selection

  const grid = document.createElement("table");
  grid.classList.add("arSpreadsheetTable", "dfmCurvesTable");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  headRow.appendChild(rowHeader("Development Year", ["dfmCurvesCorner"]));
  columns.forEach((column, index) => {
    const th = document.createElement("th");
    th.textContent = `${column.label} (${column.number})`;
    th.title = "Select this column for every development period and the tail";
    th.dataset.column = String(column.number);
    th.classList.add("dfmCurvesColumnHeader");
    if (column.column_type === "user_entry") th.classList.add("dfmCurvesUserHeader");
    headRow.appendChild(th);
    if (index === 0) headRow.appendChild(rowHeader("Include", ["dfmCurvesIncludeHeader"]));
  });
  ["Selected Estimate Number", "Selected Value", "Cumulative Value", "Cumulative Percentage", "Incremental Percentage"]
    .forEach((label) => headRow.appendChild(rowHeader(label, ["dfmCurvesDerivedHeader"])));
  thead.appendChild(headRow);
  grid.appendChild(thead);

  const tbody = document.createElement("tbody");
  const valueCell = (column, value, period, selected) => {
    const td = cell(formatFactor(value), ["dfmCurvesValue"], { column: column.number, period });
    if (selected) {
      td.classList.add("dfmCurvesSelected");
      td.setAttribute("aria-selected", "true");
    }
    if (column.column_type === "user_entry") td.classList.add("dfmCurvesUserValue");
    if (column.column_type === "prior_analysis" || column.column_type === "pattern" || column.column_type === "benchmark") {
      td.classList.add("dfmCurvesLinkedValue");
    }
    return td;
  };
  const derivedCells = (tr, selectedNumber, selectedValue, cumulative, cumulativePct, incrementalPct, period) => {
    tr.appendChild(cell(selectedNumber === null ? "" : String(selectedNumber), ["dfmCurvesEstimateNumber"], { period }));
    tr.appendChild(cell(formatFactor(selectedValue), ["dfmCurvesSelectedValue"]));
    tr.appendChild(cell(formatFactor(cumulative), ["dfmCurvesDerived"]));
    tr.appendChild(cell(formatPercent(cumulativePct), ["dfmCurvesDerived"]));
    tr.appendChild(cell(formatPercent(incrementalPct), ["dfmCurvesDerived"]));
  };

  for (let index = 0; index < periodCount; index++) {
    const tr = document.createElement("tr");
    tr.dataset.period = String(index + 1);
    tr.appendChild(rowHeader(periodRowLabel(ratioLabels, index), ["dfmCurvesRowHeader"]));
    // Leaving a period out strikes both its Include flag and the Initial
    // Selection factor the flag drops, the way ResQ shows the pair.
    const excluded = !tab.included[index];
    columns.forEach((column, columnIndex) => {
      const td = valueCell(column, column.values[index], index + 1, tab.selected_estimates[index] === column.number);
      if (columnIndex === 0 && excluded) td.classList.add("dfmCurvesExcluded");
      tr.appendChild(td);
      if (columnIndex === 0) {
        const flag = cell(excluded ? "No" : "Yes", ["dfmCurvesInclude", excluded ? "dfmCurvesExcluded" : ""], { period: index + 1 });
        flag.title = "Toggle whether this period takes part in the curve fits";
        tr.appendChild(flag);
      }
    });
    derivedCells(
      tr,
      tab.selected_estimates[index],
      table.selected_values[index],
      table.cumulative[index],
      table.cumulative_percentage[index],
      table.incremental_percentage[index],
      index + 1,
    );
    tbody.appendChild(tr);
  }
  tbody.appendChild(blankRow(totalColumns));

  // The tail row: each column's tail factor and the selected tail.
  const tailRow = document.createElement("tr");
  tailRow.dataset.period = "tail";
  tailRow.appendChild(rowHeader(tailRowLabel(ratioLabels), ["dfmCurvesRowHeader"]));
  columns.forEach((column, columnIndex) => {
    tailRow.appendChild(valueCell(column, column.tail, "tail", table.selected_tail_column === column.number));
    if (columnIndex === 0) tailRow.appendChild(cell("", ["dfmCurvesBlank"]));
  });
  derivedCells(
    tailRow,
    table.selected_tail_column,
    table.selected_tail,
    table.cumulative[periodCount],
    table.cumulative_percentage[periodCount],
    table.incremental_percentage[periodCount],
    "tail",
  );
  tbody.appendChild(tailRow);
  tbody.appendChild(blankRow(totalColumns));

  // Fit statistics under the curve columns.
  const statRow = (label, pick) => {
    const tr = document.createElement("tr");
    tr.classList.add("dfmCurvesStatRow");
    tr.appendChild(rowHeader(label, ["dfmCurvesRowHeader"]));
    columns.forEach((column, columnIndex) => {
      tr.appendChild(cell(column.fit ? pick(column) : "", ["dfmCurvesStat"]));
      if (columnIndex === 0) tr.appendChild(cell("", ["dfmCurvesBlank"]));
    });
    for (let i = 0; i < derivedCount; i++) tr.appendChild(cell("", ["dfmCurvesBlank"]));
    return tr;
  };
  const fitRow = statRow("Fit", (column) => FIT_LABELS[column.fit.result] ?? "");
  fitRow.querySelectorAll("td.dfmCurvesStat").forEach((td) => {
    if (td.textContent === "Limit" || td.textContent === "Warning") td.classList.add("dfmCurvesFitLimit");
    if (td.textContent === "Fail") td.classList.add("dfmCurvesFitFail");
  });
  tbody.appendChild(fitRow);
  tbody.appendChild(blankRow(totalColumns));
  tbody.appendChild(statRow("A", (column) => formatParameter(column.fit.a)));
  tbody.appendChild(statRow("B", (column) => formatParameter(column.fit.b)));
  tbody.appendChild(statRow("C", (column) => (column.key === "inverse_power" ? formatParameter(column.fit.c) : "")));
  tbody.appendChild(statRow("R-squared %", (column) => formatPercent(column.fit.r_squared)));
  tbody.appendChild(blankRow(totalColumns));

  // The tail pattern: the X marks the column whose run-off the tail follows,
  // and that same column stays green down the run-off rows beneath it, so the
  // choice and the periods it produces read as one block. The number comes from
  // the table rather than the tab, so the mark always names the column the
  // figures were actually built from.
  const patternColumnNumber = table.selected_tail_pattern_column;
  const patternColumn = columns.find((column) => column.number === patternColumnNumber);
  const patternTitle = patternColumn
    ? `The tail runs off along ${patternColumn.label}. Click another column to run it off along that one.`
    : "Click a column to run the tail off along it";
  const patternHeader = document.createElement("tr");
  patternHeader.classList.add("dfmCurvesPatternRow");
  patternHeader.dataset.period = "pattern";
  const patternLabel = rowHeader("Tail Pattern", ["dfmCurvesRowHeader"]);
  patternLabel.title = patternTitle;
  patternHeader.appendChild(patternLabel);
  columns.forEach((column, columnIndex) => {
    const selected = patternColumnNumber === column.number;
    const td = cell(selected ? "X" : "", ["dfmCurvesPattern", selected ? "dfmCurvesSelected" : ""], {
      column: column.number,
      period: "pattern",
    });
    td.title = patternTitle;
    if (selected) td.setAttribute("aria-selected", "true");
    patternHeader.appendChild(td);
    if (columnIndex === 0) patternHeader.appendChild(cell("", ["dfmCurvesBlank"]));
  });
  for (let i = 0; i < derivedCount; i++) patternHeader.appendChild(cell("", ["dfmCurvesBlank"]));
  tbody.appendChild(patternHeader);
  table.tail_rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.classList.add("dfmCurvesFutureRow");
    tr.appendChild(rowHeader(futureRowLabel(devs, ratioLabels, row.period), ["dfmCurvesRowHeader"]));
    columns.forEach((column, columnIndex) => {
      const value = row.values instanceof Map ? row.values.get(column.number) : row.values?.[column.number];
      const selected = patternColumnNumber === column.number;
      const td = cell(
        formatFactor(value),
        [
          "dfmCurvesFuture",
          column.column_type === "user_entry" ? "dfmCurvesUserValue" : "",
          selected ? "dfmCurvesSelected" : "",
        ],
        { column: column.number, period: "pattern" },
      );
      td.title = patternTitle;
      if (selected) td.setAttribute("aria-selected", "true");
      tr.appendChild(td);
      if (columnIndex === 0) tr.appendChild(cell("", ["dfmCurvesBlank"]));
    });
    tr.appendChild(cell("", ["dfmCurvesBlank"]));
    tr.appendChild(cell(formatFactor(row.selected_value), ["dfmCurvesSelectedValue"]));
    tr.appendChild(cell(formatFactor(row.cumulative_value), ["dfmCurvesDerived"]));
    tr.appendChild(cell(formatPercent(row.cumulative_percentage), ["dfmCurvesDerived"]));
    tr.appendChild(cell(formatPercent(row.incremental_percentage), ["dfmCurvesDerived"]));
    tbody.appendChild(tr);
  });
  grid.appendChild(tbody);
  wrap.replaceChildren(grid);
}

// ---------------------------------------------------------------------------
// Interactions
// ---------------------------------------------------------------------------

function userColumnIndex(columnNumber) {
  return Number(columnNumber) - FIXED_COLUMN_COUNT - 1;
}

function selectEstimate(period, columnNumber) {
  updateTab((tab) => {
    if (period === "tail") {
      tab.selected_tail_factor = columnNumber;
      return;
    }
    if (period === "pattern") {
      tab.selected_tail_curve = columnNumber;
      return;
    }
    tab.selected_estimates[period - 1] = columnNumber;
  });
}

function selectWholeColumn(columnNumber) {
  updateTab((tab) => {
    tab.selected_estimates = tab.selected_estimates.map(() => columnNumber);
    tab.selected_tail_factor = columnNumber;
    tab.selected_tail_curve = columnNumber;
  });
}

function toggleInclude(period) {
  updateTab((tab) => {
    tab.included[period - 1] = tab.included[period - 1] ? 0 : 1;
  });
}

function setUserValue(columnNumber, period, value) {
  const index = userColumnIndex(columnNumber);
  updateTab((tab) => {
    const column = tab.user_columns[index];
    if (!column || column.column_type !== "user_entry") return;
    if (period === "tail") column.tail = value;
    else column.values[period - 1] = value;
  });
}

function closeEditor(commit) {
  const editor = activeEditor;
  if (!editor) return;
  activeEditor = null;
  const { input, td, columnNumber, period } = editor;
  const raw = String(input.value ?? "").trim();
  input.remove();
  td.classList.remove("dfmCurvesEditing");
  if (!commit) {
    renderDfmCurvesTab();
    return;
  }
  const value = Number(raw);
  if (!raw || !Number.isFinite(value) || value <= 0) {
    renderDfmCurvesTab();
    return;
  }
  setUserValue(columnNumber, period, value);
}

function openEditor(td) {
  if (activeEditor) closeEditor(true);
  const columnNumber = Number(td.dataset.column);
  const period = td.dataset.period === "tail" ? "tail" : Number(td.dataset.period);
  const input = document.createElement("input");
  input.type = "text";
  input.className = "dfmCurvesEditor";
  input.value = td.textContent || "";
  input.setAttribute("aria-label", "User value");
  td.classList.add("dfmCurvesEditing");
  td.textContent = "";
  td.appendChild(input);
  activeEditor = { input, td, columnNumber, period };
  input.focus();
  input.select();
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      closeEditor(true);
    } else if (event.key === "Escape") {
      event.preventDefault();
      closeEditor(false);
    }
  });
  input.addEventListener("blur", () => closeEditor(true));
}

function addUserColumn() {
  updateTab((tab) => {
    tab.user_columns.push({
      label: DEFAULT_USER_COLUMN_LABEL,
      column_type: "user_entry",
      values: new Array(tab.included.length).fill(1),
      tail: 1,
    });
  });
}

function removeUserColumn(columnNumber) {
  const index = userColumnIndex(columnNumber);
  updateTab((tab) => {
    if (index < 0 || index >= tab.user_columns.length) return;
    tab.user_columns.splice(index, 1);
    const limit = FIXED_COLUMN_COUNT + tab.user_columns.length;
    const clamp = (number) => (number === columnNumber || number > limit ? 1 : number > columnNumber ? number - 1 : number);
    tab.selected_estimates = tab.selected_estimates.map(clamp);
    tab.selected_tail_factor = clamp(tab.selected_tail_factor);
    tab.selected_tail_curve = clamp(tab.selected_tail_curve);
  });
}

function renameUserColumn(columnNumber) {
  const tab = currentTab();
  const index = userColumnIndex(columnNumber);
  const column = tab?.user_columns?.[index];
  if (!column) return;
  const next = window.prompt("Column name", column.label);
  if (next === null) return;
  updateTab((current) => {
    current.user_columns[index].label = String(next).trim() || DEFAULT_USER_COLUMN_LABEL;
  });
}

function copyText(text) {
  try {
    navigator.clipboard?.writeText?.(String(text ?? ""));
  } catch {
    // The clipboard is a convenience; a refusal leaves the grid as it was.
  }
}

let contextMenu = null;

function closeCurvesContextMenu() {
  if (!contextMenu) return;
  contextMenu.remove();
  contextMenu = null;
}

function openCurvesContextMenu(event, td) {
  closeCurvesContextMenu();
  const columnNumber = Number(td?.dataset?.column);
  const tab = currentTab();
  const userColumn = Number.isFinite(columnNumber) && userColumnIndex(columnNumber) >= 0
    ? tab?.user_columns?.[userColumnIndex(columnNumber)]
    : null;
  const items = [];
  if (td?.textContent) items.push({ label: "Copy Value", onSelect: () => copyText(td.textContent) });
  items.push({ label: "Add User Column", onSelect: addUserColumn });
  if (userColumn) {
    items.push({ label: "Rename User Column", onSelect: () => renameUserColumn(columnNumber) });
    items.push({ label: "Remove User Column", onSelect: () => removeUserColumn(columnNumber) });
  }
  // The same menu surface the Ratios tab uses, built for this click and
  // dropped once it closes.
  const menu = document.createElement("div");
  menu.className = "dfmCtxMenu";
  menu.setAttribute("role", "menu");
  items.forEach((item) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "dfmCtxItem";
    button.setAttribute("role", "menuitem");
    button.textContent = item.label;
    button.addEventListener("click", () => {
      closeCurvesContextMenu();
      item.onSelect();
    });
    menu.appendChild(button);
  });
  document.body.appendChild(menu);
  contextMenu = menu;
  openContextMenu(menu, { clientX: event.clientX, clientY: event.clientY });
}

function wireCurvesContextMenuDismissal() {
  document.addEventListener("mousedown", (event) => {
    if (contextMenu && !contextMenu.contains(event.target)) closeCurvesContextMenu();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeCurvesContextMenu();
  });
}

function wireCurvesGrid() {
  const wrap = document.getElementById("dfmCurvesWrap");
  if (!wrap || wrap.dataset.curvesWired === "1") return;
  wrap.dataset.curvesWired = "1";

  wrap.addEventListener("click", (event) => {
    if (activeEditor) return;
    const header = event.target?.closest?.("th.dfmCurvesColumnHeader");
    if (header) {
      selectWholeColumn(Number(header.dataset.column));
      return;
    }
    const include = event.target?.closest?.("td.dfmCurvesInclude");
    if (include) {
      toggleInclude(Number(include.dataset.period));
      return;
    }
    // A run-off cell answers like the Tail Pattern cell above it, so the whole
    // block is one target for choosing the column the tail follows.
    const value = event.target?.closest?.("td.dfmCurvesValue, td.dfmCurvesPattern, td.dfmCurvesFuture");
    if (!value) return;
    const period = value.dataset.period === "tail" || value.dataset.period === "pattern"
      ? value.dataset.period
      : Number(value.dataset.period);
    selectEstimate(period, Number(value.dataset.column));
  });

  wrap.addEventListener("dblclick", (event) => {
    const td = event.target?.closest?.("td.dfmCurvesUserValue.dfmCurvesValue");
    if (!td) return;
    event.preventDefault();
    openEditor(td);
  });

  wrap.addEventListener("contextmenu", (event) => {
    event.preventDefault();
    openCurvesContextMenu(event, event.target?.closest?.("td, th"));
  });
}

function wireCurvesControls() {
  const future = document.getElementById("dfmCurvesFuturePeriodsInput");
  future?.addEventListener("change", () => {
    const value = Number.parseInt(future.value, 10);
    const next = Number.isFinite(value) ? Math.max(1, Math.min(MAX_FUTURE_DEVELOPMENT_PERIODS, value)) : 1;
    future.value = String(next);
    updateTab((tab) => {
      tab.future_development_periods = next;
    });
  });
  const freeFit = document.getElementById("dfmCurvesFreeFitC");
  freeFit?.addEventListener("change", () => {
    updateTab((tab) => {
      tab.free_fit_c = !!freeFit.checked;
    });
  });
  const fitting = document.getElementById("dfmCurvesFittingMethod");
  fitting?.addEventListener("change", () => {
    updateTab((tab) => {
      tab.fitting_method = fitting.value === "least_squares" ? "least_squares" : "log_regression";
    });
  });
}

export function initDfmCurvesTab() {
  if (wired) return;
  wired = true;
  wireCurvesGrid();
  wireCurvesControls();
  wireCurvesContextMenuDismissal();
  // A Ratios-tab change moves the Initial Selection, so the fits follow it.
  window.addEventListener("arcrho:dfm-owned-state-mutated", () => {
    if (isCurvesTabVisible()) renderDfmCurvesTab();
  });
}

// The persisted `curves_tab` of a loaded method, or null to start from the
// default tab; the chain then equals the Ratios tab's selection.
export function applyDfmCurvesTabPayload(curvesTab) {
  setCurvesTab(curvesTab && typeof curvesTab === "object" ? curvesTab : null);
  invalidatePersistedResultsDerivations();
  if (isCurvesTabVisible()) renderDfmCurvesTab();
}

export function buildDfmCurvesTabPayload() {
  const context = periodContext();
  if (!context) return getCurvesTab() || {};
  const tab = getNormalizedCurvesTab(context.model, context.devs);
  const table = getCurvesTable(context.model, context.devs);
  return {
    ...tab,
    selected_values: [...table.selected_values, table.selected_tail].map((value) => (
      Number.isFinite(value) ? Math.round(value * 1e6) / 1e6 : null
    )),
  };
}

export function getDfmInitialSelectionForCurves() {
  const context = periodContext();
  return context ? getSelectedRatioValues(context.model, context.devs) : [];
}
