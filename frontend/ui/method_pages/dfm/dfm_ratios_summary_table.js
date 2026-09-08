/*
===============================================================================
DFM Ratios Summary Table
Compatibility facade and render scheduler for the modular summary table.
===============================================================================
*/
import {
  registerSummaryFunctions,
  summaryRuntime,
} from "/ui/method_pages/dfm/ratios_summary/summary_runtime.js?v=20260819a";
import "/ui/method_pages/dfm/ratios_summary/summary_model.js?v=20260902a";
import "/ui/method_pages/dfm/ratios_summary/summary_formula_bar.js?v=20260831a";
import "/ui/method_pages/dfm/ratios_summary/summary_formula_bar_anchor.js?v=20260819a";
import "/ui/method_pages/dfm/ratios_summary/summary_formula_bar_drag.js?v=20260819a";
import "/ui/method_pages/dfm/ratios_summary/summary_excel.js?v=20260830b";
import "/ui/method_pages/dfm/ratios_summary/summary_entries.js?v=20260907a";
import "/ui/method_pages/dfm/ratios_summary/summary_interactions.js?v=20260831a";

export const DFM_RATIO_HIGHLIGHT_EDGE_CLASSES = Object.freeze({
  top: "dfmTableHighlightEdgeTop",
  right: "dfmTableHighlightEdgeRight",
  bottom: "dfmTableHighlightEdgeBottom",
  left: "dfmTableHighlightEdgeLeft",
});

function isRatioEditMode() {
  return document.getElementById("ratioWrap")?.dataset?.interactionMode === "edit";
}

export function refreshRatioHighlightHeaders() {
  const wrap = document.getElementById("ratioWrap");
  if (!wrap) return;
  wrap.querySelectorAll("th.arSpreadsheetSelectedLabel").forEach((header) => {
    header.classList.remove("arSpreadsheetSelectedLabel");
  });
  if (wrap.dataset.interactionMode !== "select") return;
  wrap.querySelectorAll("td.dfmTableHighlight").forEach((cell) => {
    const rowHeader = cell.parentElement?.querySelector?.("th");
    if (rowHeader) rowHeader.classList.add("arSpreadsheetSelectedLabel");
    const copyCol = Number(cell.dataset.copyC ?? cell.dataset.col ?? cell.dataset.c);
    if (!Number.isInteger(copyCol) || copyCol < 0) return;
    const columnHeader = wrap.querySelector(
      `table.ratioMainTable thead th[data-copy-col="${copyCol}"]`
    );
    if (columnHeader) columnHeader.classList.add("arSpreadsheetSelectedLabel");
  });
}

export function setSummaryTableCallbacks({
  renderRatioTable,
  onRatioStateMutated,
  toggleRatioInteractionMode,
} = {}) {
  if (typeof renderRatioTable === "function") summaryRuntime._renderRatioTable = renderRatioTable;
  if (typeof onRatioStateMutated === "function") {
    summaryRuntime._onRatioStateMutated = onRatioStateMutated;
  }
  if (typeof toggleRatioInteractionMode === "function") {
    summaryRuntime._toggleRatioInteractionMode = toggleRatioInteractionMode;
  }
}

export function resetSummaryFormulaEditState() {
  summaryRuntime.invalidateDfmExcelRefresh();
  summaryRuntime.cancelActiveSummaryFormulaCommit();
  summaryRuntime.cancelFormulaBarDisplayRefresh();
  summaryRuntime.clearFormulaBarFocusRestoreHandler();
  summaryRuntime.clearSummaryFormulaBarValidationError();
  summaryRuntime.summaryCopyHighlight?.destroy?.();
  summaryRuntime.summaryCopyHighlight = null;
  summaryRuntime.summarySelectionDestroy?.();
  summaryRuntime.summarySelectionDestroy = null;
  summaryRuntime.summaryFormulaEditState = null;
  summaryRuntime.summaryFormulaBarHoverCell = null;
  summaryRuntime.summaryFormulaBarHoverKey = "";
  summaryRuntime.summaryFormulaBarVisibleKey = "";
  // summaryFormulaBarSuppressedKey deliberately survives: a re-render is not the
  // user changing their mind about a bar they toggled off, and Edit mode
  // re-renders on the same click that toggles.
  summaryRuntime.summaryFormulaBarState = {
    mode: "display",
    input: null,
    generation: summaryRuntime.summaryFormulaBarState.generation + 1,
  };
}

export function updateRatioSummary() {
  const wrap = document.getElementById("ratioWrap");
  const model = summaryRuntime.state.model;
  if (!wrap || !model || !Array.isArray(model.values) || !Array.isArray(model.mask)) return;
  summaryRuntime.recalculateUserEntryDependencies();
  const cells = wrap.querySelectorAll("td.ratioCell[data-r]");
  if (!cells.length) return;

  const devs = summaryRuntime.getEffectiveDevLabelsForModel(model);
  cells.forEach((cell) => {
    const col = Number.parseInt(cell.dataset.c, 10);
    const rowType = cell.dataset.r;
    const config = summaryRuntime.summaryRowMap.get(rowType);
    const isSummary = !!config;

    if (!Number.isFinite(col) || col < 0) return;
    cell.classList.remove("userEntryEditable", "excelLinked", "excelLinkError");
    cell.title = "";
    if (config && summaryRuntime.isUserEntryConfig(config)) {
      const value = summaryRuntime.getUserEntryValueForCol(config, col);
      cell.textContent = summaryRuntime.formatUserEntryFormulaEvaluationValue(value);
      cell.classList.remove("na", "ratioPlaceholder", "strike");
      cell.classList.add("userEntryEditable");
      const inputText = String(summaryRuntime.getUserEntryInputForCol(config, col) || "");
      if (summaryRuntime.containsExcelRef(inputText)) {
        cell.classList.add("excelLinked");
        cell.title = inputText;
        // A reference that failed validation stays red through every re-render;
        // the record lives on the runtime, not on the discarded cell.
        const failure = summaryRuntime._dfmExcelInvalidTargets?.get(`${rowType}\u001f${col}`);
        if (failure) {
          cell.classList.add("excelLinkError");
          cell.title = `${inputText}\n${failure.error}`;
        }
      }
      return;
    }
    if (col >= devs.length - 1) {
      if (isSummary && summaryRuntime.summaryRowOwnsTail(config)) {
        // A frozen benchmark row shows its own tail factor, as ResQ does.
        const tail = summaryRuntime.getSummaryRowTailFactor(config, col);
        cell.textContent = summaryRuntime.formatRatio(tail, summaryRuntime.getDfmDecimalPlaces());
        cell.classList.remove("na", "ratioPlaceholder", "strike");
      } else if (isSummary) {
        cell.textContent = "1.0000";
        cell.classList.remove("na");
        cell.classList.add("ratioPlaceholder");
        cell.classList.remove("strike");
      } else {
        cell.textContent = "";
        cell.classList.add("na");
        cell.classList.remove("ratioPlaceholder", "strike");
      }
      return;
    }

    if (!config) return;
    summaryRuntime.ratioStrikeSet.delete(`${rowType},${col}`);
    const excluded = summaryRuntime.buildExcludedSetForColumn(
      model,
      col,
      config,
      summaryRuntime.ratioStrikeSet
    );
    const summary = summaryRuntime.computeAverageForColumn(
      model,
      col,
      excluded,
      config,
      summaryRuntime.ratioStrikeSet
    );
    if (summary.totalValid > 0 && summary.totalIncluded === 0) {
      cell.textContent = "1.0000";
      cell.classList.remove("na", "ratioPlaceholder", "strike");
      return;
    }
    const isVolume = String(config.base || "volume").toLowerCase() === "volume";
    const hasValue = summary.value !== null && (
      isVolume ? summary.sumA : summary.totalIncluded > 0
    );
    if (hasValue) {
      const rounded = summaryRuntime.roundRatio(summary.value, 6);
      cell.textContent = summaryRuntime.formatRatio(
        rounded,
        summaryRuntime.getDfmDecimalPlaces()
      );
      cell.classList.remove("na", "ratioPlaceholder");
    } else {
      cell.textContent = "1.0000";
      cell.classList.remove("na");
      cell.classList.add("ratioPlaceholder");
    }
    cell.classList.remove("strike");
  });

  const summaryTable = wrap.querySelector("table.ratioSummaryTable");
  const selectedTable = wrap.querySelector("table.ratioSelectedTable");
  if (summaryTable && selectedTable) {
    summaryRuntime.ensureSelectedRowValues(summaryTable, selectedTable);
    summaryRuntime.applyUserEntryReferenceHighlights(summaryTable);
    summaryRuntime.applyExcelRangeHighlights(summaryTable);
  }
}

export function scheduleRatioSummaryUpdate() {
  if (summaryRuntime.getRatioSummaryRaf()) return;
  summaryRuntime.setRatioSummaryRaf(requestAnimationFrame(() => {
    summaryRuntime.setRatioSummaryRaf(null);
    updateRatioSummary();
  }));
}

const delegate = (name) => (...args) => summaryRuntime[name](...args);

export const cancelDfmExcelFreshnessCheck = delegate("cancelDfmExcelFreshnessCheck");
export const buildRatioSelectionPattern = delegate("buildRatioSelectionPattern");
export const buildAverageSelectionPayload = delegate("buildAverageSelectionPayload");
export const applyRatioSelectionPattern = delegate("applyRatioSelectionPattern");
export const applySelectedSummaryFromSaved = delegate("applySelectedSummaryFromSaved");
export const applyAverageSelectionFromSaved = delegate("applyAverageSelectionFromSaved");
export const wireSummaryRowDrag = delegate("wireSummaryRowDrag");
export const clearSummaryTableHighlight = delegate("clearSummaryTableHighlight");
export const applyUserEntryReferenceHighlights = delegate("applyUserEntryReferenceHighlights");
export const isUserEntryConfig = delegate("isUserEntryConfig");
export const getUserEntryValueForCol = delegate("getUserEntryValueForCol");
export const refreshAllExcelLinks = delegate("refreshAllExcelLinks");
export const checkDfmExcelLinkFreshness = delegate("checkDfmExcelLinkFreshness");
export const getDfmExternalLinkRecords = delegate("getDfmExternalLinkRecords");
export const breakDfmExternalLinks = delegate("breakDfmExternalLinks");
export const breakDfmExternalLink = delegate("breakDfmExternalLink");
export const recalculateUserEntryDependencies = delegate("recalculateUserEntryDependencies");
export const wireSummaryContextMenu = delegate("wireSummaryContextMenu");
export const applySummarySelection = delegate("applySummarySelection");
export const selectSummaryCell = delegate("selectSummaryCell");
export const initDefaultSummarySelection = delegate("initDefaultSummarySelection");
export const wireSummarySelection = delegate("wireSummarySelection");

Object.assign(summaryRuntime, {
  DFM_RATIO_HIGHLIGHT_EDGE_CLASSES,
  isRatioEditMode,
  refreshRatioHighlightHeaders,
});
registerSummaryFunctions({
  updateRatioSummary,
  scheduleRatioSummaryUpdate,
});
