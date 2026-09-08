/*
===============================================================================
DFM Ratios Summary User Entries
===============================================================================
*/
import {
  registerSummaryFunctions,
  summaryRuntime,
} from "/ui/method_pages/dfm/ratios_summary/summary_runtime.js?v=20260819a";
import { containsDfmDatasetReference } from "/ui/method_pages/dfm/dfm_dataset_reference.js?v=20260811b";
import {
  resolveDfmDatasetReferencesInFormulaDetailed,
  substituteCachedDfmDatasetReferencesInFormula,
} from "/ui/method_pages/dfm/dfm_dataset_formula.js?v=20260820a";

const {
  state, calcRatio, roundRatio, averageRowReferenceValue, formatRatio, computeAverageForColumn,
  ratioStrikeSet, selectedSummaryByCol, summaryRowConfigs, summaryRowMap, BASE_SUMMARY_ROWS,
  getShowNaBorders, getRatioSummaryRaf, setRatioSummaryRaf,
  getLastSummaryCtxRowId, setLastSummaryCtxRowId,
  getEffectiveDevLabelsForModel, getRatioHeaderLabels, buildSummaryRows,
  buildExcludedSetForColumn, parsePeriodsValue, parseExcludeValue, getDfmDecimalPlaces,
  getSummaryConfigKey, loadCustomSummaryRows, saveCustomSummaryRows,
  readExcelCell, readExcelCellsBatch, openExcelWorkbook,
  buildExcelRangeSourceCells, containsExcelRef, excelColumnFromIndex, findExcelRefsInline,
  formatExcelRef, normalizeExcelReferenceAddressCase, parseStandaloneExcelRange,
  collectDfmExternalLinkGroupsModel, getDfmExternalLinkHardCodeTargets, getDfmExternalLinkRangeTargets,
  DFM_FORMULA_VALIDATION_TIMEOUT_MS, beginFormulaValidationLease, clearFormulaValidationError,
  computeFormulaValidationTooltipLayout, revealAndFocusFormulaInput, showFormulaValidationError,
  wireSelectableTable, openDfmSummaryPlotWindow, hasDfmCellNote, showDfmCellNoteEditor,
  beginRatioHistoryAction, commitRatioHistoryAction,
} = summaryRuntime;

const getAvgModalEl = (...args) => summaryRuntime.getAvgModalEl(...args);
const isSummaryFormulaEditSessionActive = (...args) => summaryRuntime.isSummaryFormulaEditSessionActive(...args);
const normalizeAverageType = (...args) => summaryRuntime.normalizeAverageType(...args);
const isUserEntryConfig = (...args) => summaryRuntime.isUserEntryConfig(...args);
const getCurrentRatioColumnCount = (...args) => summaryRuntime.getCurrentRatioColumnCount(...args);
const sanitizeUserEntryValue = (...args) => summaryRuntime.sanitizeUserEntryValue(...args);
const findReferencedLabels = (...args) => summaryRuntime.findReferencedLabels(...args);
const updateActiveSummaryFormulaReferenceUi = (...args) => summaryRuntime.updateActiveSummaryFormulaReferenceUi(...args);
const applyUserEntryReferenceHighlights = (...args) => summaryRuntime.applyUserEntryReferenceHighlights(...args);
const evaluateSimpleMathExpression = (...args) => summaryRuntime.evaluateSimpleMathExpression(...args);
const stripFormulaEquals = (...args) => summaryRuntime.stripFormulaEquals(...args);
const parseSummaryArrayFormula = (...args) => summaryRuntime.parseSummaryArrayFormula(...args);
const normalizeUserEntryValues = (...args) => summaryRuntime.normalizeUserEntryValues(...args);
const normalizeUserEntryInputs = (...args) => summaryRuntime.normalizeUserEntryInputs(...args);
const normalizeUserEntryDisplayInputs = (...args) => summaryRuntime.normalizeUserEntryDisplayInputs(...args);
const getUserEntryValueForCol = (...args) => summaryRuntime.getUserEntryValueForCol(...args);
const getUserEntryInputForCol = (...args) => summaryRuntime.getUserEntryInputForCol(...args);
const getUserEntryDisplayInputForCol = (...args) => summaryRuntime.getUserEntryDisplayInputForCol(...args);
const summaryTableHasUserEntryRows = (...args) => summaryRuntime.summaryTableHasUserEntryRows(...args);
const setModalValidationError = (...args) => summaryRuntime.setModalValidationError(...args);
const clearModalValidationError = (...args) => summaryRuntime.clearModalValidationError(...args);
const hideAvgModal = (...args) => summaryRuntime.hideAvgModal(...args);
const computeAutoNameWithExclude = (...args) => summaryRuntime.computeAutoNameWithExclude(...args);
const scrollSummaryFormulaInputToEnd = (...args) => summaryRuntime.scrollSummaryFormulaInputToEnd(...args);
const updateFormulaBarDisplayMode = (...args) => summaryRuntime.updateFormulaBarDisplayMode(...args);
const positionSummaryFormulaBar = (...args) => summaryRuntime.positionSummaryFormulaBar(...args);
const clearSummaryFormulaBarValidationError = (...args) => summaryRuntime.clearSummaryFormulaBarValidationError(...args);
const showSummaryFormulaBarValidationError = (...args) => summaryRuntime.showSummaryFormulaBarValidationError(...args);
const isSummaryFormulaBarInputEditing = (...args) => summaryRuntime.isSummaryFormulaBarInputEditing(...args);
const summaryFormulaBarTargetKey = (...args) => summaryRuntime.summaryFormulaBarTargetKey(...args);
const ensureSummaryFormulaBarEl = (...args) => summaryRuntime.ensureSummaryFormulaBarEl(...args);
const setStatusBarText = (...args) => summaryRuntime.setStatusBarText(...args);
const invalidateDfmExcelRefresh = (...args) => summaryRuntime.invalidateDfmExcelRefresh(...args);
const commitExcelFormulaAsync = (...args) => summaryRuntime.commitExcelFormulaAsync(...args);
const hideSummaryFormulaBar = (...args) => summaryRuntime.hideSummaryFormulaBar(...args);
const setUserEntryCellDisplayValue = (...args) => summaryRuntime.setUserEntryCellDisplayValue(...args);
const restoreSupersededExcelRange = (...args) => summaryRuntime.restoreSupersededExcelRange(...args);
const getSummaryArrayFormulaDestination = (...args) => summaryRuntime.getSummaryArrayFormulaDestination(...args);
const applyExcelRangeHighlights = (...args) => summaryRuntime.applyExcelRangeHighlights(...args);
const commitExcelRangeFormulaAsync = (...args) => summaryRuntime.commitExcelRangeFormulaAsync(...args);
const ensureSelectedRowValues = (...args) => summaryRuntime.ensureSelectedRowValues(...args);
const isRatioEditMode = (...args) => summaryRuntime.isRatioEditMode(...args);
const refreshRatioHighlightHeaders = (...args) => summaryRuntime.refreshRatioHighlightHeaders(...args);

function parseUserEntryClipboardGrid(rawText) {
  const normalized = String(rawText ?? "").replace(/\r\n?/g, "\n").replace(/\n+$/, "");
  if (!normalized) return { ok: false, error: "The clipboard does not contain a value." };
  const rows = normalized.split("\n").map((row) => row.split("\t"));
  const width = rows[0]?.length || 0;
  if (!width || rows.some((row) => row.length !== width)) {
    return { ok: false, error: "Paste a rectangular range of Excel cells." };
  }
  return { ok: true, rows, width };
}

function parseUserEntryClipboardValue(raw, referenceValues) {
  const text = String(raw ?? "").trim();
  if (!text) return null;

  const evaluated = evaluateSimpleMathExpression(text, referenceValues);
  if (Number.isFinite(evaluated) && evaluated > 0) {
    return { input: text, value: roundRatio(evaluated, 6) };
  }

  const compact = text.replace(/\u00a0/g, "").replace(/,/g, "");
  const formattedNumber = /^([+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)(%)?$/.exec(compact);
  if (!formattedNumber) return null;
  const numeric = Number(formattedNumber[1]);
  const value = formattedNumber[2] ? numeric / 100 : numeric;
  if (!Number.isFinite(value) || value <= 0) return null;
  const rounded = roundRatio(value, 6);
  return { input: String(rounded), value: rounded };
}

function pasteUserEntryClipboardGrid(summaryTable, selectedTable, startCell, rawText) {
  if (!summaryTable || !startCell) return false;
  if (startCell.classList.contains("excelRangeSpillCell")) {
    showSummaryFormulaBarValidationError("Edit the first cell of the Excel-linked range instead.");
    return true;
  }
  const startRow = startCell.closest("tr[data-row-id]");
  const startRowId = String(startRow?.dataset?.rowId || "");
  const startCol = Number(startCell.dataset.col);
  if (!startRow || !startRowId || !Number.isFinite(startCol) || startCol < 0) return false;
  if (!isUserEntryConfig(summaryRowMap.get(startRowId))) return false;

  const parsedGrid = parseUserEntryClipboardGrid(rawText);
  if (!parsedGrid.ok) {
    showSummaryFormulaBarValidationError(parsedGrid.error);
    return true;
  }

  const tableRows = Array.from(summaryTable.querySelectorAll("tr[data-row-id]"));
  const startRowIndex = tableRows.indexOf(startRow);
  const entries = [];
  for (let rowOffset = 0; rowOffset < parsedGrid.rows.length; rowOffset++) {
    const targetRow = tableRows[startRowIndex + rowOffset];
    if (!targetRow) {
      showSummaryFormulaBarValidationError("The pasted range extends beyond the available Average Formula rows.");
      return true;
    }
    const rowId = String(targetRow.dataset.rowId || "");
    if (!isUserEntryConfig(summaryRowMap.get(rowId))) {
      showSummaryFormulaBarValidationError("Every destination row in the pasted range must be a User Entry row.");
      return true;
    }
    for (let colOffset = 0; colOffset < parsedGrid.width; colOffset++) {
      const col = startCol + colOffset;
      const cell = targetRow.querySelector(`td.summaryCell[data-col="${col}"]`);
      if (!cell) {
        showSummaryFormulaBarValidationError("The pasted range extends beyond the available development columns.");
        return true;
      }
      const parsedValue = parseUserEntryClipboardValue(
        parsedGrid.rows[rowOffset][colOffset],
        buildSummaryReferenceValues(summaryTable, col)
      );
      if (!parsedValue) {
        showSummaryFormulaBarValidationError(
          `Clipboard value at row ${rowOffset + 1}, column ${colOffset + 1} must be a number greater than 0.`
        );
        return true;
      }
      entries.push({ cell, rowId, col, ...parsedValue });
    }
  }

  clearSummaryReferenceUi(summaryTable);
  clearSummaryFormulaBarValidationError();
  summaryRuntime.summaryFormulaEditState = null;
  entries.forEach((entry) => {
    restoreSupersededExcelRange(summaryTable, entry.rowId, entry.col, entry.input);
    setUserEntryCellEntry(entry.rowId, entry.col, entry.input, entry.value, { persist: false });
    setUserEntryCellDisplayValue(entry.cell, entry.value);
  });
  persistUserEntryRowsFromState();
  ensureSelectedRowValues(summaryTable, selectedTable);
  applyUserEntryReferenceHighlights(summaryTable);
  applyExcelRangeHighlights(summaryTable);
  summaryRuntime.summaryCopyHighlight?.selectCell?.(startCell, false);
  summaryRuntime.summaryActiveCellState = { rowId: startRowId, col: startCol };
  updateSummaryFormulaBarForCell(startCell);
  summaryRuntime._onRatioStateMutated();
  const count = entries.length;
  setStatusBarText(`Pasted ${count} value${count === 1 ? "" : "s"} into User Entry.`);
  return true;
}

function commitUserEntryArrayFormula(summaryTable, selectedTable, rowId, startCol, raw) {
  const parsedArray = parseSummaryArrayFormula(raw);
  if (!parsedArray) return { handled: false, ok: true };
  if (!parsedArray.ok) return { handled: true, ok: false, error: parsedArray.error };
  if (containsExcelRef(raw)) {
    return {
      handled: true,
      ok: false,
      error: "Array formulas currently support numbers and DFM row-reference math, but not Excel cell links inside the array.",
    };
  }
  if (containsDfmDatasetReference(raw)) {
    return {
      handled: true,
      ok: false,
      error: "Array formulas do not support ArcRho dataset references.",
    };
  }

  const availableCells = getSummaryArrayFormulaDestination(
    summaryTable,
    rowId,
    startCol,
    parsedArray.expressions.length,
  ).entries;
  const applyCount = availableCells.length;
  if (applyCount <= 0) {
    return { handled: true, ok: false, error: "Array formula has no cells available to fill." };
  }

  const nextEntries = [];
  for (let i = 0; i < applyCount; i++) {
    const targetCol = availableCells[i].col;
    const expr = String(parsedArray.expressions[i] || "").trim();
    const refValues = buildSummaryReferenceValues(summaryTable, targetCol);
    const value = evaluateSimpleMathExpression(expr, refValues);
    if (!Number.isFinite(value) || value <= 0) {
      return {
        handled: true,
        ok: false,
        error: "Each array formula item must evaluate to a number > 0.",
      };
    }
    const nextValue = roundRatio(value, 6);
    nextEntries.push({
      cell: availableCells[i].cell,
      col: targetCol,
      value: nextValue,
      input: i === 0 ? String(raw || "").trim() : String(nextValue),
    });
  }

  restoreSupersededExcelRange(summaryTable, rowId, startCol, raw);
  nextEntries.forEach((entry) => {
    setUserEntryCellEntry(rowId, entry.col, entry.input, entry.value, { persist: false });
    setUserEntryCellDisplayValue(entry.cell, entry.value);
    selectedSummaryByCol.set(entry.col, String(rowId));
    summaryTable.querySelectorAll(`td.summaryCell[data-col="${entry.col}"]`)
      .forEach((el) => el.classList.remove("ratioSelectedCell"));
    entry.cell.classList.add("ratioSelectedCell");
  });
  persistUserEntryRowsFromState();

  const firstCell = nextEntries[0]?.cell || null;
  summaryTable.querySelectorAll("td.summaryCell.summaryActiveCell")
    .forEach((el) => el.classList.remove("summaryActiveCell"));
  if (firstCell) {
    firstCell.classList.add("summaryActiveCell");
    summaryRuntime.summaryCopyHighlight?.selectCell?.(firstCell, false);
    summaryRuntime.summaryActiveCellState = { rowId: String(rowId), col: nextEntries[0].col };
  }
  if (selectedTable) ensureSelectedRowValues(summaryTable, selectedTable);
  applyUserEntryReferenceHighlights(summaryTable);
  applyExcelRangeHighlights(summaryTable);
  clearSummaryReferenceUi(summaryTable);
  summaryRuntime.summaryFormulaEditState = null;
  updateSummaryFormulaBarForCell(firstCell);
  summaryRuntime._onRatioStateMutated();
  return { handled: true, ok: true };
}

function isSummaryFormulaCommitPending(inputEl) {
  return inputEl?.dataset?.formulaCommitPending === "1";
}

async function commitSummaryFormulaInput(inputEl) {
  const summaryTable = document.querySelector("#ratioWrap table.ratioSummaryTable");
  const selectedTable = document.querySelector("#ratioWrap table.ratioSelectedTable");
  if (!inputEl || !summaryTable) return true;
  const rowId = String(inputEl.dataset.rowId || "");
  const col = Number(inputEl.dataset.col);
  if (!rowId || !Number.isFinite(col) || col < 0) return true;
  const cfg = summaryRowMap.get(rowId);
  if (!cfg || !isUserEntryConfig(cfg)) return true;
  if (isSummaryFormulaCommitPending(inputEl)) return false;

  const generation = ++summaryRuntime.summaryFormulaCommitGeneration;
  const validationLease = beginFormulaValidationLease(inputEl, {
    timeoutMs: DFM_FORMULA_VALIDATION_TIMEOUT_MS,
  });
  summaryRuntime.summaryFormulaCommitLease = validationLease;
  const isCurrent = () => (
    generation === summaryRuntime.summaryFormulaCommitGeneration &&
    summaryRuntime.summaryFormulaCommitLease === validationLease &&
    inputEl.isConnected
  );
  clearSummaryFormulaBarValidationError();
  try {
    const raw = normalizeExcelReferenceAddressCase(String(inputEl.value || "").trim());
    inputEl.value = raw;
    const excelRange = parseStandaloneExcelRange(raw);
    if (excelRange) {
      return await commitExcelRangeFormulaAsync(rowId, col, raw, excelRange, {
        signal: validationLease.signal,
        isCurrent,
      });
    }
    const arrayCommit = commitUserEntryArrayFormula(summaryTable, selectedTable, rowId, col, raw);
    if (arrayCommit.handled) {
      if (!arrayCommit.ok) {
        showSummaryFormulaBarValidationError(arrayCommit.error || "Could not apply array formula.", inputEl);
      }
      return !!arrayCommit.ok;
    }
    // Check if expression contains any Excel references (standalone or inline)
    if (containsExcelRef(raw)) {
      return await commitExcelFormulaAsync(rowId, col, raw, {
        signal: validationLease.signal,
        isCurrent,
      });
    }
    const resolvedDatasetFormula = await resolveDfmDatasetReferencesInFormulaDetailed(raw, {
      signal: validationLease.signal,
    });
    if (!isCurrent()) return false;
    const refValues = buildSummaryReferenceValues(summaryTable, col);
    const parsed = stripFormulaEquals(resolvedDatasetFormula.resolvedFormula)
      ? evaluateSimpleMathExpression(resolvedDatasetFormula.resolvedFormula, refValues)
      : 1;
    if (!Number.isFinite(parsed) || parsed <= 0) {
      showSummaryFormulaBarValidationError(
        "Enter a number > 0, a DFM row formula, or an ArcRho dataset reference.",
        inputEl
      );
      return false;
    }
    const nextValue = roundRatio(parsed, 6);
    restoreSupersededExcelRange(summaryTable, rowId, col, raw);
    setUserEntryCellEntry(rowId, col, stripFormulaEquals(raw) ? raw : "1", nextValue, {
      displayInput: resolvedDatasetFormula.displayFormula === raw ? "" : resolvedDatasetFormula.displayFormula,
    });
    persistUserEntryRowsFromState();
    const cell = summaryTable.querySelector(`td.summaryCell[data-r="${rowId}"][data-col="${col}"]`);
    if (cell) setUserEntryCellDisplayValue(cell, nextValue);
    if (selectedTable) ensureSelectedRowValues(summaryTable, selectedTable);
    applyUserEntryReferenceHighlights(summaryTable);
    applyExcelRangeHighlights(summaryTable);
    clearSummaryReferenceUi(summaryTable);
    summaryRuntime.summaryFormulaEditState = null;
    clearSummaryFormulaBarValidationError();
    updateSummaryFormulaBarForCell(cell);
    summaryRuntime._onRatioStateMutated();
    return true;
  } catch (error) {
    if (!isCurrent()) return false;
    if (error?.name === "AbortError") {
      if (validationLease.timedOut) {
        showSummaryFormulaBarValidationError(
          "Linked formula validation timed out after 30 seconds. Check the source and try again.",
          inputEl
        );
      }
      return false;
    }
    showSummaryFormulaBarValidationError(error?.message || "Formula validation failed.", inputEl);
    return false;
  } finally {
    validationLease.finish();
    if (summaryRuntime.summaryFormulaCommitLease === validationLease) summaryRuntime.summaryFormulaCommitLease = null;
  }
}

function updateSummaryFormulaBarForCell(cell) {
  const summaryTable =
    cell?.closest?.("table.ratioSummaryTable") ||
    document.querySelector("#ratioWrap table.ratioSummaryTable");
  if (!summaryTable) {
    hideSummaryFormulaBar();
    return;
  }
  if (!summaryTableHasUserEntryRows(summaryTable)) {
    hideSummaryFormulaBar();
    return;
  }

  const el = ensureSummaryFormulaBarEl(summaryTable);
  const inputEl = el.querySelector("#dfmSummaryFormulaBarInput");
  let inputRaw = "";
  let targetCell = cell;
  if (!targetCell || !summaryTable.contains(targetCell)) {
    const stateCell = summaryTable.querySelector(
      `td.summaryCell[data-r="${summaryRuntime.summaryActiveCellState.rowId}"][data-col="${summaryRuntime.summaryActiveCellState.col}"]`
    );
    targetCell = stateCell || null;
  }
  if (targetCell) {
    const rowId = String(targetCell.dataset.r || "");
    const col = Number(targetCell.dataset.col);
    if (rowId && Number.isFinite(col) && col >= 0) {
      const isExcelRangeCell = !!targetCell.dataset.excelRangeFormula;
      const anchorCol = Number(targetCell.dataset.excelRangeAnchorCol);
      const editRowId = isExcelRangeCell
        ? String(targetCell.dataset.excelRangeAnchorRowId || rowId)
        : rowId;
      const editCol = isExcelRangeCell && Number.isFinite(anchorCol) && anchorCol >= 0
        ? anchorCol
        : col;
      const cfg = summaryRowMap.get(editRowId);
      if (cfg && isUserEntryConfig(cfg)) {
        inputRaw = isExcelRangeCell
          ? String(targetCell.dataset.excelRangeFormula || "").trim()
          : String(getUserEntryInputForCol(cfg, editCol) || "").trim();
        const displayInputRaw = isExcelRangeCell
          ? ""
          : String(getUserEntryDisplayInputForCol(cfg, editCol) || "").trim();
        const labelEl = el.querySelector("#dfmSummaryFormulaBarLabelText");
        if (labelEl) {
          const rowLabel = String(cfg.label || cfg.id || "f(x)");
          labelEl.textContent = rowLabel;
        }
        if (inputEl) {
          const inputHasFocus = document.activeElement === inputEl;
          const sameTarget =
            String(inputEl.dataset.rowId || "") === editRowId &&
            Number(inputEl.dataset.col) === editCol;
          const editingSameTarget = sameTarget && isSummaryFormulaBarInputEditing(inputEl);
          if ((!inputHasFocus && !editingSameTarget) || !sameTarget) {
            const body = (inputRaw || "").replace(/^=\s*/, "");
            inputEl.value = "= " + body;
            scrollSummaryFormulaInputToEnd(inputEl);
          }
          if (!sameTarget) clearSummaryFormulaBarValidationError();
          inputEl.dataset.rowId = editRowId;
          inputEl.dataset.col = String(editCol);
          if (displayInputRaw) inputEl.dataset.displayFormula = displayInputRaw;
          else delete inputEl.dataset.displayFormula;
          inputEl.disabled = false;
          inputEl.placeholder = "Enter value or formula";
        }

      } else {
        hideSummaryFormulaBar();
        return;
      }
    }
  } else {
    hideSummaryFormulaBar();
    return;
  }

  // A target the user toggled off stays off until they pick a different one.
  const targetKey = summaryFormulaBarTargetKey(targetCell);
  if (summaryRuntime.summaryFormulaBarSuppressedKey === targetKey) {
    hideSummaryFormulaBar({ keepHoverTarget: true });
    return;
  }
  summaryRuntime.summaryFormulaBarVisibleKey = targetKey;
  // Showing the bar for any other target ends a hand-placed position, so moving
  // to another cell and coming back both restore the anchored one.
  summaryRuntime.syncSummaryFormulaBarDragPlacementTarget?.(el, targetKey);

  el.classList.add("isOpen");
  const isEditing = isSummaryFormulaBarInputEditing(inputEl);
  updateFormulaBarDisplayMode(el, isEditing);
  positionSummaryFormulaBar(el, summaryTable, targetCell);
  window.requestAnimationFrame(() => positionSummaryFormulaBar(el, summaryTable, targetCell));
}

function refreshSummaryFormulaBar() {
  // A hovered dynamic array outranks the active cell, so keep the bar on it.
  const hoverCell = summaryRuntime.summaryFormulaBarHoverCell;
  updateSummaryFormulaBarForCell(hoverCell?.isConnected ? hoverCell : null);
}

function handleSummaryTableSelectionChange(summaryTable, selection) {
  refreshRatioHighlightHeaders();
  // The reference colours belong to the highlighted cell, so they move with it.
  applyUserEntryReferenceHighlights(summaryTable);
  if (isRatioEditMode() || isSummaryFormulaEditSessionActive(summaryTable)) return;
  const active = selection?.activeCell;
  const cell = active
    ? summaryTable.querySelector(
      `td.summaryCell[data-copy-r="${active.r}"][data-copy-c="${active.c}"]`,
    )
    : null;
  if (!cell) {
    summaryRuntime.summaryActiveCellState = { rowId: "", col: -1 };
    hideSummaryFormulaBar();
    return;
  }
  const rowId = String(cell.dataset.r || "");
  const col = Number(cell.dataset.col);
  if (!rowId || !Number.isFinite(col) || col < 0) {
    hideSummaryFormulaBar();
    return;
  }
  summaryRuntime.summaryActiveCellState = { rowId, col };
  updateSummaryFormulaBarForCell(cell);
}

function clearSummaryReferenceUi(summaryTable) {
  if (!summaryTable) return;
  summaryTable.querySelectorAll("td.summaryCell.summaryRefHover")
    .forEach((el) => el.classList.remove("summaryRefHover"));
  summaryTable.querySelectorAll("td.summaryCell.summaryRefCandidate")
    .forEach((el) => el.classList.remove("summaryRefCandidate"));
  summaryTable.querySelectorAll("td.summaryCell.summaryFormulaActiveRefCell")
    .forEach((el) => el.classList.remove("summaryFormulaActiveRefCell"));
  summaryTable.querySelectorAll("td.summaryCell.summaryFormulaRefDragTarget")
    .forEach((el) => el.classList.remove("summaryFormulaRefDragTarget"));
  summaryTable.querySelectorAll("td.summaryCell.summaryFormulaRefDragReady")
    .forEach((el) => el.classList.remove("summaryFormulaRefDragReady"));
}

// Reference values are computed by the same engine that recalculates User Entry
// rows, never read back off the displayed cell text, which carries a thousands
// separator and an empty string where a row has no value. Each one is then read
// at the Decimal Places the Ratios tab prints it at, so a reviewer multiplying
// the digits on screen reaches the User Entry factor exactly.
function buildSummaryReferenceValues(_summaryTable, col) {
  const out = new Map();
  if (!Number.isFinite(col) || col < 0) return out;
  const model = state.model;
  if (!model || !Array.isArray(model.values) || !Array.isArray(model.mask)) return out;
  const rows = Array.isArray(summaryRowConfigs) ? summaryRowConfigs : [];
  const devs = getEffectiveDevLabelsForModel(model);
  const lastCol = Math.max(0, devs.length - 1);
  const labelToId = new Map(
    rows.map((cfg) => [String(cfg?.label || cfg?.id || "").trim(), String(cfg?.id || "")]).filter(([k, v]) => k && v)
  );
  const cache = new Map();
  const visiting = new Set();
  const decimals = getDfmDecimalPlaces();
  labelToId.forEach((rowId, label) => {
    const value = averageRowReferenceValue(
      computeSummaryRowValueForColumn(model, col, rowId, cache, visiting, labelToId, lastCol),
      decimals,
    );
    if (Number.isFinite(value)) out.set(label, Number(value));
  });
  return out;
}

function insertAtInputCursor(input, text) {
  if (!input) return;
  const start = Number.isFinite(input.selectionStart) ? input.selectionStart : input.value.length;
  const end = Number.isFinite(input.selectionEnd) ? input.selectionEnd : input.value.length;
  const before = input.value.slice(0, start);
  const after = input.value.slice(end);
  input.value = `${before}${text}${after}`;
  const nextPos = start + text.length;
  input.setSelectionRange(nextPos, nextPos);
}

function beginSummaryFormulaEditSession(summaryTable, cell, input, col) {
  if (!summaryTable || !cell || !input) return;
  if (!Number.isFinite(col) || col < 0) return;
  const rowId = String(cell.dataset.r || "");
  if (!rowId) return;
  const cfg = summaryRowMap.get(rowId);
  const fallbackOriginal = cfg && isUserEntryConfig(cfg)
    ? String(getUserEntryInputForCol(cfg, col) || "").trim()
    : "";
  const keepOriginal =
    summaryRuntime.summaryFormulaEditState &&
    summaryRuntime.summaryFormulaEditState.summaryTable === summaryTable &&
    summaryRuntime.summaryFormulaEditState.cell === cell &&
    Number(summaryRuntime.summaryFormulaEditState.col) === col
      ? String(summaryRuntime.summaryFormulaEditState.originalInput ?? fallbackOriginal)
      : fallbackOriginal;
  summaryRuntime.summaryFormulaEditState = {
    summaryTable,
    cell,
    input,
    col,
    rowId,
    originalInput: keepOriginal,
  };
  updateActiveSummaryFormulaReferenceUi(summaryTable);
  // The draft decides which cells read as referenced, so the fill follows the
  // formula as it is typed or dragged rather than waiting for the commit.
  applyUserEntryReferenceHighlights(summaryTable);
}

function cancelSummaryFormulaEditSession() {
  const state = summaryRuntime.summaryFormulaEditState;
  if (!state) return;
  const { summaryTable, cell, input, originalInput } = state;
  if (input && document.body.contains(input)) {
    input.value = String(originalInput ?? "");
  }
  clearSummaryReferenceUi(summaryTable);
  summaryRuntime.summaryFormulaEditState = null;
  applyUserEntryReferenceHighlights(summaryTable);
  updateSummaryFormulaBarForCell(cell);
}

function setUserEntryCellEntry(rowId, col, inputRaw, value, options = {}) {
  const persist = options?.persist !== false;
  if (!rowId || !Number.isFinite(col) || col < 0) return false;
  const cfg = summaryRowMap.get(String(rowId));
  if (!cfg || !isUserEntryConfig(cfg)) return false;
  if (!summaryRuntime._applyingDfmExcelRefresh) invalidateDfmExcelRefresh();

  const nextInput = String(inputRaw ?? "").trim() || "1";
  const nextValue = sanitizeUserEntryValue(value);
  const colCount = getCurrentRatioColumnCount();
  const values = normalizeUserEntryValues(cfg.values, Math.max(colCount, col + 1));
  const inputs = normalizeUserEntryInputs(cfg.inputs ?? cfg.formulas, values, Math.max(colCount, col + 1));
  const displayInputs = normalizeUserEntryDisplayInputs(cfg.displayInputs, Math.max(colCount, col + 1));
  const previousInput = String(inputs[col] ?? "").trim();
  values[col] = nextValue;
  inputs[col] = nextInput;
  if (Object.prototype.hasOwnProperty.call(options || {}, "displayInput")) {
    displayInputs[col] = String(options.displayInput ?? "").trim();
  } else if (previousInput !== nextInput) {
    displayInputs[col] = "";
  }
  cfg.values = values;
  cfg.inputs = inputs;
  cfg.displayInputs = displayInputs;
  if (Object.prototype.hasOwnProperty.call(cfg, "formulas")) delete cfg.formulas;

  if (!persist) return true;
  const cfgKey = getSummaryConfigKey();
  if (!cfgKey) return true;
  const customRows = loadCustomSummaryRows(cfgKey);
  const idx = customRows.findIndex((row) => String(row?.id || "") === String(rowId));
  if (idx < 0) return true;
  const { formulas: _legacyFormulas, ...baseRow } = customRows[idx] || {};
  customRows[idx] = {
    ...baseRow,
    averageType: "user_entry",
    base: "simple",
    periods: "all",
    exclude: 0,
    values,
    inputs,
    displayInputs,
  };
  saveCustomSummaryRows(cfgKey, customRows);
  return true;
}

function persistUserEntryRowsFromState() {
  const cfgKey = getSummaryConfigKey();
  if (!cfgKey) return;
  const customRows = loadCustomSummaryRows(cfgKey);
  if (!Array.isArray(customRows) || !customRows.length) return;
  let changed = false;
  const colCount = getCurrentRatioColumnCount();
  const nextRows = customRows.map((row) => {
    const rowId = String(row?.id || "");
    const cfg = summaryRowMap.get(rowId);
    if (!cfg || !isUserEntryConfig(cfg)) return row;
    const values = normalizeUserEntryValues(cfg.values, colCount);
    const inputs = normalizeUserEntryInputs(cfg.inputs ?? cfg.formulas, values, colCount);
    const displayInputs = normalizeUserEntryDisplayInputs(cfg.displayInputs, colCount);
    const { formulas: _legacyFormulas, ...baseRow } = row || {};
    const nextRow = {
      ...baseRow,
      averageType: "user_entry",
      base: "simple",
      periods: "all",
      exclude: 0,
      values,
      inputs,
      displayInputs,
    };
    if (!changed) changed = JSON.stringify(row) !== JSON.stringify(nextRow);
    return nextRow;
  });
  if (changed) saveCustomSummaryRows(cfgKey, nextRows);
}

function computeSummaryRowValueForColumn(model, col, rowId, cache, visiting, labelToId, lastCol) {
  const key = String(rowId || "");
  if (!key) return 1;
  if (cache.has(key)) return cache.get(key);
  if (visiting.has(key)) return 1;

  const cfg = summaryRowMap.get(key);
  if (!cfg) {
    cache.set(key, 1);
    return 1;
  }
  if (col >= lastCol) {
    // The "- Ult" column is the row's entered tail factor; no formula runs there.
    const tail = summaryRuntime.getSummaryRowTailFactor(cfg, col);
    cache.set(key, tail);
    return tail;
  }

  let value = 1;
  if (isUserEntryConfig(cfg)) {
    // A referenced row enters the formula at the precision the tab prints it at.
    const decimals = getDfmDecimalPlaces();
    const storedInput = String(getUserEntryInputForCol(cfg, col) || "").trim();
    // Substitute dataset references with their last-resolved session values so
    // formulas that mix dataset references with average-formula row references
    // still re-evaluate when a referenced row changes. Until every dataset
    // reference has a resolved value, keep the stored value.
    const datasetSubstitution = substituteCachedDfmDatasetReferencesInFormula(storedInput);
    if (!datasetSubstitution.ok) {
      const stored = sanitizeUserEntryValue(getUserEntryValueForCol(cfg, col));
      cache.set(key, stored);
      return stored;
    }
    const inputRaw = datasetSubstitution.formula;
    // Determine which labels are actually referenced in this formula
    const allLabels = Array.from(labelToId.keys());
    const referencedLabels = findReferencedLabels(inputRaw, allLabels);

    if (containsExcelRef(inputRaw)) {
      // Substitute Excel refs with cached values, then evaluate with current row refs
      let expr = inputRaw.startsWith("=") ? inputRaw : "=" + inputRaw;
      const xlRefs = findExcelRefsInline(expr);
      let allCached = true;
      for (const ref of xlRefs) {
        if (summaryRuntime._xlCellValueCache.has(ref.match)) {
          expr = expr.split(ref.match).join(String(summaryRuntime._xlCellValueCache.get(ref.match)));
        } else {
          allCached = false;
        }
      }
      if (allCached) {
        visiting.add(key);
        const refValues = new Map();
        for (const label of referencedLabels) {
          const depId = labelToId.get(label);
          if (!depId || String(depId) === key) continue;
          const depValue = averageRowReferenceValue(
            computeSummaryRowValueForColumn(model, col, depId, cache, visiting, labelToId, lastCol),
            decimals,
          );
          if (Number.isFinite(depValue)) refValues.set(label, depValue);
        }
        visiting.delete(key);
        const parsed = evaluateSimpleMathExpression(expr, refValues);
        value = Number.isFinite(parsed) && parsed > 0 ? roundRatio(parsed, 6) : sanitizeUserEntryValue(getUserEntryValueForCol(cfg, col));
      } else {
        // No cached Excel values yet; keep the stored value
        value = sanitizeUserEntryValue(getUserEntryValueForCol(cfg, col));
      }
    } else {
      visiting.add(key);
      const refValues = new Map();
      for (const label of referencedLabels) {
        const depId = labelToId.get(label);
        if (!depId || String(depId) === key) continue;
        const depValue = averageRowReferenceValue(
          computeSummaryRowValueForColumn(model, col, depId, cache, visiting, labelToId, lastCol),
          decimals,
        );
        if (Number.isFinite(depValue)) refValues.set(label, depValue);
      }
      visiting.delete(key);
      const parsed = inputRaw ? evaluateSimpleMathExpression(inputRaw, refValues) : 1;
      if (Number.isFinite(parsed) && parsed > 0) {
        value = roundRatio(parsed, 6);
      } else {
        // If evaluation failed (e.g. dependency has Excel ref not yet cached),
        // keep the current stored value instead of resetting to 1
        const stored = sanitizeUserEntryValue(getUserEntryValueForCol(cfg, col));
        value = stored;
      }
    }
  } else {
    const excluded = buildExcludedSetForColumn(model, col, cfg, ratioStrikeSet);
    const summary = computeAverageForColumn(model, col, excluded, cfg, ratioStrikeSet);
    if (summary.totalValid > 0 && summary.totalIncluded === 0) {
      value = 1;
    } else {
      const isVolume = String(cfg.base || "volume").toLowerCase() === "volume";
      const hasValue =
        summary.value !== null &&
        (isVolume ? summary.sumA : summary.totalIncluded > 0);
      value = hasValue ? roundRatio(summary.value, 6) : 1;
    }
  }

  cache.set(key, value);
  return value;
}

export function recalculateUserEntryDependencies() {
  if (summaryRuntime.summaryFormulaEditState?.input && document.body.contains(summaryRuntime.summaryFormulaEditState.input)) {
    return false;
  }
  const model = state.model;
  if (!model || !Array.isArray(model.values) || !Array.isArray(model.mask)) return false;
  const rows = Array.isArray(summaryRowConfigs) ? summaryRowConfigs : [];
  const userRows = rows.filter((cfg) => isUserEntryConfig(cfg));
  if (!userRows.length) return false;

  const devs = getEffectiveDevLabelsForModel(model);
  const colCount = getRatioHeaderLabels(devs).length;
  const lastCol = Math.max(0, devs.length - 1);
  const labelToId = new Map(
    rows.map((cfg) => [String(cfg?.label || cfg?.id || ""), String(cfg?.id || "")]).filter(([k, v]) => k && v)
  );
  let changed = false;

  for (let col = 0; col < colCount; col++) {
    const cache = new Map();
    const visiting = new Set();
    rows.forEach((cfg) => {
      const rowId = String(cfg?.id || "");
      if (!rowId) return;
      computeSummaryRowValueForColumn(model, col, rowId, cache, visiting, labelToId, lastCol);
    });
    userRows.forEach((cfg) => {
      const rowId = String(cfg?.id || "");
      if (!rowId) return;
      const nextValue = sanitizeUserEntryValue(cache.get(rowId));
      const currentValue = sanitizeUserEntryValue(getUserEntryValueForCol(cfg, col));
      const inputRaw = String(getUserEntryInputForCol(cfg, col) || "").trim() || String(currentValue);
      if (Math.abs(nextValue - currentValue) > 1e-12) changed = true;
      setUserEntryCellEntry(rowId, col, inputRaw, nextValue, { persist: false });
    });
  }

  if (changed) persistUserEntryRowsFromState();
  return changed;
}


function wireAvgModal() {
  const modal = getAvgModalEl();
  if (!modal || modal.dataset.wired === "1") return;
  modal.dataset.wired = "1";

  const nameInput = modal.querySelector("#dfmAvgName");
  const typeSelect = modal.querySelector("#dfmAvgType");
  const baseSelect = modal.querySelector("#dfmAvgBase");
  const periodInput = modal.querySelector("#dfmAvgPeriods");
  const excludeInput = modal.querySelector("#dfmAvgExclude");
  const addBtn = modal.querySelector("#dfmAvgAdd");
  const cancelBtn = modal.querySelector("#dfmAvgCancel");

  const syncName = () => {
    if (normalizeAverageType(typeSelect?.value) === "user_entry") return;
    const base = baseSelect?.value || "User Entry";
    const periods = parsePeriodsValue(periodInput?.value);
    const excludeCount = parseExcludeValue(excludeInput?.value);
    if (nameInput) nameInput.value = computeAutoNameWithExclude(base, periods, excludeCount);
  };

  const applyTypeState = () => {
    const isUserEntry = normalizeAverageType(typeSelect?.value) === "user_entry";
    [baseSelect, periodInput, excludeInput].forEach((el) => {
      if (el) el.disabled = isUserEntry;
    });
    [baseSelect, periodInput, excludeInput].forEach((el) => {
      const field = el?.closest?.(".dfmModalField");
      if (field) field.classList.toggle("disabled", isUserEntry);
    });
    if (isUserEntry) {
      if (baseSelect) baseSelect.value = "simple";
      if (periodInput) periodInput.value = "";
      if (excludeInput) excludeInput.value = "None";
      if (nameInput && !String(nameInput.value || "").trim()) nameInput.value = "User Entry";
      return;
    }
    syncName();
  };

  const normalizePeriodsInput = () => {
    if (!periodInput) return;
    const raw = String(periodInput.value || "");
    if (!raw) return;
    if (/^all$/i.test(raw.trim())) {
      periodInput.value = "";
      return;
    }
    const digits = raw.replace(/[^\d]/g, "");
    if (digits !== raw) periodInput.value = digits;
  };

  const applyPeriodDelta = (dir) => {
    if (!periodInput) return;
    const raw = String(periodInput.value || "").trim();
    if (!raw) {
      periodInput.value = "2";
    } else {
      const current = parseInt(raw, 10);
      const base = Number.isFinite(current) ? current : 2;
      const next = Math.max(2, base + dir);
      periodInput.value = String(next);
    }
    syncName();
  };

  const normalizeExcludeInput = () => {
    if (!excludeInput) return;
    const raw = String(excludeInput.value || "").trim();
    if (!raw) return;
    if (/^none$/i.test(raw)) {
      excludeInput.value = "None";
      return;
    }
    const digits = raw.replace(/[^\d]/g, "");
    if (digits !== raw) excludeInput.value = digits;
  };

  nameInput?.addEventListener("input", () => {
    clearModalValidationError(modal, "#dfmAvgName", "#dfmAvgError");
  });
  typeSelect?.addEventListener("change", applyTypeState);
  baseSelect?.addEventListener("change", syncName);
  periodInput?.addEventListener("input", () => {
    normalizePeriodsInput();
    syncName();
  });
  periodInput?.addEventListener("change", () => {
    normalizePeriodsInput();
    syncName();
  });
  excludeInput?.addEventListener("input", () => {
    normalizeExcludeInput();
    syncName();
  });
  excludeInput?.addEventListener("change", () => {
    normalizeExcludeInput();
    syncName();
  });
  periodInput?.addEventListener("wheel", (e) => {
    if (periodInput.disabled) return;
    e.preventDefault();
    const dir = e.deltaY < 0 ? 1 : -1;
    applyPeriodDelta(dir);
  }, { passive: false });

  applyTypeState();

  cancelBtn?.addEventListener("click", () => hideAvgModal());
  modal.querySelector(".dfmModalBackdrop")?.addEventListener("click", () => hideAvgModal());

  addBtn?.addEventListener("click", () => {
    const averageType = normalizeAverageType(typeSelect?.value);
    const isUserEntry = averageType === "user_entry";
    const base = isUserEntry ? "simple" : (baseSelect?.value || "simple").toLowerCase();
    const periods = isUserEntry ? "all" : parsePeriodsValue(periodInput?.value);
    const excludeCount = isUserEntry ? 0 : parseExcludeValue(excludeInput?.value);
    const fallbackName = isUserEntry ? "User Entry" : computeAutoNameWithExclude(base, periods, excludeCount);
    const label = nameInput?.value?.trim() || fallbackName;
    const cfgKey = getSummaryConfigKey();
    if (!cfgKey) {
      hideAvgModal();
      return;
    }
    const customRows = summaryRowConfigs.length
      ? summaryRowConfigs.map((row) => ({ ...row }))
      : BASE_SUMMARY_ROWS.map((row) => ({ ...row }));
    const normalizedLabel = label.trim();
    const nameExists = summaryRowConfigs.some((row) =>
      String(row.label || "").trim().toLowerCase() === normalizedLabel.toLowerCase()
    );
    if (nameExists) {
      setModalValidationError(
        modal,
        "#dfmAvgName",
        "#dfmAvgError",
        "Average formula name already exists."
      );
      return;
    }
    const nextRow = {
      id: `custom_${Date.now()}`,
      label,
      averageType,
      base,
      periods,
      exclude: excludeCount,
    };
    if (isUserEntry) {
      const colCount = getCurrentRatioColumnCount();
      nextRow.values = new Array(Math.max(0, colCount)).fill(1);
      nextRow.inputs = new Array(Math.max(0, colCount)).fill("1");
    }
    customRows.push(nextRow);
    saveCustomSummaryRows(cfgKey, customRows);
    hideAvgModal();
    summaryRuntime._renderRatioTable();
  });
}

registerSummaryFunctions({
  parseUserEntryClipboardGrid,
  parseUserEntryClipboardValue,
  pasteUserEntryClipboardGrid,
  commitUserEntryArrayFormula,
  isSummaryFormulaCommitPending,
  commitSummaryFormulaInput,
  updateSummaryFormulaBarForCell,
  refreshSummaryFormulaBar,
  handleSummaryTableSelectionChange,
  clearSummaryReferenceUi,
  buildSummaryReferenceValues,
  insertAtInputCursor,
  beginSummaryFormulaEditSession,
  cancelSummaryFormulaEditSession,
  setUserEntryCellEntry,
  persistUserEntryRowsFromState,
  computeSummaryRowValueForColumn,
  recalculateUserEntryDependencies,
  wireAvgModal,
});
