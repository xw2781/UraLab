/*
===============================================================================
DFM Ratios Summary Formula Bar
===============================================================================
*/
import { attachArcrhoTooltip } from "/ui/shared/components/tooltip/tooltip.js?v=20260812a";
import { installDfmDatasetAutocomplete } from "/ui/method_pages/dfm/dfm_dataset_autocomplete.js?v=20260814b";
import {
  getCachedDfmDatasetReferenceValues,
  resolveDfmDatasetReferencesInFormulaDetailed,
} from "/ui/method_pages/dfm/dfm_dataset_formula.js?v=20260820a";
import {
  formatFormulaText,
  stripRoundWrappers,
  tokenizeFormula,
} from "/ui/shared/components/formula_bar/formula_text.js?v=20260908a";
import {
  registerSummaryFunctions,
  summaryRuntime,
} from "/ui/method_pages/dfm/ratios_summary/summary_runtime.js?v=20260819a";

const {
  state, calcRatio, roundRatio, formatRatio, computeAverageForColumn,
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

const updateActiveSummaryFormulaReferenceUi = (...args) => summaryRuntime.updateActiveSummaryFormulaReferenceUi(...args);
const formatUserEntryFormulaEvaluationValue = (...args) => (
  summaryRuntime.formatUserEntryFormulaEvaluationValue(...args)
);
const isSummaryFormulaCommitPending = (...args) => summaryRuntime.isSummaryFormulaCommitPending(...args);
const commitSummaryFormulaInput = (...args) => summaryRuntime.commitSummaryFormulaInput(...args);
const updateSummaryFormulaBarForCell = (...args) => summaryRuntime.updateSummaryFormulaBarForCell(...args);
const refreshSummaryFormulaBar = (...args) => summaryRuntime.refreshSummaryFormulaBar(...args);
const beginSummaryFormulaEditSession = (...args) => summaryRuntime.beginSummaryFormulaEditSession(...args);
const cancelSummaryFormulaEditSession = (...args) => summaryRuntime.cancelSummaryFormulaEditSession(...args);

function scrollSummaryFormulaInputToEnd(inputEl) {
  if (!inputEl) return;
  window.requestAnimationFrame(() => {
    try {
      inputEl.scrollLeft = inputEl.scrollWidth;
    } catch (_err) {
      // no-op: some browsers may not expose scroll metrics on detached inputs
    }
  });
}

/** Ask the containing Project Instance to open a dataset explicitly in DSV. */
function openDfmFormulaDataset(datasetName, windowRef = window) {
  const name = String(datasetName || "").trim();
  const parentWindow = windowRef?.parent;
  if (!name || !parentWindow || parentWindow === windowRef) return false;
  parentWindow.postMessage({
    type: "arcrho:project-instance-open-dependent-dataset",
    datasetName: name,
    openMethod: false,
  }, "*");
  return true;
}

// A dataset factor of exactly 1 leaves the User Entry value where it was, so
// its pill is drawn quiet grey: the reference is still live and clickable, it
// just is not moving the number. The tolerance only absorbs binary rounding.
const NEUTRAL_DATASET_FACTOR_TOLERANCE = 1e-9;

function isNeutralDatasetFactor(value) {
  return Number.isFinite(value) && Math.abs(value - 1) <= NEUTRAL_DATASET_FACTOR_TOLERANCE;
}

/**
 * Render colorized formula display in the overlay div.
 * - Excel refs → dark green
 * - Quoted row references → one palette colour each, matching the cells they name
 * - Dataset references → clickable DSV pills, grey when the value is 1
 * - Operators get spaces around them
 * - Always shows leading '='
 */
function renderFormulaBarDisplay(displayEl, rawText, sourceText = rawText) {
  if (!displayEl) return;
  const tokens = tokenizeFormula(stripRoundWrappers(rawText));
  if (!tokens.length) {
    displayEl.textContent = "";
    return;
  }

  // Optional: this module is also loaded standalone, without the model module.
  const referenceColors = summaryRuntime.buildSummaryFormulaReferenceColorsByLabel?.(sourceText)
    || new Map();
  displayEl.innerHTML = "";
  const sourceDatasetTokens = tokenizeFormula(sourceText).filter((token) => token.datasetName);
  // Values come from the session cache the page warms when it opens, keyed by
  // the reference finder rather than by this tokenizer, so they are only
  // trusted when both readings of the formula found the same references.
  const cachedValues = getCachedDfmDatasetReferenceValues(sourceText);
  const datasetValues = cachedValues.length === sourceDatasetTokens.length ? cachedValues : [];
  const unresolvedPills = [];
  let datasetIndex = 0;
  for (const tok of tokens) {
    if (tok.datasetCoordinate) continue;
    if (tok.type === "excel") {
      const span = document.createElement("span");
      span.className = "fmtExcelRef";
      span.textContent = tok.text;
      displayEl.appendChild(span);
    } else if (tok.type === "ref") {
      const span = document.createElement("span");
      const label = tok.text.slice(1, -1);
      span.className = "fmtRowRef";
      const colorClass = referenceColors.get(label.trim().toLowerCase());
      if (colorClass) span.classList.add(colorClass);
      span.textContent = label;
      displayEl.appendChild(span);
    } else if (tok.type === "bracket" && tok.datasetName) {
      const referenceIndex = datasetIndex;
      const sourceToken = sourceDatasetTokens[referenceIndex] || tok;
      datasetIndex += 1;
      const button = document.createElement("button");
      button.type = "button";
      button.className = "fmtDatasetRef";
      const cachedValue = datasetValues[referenceIndex];
      if (isNeutralDatasetFactor(cachedValue)) button.classList.add("isNeutral");
      else if (!Number.isFinite(cachedValue)) unresolvedPills.push({ button, referenceIndex });
      button.textContent = `${tok.datasetName} @ ${tok.datasetCoordinateLabel}`;
      button.dataset.datasetName = tok.datasetName;
      button.dataset.coordinateLabel = tok.datasetCoordinateLabel;
      button.setAttribute(
        "aria-label",
        `Open dataset ${tok.datasetName} at ${tok.datasetCoordinateLabel} in Dataset Viewer`,
      );
      let tooltipValuePromise = null;
      attachArcrhoTooltip(button, async () => {
        if (!tooltipValuePromise) {
          const referenceFormula = `=[${sourceToken.datasetName}][${sourceToken.datasetCoordinateLabel}]`;
          tooltipValuePromise = resolveDfmDatasetReferencesInFormulaDetailed(referenceFormula)
            .then((resolved) => {
              const value = Number(String(resolved?.resolvedFormula || "").replace(/^=\s*/, ""));
              return Number.isFinite(value)
                ? formatUserEntryFormulaEvaluationValue(value)
                : "Value unavailable";
            })
            .catch(() => "Value unavailable");
        }
        return tooltipValuePromise;
      });
      button.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        if (!openDfmFormulaDataset(tok.datasetName)) {
          setStatusBarText(`Could not open dataset ${tok.datasetName}.`);
        }
      });
      displayEl.appendChild(button);
    } else if (tok.type === "op") {
      displayEl.appendChild(document.createTextNode(" " + tok.text + " "));
    } else {
      const t = tok.text.trim();
      if (t) displayEl.appendChild(document.createTextNode(t === "=" ? "= " : t));
    }
  }

  // Nothing is known about a reference until it resolves once, and until then
  // its pill reads as blue. One batched read for the whole formula fills the
  // cache the next render also uses, so this costs a single request in the
  // narrow window before the page-open warm-up lands.
  if (unresolvedPills.length) {
    resolveDfmDatasetReferencesInFormulaDetailed(sourceText).then(() => {
      const resolvedValues = getCachedDfmDatasetReferenceValues(sourceText);
      if (resolvedValues.length !== sourceDatasetTokens.length) return;
      for (const pill of unresolvedPills) {
        pill.button.classList.toggle("isNeutral", isNeutralDatasetFactor(resolvedValues[pill.referenceIndex]));
      }
    }).catch(() => {
      // Best-effort: an unreadable reference simply keeps its blue pill.
    });
  }
}

/** Show/hide display overlay vs input based on focus state. */
function updateFormulaBarDisplayMode(barEl, isEditing) {
  if (!barEl) return;
  const input = barEl.querySelector("#dfmSummaryFormulaBarInput");
  const display = barEl.querySelector("#dfmSummaryFormulaBarDisplay");
  if (!input || !display) return;
  if (isEditing) {
    input.style.display = "";
    display.style.display = "none";
  } else {
    // Format the raw input with proper spacing and leading '='
    const raw = String(input.value || "").trim();
    if (raw) {
      input.value = formatFormulaText(raw);
    }
    input.style.display = "none";
    display.style.display = "";
    renderFormulaBarDisplay(display, input.dataset.displayFormula || input.value, input.value);
  }
  // The two modes need different widths, so the bar is re-measured on every swap.
  // Optional: this module is also loaded standalone, without the anchor module.
  summaryRuntime.repositionSummaryFormulaBar?.(barEl);
}

function positionSummaryFormulaBarValidationTooltip() {
  const { bar, input, display, error } = getSummaryFormulaBarParts();
  if (!bar || !error || error.hidden) return;

  error.style.visibility = "hidden";
  const host = bar.closest?.("#ratioWrapHost") || document.getElementById("ratioWrapHost");
  const ratiosPage = document.getElementById("dfmRatiosPage");
  if (
    !host
    || !bar.isConnected
    || !bar.classList.contains("isOpen")
    || ratiosPage?.getClientRects?.().length === 0
  ) return;

  const popout = bar.closest?.(".tabPopoutWindow");
  const computedPopoutZ = popout ? window.getComputedStyle?.(popout)?.zIndex : "";
  const popoutZ = Number.parseInt(
    popout?.style?.zIndex || computedPopoutZ || "",
    10,
  );
  const tooltipZ = Number.isFinite(popoutZ)
    ? Math.min(summaryRuntime.SUMMARY_FORMULA_BAR_TOOLTIP_MAX_Z_INDEX, popoutZ + 1)
    : summaryRuntime.SUMMARY_FORMULA_BAR_TOOLTIP_Z_INDEX;
  error.style.zIndex = String(tooltipZ);

  const barRect = bar.getBoundingClientRect();
  const anchorEl = input?.getClientRects?.().length ? input : display;
  const anchorRect = anchorEl?.getBoundingClientRect?.() || barRect;
  const hostRect = host.getBoundingClientRect();
  const viewportWidth = Math.max(0, Number(window.innerWidth || document.documentElement?.clientWidth || 0));
  const viewportHeight = Math.max(0, Number(window.innerHeight || document.documentElement?.clientHeight || 0));
  const layoutInput = { barRect, anchorRect, hostRect, viewportWidth, viewportHeight };
  const widthLayout = computeFormulaValidationTooltipLayout({
    ...layoutInput,
    tooltipRect: { width: 0, height: 0 },
  });
  error.style.maxWidth = `${widthLayout.maxWidth}px`;

  const layout = computeFormulaValidationTooltipLayout({
    ...layoutInput,
    tooltipRect: error.getBoundingClientRect(),
  });
  error.style.left = `${Math.round(layout.left)}px`;
  error.style.top = `${Math.round(layout.top)}px`;
  error.style.setProperty("--dfm-summary-formula-tooltip-arrow-x", `${Math.round(layout.arrowX)}px`);
  error.dataset.placement = layout.placement;
  error.style.visibility = layout.visible ? "visible" : "hidden";
}

function scheduleSummaryFormulaBarValidationTooltipPosition() {
  const error = document.getElementById("dfmSummaryFormulaBarError");
  if (!error || error.hidden || summaryRuntime.formulaBarTooltipRaf) return;
  summaryRuntime.formulaBarTooltipRaf = window.requestAnimationFrame(() => {
    summaryRuntime.formulaBarTooltipRaf = 0;
    positionSummaryFormulaBarValidationTooltip();
  });
}

function scheduleSummaryFormulaBarResizeRefresh() {
  if (summaryRuntime.formulaBarResizeRaf) return;
  summaryRuntime.formulaBarResizeRaf = window.requestAnimationFrame(() => {
    summaryRuntime.formulaBarResizeRaf = 0;
    refreshSummaryFormulaBar();
    scheduleSummaryFormulaBarValidationTooltipPosition();
  });
}

// A resize can also be a zoom change, so the measured text width is re-taken;
// a scroll leaves it valid and keeps the cheap path.
function handleSummaryFormulaBarViewportResize() {
  summaryRuntime.invalidateSummaryFormulaBarWidthCache();
  scheduleSummaryFormulaBarResizeRefresh();
}

function wireSummaryFormulaBarResizeWatcher(summaryTable) {
  const host = summaryTable?.closest?.("#ratioWrapHost") || document.getElementById("ratioWrapHost");
  if (summaryRuntime.formulaBarScrollHost && summaryRuntime.formulaBarScrollHost !== host) {
    summaryRuntime.formulaBarScrollHost.removeEventListener("scroll", scheduleSummaryFormulaBarResizeRefresh);
    summaryRuntime.formulaBarScrollHost = null;
  }
  if (host && summaryRuntime.formulaBarScrollHost !== host) {
    host.addEventListener("scroll", scheduleSummaryFormulaBarResizeRefresh, { passive: true });
    summaryRuntime.formulaBarScrollHost = host;
  }
  if (host && window.ResizeObserver) {
    if (summaryRuntime.formulaBarResizeObserver?.target !== host) {
      summaryRuntime.formulaBarResizeObserver?.observer?.disconnect?.();
      const observer = new ResizeObserver(handleSummaryFormulaBarViewportResize);
      observer.observe(host);
      summaryRuntime.formulaBarResizeObserver = { observer, target: host };
    }
  }
  if (!summaryRuntime.formulaBarResizeWired) {
    summaryRuntime.formulaBarResizeWired = true;
    window.addEventListener("resize", handleSummaryFormulaBarViewportResize);
    window.addEventListener(
      "pointerdown",
      scheduleSummaryFormulaBarValidationTooltipPosition,
      { capture: true, passive: true },
    );
  }
}

function getSummaryFormulaBarParts(barEl = null) {
  const bar = barEl || document.getElementById("dfmSummaryFormulaBar");
  return {
    bar,
    input: bar?.querySelector?.("#dfmSummaryFormulaBarInput") || null,
    display: bar?.querySelector?.("#dfmSummaryFormulaBarDisplay") || null,
    error: bar?.querySelector?.("#dfmSummaryFormulaBarError")
      || document.getElementById("dfmSummaryFormulaBarError")
      || null,
    state: bar?.querySelector?.("#dfmSummaryFormulaBarState") || null,
  };
}

function clearSummaryFormulaBarValidationError() {
  const { bar, input, error } = getSummaryFormulaBarParts();
  if (summaryRuntime.formulaValidationErrorInput && summaryRuntime.formulaValidationErrorInput !== input) {
    clearFormulaValidationError({ inputEl: summaryRuntime.formulaValidationErrorInput, errorEl: error });
  }
  clearFormulaValidationError({
    barEl: bar,
    inputEl: summaryRuntime.formulaValidationErrorInput || input,
    errorEl: error,
  });
  summaryRuntime.formulaValidationErrorInput = null;
}

function showSummaryFormulaBarValidationError(message, inputEl = null) {
  const { bar, input, error } = getSummaryFormulaBarParts();
  const targetInput = inputEl || input;
  if (summaryRuntime.formulaValidationErrorInput && summaryRuntime.formulaValidationErrorInput !== targetInput) {
    clearFormulaValidationError({ inputEl: summaryRuntime.formulaValidationErrorInput, errorEl: error });
  }
  const text = showFormulaValidationError({
    barEl: bar,
    inputEl: targetInput,
    errorEl: error,
    message,
  });
  summaryRuntime.formulaValidationErrorInput = targetInput;
  positionSummaryFormulaBarValidationTooltip();
  scheduleSummaryFormulaBarValidationTooltipPosition();
  return text;
}

function cancelFormulaBarDisplayRefresh() {
  if (!summaryRuntime.summaryFormulaBarDisplayRaf) return;
  window.cancelAnimationFrame(summaryRuntime.summaryFormulaBarDisplayRaf);
  summaryRuntime.summaryFormulaBarDisplayRaf = 0;
}

function clearFormulaBarFocusRestoreHandler() {
  if (!summaryRuntime.summaryFormulaBarFocusRestoreHandler) return;
  window.removeEventListener("focus", summaryRuntime.summaryFormulaBarFocusRestoreHandler);
  summaryRuntime.summaryFormulaBarFocusRestoreHandler = null;
}

function isSummaryFormulaBarInputEditing(inputEl) {
  return !!(
    inputEl &&
    inputEl.isConnected &&
    summaryRuntime.summaryFormulaBarState.input === inputEl &&
    summaryRuntime.summaryFormulaBarState.mode !== "display"
  );
}

function setSummaryFormulaBarMode(mode, inputEl = null) {
  const nextMode = mode === "validating" ? "validating" : (mode === "editing" ? "editing" : "display");
  const currentInput = inputEl || getSummaryFormulaBarParts().input;
  summaryRuntime.summaryFormulaBarState = {
    mode: nextMode,
    input: nextMode === "display" ? null : currentInput,
    generation: summaryRuntime.summaryFormulaBarState.generation + 1,
  };
  const { bar, state } = getSummaryFormulaBarParts(currentInput?.closest?.(".dfmSummaryFormulaBar"));
  bar?.classList?.toggle("isValidating", nextMode === "validating");
  if (state) {
    state.hidden = nextMode !== "validating";
    state.textContent = nextMode === "validating" ? "Validating…" : "";
    // The chip takes room of its own: make space for it rather than squeezing
    // the formula while a commit is in flight.
    summaryRuntime.repositionSummaryFormulaBar?.(bar);
  }
}

function scheduleFormulaBarDisplayMode(barEl, inputEl) {
  cancelFormulaBarDisplayRefresh();
  const generation = summaryRuntime.summaryFormulaBarState.generation;
  summaryRuntime.summaryFormulaBarDisplayRaf = window.requestAnimationFrame(() => {
    summaryRuntime.summaryFormulaBarDisplayRaf = 0;
    if (generation !== summaryRuntime.summaryFormulaBarState.generation) return;
    const { bar, input } = getSummaryFormulaBarParts(barEl);
    if (!bar || !input || input !== inputEl || !input.isConnected) return;
    updateFormulaBarDisplayMode(bar, isSummaryFormulaBarInputEditing(input));
  });
}

function captureFormulaInputSelection(inputEl) {
  const valueLength = String(inputEl?.value || "").length;
  const start = Number.isInteger(inputEl?.selectionStart) ? inputEl.selectionStart : valueLength;
  const end = Number.isInteger(inputEl?.selectionEnd) ? inputEl.selectionEnd : start;
  return {
    selectionStart: Math.max(2, start),
    selectionEnd: Math.max(2, end),
  };
}

function restoreFormulaBarEditingAfterValidation(barEl, inputEl, selection = {}) {
  cancelFormulaBarDisplayRefresh();
  clearFormulaBarFocusRestoreHandler();
  const { bar, input, display } = getSummaryFormulaBarParts(barEl);
  if (!bar || !input || input !== inputEl || !input.isConnected) return;
  setSummaryFormulaBarMode("editing", input);
  updateFormulaBarDisplayMode(bar, true);

  const restore = () => {
    summaryRuntime.summaryFormulaBarFocusRestoreHandler = null;
    if (!isSummaryFormulaBarInputEditing(input) || !input.isConnected) return;
    revealAndFocusFormulaInput({
      inputEl: input,
      displayEl: display,
      selectionStart: selection.selectionStart,
      selectionEnd: selection.selectionEnd,
    });
  };

  if (document.hasFocus()) {
    window.requestAnimationFrame(restore);
  } else {
    summaryRuntime.summaryFormulaBarFocusRestoreHandler = restore;
    window.addEventListener("focus", restore, { once: true });
  }
}

function cancelActiveSummaryFormulaCommit() {
  summaryRuntime.summaryFormulaCommitGeneration += 1;
  const lease = summaryRuntime.summaryFormulaCommitLease;
  lease?.cancel?.();
  summaryRuntime.summaryFormulaCommitLease = null;
}

function ensureSummaryFormulaBarValidationTooltip() {
  let error = document.getElementById("dfmSummaryFormulaBarError");
  if (!error) {
    error = document.createElement("div");
    error.id = "dfmSummaryFormulaBarError";
    error.className = "dfmSummaryFormulaBarError";
    error.setAttribute("role", "alert");
    error.setAttribute("aria-live", "assertive");
    error.setAttribute("aria-atomic", "true");
    error.hidden = true;
  }
  if (document.body && error.parentElement !== document.body) {
    document.body.appendChild(error);
  }
  return error;
}

function ensureSummaryFormulaBarEl(summaryTable) {
  ensureSummaryFormulaBarValidationTooltip();
  let el = document.getElementById("dfmSummaryFormulaBar");
  if (!el) {
    el = document.createElement("div");
    el.id = "dfmSummaryFormulaBar";
    el.className = "arFormulaBar dfmSummaryFormulaBar";
    const fxIcon = document.createElement("span");
    fxIcon.className = "arFormulaBarFxIcon";
    fxIcon.textContent = "fx";
    const label = document.createElement("span");
    label.id = "dfmSummaryFormulaBarLabelText";
    label.className = "dfmSummaryFormulaBarLabel";
    label.textContent = "f(x)";
    const input = document.createElement("input");
    input.id = "dfmSummaryFormulaBarInput";
    input.className = "arFormulaBarInput dfmSummaryFormulaBarInput";
    input.type = "text";
    input.autocomplete = "off";
    input.spellcheck = false;
    const display = document.createElement("div");
    display.id = "dfmSummaryFormulaBarDisplay";
    display.className = "arFormulaBarDisplay dfmSummaryFormulaBarDisplay";
    const validationState = document.createElement("span");
    validationState.id = "dfmSummaryFormulaBarState";
    validationState.className = "dfmSummaryFormulaBarState";
    validationState.setAttribute("aria-live", "polite");
    validationState.hidden = true;
    el.appendChild(fxIcon);
    el.appendChild(label);
    el.appendChild(input);
    el.appendChild(display);
    el.appendChild(validationState);
  }
  if (el.dataset.wired !== "1") {
    const input = el.querySelector("#dfmSummaryFormulaBarInput");
    installDfmDatasetAutocomplete(input);
    // The badge is the bar's drag handle; it carries no tooltip of its own so a
    // bubble cannot sit under the pointer that is about to move the bar.
    summaryRuntime.wireSummaryFormulaBarDragHandle?.(el, el.querySelector(".arFormulaBarFxIcon"));
    const FORMULA_PREFIX = "= ";
    const PREFIX_LEN = FORMULA_PREFIX.length; // 2
    input?.addEventListener("focus", () => {
      setSummaryFormulaBarMode("editing", input);
      updateFormulaBarDisplayMode(el, true);
      // Ensure leading "= " prefix is present
      if (!input.value.startsWith(FORMULA_PREFIX)) {
        const body = input.value.replace(/^=\s*/, "");
        input.value = FORMULA_PREFIX + body;
      }
      const summaryTableEl = document.querySelector("#ratioWrap table.ratioSummaryTable");
      const rowId = String(input.dataset.rowId || "");
      const col = Number(input.dataset.col);
      if (!summaryTableEl || !rowId || !Number.isFinite(col) || col < 0) return;
      const cell = summaryTableEl.querySelector(`td.summaryCell[data-r="${rowId}"][data-col="${col}"]`);
      if (!cell) return;
      beginSummaryFormulaEditSession(summaryTableEl, cell, input, col);
      updateActiveSummaryFormulaReferenceUi(summaryTableEl);
      scrollSummaryFormulaInputToEnd(input);
    });
    // Prevent cursor from moving before the prefix
    input?.addEventListener("click", () => {
      if (input.selectionStart < PREFIX_LEN) input.setSelectionRange(PREFIX_LEN, PREFIX_LEN);
    });
    input?.addEventListener("input", () => {
      delete input.dataset.skipFormulaBlurCommit;
      setSummaryFormulaBarMode("editing", input);
      clearSummaryFormulaBarValidationError();
      // Keep the leading "= " undeletable
      if (!input.value.startsWith(FORMULA_PREFIX)) {
        const cleaned = input.value.replace(/^=\s*/, "");
        input.value = FORMULA_PREFIX + cleaned;
        input.setSelectionRange(PREFIX_LEN, PREFIX_LEN);
      }
      const normalizedReference = normalizeExcelReferenceAddressCase(input.value);
      if (normalizedReference !== input.value) {
        const selectionStart = input.selectionStart;
        const selectionEnd = input.selectionEnd;
        input.value = normalizedReference;
        if (Number.isInteger(selectionStart) && Number.isInteger(selectionEnd)) {
          input.setSelectionRange(selectionStart, selectionEnd);
        }
      }
      const summaryTableEl = document.querySelector("#ratioWrap table.ratioSummaryTable");
      const rowId = String(input.dataset.rowId || "");
      const col = Number(input.dataset.col);
      if (summaryTableEl && rowId && Number.isFinite(col) && col >= 0) {
        const cell = summaryTableEl.querySelector(`td.summaryCell[data-r="${rowId}"][data-col="${col}"]`);
        if (cell) {
          beginSummaryFormulaEditSession(summaryTableEl, cell, input, col);
          updateSummaryFormulaBarForCell(cell);
          updateActiveSummaryFormulaReferenceUi(summaryTableEl);
        }
      }
    });
    input?.addEventListener("keydown", async (e) => {
      // Prevent deleting the leading "= " prefix
      if (e.key === "Backspace" && input.selectionStart <= PREFIX_LEN && input.selectionEnd <= PREFIX_LEN) {
        e.preventDefault();
        return;
      }
      if (e.key === "Delete" && input.selectionStart < PREFIX_LEN && input.selectionEnd <= PREFIX_LEN) {
        e.preventDefault();
        return;
      }
      // Prevent selecting/replacing the prefix via Home or Ctrl+A
      if (e.key === "Home") {
        e.preventDefault();
        input.setSelectionRange(PREFIX_LEN, e.shiftKey ? input.selectionEnd : PREFIX_LEN);
        return;
      }
      if (e.key === "ArrowLeft" && input.selectionStart <= PREFIX_LEN && !e.shiftKey) {
        e.preventDefault();
        return;
      }
      if (e.key === "a" && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        input.setSelectionRange(PREFIX_LEN, input.value.length);
        return;
      }
      if (e.key === "Enter") {
        e.preventDefault();
        if (isSummaryFormulaCommitPending(input)) return;
        const selection = captureFormulaInputSelection(input);
        setSummaryFormulaBarMode("validating", input);
        const validationStateGeneration = summaryRuntime.summaryFormulaBarState.generation;
        const ok = await commitSummaryFormulaInput(input);
        if (
          summaryRuntime.summaryFormulaBarState.generation !== validationStateGeneration ||
          summaryRuntime.summaryFormulaBarState.input !== input ||
          summaryRuntime.summaryFormulaBarState.mode !== "validating"
        ) return;
        if (ok) {
          setSummaryFormulaBarMode("display", input);
          if (document.activeElement === input) {
            input.dataset.skipFormulaBlurCommit = "1";
            input.blur();
          } else {
            scheduleFormulaBarDisplayMode(el, input);
          }
        } else {
          restoreFormulaBarEditingAfterValidation(el, input, selection);
        }
      } else if (e.key === "Escape") {
        e.preventDefault();
        cancelActiveSummaryFormulaCommit();
        cancelSummaryFormulaEditSession();
        clearSummaryFormulaBarValidationError();
        setSummaryFormulaBarMode("display", input);
        input.dataset.skipFormulaBlurCommit = "1";
        input.blur();
      }
    });
    input?.addEventListener("blur", async () => {
      if (input.dataset.skipFormulaBlurCommit === "1") {
        delete input.dataset.skipFormulaBlurCommit;
        scheduleFormulaBarDisplayMode(el, input);
        return;
      }
      if (isSummaryFormulaCommitPending(input)) {
        scheduleFormulaBarDisplayMode(el, input);
        return;
      }
      const selection = captureFormulaInputSelection(input);
      setSummaryFormulaBarMode("validating", input);
      const validationStateGeneration = summaryRuntime.summaryFormulaBarState.generation;
      const ok = await commitSummaryFormulaInput(input);
      if (
        summaryRuntime.summaryFormulaBarState.generation !== validationStateGeneration ||
        summaryRuntime.summaryFormulaBarState.input !== input ||
        summaryRuntime.summaryFormulaBarState.mode !== "validating"
      ) return;
      if (!ok) {
        restoreFormulaBarEditingAfterValidation(el, input, selection);
        return;
      }
      setSummaryFormulaBarMode("display", input);
      scheduleFormulaBarDisplayMode(el, input);
    });
    const displayDiv = el.querySelector("#dfmSummaryFormulaBarDisplay");
    displayDiv?.addEventListener("click", () => {
      if (input && !input.disabled && !input.readOnly && !isSummaryFormulaCommitPending(input)) {
        setSummaryFormulaBarMode("editing", input);
        updateFormulaBarDisplayMode(el, true);
        input.focus({ preventScroll: true });
      }
    });
    el.dataset.wired = "1";
  }
  // The bar floats over the grid, so it lives in the scrolling host rather than
  // in the table's flow: absolute children of the host scroll with the tables.
  const host = summaryTable?.closest?.("#ratioWrapHost") || document.getElementById("ratioWrapHost");
  if (host && el.parentElement !== host) {
    host.appendChild(el);
  }
  wireSummaryFormulaBarResizeWatcher(summaryTable);
  return el;
}

function setStatusBarText(text) {
  // Status bar lives in the parent document (DFM runs in an iframe)
  const doc = window.parent?.document || document;
  const el = doc.getElementById("statusText") || doc.getElementById("statusBar");
  if (el) el.textContent = text || "";
}

registerSummaryFunctions({
  scrollSummaryFormulaInputToEnd,
  tokenizeFormula,
  formatFormulaText,
  stripRoundWrappers,
  openDfmFormulaDataset,
  renderFormulaBarDisplay,
  updateFormulaBarDisplayMode,
  positionSummaryFormulaBarValidationTooltip,
  scheduleSummaryFormulaBarValidationTooltipPosition,
  scheduleSummaryFormulaBarResizeRefresh,
  handleSummaryFormulaBarViewportResize,
  wireSummaryFormulaBarResizeWatcher,
  getSummaryFormulaBarParts,
  clearSummaryFormulaBarValidationError,
  showSummaryFormulaBarValidationError,
  cancelFormulaBarDisplayRefresh,
  clearFormulaBarFocusRestoreHandler,
  isSummaryFormulaBarInputEditing,
  setSummaryFormulaBarMode,
  scheduleFormulaBarDisplayMode,
  captureFormulaInputSelection,
  restoreFormulaBarEditingAfterValidation,
  cancelActiveSummaryFormulaCommit,
  ensureSummaryFormulaBarValidationTooltip,
  ensureSummaryFormulaBarEl,
  setStatusBarText,
});
