import {
  createSpreadsheetTableController,
  getTopLeftRangeCell,
  normalizeRange,
} from "/ui/shared/components/spreadsheet/spreadsheet_table.js?v=20260715a";
import {
  getDatasetGridSelectionLayout,
  getDisplayDatasetModel,
  setDatasetGridEditConfig,
} from "/ui/shared/tabs/data/dataset_grid_view.js?v=20260907c";
import { parseExcelReference } from "/ui/shared/integrations/excel_reference.js?v=20260715a";
import { createFormulaHoverEditor } from "/ui/shared/components/formula_hover/formula_hover.js?v=20260907a";
import {
  buildInternalDatasetReferenceText,
  insertPickedDatasetReference,
  isInternalReferencePickDraft,
} from "/ui/shared/dataset/dataset_internal_reference.js?v=20260830a";
import { classifyDatasetFormula } from "/ui/shared/dataset/dataset_formula.js?v=20260830a";
import { showPageMessageBox } from "/ui/shared/components/message_box/message_box.js?v=20260831a";

export function wireDatasetGridInteractions(deps) {
  const {
    state,
    renderTable,
    isReadOnly = () => false,
    readOnlyMessage = () => "Generated datasets are read-only.",
    showReadOnlyNotice = (message) => showPageMessageBox({
      title: "Read-only view",
      message,
      tone: "warn",
    }),
    setStatus = () => {},
    notifyDatasetUpdated = () => {},
    refreshDatasetSettingsDirty = () => {},
    commitExternalReference = async () => ({ handled: false, ok: false }),
    commitInternalReference = async () => ({ handled: false, ok: false }),
    commitFormulaReference = async () => ({ handled: false, ok: false }),
    cancelExternalReference = () => {},
    hardCodeExternalLinkCells = () => 0,
    decorateExternalLinkCell = () => {},
    getExternalLinkCellInfo = () => null,
    beginReferencePick = () => {},
    endReferencePick = () => {},
    publishReferencePick = () => {},
  } = deps;

  // Digits typed against a multi-cell selection accumulate here until the selection changes.
  let rangeFillSession = null;
  // Cross-window reference pick: armed while an edit in this window holds a
  // formula draft that can accept a [Dataset][rows] reference picked from
  // another open Dataset window. The owner says which editor is waiting for it
  // — a cell being typed into, or the floating formula bar — so a picked range
  // goes back to the right one. The gesture flag tracks a pick drag in the
  // window doing the picking.
  let referencePickArmed = false;
  let referencePickOwner = "";
  let referencePickGesture = false;
  // Whether this window has offered a range yet. Until it has, the selection it
  // happened to be carrying is not part of anyone's formula and is left alone.
  let referencePickOffered = false;
  // Whether the grid is currently wearing the pick treatment, so the sweep that
  // applies it can be skipped entirely on the ordinary selection changes that
  // have nothing to do with a pick.
  let referencePickDecorated = false;
  // Whether a refusal is already on screen. A locked grid refuses every
  // keystroke of a typed number, and the reader has to be told only once.
  let readOnlyNoticeOpen = false;

  // Why an edit was refused belongs in the window the reader is looking at:
  // the status line lives in the shell, below every page, where a notice about
  // this grid is easily missed.
  function reportReadOnlyRefusal() {
    const message = readOnlyMessage();
    if (readOnlyNoticeOpen) return message;
    readOnlyNoticeOpen = true;
    try {
      void Promise.resolve(showReadOnlyNotice(message))
        .catch(() => setStatus(message))
        .finally(() => {
          readOnlyNoticeOpen = false;
        });
    } catch {
      readOnlyNoticeOpen = false;
      setStatus(message);
    }
    return message;
  }

  const formulaHover = createFormulaHoverEditor({
    onCommit: commitHoveredExternalFormula,
    onDismiss: () => document.getElementById("keySink")?.focus?.({ preventScroll: true }),
    onEditStart: cancelExternalReference,
    onStatus: setStatus,
    onDraftChange: (value) => syncReferencePickSession(value, "bar"),
    onClosed: () => {
      if (referencePickOwner === "bar") stopReferencePickSession();
    },
    // Picking cells in another window takes focus out of this one; the formula
    // bar has to survive that or there is nothing to pick into.
    shouldStayOpenUnfocused: () => referencePickArmed && referencePickOwner === "bar",
  });

  const spreadsheetTable = createSpreadsheetTableController({
    getRoot: () => document.getElementById("tableWrap"),
    getBounds: () => {
      const { maxRow, maxCol } = getDatasetGridSelectionLayout();
      return { maxRow, maxCol };
    },
    readSelection: () => ({
      ranges: state.selRanges || [],
      activeCell: state.activeCell,
      anchorCell: state.selectionAnchor,
    }),
    writeSelection: ({ ranges, activeCell, anchorCell }) => {
      state.selRanges = ranges;
      state.activeCell = activeCell;
      state.selectionAnchor = anchorCell;
    },
    onAfterWrite: () => {
      resetRangeFillSession();
      applyReferencePickDecoration();
    },
    cellSelector: "td[data-r][data-c]",
    rowHeaderSelector: "th.rowhdr[data-r]",
    columnHeaderSelector: "th.colhdr[data-c]",
    selectedClasses: ["sel"],
    activeClasses: ["active"],
    anchorClasses: ["selectionAnchor", "arSpreadsheetSelectionAnchor"],
    rowSelectedLabelClasses: ["activeRow", "arSpreadsheetSelectedLabel"],
    columnSelectedLabelClasses: ["activeCol", "arSpreadsheetSelectedLabel"],
    getCellValue: ({ r, c }, cell) => (
      cell?.dataset?.copyValue ?? getDisplayDatasetModel()?.values?.[r]?.[c] ?? ""
    ),
    lineSeparator: "\n",
    scrollCellIntoView: scrollDatasetCellIntoView,
  });

  setDatasetGridEditConfig({
    isEditableCell: (displayR, displayC) => !!canEditDisplayCell(displayR, displayC, { silent: true }),
    isEditingCell: (displayR, displayC) => state.editingCell?.r === displayR && state.editingCell?.c === displayC,
    onCellFocus: (displayR, displayC) => {
      formulaHover.hide?.();
      cancelExternalReference();
      state.activeCell = { r: displayR, c: displayC };
      applySelectionFromState();
    },
    onCellInput: (displayR, displayC, rawValue, input, td) => {
      if (isExternalReferenceDraft(rawValue)) {
        if (state.editingCell) state.editingCell.pendingExternalReference = String(rawValue || "");
        syncReferencePickSession(rawValue);
        return;
      }
      if (state.editingCell) delete state.editingCell.pendingExternalReference;
      syncReferencePickSession(rawValue);
      const nextValue = setDisplayCellValue(displayR, displayC, rawValue, { silentInvalid: true });
      syncInputCellDisplay(td, input, nextValue);
    },
    onCellPaste: (displayR, displayC, event) => {
      const data = event.clipboardData?.getData("text/plain") || "";
      if (!data.includes("\t") && !data.includes("\n") && !data.includes("\r")) return;
      event.preventDefault();
      applyPastedGridText(data, { r: displayR, c: displayC });
    },
    onCellCommit: async (displayR, displayC, rawValue, input, td) => {
      const edit = state.editingCell;
      if (edit?.r !== displayR || edit?.c !== displayC || edit.commitPending) return;
      // While a cross-window pick is armed, focus leaving this window is part
      // of picking cells elsewhere, never a commit; Enter or an in-window
      // blur still commits because the document keeps focus for those.
      if (referencePickArmed && document.hasFocus?.() === false) return;
      if (isExternalReferenceDraft(rawValue)) {
        edit.commitPending = true;
        input.readOnly = true;
        input.setAttribute("aria-busy", "true");
        const result = await commitReferenceDraft({
          displayRow: displayR,
          displayColumn: displayC,
          reference: rawValue,
        });
        if (state.editingCell !== edit) return;
        edit.commitPending = false;
        input.readOnly = false;
        input.removeAttribute("aria-busy");
        if (!result?.ok) {
          if (!result?.aborted && !result?.stale) {
            setStatus(result?.error || "The linked values could not be loaded.");
            requestAnimationFrame(() => input.isConnected && input.focus({ preventScroll: true }));
          }
          return;
        }
        stopReferencePickSession();
        state.editingCell = null;
        renderTable();
        notifyDatasetUpdated();
        applySelectionFromState();
        setStatus(result.message || linkedCellsMessage(result));
        return;
      }
      stopReferencePickSession();
      const nextValue = setDisplayCellValue(displayR, displayC, rawValue, { hardCodeLinks: true });
      syncInputCellDisplay(td, input, nextValue);
      state.editingCell = null;
      renderTable();
      notifyDatasetUpdated();
      applySelectionFromState();
    },
    onCellKeyDown: (displayR, displayC, event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        event.stopPropagation();
        event.currentTarget?.blur?.();
        return;
      }
      if (event.key !== "Escape") return;
      event.preventDefault();
      event.stopPropagation();
      cancelCellEdit(displayR, displayC);
    },
    onCellContextMenu: (displayR, displayC) => prepareContextSelection(displayR, displayC),
    canPasteSelection: () => hasEditableSelectionTarget(),
    canClearData: () => !isReadOnly() && !!getDisplayDatasetModel(),
    onContextAction: (action) => handleGridContextAction(action),
    onTableRendered: () => {
      formulaHover.hide?.();
      applySelectionFromState();
    },
    decorateCell: (cell, displayR, displayC) => {
      decorateExternalLinkCell(cell, displayR, displayC);
      const info = getExternalLinkCellInfo(displayR, displayC);
      // A notice stands in for the formula on every cell of a linked dataset
      // that is being shown at a coarser period than it is stored at. It sits
      // on the cell under the pointer, since no cell here is the one the link
      // names, and it is keyed by that cell so the bar still behaves as a
      // per-cell control.
      if (info?.note) {
        formulaHover.attach(cell, info, { key: `note:${displayR},${displayC}` });
        return;
      }
      if (!info?.reference) return;
      formulaHover.attach(cell, {
        ...info,
        formula: info.reference,
        readOnly: isReadOnly(),
      }, {
        resolveAnchor: () => resolveExternalFormulaAnchor(info, cell),
        positionRect: () => resolveExternalFormulaRangeRect(info, cell),
      });
    },
  });
  wireArrowKeyNavigation();
  wireRectSelectionAndCopy();

  function sameExternalFormulaRange(left, right) {
    return !!(
      left?.reference
      && right?.reference === left.reference
      && right.anchorDisplayRow === left.anchorDisplayRow
      && right.anchorDisplayColumn === left.anchorDisplayColumn
    );
  }

  function externalFormulaRangeCells(info) {
    return Array.from(document.querySelectorAll?.("#tableWrap td[data-r][data-c]") || []).filter((cell) => {
      const row = Number(cell.dataset?.r);
      const column = Number(cell.dataset?.c);
      return Number.isInteger(row)
        && Number.isInteger(column)
        && sameExternalFormulaRange(info, getExternalLinkCellInfo(row, column));
    });
  }

  function resolveExternalFormulaAnchor(info, fallbackCell = null) {
    const selector = `#tableWrap td[data-r="${info.anchorDisplayRow}"][data-c="${info.anchorDisplayColumn}"]`;
    return document.querySelector(selector) || fallbackCell;
  }

  function resolveExternalFormulaRangeRect(info, fallbackCell = null) {
    const cells = externalFormulaRangeCells(info);
    if (!cells.length && fallbackCell) cells.push(fallbackCell);
    const rects = cells
      .map((cell) => cell.getBoundingClientRect?.())
      .filter(Boolean);
    if (!rects.length) return null;
    const left = Math.min(...rects.map((rect) => rect.left));
    const top = Math.min(...rects.map((rect) => rect.top));
    const right = Math.max(...rects.map((rect) => rect.right));
    const bottom = Math.max(...rects.map((rect) => rect.bottom));
    return { left, top, right, bottom, width: right - left, height: bottom - top };
  }

  function wireArrowKeyNavigation() {
    if (window.__arcRhoArrowNavWired) return;
    window.__arcRhoArrowNavWired = true;

    document.addEventListener("keydown", (e) => {
      const delta = {
        ArrowUp: [-1, 0],
        ArrowDown: [1, 0],
        ArrowLeft: [0, -1],
        ArrowRight: [0, 1],
      }[e.key];
      if (!delta || isTypingTarget(e.target) || !state.activeCell) return;
      if (spreadsheetTable.move(delta[0], delta[1], {
        extend: e.shiftKey,
        jump: e.ctrlKey || e.metaKey,
      })) {
        e.preventDefault();
      }
    });
  }

  function scrollDatasetCellIntoView({ r, c }) {
    const td = document.querySelector(`#tableWrap td[data-r="${r}"][data-c="${c}"]`);
    const wrap = document.getElementById("tableWrap");
    if (!td || !wrap) return;
    const tdRect = td.getBoundingClientRect();
    const wrapRect = wrap.getBoundingClientRect();
    const stickyLeft = wrap.querySelector("tbody th, tbody td:first-child")?.getBoundingClientRect().width || 0;
    const stickyTop = wrap.querySelector("thead th")?.getBoundingClientRect().height || 0;
    const leftDelta = tdRect.left - (wrapRect.left + stickyLeft);
    const rightDelta = tdRect.right - wrapRect.right;
    const topDelta = tdRect.top - (wrapRect.top + stickyTop);
    const bottomDelta = tdRect.bottom - wrapRect.bottom;
    if (leftDelta < 0) wrap.scrollLeft += leftDelta;
    else if (rightDelta > 0) wrap.scrollLeft += rightDelta;
    if (topDelta < 0) wrap.scrollTop += topDelta;
    else if (bottomDelta > 0) wrap.scrollTop += bottomDelta;
  }

  function rcFromTd(td) {
    const r = Number(td?.dataset?.r);
    const c = Number(td?.dataset?.c);
    if (!Number.isInteger(r) || !Number.isInteger(c)) return null;
    return { r, c };
  }

  function isTypingTarget(t) {
    if (!t) return false;
    return !!(
      t.closest
        ? t.closest("input, textarea, select, option, button, [contenteditable='true']")
        : (t.matches && t.matches("input, textarea, select, option, button, [contenteditable='true']"))
    ) || !!t.isContentEditable;
  }

  function displayToActualCell(displayR, displayC) {
    return document.getElementById("transposedChk")?.checked === true
      ? { r: displayC, c: displayR }
      : { r: displayR, c: displayC };
  }

  function parseEditableCellValue(rawInput) {
    const raw = String(rawInput ?? "").trim().replace(/,/g, "");
    if (raw === "") return { ok: true, value: null };
    let value = null;
    if (raw.endsWith("%")) {
      const pct = Number(raw.slice(0, -1));
      value = Number.isFinite(pct) ? pct / 100 : NaN;
    } else {
      value = Number(raw);
    }
    return Number.isFinite(value) ? { ok: true, value } : { ok: false, value: null };
  }

  function isExternalReferenceDraft(rawInput) {
    const raw = String(rawInput ?? "").trim();
    return !!raw && (raw.startsWith("=") || !!parseExcelReference(raw));
  }

  function linkedCellsMessage(result) {
    const count = Number(result?.affectedCellCount) || 0;
    return `Linked ${count} dataset cell${count === 1 ? "" : "s"}.`;
  }

  /**
   * One door for every formula draft — typed into a cell, committed from the
   * floating formula bar, or pasted. The draft is read once by the shared
   * grammar, and a standalone Excel or dataset link keeps its own controller
   * while anything with arithmetic in it is calculated as a formula.
   */
  async function commitReferenceDraft({ displayRow, displayColumn, reference }) {
    const classified = classifyDatasetFormula(reference);
    if (classified.kind === "invalid") return { handled: true, ok: false, error: classified.error };
    const commit = classified.kind === "excel"
      ? commitExternalReference
      : (classified.kind === "internal" ? commitInternalReference : commitFormulaReference);
    setStatus(classified.kind === "excel"
      ? "Loading linked values from Excel..."
      : (classified.kind === "internal"
        ? "Loading linked values from the referenced dataset..."
        : "Calculating the formula..."));
    return commit({ displayRow, displayColumn, reference });
  }

  function syncReferencePickSession(rawValue, owner = "cell") {
    const editing = owner === "bar" ? !!formulaHover.isEditing?.() : !!state.editingCell;
    const armed = editing && isInternalReferencePickDraft(rawValue);
    // One editor at a time: a draft that has stopped looking like a reference
    // only ends the pick if it is the draft the pick was armed for.
    if (!armed && referencePickOwner && referencePickOwner !== owner) return;
    if (armed) referencePickOwner = owner;
    if (armed === referencePickArmed) return;
    referencePickArmed = armed;
    if (armed) {
      beginReferencePick();
      return;
    }
    referencePickOwner = "";
    endReferencePick();
  }

  function stopReferencePickSession() {
    if (!referencePickArmed) return;
    referencePickArmed = false;
    referencePickOwner = "";
    endReferencePick();
  }

  /**
   * A reference picked in another Dataset window lands here, routed by the
   * Project Instance host. The picked rectangle goes where the draft has room
   * for it: it replaces the dataset reference the draft ends with, so repeated
   * picks re-aim it the way Excel re-aims the reference under the caret, or
   * follows the operator a formula is waiting on.
   */
  function applyDatasetReferencePick(message = {}) {
    if (!referencePickArmed) return false;
    const text = buildInternalDatasetReferenceText({
      datasetName: message.datasetName,
      rowStart: message.rowStart,
      rowEnd: message.rowEnd,
      colStart: message.colStart,
      colEnd: message.colEnd,
      isVector: String(message.dataFormat || "").trim().toLowerCase() === "vector",
    });
    if (!text) return false;
    if (referencePickOwner === "bar") {
      const draft = insertPickedDatasetReference(formulaHover.getDraft?.(), text);
      return !!draft && formulaHover.setDraft(draft, { focus: !!message.final });
    }
    const edit = state.editingCell;
    if (!edit) return false;
    const input = document.querySelector(`#tableWrap .dsCellInput[data-r="${edit.r}"][data-c="${edit.c}"]`);
    if (!input) return false;
    const draft = insertPickedDatasetReference(input.value, text);
    if (!draft) return false;
    input.value = draft;
    edit.pendingExternalReference = draft;
    if (message.final) {
      requestAnimationFrame(() => {
        if (!input.isConnected) return;
        input.focus({ preventScroll: true });
        try {
          input.setSelectionRange(input.value.length, input.value.length);
        } catch { /* keep browser default cursor placement */ }
      });
    }
    return true;
  }

  /**
   * The other half of the pick: this window's selection reported to the
   * window whose formula is being edited, in untransposed dataset
   * coordinates clamped to the value grid (total rows and columns fall off).
   */
  function publishReferencePickSelection(final) {
    if (!state.referencePickRequester || !state.model) return;
    const range = Array.isArray(state.selRanges) && state.selRanges.length
      ? state.selRanges[state.selRanges.length - 1]
      : null;
    if (!range) return;
    const cornerA = displayToActualCell(range.r0, range.c0);
    const cornerB = displayToActualCell(range.r1, range.c1);
    const rowLimit = (state.model.origin_labels?.length || 0) - 1;
    const colLimit = (state.model.dev_labels?.length || 0) - 1;
    if (rowLimit < 0 || colLimit < 0) return;
    const rowStart = Math.max(0, Math.min(cornerA.r, cornerB.r));
    const rowEnd = Math.min(rowLimit, Math.max(cornerA.r, cornerB.r));
    const colStart = Math.max(0, Math.min(cornerA.c, cornerB.c));
    const colEnd = Math.min(colLimit, Math.max(cornerA.c, cornerB.c));
    if (rowEnd < rowStart || colEnd < colStart) return;
    referencePickOffered = true;
    publishReferencePick({ rowStart, rowEnd, colStart, colEnd, final: !!final });
    applyReferencePickDecoration();
  }

  async function commitHoveredExternalFormula({ formula, context }) {
    if (isReadOnly()) {
      const error = reportReadOnlyRefusal();
      return { ok: false, error };
    }
    const displayRow = Number(context?.anchorDisplayRow);
    const displayColumn = Number(context?.anchorDisplayColumn);
    if (!Number.isInteger(displayRow) || !Number.isInteger(displayColumn)) {
      const error = "The linked range anchor is unavailable.";
      setStatus(error);
      return { ok: false, error };
    }

    cancelExternalReference();
    const result = await commitReferenceDraft({
      displayRow,
      displayColumn,
      reference: formula,
    });
    if (!result?.ok) {
      if (!result?.aborted && !result?.stale) {
        setStatus(result?.error || "The linked values could not be loaded.");
      }
      return result;
    }

    state.editingCell = null;
    renderTable();
    notifyDatasetUpdated();
    applySelectionFromState();
    setStatus(result.message || linkedCellsMessage(result));
    return result;
  }

  function canEditDisplayCell(displayR, displayC, options = {}) {
    if (isReadOnly()) {
      if (!options?.silent) reportReadOnlyRefusal();
      return null;
    }
    const model = getDisplayDatasetModel();
    const sourceModel = state.model;
    if (!model || !sourceModel) return null;
    if (displayR < 0 || displayC < 0) return null;
    if (displayR >= (model.origin_labels?.length || 0) || displayC >= (model.dev_labels?.length || 0)) return null;
    const actual = displayToActualCell(displayR, displayC);
    if (!sourceModel.mask?.[actual.r]?.[actual.c]) return null;
    if (!Array.isArray(sourceModel.values?.[actual.r])) return null;
    return actual;
  }

  function setDisplayCellValue(displayR, displayC, rawValue, options = {}) {
    const actual = canEditDisplayCell(displayR, displayC);
    if (!actual) return null;
    const parsed = parseEditableCellValue(rawValue);
    if (!parsed.ok) {
      if (!options?.silentInvalid) setStatus("Enter a numeric value.");
      return null;
    }
    if (options?.hardCodeLinks) hardCodeExternalLinkCells([actual]);
    state.model.values[actual.r][actual.c] = parsed.value;
    state.dirty.set(`${actual.r},${actual.c}`, parsed.value);
    return parsed.value;
  }

  function restoreDirtyValue(key, edit) {
    if (!state.dirty || typeof state.dirty.set !== "function") return;
    if (edit.hadDirtyValue) {
      state.dirty.set(key, edit.previousDirtyValue);
    } else if (typeof state.dirty.delete === "function") {
      state.dirty.delete(key);
    }
  }

  function cancelCellEdit(displayR, displayC) {
    const edit = state.editingCell;
    if (!edit || edit.r !== displayR || edit.c !== displayC) return false;
    stopReferencePickSession();
    if (edit.commitPending || edit.pendingExternalReference) cancelExternalReference();
    const actualR = Number.isInteger(edit.actualR) ? edit.actualR : null;
    const actualC = Number.isInteger(edit.actualC) ? edit.actualC : null;
    if (actualR !== null && actualC !== null && Array.isArray(state.model?.values?.[actualR])) {
      state.model.values[actualR][actualC] = edit.previousValue;
      restoreDirtyValue(`${actualR},${actualC}`, edit);
    }
    state.editingCell = null;
    renderTable();
    notifyDatasetUpdated();
    applySelectionFromState();
    setStatus("Edit canceled.");
    return true;
  }

  function syncInputCellDisplay(td, input, value) {
    if (td) {
      td.dataset.copyValue = value == null ? "" : String(value);
    }
    if (input) {
      input.classList.toggle("dsCellInputBlank", value == null);
    }
  }

  function getPrimaryEditCell() {
    const ranges = Array.isArray(state.selRanges) ? state.selRanges : [];
    return getTopLeftRangeCell(ranges) || state.activeCell;
  }

  function selectedRanges() {
    if (Array.isArray(state.selRanges) && state.selRanges.length) return state.selRanges;
    if (!state.activeCell) return [];
    return [normalizeRange(state.activeCell.r, state.activeCell.c, state.activeCell.r, state.activeCell.c)];
  }

  function fillCells(ranges, value, describe) {
    if (isReadOnly()) {
      reportReadOnlyRefusal();
      return 0;
    }
    if (!ranges.length) return 0;

    const seen = new Set();
    let applied = 0;
    for (const range of ranges) {
      for (let r = range.r0; r <= range.r1; r += 1) {
        for (let c = range.c0; c <= range.c1; c += 1) {
          const key = `${r},${c}`;
          if (seen.has(key)) continue;
          seen.add(key);
          const actual = canEditDisplayCell(r, c, { silent: true });
          if (!actual) continue;
          hardCodeExternalLinkCells([actual]);
          state.model.values[actual.r][actual.c] = value;
          state.dirty.set(`${actual.r},${actual.c}`, value);
          applied += 1;
        }
      }
    }
    if (!applied) return 0;
    state.editingCell = null;
    renderTable();
    notifyDatasetUpdated();
    applySelectionFromState();
    setStatus(describe(applied));
    return applied;
  }

  function fillSelectedCells(value, describe) {
    return fillCells(selectedRanges(), value, describe);
  }

  function describeZeroed(applied) {
    return `Set ${applied} cell${applied === 1 ? "" : "s"} to 0.`;
  }

  function zeroSelectedCells() {
    return fillSelectedCells(0, describeZeroed);
  }

  // `Clear data` on the context menu: every cell the grid shows goes to 0,
  // which is what lets the length controls open up again on a hand-entered
  // dataset.
  function clearAllCells() {
    const model = getDisplayDatasetModel();
    const rows = model?.origin_labels?.length || 0;
    const cols = model?.dev_labels?.length || 0;
    const ranges = rows && cols ? [normalizeRange(0, 0, rows - 1, cols - 1)] : [];
    return fillCells(ranges, 0, describeZeroed);
  }

  function selectionSignature() {
    return selectedRanges().map((range) => `${range.r0}:${range.c0}:${range.r1}:${range.c1}`).join("|");
  }

  function selectionSpansManyCells() {
    let count = 0;
    for (const range of selectedRanges()) {
      count += (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
      if (count > 1) return true;
    }
    return false;
  }

  function resetRangeFillSession() {
    rangeFillSession = null;
  }

  function nextRangeFillText(current, key) {
    if (key === "-") return current ? null : "-";
    if (key !== ".") return `${current}${key}`;
    if (current.includes(".")) return null;
    return current === "" || current === "-" ? `${current}0.` : `${current}.`;
  }

  function typeIntoSelectedRange(key) {
    const signature = selectionSignature();
    if (!signature) return false;
    const current = rangeFillSession?.signature === signature ? rangeFillSession.text : "";
    const text = nextRangeFillText(current, key);
    if (text === null) return false;
    rangeFillSession = { signature, text };
    const parsed = parseEditableCellValue(text);
    // A lone leading sign holds the session open until the first digit arrives.
    if (!parsed.ok) return true;
    const applied = fillSelectedCells(
      parsed.value,
      (count) => `Set ${count} cell${count === 1 ? "" : "s"} to ${text}.`,
    );
    if (applied) return true;
    resetRangeFillSession();
    return false;
  }

  function parseClipboardRows(text) {
    return String(text || "")
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n")
      .split("\n")
      .filter((row, index, arr) => index < arr.length - 1 || row !== "")
      .map((row) => row.split("\t"));
  }

  function applyPastedGridText(text, start) {
    if (isReadOnly()) {
      reportReadOnlyRefusal();
      return 0;
    }
    const model = getDisplayDatasetModel();
    const sourceModel = state.model;
    if (!model || !sourceModel || !start) return 0;

    const rows = parseClipboardRows(text);
    if (!rows.length) return 0;

    if (rows.length === 1 && rows[0].length === 1 && isExternalReferenceDraft(rows[0][0])) {
      const rawReference = rows[0][0];
      void (async () => {
        const result = await commitReferenceDraft({
          displayRow: start.r,
          displayColumn: start.c,
          reference: rawReference,
        });
        if (!result?.ok) {
          if (!result?.aborted && !result?.stale) setStatus(result?.error || "The linked values could not be loaded.");
          return;
        }
        state.editingCell = null;
        state.activeCell = { r: start.r, c: start.c };
        state.selectionAnchor = { r: start.r, c: start.c };
        state.selRanges = [normalizeRange(start.r, start.c, start.r, start.c)];
        renderTable();
        notifyDatasetUpdated();
        applySelectionFromState();
        setStatus(result.message || linkedCellsMessage(result));
      })();
      return 1;
    }

    if (rows.length === 1 && rows[0].length === 1 && Array.isArray(state.selRanges) && state.selRanges.length) {
      const parsed = parseEditableCellValue(rows[0][0]);
      if (!parsed.ok) return 0;
      const seen = new Set();
      let applied = 0;
      for (const range of state.selRanges) {
        for (let r = range.r0; r <= range.r1; r += 1) {
          for (let c = range.c0; c <= range.c1; c += 1) {
            const actual = canEditDisplayCell(r, c, { silent: true });
            if (!actual) continue;
            const key = `${actual.r},${actual.c}`;
            if (seen.has(key)) continue;
            seen.add(key);
            hardCodeExternalLinkCells([actual]);
            state.model.values[actual.r][actual.c] = parsed.value;
            state.dirty.set(key, parsed.value);
            applied += 1;
          }
        }
      }
      if (!applied) return 0;
      state.editingCell = null;
      renderTable();
      notifyDatasetUpdated();
      applySelectionFromState();
      setStatus(`Pasted ${applied} cell${applied === 1 ? "" : "s"}.`);
      return applied;
    }

    let applied = 0;
    for (let rr = 0; rr < rows.length; rr += 1) {
      for (let cc = 0; cc < rows[rr].length; cc += 1) {
        const displayR = start.r + rr;
        const displayC = start.c + cc;
        if (displayR < 0 || displayC < 0) continue;
        if (displayR >= (model.origin_labels?.length || 0) || displayC >= (model.dev_labels?.length || 0)) continue;

        const actual = displayToActualCell(displayR, displayC);
        if (!sourceModel.mask?.[actual.r]?.[actual.c]) continue;

        const parsed = parseEditableCellValue(rows[rr][cc]);
        if (!parsed.ok) continue;

        if (!Array.isArray(sourceModel.values[actual.r])) continue;
        hardCodeExternalLinkCells([actual]);
        sourceModel.values[actual.r][actual.c] = parsed.value;
        state.dirty.set(`${actual.r},${actual.c}`, parsed.value);
        applied += 1;
      }
    }
    if (!applied) return 0;
    state.activeCell = { r: start.r, c: start.c };
    state.selectionAnchor = { r: start.r, c: start.c };
    const lastDisplayRow = Math.min(
      start.r + rows.length - 1,
      Math.max(0, (model.origin_labels?.length || 0) - 1),
    );
    const lastDisplayColumn = Math.min(
      start.c + Math.max(0, rows.reduce((max, row) => Math.max(max, row.length), 0) - 1),
      Math.max(0, (model.dev_labels?.length || 0) - 1),
    );
    state.selRanges = [normalizeRange(
      start.r,
      start.c,
      lastDisplayRow,
      lastDisplayColumn,
    )];
    renderTable();
    notifyDatasetUpdated();
    applySelectionFromState();
    setStatus(`Pasted ${applied} cell${applied === 1 ? "" : "s"}.`);
    return applied;
  }

  function focusCellInput(displayR, displayC, initialText = null) {
    const actual = canEditDisplayCell(displayR, displayC);
    if (!actual) return false;
    resetRangeFillSession();
    formulaHover.hide?.();
    cancelExternalReference();
    const dirtyKey = `${actual.r},${actual.c}`;
    const hadDirtyValue = !!state.dirty?.has?.(dirtyKey);
    state.editingCell = {
      r: displayR,
      c: displayC,
      actualR: actual.r,
      actualC: actual.c,
      previousValue: state.model?.values?.[actual.r]?.[actual.c],
      hadDirtyValue,
      previousDirtyValue: hadDirtyValue ? state.dirty.get(dirtyKey) : undefined,
    };
    state.activeCell = { r: displayR, c: displayC };
    renderTable();
    applySelectionFromState();

    const input = document.querySelector(`#tableWrap .dsCellInput[data-r="${displayR}"][data-c="${displayC}"]`);
    if (!input) {
      state.editingCell = null;
      canEditDisplayCell(displayR, displayC);
      renderTable();
      applySelectionFromState();
      return false;
    }
    if (initialText !== null) {
      input.value = String(initialText);
      if (isExternalReferenceDraft(input.value)) {
        state.editingCell.pendingExternalReference = input.value;
        syncReferencePickSession(input.value);
      } else {
        setDisplayCellValue(displayR, displayC, input.value, { silentInvalid: true });
        const actual = displayToActualCell(displayR, displayC);
        syncInputCellDisplay(input.closest("td"), input, state.model?.values?.[actual.r]?.[actual.c]);
        notifyDatasetUpdated();
      }
    }
    requestAnimationFrame(() => {
      input.focus({ preventScroll: true });
      if (initialText === null) {
        try { input.select(); } catch { /* number inputs do not expose text selection consistently */ }
      } else {
        try { input.setSelectionRange(input.value.length, input.value.length); } catch { /* keep browser default cursor placement */ }
      }
    });
    return true;
  }

  function applySelectionFromState() {
    spreadsheetTable.applyDom();
    applyReferencePickDecoration();
  }

  /**
   * The rectangle this window is offering to the window that is writing a
   * formula: the last range the user drew, which is also the one published.
   */
  function referencePickRange() {
    if (!state.referencePickRequester || !referencePickOffered) return null;
    const ranges = Array.isArray(state.selRanges) ? state.selRanges : [];
    return ranges.length ? ranges[ranges.length - 1] : null;
  }

  /**
   * Mark what this window is offering, the way a spreadsheet does while a
   * formula is reading from it: the picked range is ringed by moving dashes,
   * and the whole grid says its cells can be pointed at.
   */
  function applyReferencePickDecoration() {
    const picking = !!state.referencePickRequester;
    // No pick running and none painted, which is every ordinary selection.
    if (!picking && !referencePickDecorated) return;
    const wrap = document.getElementById("tableWrap");
    if (!wrap) return;
    if (!picking) referencePickOffered = false;
    referencePickDecorated = picking;
    wrap.querySelector("table")?.classList?.toggle("isReferencePickSource", picking);
    const range = referencePickRange();
    wrap.querySelectorAll("td[data-r][data-c]").forEach((cell) => {
      const row = Number(cell.dataset?.r);
      const column = Number(cell.dataset?.c);
      const inside = !!range
        && row >= range.r0 && row <= range.r1
        && column >= range.c0 && column <= range.c1;
      const top = inside && row === range.r0;
      const bottom = inside && row === range.r1;
      const left = inside && column === range.c0;
      const right = inside && column === range.c1;
      // Only the perimeter carries the dashes, so a cell in the middle of a
      // large range is left without an overlay to animate.
      cell.classList.toggle("arReferencePickCell", top || bottom || left || right);
      cell.classList.toggle("arReferencePickEdgeTop", top);
      cell.classList.toggle("arReferencePickEdgeBottom", bottom);
      cell.classList.toggle("arReferencePickEdgeLeft", left);
      cell.classList.toggle("arReferencePickEdgeRight", right);
      if (!picking) cell.classList.remove("arReferencePickHover");
    });
  }

  /** The cell under the pointer, lit up before it has been picked. */
  function setReferencePickHover(cell) {
    const wrap = document.getElementById("tableWrap");
    if (!wrap) return;
    wrap.querySelectorAll("td.arReferencePickHover").forEach((hovered) => {
      if (hovered !== cell) hovered.classList.remove("arReferencePickHover");
    });
    if (cell && state.referencePickRequester) cell.classList.add("arReferencePickHover");
  }

  function clearGridSelection() {
    state.dragSel = null;
    spreadsheetTable.clear();
  }

  function prepareContextSelection(r, c) {
    spreadsheetTable.prepareContextCell({ r, c });
    window.__arcRhoCopyActiveGridSelection = copyActiveRangeToClipboard;
  }

  function hasEditableSelectionTarget() {
    const ranges = Array.isArray(state.selRanges) ? state.selRanges : [];
    for (const range of ranges) {
      for (let r = range.r0; r <= range.r1; r += 1) {
        for (let c = range.c0; c <= range.c1; c += 1) {
          if (canEditDisplayCell(r, c, { silent: true })) return true;
        }
      }
    }
    return false;
  }

  async function pasteSelectionFromClipboard() {
    if (!navigator.clipboard?.readText) {
      setStatus("Clipboard paste is not available in this browser.");
      return 0;
    }
    try {
      const text = await navigator.clipboard.readText();
      const start = getTopLeftRangeCell(state.selRanges || []) || state.activeCell;
      return text && start ? applyPastedGridText(text, start) : 0;
    } catch (error) {
      setStatus(`Paste failed: ${String(error?.message || error)}`);
      return 0;
    }
  }

  async function handleGridContextAction(action) {
    if (action === "paste") return pasteSelectionFromClipboard();
    if (action === "clear_data") return clearAllCells();
    if (action === "toggle_subtotal") {
      state.showSubtotal = state.showSubtotal === false;
      state.activeCell = null;
      state.selRanges = [];
      renderTable();
      refreshDatasetSettingsDirty();
      return true;
    }
    if (action === "remove_highlights") {
      clearGridSelection();
      return true;
    }
    return false;
  }

  async function copyActiveRangeToClipboard() {
    return spreadsheetTable.copy();
  }

  function wireRectSelectionAndCopy() {
    if (window.__arcRhoRectSelWired) return;
    window.__arcRhoRectSelWired = true;

    // state containers
    if (!Array.isArray(state.selRanges)) state.selRanges = [];
    state.dragSel = null;
    window.__arcRhoDatasetCopyActiveGridSelection = copyActiveRangeToClipboard;
    window.__arcRhoCopyActiveGridSelection = copyActiveRangeToClipboard;

    const wrap = document.getElementById("tableWrap");
    if (!wrap) return;

    // While another window is waiting for a range, the cell under the pointer
    // lights up before it is picked, so it is clear what a click would take.
    wrap.addEventListener("mouseover", (e) => {
      if (!state.referencePickRequester) return;
      setReferencePickHover(e.target.closest?.("td[data-r][data-c]") || null);
    });
    wrap.addEventListener("mouseleave", () => setReferencePickHover(null));

    // start drag
    wrap.addEventListener("mousedown", (e) => {
      // left button only
      if (e.button !== 0) return;
      if (isTypingTarget(e.target)) return;

      // NEW: leave dropdown/input focus when interacting with grid
      const ae = document.activeElement;
      if (ae && isTypingTarget(ae)) {
        try { ae.blur(); } catch {}
      }

      const td = e.target.closest('td[data-r][data-c]');
      if (!td) return;

      e.preventDefault(); // stop text selection

      const rc = rcFromTd(td);
      if (!rc) return;
      window.__arcRhoCopyActiveGridSelection = copyActiveRangeToClipboard;

      const append = !!(e.ctrlKey || e.metaKey) && !e.shiftKey;
      const baseRanges = append ? spreadsheetTable.selection().ranges : [];
      spreadsheetTable.selectCell(rc, { append, extend: e.shiftKey });
      const selected = spreadsheetTable.selection();

      state.dragSel = {
        anchor: selected.anchorCell || { r: rc.r, c: rc.c },
        append,
        baseRanges,
      };
      if (state.referencePickRequester) {
        referencePickGesture = true;
        publishReferencePickSelection(false);
      }
    });

    // drag over (use mouseover to avoid heavy mousemove)
    wrap.addEventListener("mouseover", (e) => {
      if (!state.dragSel) return;

      const td = e.target.closest('td[data-r][data-c]');
      if (!td) return;

      const rc = rcFromTd(td);
      if (!rc) return;

      const { anchor, append, baseRanges } = state.dragSel;
      spreadsheetTable.setRange(anchor, rc, { append, baseRanges });
      if (referencePickGesture) publishReferencePickSelection(false);
    });

    // end drag anywhere
    document.addEventListener("mouseup", () => {
      state.dragSel = null;
      if (referencePickGesture) {
        referencePickGesture = false;
        publishReferencePickSelection(true);
      }
    });

    // Click row header -> select entire row
    // Click row header -> select / deselect entire row
    wrap.addEventListener("click", (e) => {
      const th = e.target.closest("th.rowhdr[data-r]");
      if (!th) return;
      const r = Number(th.dataset.r);
      window.__arcRhoCopyActiveGridSelection = copyActiveRangeToClipboard;
      spreadsheetTable.selectRow(r, {
        append: (e.ctrlKey || e.metaKey) && !e.shiftKey,
        extend: e.shiftKey,
        toggle: !e.shiftKey,
      });
    });

    // Click column header -> select / deselect entire column
    wrap.addEventListener("click", (e) => {
      const th = e.target.closest("th.colhdr[data-c]");
      if (!th) return;
      const c = Number(th.dataset.c);
      window.__arcRhoCopyActiveGridSelection = copyActiveRangeToClipboard;
      spreadsheetTable.selectColumn(c, {
        append: (e.ctrlKey || e.metaKey) && !e.shiftKey,
        extend: e.shiftKey,
        toggle: !e.shiftKey,
      });
    });

    // Ctrl+C copy
    document.addEventListener("keydown", (e) => {
      if (isTypingTarget(e.target)) return;

      const isCopy = (e.key === "c" || e.key === "C") && (e.ctrlKey || e.metaKey);
      if (!isCopy) return;

      if (!state.selRanges || !state.selRanges.length) return;
      if (window.__arcRhoCopyActiveGridSelection !== copyActiveRangeToClipboard) return;

      e.preventDefault();
      copyActiveRangeToClipboard();
    });

    document.addEventListener("keydown", (e) => {
      if (isTypingTarget(e.target)) return;
      if (e.key === "Escape" && (state.activeCell || state.selRanges?.length)) {
        e.preventDefault();
        clearGridSelection();
        return;
      }
      if (!state.activeCell) return;
      if (e.ctrlKey || e.metaKey || e.altKey) return;

      if (e.key === "F2") {
        const cell = getPrimaryEditCell();
        if (!cell) return;
        const info = getExternalLinkCellInfo(cell.r, cell.c);
        if (!info?.reference && !info?.note) return;
        const hoveredCell = document.querySelector(`#tableWrap td[data-r="${cell.r}"][data-c="${cell.c}"]`);
        if (info.note) {
          if (!hoveredCell) return;
          e.preventDefault();
          formulaHover.open(hoveredCell, info, { key: `note:${cell.r},${cell.c}` });
          return;
        }
        const anchor = resolveExternalFormulaAnchor(info, hoveredCell);
        if (!anchor) return;
        e.preventDefault();
        formulaHover.open(anchor, {
          ...info,
          formula: info.reference,
          readOnly: isReadOnly(),
        }, {
          focus: true,
          positionRect: () => resolveExternalFormulaRangeRect(info, anchor),
        });
        return;
      }

      if (e.key === "Delete" || e.key === "Backspace") {
        e.preventDefault();
        resetRangeFillSession();
        zeroSelectedCells();
        return;
      }

      if (e.key === "Enter" && rangeFillSession) {
        e.preventDefault();
        resetRangeFillSession();
        return;
      }

      if (/^[0-9.-]$/.test(e.key || "") && selectionSpansManyCells()) {
        if (typeIntoSelectedRange(e.key)) e.preventDefault();
        return;
      }

      if (/^[0-9]$/.test(e.key || "") || e.key === "=") {
        const cell = getPrimaryEditCell();
        if (!cell) return;
        e.preventDefault();
        focusCellInput(cell.r, cell.c, e.key);
      }
    });

    document.addEventListener("paste", (e) => {
      if (isTypingTarget(e.target)) return;
      if (isReadOnly()) {
        reportReadOnlyRefusal();
        return;
      }
      const text = String(e.clipboardData?.getData("text/plain") || "");
      if (!text) return;
      const start = getTopLeftRangeCell(state.selRanges || []) || state.activeCell;
      if (!start) return;
      const applied = applyPastedGridText(text, start);
      if (!applied) return;
      e.preventDefault();
    });
  }

  return {
    applyDatasetReferencePick,
    applySelectionFromState,
    copyActiveRangeToClipboard,
  };
}
