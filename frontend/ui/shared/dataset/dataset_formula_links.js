/*
===============================================================================
Dataset Formula Links
Per-cell links whose values are calculated from a formula rather than copied
from one source: the third sibling of the Excel `external_links` and ArcRho
`internal_links` controllers. A link stores the canonical formula text and the
target cells it owns, each mapped to a cell of the formula's result matrix:

    { formula: "=[C 82 - Prior Qtr Selected][1:7] * 2",
      target_cells: [{ row, column, result_row, result_column }, ...] }

Committing or refreshing resolves every dataset reference on the app server
(one read per unique dataset) and every Excel reference through one workbook
read, evaluates the formula with Excel's array rules, and spills the result
from the anchor cell exactly like a range link. Values are snapshots; the
formula is re-evaluated only when the user asks.
===============================================================================
*/
import { readExcelCellsBatch } from "/ui/shared/integrations/excel_api.js?v=20260819a";
import {
  excelColumnFromIndex,
  parseExcelCellAddress,
} from "/ui/shared/integrations/excel_reference.js?v=20260715a";
import {
  applyDatasetLinkOutlineClasses,
  buildDatasetExternalLinkTargets,
  buildDatasetLinkOutline,
  describeTargetDestination,
} from "/ui/shared/dataset/dataset_external_links.js?v=20260907b";
import {
  classifyDatasetFormula,
  evaluateDatasetFormula,
  parseDatasetFormula,
} from "/ui/shared/dataset/dataset_formula.js?v=20260830a";
import { formatInternalDatasetReference } from "/ui/shared/dataset/dataset_internal_reference.js?v=20260830a";

function targetCellKey(target) {
  return `${target.row},${target.column}`;
}

function nonnegativeInt(value) {
  const numeric = Number(value);
  return Number.isInteger(numeric) && numeric >= 0 ? numeric : null;
}

function cloneLinks(links) {
  return links.map((link) => ({
    formula: link.formula,
    target_cells: link.target_cells.map((target) => ({ ...target })),
  }));
}

function linksSnapshot(links) {
  return JSON.stringify(links);
}

function canonicalFormula(value) {
  const parsed = parseDatasetFormula(value);
  return parsed.ok && parsed.references.length ? parsed.canonical : "";
}

export function normalizeDatasetFormulaLinks(value) {
  const source = Array.isArray(value) ? value : [];
  const normalized = [];
  const seenLinks = new Set();
  const ownedTargets = new Set();
  source.forEach((item) => {
    const formula = canonicalFormula(item?.formula);
    if (!formula) return;
    const rawTargets = Array.isArray(item?.target_cells)
      ? item.target_cells
      : (Array.isArray(item?.targetCells) ? item.targetCells : []);
    const targetCells = [];
    const seenTargets = new Set();
    const seenResults = new Set();
    let invalidTargets = false;
    rawTargets.forEach((target) => {
      const row = nonnegativeInt(target?.row);
      const column = nonnegativeInt(target?.column);
      const resultRow = nonnegativeInt(target?.result_row ?? target?.resultRow);
      const resultColumn = nonnegativeInt(target?.result_column ?? target?.resultColumn);
      if (row === null || column === null || resultRow === null || resultColumn === null) {
        invalidTargets = true;
        return;
      }
      const key = `${row},${column}`;
      const resultKey = `${resultRow},${resultColumn}`;
      if (seenTargets.has(key) || seenResults.has(resultKey)) {
        invalidTargets = true;
        return;
      }
      seenTargets.add(key);
      seenResults.add(resultKey);
      targetCells.push({ row, column, result_row: resultRow, result_column: resultColumn });
    });
    if (invalidTargets || !targetCells.length) return;
    const linkKey = `${formula}${targetCells.map(targetCellKey).join(";")}`;
    if (seenLinks.has(linkKey)) return;
    if (targetCells.some((target) => ownedTargets.has(targetCellKey(target)))) return;
    seenLinks.add(linkKey);
    targetCells.forEach((target) => ownedTargets.add(targetCellKey(target)));
    normalized.push({ formula, target_cells: targetCells });
  });
  return normalized;
}

function displayToActualCell(row, column, transposed) {
  return transposed ? { row: column, column: row } : { row, column };
}

function actualToDisplayCell(row, column, transposed) {
  return transposed ? { row: column, column: row } : { row, column };
}

function targetDestinationLabel(model, target) {
  const origin = String(model?.origin_labels?.[target.row] ?? `Row ${target.row + 1}`);
  const development = String(model?.dev_labels?.[target.column] ?? "");
  return development ? `${origin} / ${development}` : origin;
}

function targetValuePreview(model, targets, isRange) {
  const first = targets[0];
  if (!first) return "";
  const value = model?.values?.[first.row]?.[first.column];
  const text = value === null || value === undefined ? "" : String(value);
  return isRange ? `${text}...` : text;
}

function valuesEqual(left, right) {
  if (left == null || right == null) return left == null && right == null;
  return Number(left) === Number(right);
}

function linkRecordId(link) {
  return `${link.formula}${link.target_cells.map(targetCellKey).join(";")}`;
}

/** Just the coordinates an internal reference token names, e.g. "[1:12]". */
function internalReferenceCoordinateText(parsed) {
  const formatted = formatInternalDatasetReference(parsed);
  const open = formatted.lastIndexOf("][");
  return open >= 0 ? `[${formatted.slice(open + 2, -1)}]` : "";
}

/** Just the sheet and address an Excel reference token names, e.g. "Sheet1!A1:A12". */
function excelReferenceAddressText(parsed) {
  const sheet = String(parsed.sheet || "").trim();
  const cell = String(parsed.cell || "").trim();
  const endCell = String(parsed.endCell || cell).trim();
  const address = cell && endCell && cell !== endCell ? `${cell}:${endCell}` : cell;
  return sheet && address ? `${sheet}!${address}` : (address || sheet);
}

/**
 * The sources a formula reads, one per dataset or workbook however many
 * references it makes to each: the Links tab lists a formula once per source,
 * under that source's own kind, showing that source's own reference rather
 * than the whole formula.
 */
function formulaComponents(parsed) {
  const components = new Map();
  for (const reference of parsed.references) {
    const component = reference.kind === "internal"
      ? {
        datasetName: String(reference.parsed.datasetName || ""),
        reference: internalReferenceCoordinateText(reference.parsed),
      }
      : {
        workbookPath: String(reference.parsed.bookPath || ""),
        reference: excelReferenceAddressText(reference.parsed),
      };
    const key = component.datasetName || component.workbookPath;
    if (key && !components.has(key)) components.set(key, component);
  }
  return Array.from(components.values());
}

// A component row's id carries its formula's id, so refreshing or breaking
// any one row acts on the whole formula.
const COMPONENT_ID_SEPARATOR = "#component-";

function requestedLinkIds(ids) {
  return new Set(
    (Array.isArray(ids) ? ids : [ids])
      .map((id) => String(id || "").split(COMPONENT_ID_SEPARATOR)[0])
      .filter(Boolean),
  );
}

function excelCellValue(result) {
  if (!result?.ok) return { ok: false, error: String(result?.error || "Excel cell read failed.") };
  if (result.value === null || result.value === undefined || result.value === "") return { ok: true, value: null };
  const value = Number(result.value);
  return Number.isFinite(value)
    ? { ok: true, value }
    : { ok: false, error: `Excel returned a non-numeric value: ${String(result.value)}` };
}

function excelRangeCells(parsed) {
  const start = parseExcelCellAddress(parsed.cell);
  const end = parseExcelCellAddress(parsed.endCell || parsed.cell);
  const row0 = Math.min(start.row, end.row);
  const row1 = Math.max(start.row, end.row);
  const col0 = Math.min(start.col, end.col);
  const col1 = Math.max(start.col, end.col);
  const cells = [];
  for (let row = row0; row <= row1; row += 1) {
    for (let col = col0; col <= col1; col += 1) {
      cells.push(`${excelColumnFromIndex(col)}${row + 1}`);
    }
  }
  return { rows: row1 - row0 + 1, cols: col1 - col0 + 1, cells };
}

function matrixFromCells(rows, cols, flat) {
  const values = [];
  for (let row = 0; row < rows; row += 1) values.push(flat.slice(row * cols, (row + 1) * cols));
  return { rows, cols, values };
}

function referenceKey(token) {
  return `${token.kind}${token.canonical}`;
}

export function createDatasetFormulaLinksController({
  state,
  resolveReferences,
  readCellsBatch = readExcelCellsBatch,
  isReadOnly = () => false,
  isTransposed = () => false,
  isAtLinkedShape = () => true,
  onInventoryChanged = () => {},
  onTargetsClaimed = () => {},
} = {}) {
  let links = [];
  let savedLinks = [];
  let requestGeneration = 0;
  let requestController = null;
  let targetDecorationIndex = null;
  // Target cell key -> the failure that left its stored value in place, so a
  // re-render repaints the same cells red without another evaluation.
  let failuresByTargetKey = new Map();

  function ownedTargetKeys() {
    const keys = new Set();
    links.forEach((link) => link.target_cells.forEach((target) => keys.add(targetCellKey(target))));
    return keys;
  }

  function notifyInventoryChanged() {
    targetDecorationIndex = null;
    if (failuresByTargetKey.size) {
      const owned = ownedTargetKeys();
      failuresByTargetKey.forEach((_failure, key) => {
        if (!owned.has(key)) failuresByTargetKey.delete(key);
      });
    }
    onInventoryChanged();
  }

  function abort() {
    requestGeneration += 1;
    if (requestController) requestController.abort();
    requestController = null;
  }

  function load(value) {
    abort();
    failuresByTargetKey = new Map();
    links = normalizeDatasetFormulaLinks(value);
    savedLinks = cloneLinks(links);
    notifyInventoryChanged();
  }

  function clear() {
    load([]);
  }

  function serialize() {
    return cloneLinks(links);
  }

  function isDirty() {
    return linksSnapshot(links) !== linksSnapshot(savedLinks);
  }

  function markClean(value = links) {
    links = normalizeDatasetFormulaLinks(value);
    savedLinks = cloneLinks(links);
    notifyInventoryChanged();
  }

  function restoreSaved() {
    abort();
    failuresByTargetKey = new Map();
    links = cloneLinks(savedLinks);
    notifyInventoryChanged();
  }

  function listFailures() {
    return Array.from(failuresByTargetKey.values()).map((failure) => ({ ...failure }));
  }

  function getTargetDecorationIndex() {
    const transposed = !!isTransposed();
    if (targetDecorationIndex?.transposed === transposed) return targetDecorationIndex;
    const targets = new Map();
    const outlineGaps = new Map();
    links.forEach((link) => {
      const isRange = link.target_cells.length > 1;
      const outline = buildDatasetLinkOutline(link.target_cells, transposed);
      link.target_cells.forEach((target) => {
        const display = actualToDisplayCell(target.row, target.column, transposed);
        targets.set(targetCellKey(target), {
          link,
          target,
          isRange,
          ...(outline ? outline.edgesAt(display.row, display.column) : {}),
        });
      });
      if (!isRange || !outline) return;
      outline.gapCells.forEach((gap) => {
        outlineGaps.set(targetCellKey(gap.cell), gap.edges);
      });
    });
    targetDecorationIndex = { transposed, targets, outlineGaps };
    return targetDecorationIndex;
  }

  function linksForTargetCells(targetCells) {
    const keys = new Set((Array.isArray(targetCells) ? targetCells : []).map(targetCellKey));
    if (!keys.size) return new Set();
    const indexes = new Set();
    links.forEach((link, index) => {
      if (link.target_cells.some((target) => keys.has(targetCellKey(target)))) indexes.add(index);
    });
    return indexes;
  }

  function removeLinkIndexes(indexes) {
    if (!(indexes instanceof Set) || !indexes.size) return 0;
    const previousCount = links.length;
    links = links.filter((_link, index) => !indexes.has(index));
    const removed = previousCount - links.length;
    if (removed) notifyInventoryChanged();
    return removed;
  }

  // A saved link names a cell of the grid its dataset was displayed at. While
  // the window is showing it at other lengths, every cell on screen stands for
  // other cells entirely, so there is no square to paint, name, or release: the
  // whole inventory stands still and comes back when the lengths do.
  function hardCodeTargetCells(targetCells) {
    if (!isAtLinkedShape()) return 0;
    return removeLinkIndexes(linksForTargetCells(targetCells));
  }

  function getCellLinkInfo(displayRow, displayColumn) {
    if (!state?.model || !isAtLinkedShape()) return null;
    const actual = displayToActualCell(displayRow, displayColumn, !!isTransposed());
    const decoration = getTargetDecorationIndex().targets.get(targetCellKey(actual));
    const link = decoration?.link;
    if (!link) return null;
    const anchor = link.target_cells[0];
    if (!anchor) return null;
    const transposed = !!isTransposed();
    return {
      id: linkRecordId(link),
      reference: link.formula,
      sourceKind: "formula",
      anchorDisplayRow: transposed ? anchor.column : anchor.row,
      anchorDisplayColumn: transposed ? anchor.row : anchor.column,
    };
  }

  function decorateCell(cell, displayRow, displayColumn) {
    if (!cell || !state?.model || !isAtLinkedShape()) return;
    const actual = displayToActualCell(displayRow, displayColumn, !!isTransposed());
    const key = targetCellKey(actual);
    const index = getTargetDecorationIndex();
    const decoration = index.targets.get(key);
    const link = decoration?.link;
    const failure = failuresByTargetKey.get(key);
    // A blank the mask left inside a calculated rectangle still carries that
    // rectangle's edge, so the frame closes across the empty corner.
    const outline = decoration?.isRange ? decoration : index.outlineGaps.get(key);
    cell.classList.toggle("arFormulaLinkCell", !!link || !!outline);
    cell.classList.toggle("arFormulaLinkErrorCell", !!failure);
    applyDatasetLinkOutlineClasses(cell, outline, "formula");
    if (link) {
      cell.dataset.formulaLinkReference = link.formula;
    } else {
      delete cell.dataset.formulaLinkReference;
    }
  }

  // The Value and Destination columns are read off the grid on screen. A view
  // at other lengths is not the grid the link points into, so they are left
  // blank rather than quoting a number and a period the link never named.
  function listRecords() {
    const atLinkedShape = !!isAtLinkedShape();
    return links.flatMap((link) => {
      const targets = link.target_cells;
      const parsed = parseDatasetFormula(link.formula);
      const linkId = linkRecordId(link);
      const record = {
        sourceKind: "formula",
        formula: link.formula,
        reference: link.formula,
        value: atLinkedShape ? targetValuePreview(state?.model, targets, targets.length > 1) : "",
        destination: (atLinkedShape ? describeTargetDestination(state?.model, targets) : "") || "Data",
        affectedCellCount: targets.length,
        readOnly: !!isReadOnly(),
      };
      // One row per source the formula reads, each filed under that source's
      // kind; a formula the grammar rejects keeps a single row.
      const components = parsed.ok ? formulaComponents(parsed) : [];
      if (!components.length) return [{ id: linkId, ...record }];
      return components.map((component, index) => ({
        id: `${linkId}${COMPONENT_ID_SEPARATOR}${index}`,
        ...record,
        ...component,
      }));
    });
  }

  /**
   * Resolve every reference the formula reads and evaluate it. Dataset
   * references go to the app server in one request; Excel references go to
   * the workbook reader in one batch. Returns the result matrix or a failure
   * that names the reference at fault.
   */
  async function evaluateFormula(parsed, generation, signal) {
    const matrices = new Map();
    const internalReferences = parsed.references.filter((token) => token.kind === "internal");
    if (internalReferences.length) {
      let resp;
      try {
        resp = await resolveReferences(internalReferences.map((token) => `=${token.canonical}`));
      } catch (error) {
        return { ok: false, error: String(error?.message || error || "Dataset reference resolve failed.") };
      }
      if (generation !== requestGeneration) return { ok: false, stale: true };
      if (!resp?.ok) {
        return {
          ok: false,
          error: String(resp?.data?.detail || resp?.data?.error || "The dataset reference could not be resolved."),
        };
      }
      const results = Array.isArray(resp.data?.results) ? resp.data.results : [];
      for (let index = 0; index < internalReferences.length; index += 1) {
        const result = results[index];
        const rows = Number(result?.row_count) || 0;
        const cols = Number(result?.column_count) || 0;
        if (!Array.isArray(result?.cells) || result.cells.length !== rows * cols) {
          return { ok: false, error: `${internalReferences[index].text} could not be resolved.` };
        }
        matrices.set(
          referenceKey(internalReferences[index]),
          matrixFromCells(rows, cols, result.cells.map((cell) => cell.value ?? null)),
        );
      }
    }
    const excelReferences = parsed.references.filter((token) => token.kind === "excel");
    if (excelReferences.length) {
      const items = [];
      const spans = excelReferences.map((token) => {
        const range = excelRangeCells(token.parsed);
        const start = items.length;
        range.cells.forEach((cell) => items.push({
          book_path: token.parsed.bookPath,
          sheet: token.parsed.sheet,
          cell,
        }));
        return { token, range, start };
      });
      let resp;
      try {
        resp = await readCellsBatch(items, { signal });
      } catch (error) {
        if (error?.name === "AbortError") return { ok: false, aborted: true };
        return { ok: false, error: String(error?.message || error || "Excel read failed.") };
      }
      if (generation !== requestGeneration) return { ok: false, stale: true };
      if (!resp?.ok || !Array.isArray(resp.results) || resp.results.length !== items.length) {
        return { ok: false, error: String(resp?.error || "Excel range read failed.") };
      }
      for (const span of spans) {
        const flat = [];
        for (let offset = 0; offset < span.range.cells.length; offset += 1) {
          const parsedValue = excelCellValue(resp.results[span.start + offset]);
          if (!parsedValue.ok) {
            return { ok: false, error: `${span.range.cells[offset]}: ${parsedValue.error}` };
          }
          flat.push(parsedValue.value);
        }
        matrices.set(referenceKey(span.token), matrixFromCells(span.range.rows, span.range.cols, flat));
      }
    }
    return evaluateDatasetFormula(parsed.tree, (token) => matrices.get(referenceKey(token)));
  }

  async function commitReference({ displayRow, displayColumn, reference } = {}) {
    if (isReadOnly()) return { handled: true, ok: false, error: "This dataset is read-only." };
    const classified = classifyDatasetFormula(reference);
    if (classified.kind === "invalid") return { handled: true, ok: false, error: classified.error };
    if (classified.kind !== "formula") return { handled: false, ok: false };

    abort();
    const generation = requestGeneration;
    requestController = new AbortController();
    let result;
    try {
      result = await evaluateFormula(classified, generation, requestController.signal);
    } finally {
      if (generation === requestGeneration) requestController = null;
    }
    if (!result.ok) return { handled: true, ...result };
    const targetResult = buildDatasetExternalLinkTargets({
      model: state?.model,
      transposed: !!isTransposed(),
      startRow: displayRow,
      startColumn: displayColumn,
      rowCount: result.rows,
      columnCount: result.cols,
    });
    if (!targetResult.ok) {
      return {
        handled: true,
        ok: false,
        error: targetResult.error === "The Excel range does not overlap an editable dataset cell."
          ? "The formula result does not overlap an editable dataset cell."
          : targetResult.error,
      };
    }
    const targets = targetResult.targets.map((target) => ({
      row: target.row,
      column: target.column,
      result_row: target.rowOffset,
      result_column: target.columnOffset,
    }));
    const values = targetResult.targets.map((target) => result.values[target.rowOffset][target.columnOffset]);

    const overlapping = linksForTargetCells(targets);
    if (overlapping.size) {
      links = links.filter((_link, index) => !overlapping.has(index));
    }
    onTargetsClaimed(targets.map((target) => ({ row: target.row, column: target.column })));
    let changedCount = 0;
    targets.forEach((target, index) => {
      const value = values[index];
      const previous = state.model.values[target.row][target.column];
      if (!valuesEqual(previous, value)) changedCount += 1;
      state.model.values[target.row][target.column] = value;
      state.dirty.set(targetCellKey(target), value);
    });
    links.push({ formula: classified.canonical, target_cells: targets });
    notifyInventoryChanged();
    return {
      handled: true,
      ok: true,
      changedCount,
      affectedCellCount: targets.length,
      reference: classified.canonical,
      message: `Calculated ${targets.length} dataset cell${targets.length === 1 ? "" : "s"} from the formula.`,
    };
  }

  function recordLinkFailures(link, error) {
    const failures = [];
    link.target_cells.forEach((target) => {
      const key = targetCellKey(target);
      if (!error) {
        failuresByTargetKey.delete(key);
        return;
      }
      const failure = {
        reference: link.formula,
        destination: targetDestinationLabel(state?.model, target),
        error,
      };
      failuresByTargetKey.set(key, failure);
      failures.push(failure);
    });
    return failures;
  }

  /**
   * Re-evaluate saved formulas and apply the current results. Each formula
   * evaluates on its own so one broken reference cannot fail the rest; a
   * link is refreshed whole or not at all.
   */
  async function refreshAll(ids = null, options = {}) {
    const markRefreshedCellsDirty = options?.markRefreshedCellsDirty === true;
    const requestedIds = Array.isArray(ids) ? requestedLinkIds(ids) : null;
    const scopedLinks = requestedIds
      ? links.filter((link) => requestedIds.has(linkRecordId(link)))
      : links.slice();
    if (!scopedLinks.length || !state?.model) {
      return { linkedCellCount: 0, changedCount: 0, failedCount: 0, failures: [] };
    }
    abort();
    const generation = requestGeneration;
    requestController = new AbortController();
    const signal = requestController.signal;
    let results;
    try {
      results = await Promise.all(scopedLinks.map((link) => {
        const parsed = parseDatasetFormula(link.formula);
        return parsed.ok ? evaluateFormula(parsed, generation, signal) : Promise.resolve(parsed);
      }));
    } finally {
      if (generation === requestGeneration) requestController = null;
    }
    if (generation !== requestGeneration || results.some((result) => result.stale)) {
      return { linkedCellCount: 0, changedCount: 0, failedCount: 0, failures: [], stale: true };
    }
    if (results.some((result) => result.aborted)) {
      return { linkedCellCount: 0, changedCount: 0, failedCount: 0, failures: [], aborted: true };
    }
    const failures = [];
    let linkedCellCount = 0;
    let changedCount = 0;
    let failedCount = 0;
    scopedLinks.forEach((link, index) => {
      const count = link.target_cells.length;
      linkedCellCount += count;
      const result = results[index];
      const validTargets = link.target_cells.every((target) => (
        state.model?.mask?.[target.row]?.[target.column] === true
        && Array.isArray(state.model?.values?.[target.row])
      ));
      let linkError = result.ok
        ? (validTargets ? "" : "The linked dataset cell is no longer part of this dataset.")
        : String(result.error || "The formula could not be evaluated.");
      if (!linkError && link.target_cells.some((target) => (
        target.result_row >= result.rows || target.result_column >= result.cols
      ))) {
        linkError = "The formula result is smaller than the range it fills.";
      }
      failures.push(...recordLinkFailures(link, linkError));
      if (linkError) {
        failedCount += count;
        return;
      }
      link.target_cells.forEach((target) => {
        const value = result.values[target.result_row][target.result_column];
        const previous = state.model.values[target.row][target.column];
        const changed = !valuesEqual(previous, value);
        if (changed) {
          changedCount += 1;
          state.model.values[target.row][target.column] = value;
        }
        if (changed || markRefreshedCellsDirty) {
          state.dirty.set(targetCellKey(target), value);
        }
      });
    });
    if (failures.length || failuresByTargetKey.size) targetDecorationIndex = null;
    return { linkedCellCount, changedCount, failedCount, failures };
  }

  function breakLinks(ids) {
    if (isReadOnly()) return { ok: false, error: "This dataset is read-only." };
    const requestedIds = requestedLinkIds(ids);
    const indexes = new Set();
    let affectedCellCount = 0;
    links.forEach((link, index) => {
      if (!requestedIds.has(linkRecordId(link))) return;
      indexes.add(index);
      affectedCellCount += link.target_cells.length;
    });
    if (!indexes.size) return { ok: false, error: "The formula link is no longer available." };
    const removed = removeLinkIndexes(indexes);
    return {
      ok: removed > 0,
      affectedCellCount,
      message: removed > 0
        ? `${removed === 1 ? "Link" : `${removed} links`} broken. Current dataset values are now hard-coded.`
        : "",
    };
  }

  function breakLink(id) {
    return breakLinks([id]);
  }

  // Whether the dataset holds any link at all, whatever the window is
  // showing. The Data tab asks this to decide whether a view that cannot
  // paint links has any to explain away.
  function hasLinks() {
    return links.length > 0;
  }

  return {
    abort,
    breakLink,
    breakLinks,
    clear,
    commitReference,
    decorateCell,
    hasLinks,
    hardCodeTargetCells,
    getCellLinkInfo,
    getLinkFailures: listFailures,
    isDirty,
    listRecords,
    load,
    markClean,
    refreshAll,
    restoreSaved,
    serialize,
  };
}
