/*
===============================================================================
Internal Dataset Links
Per-cell links from this dataset's editable grid into another dataset of the
same reserving class, the ArcRho sibling of the Excel `external_links`
controller. A link stores its standalone reference text plus the target cells
it owns, each mapped to a source cell of the referenced dataset:

    { reference: "=[C 82 - Prior Qtr Selected][1:6]",
      target_cells: [{ row, column, source_row, source_column }, ...] }

Values are snapshots taken when the link is committed or refreshed; the app
server resolves the reference (one dataset read per unique name) and this
controller spills the returned rectangle into the grid exactly like an Excel
range link.
===============================================================================
*/
import {
  applyDatasetLinkOutlineClasses,
  buildDatasetExternalLinkTargets,
  buildDatasetLinkOutline,
  describeTargetDestination,
} from "/ui/shared/dataset/dataset_external_links.js?v=20260907b";
import {
  formatInternalDatasetReference,
  parseInternalDatasetReference,
} from "/ui/shared/dataset/dataset_internal_reference.js?v=20260830a";

function targetCellKey(target) {
  return `${target.row},${target.column}`;
}

function nonnegativeInt(value) {
  const numeric = Number(value);
  return Number.isInteger(numeric) && numeric >= 0 ? numeric : null;
}

function cloneLinks(links) {
  return links.map((link) => ({
    reference: link.reference,
    target_cells: link.target_cells.map((target) => ({ ...target })),
  }));
}

function linksSnapshot(links) {
  return JSON.stringify(links);
}

function canonicalReference(value) {
  const parsed = parseInternalDatasetReference(value);
  return parsed.ok ? formatInternalDatasetReference(parsed) : "";
}

export function normalizeDatasetInternalLinks(value) {
  const source = Array.isArray(value) ? value : [];
  const normalized = [];
  const seenLinks = new Set();
  const ownedTargets = new Set();
  source.forEach((item) => {
    const reference = canonicalReference(item?.reference);
    if (!reference) return;
    const rawTargets = Array.isArray(item?.target_cells)
      ? item.target_cells
      : (Array.isArray(item?.targetCells) ? item.targetCells : []);
    const targetCells = [];
    const seenTargets = new Set();
    const seenSources = new Set();
    let invalidTargets = false;
    rawTargets.forEach((target) => {
      const row = nonnegativeInt(target?.row);
      const column = nonnegativeInt(target?.column);
      const sourceRow = nonnegativeInt(target?.source_row ?? target?.sourceRow);
      const sourceColumn = nonnegativeInt(target?.source_column ?? target?.sourceColumn);
      if (row === null || column === null || sourceRow === null || sourceColumn === null) {
        invalidTargets = true;
        return;
      }
      const key = `${row},${column}`;
      const sourceKey = `${sourceRow},${sourceColumn}`;
      if (seenTargets.has(key) || seenSources.has(sourceKey)) {
        invalidTargets = true;
        return;
      }
      seenTargets.add(key);
      seenSources.add(sourceKey);
      targetCells.push({ row, column, source_row: sourceRow, source_column: sourceColumn });
    });
    if (invalidTargets || !targetCells.length) return;
    const linkKey = `${reference}${targetCells.map(targetCellKey).join(";")}`;
    if (seenLinks.has(linkKey)) return;
    if (targetCells.some((target) => ownedTargets.has(targetCellKey(target)))) return;
    seenLinks.add(linkKey);
    targetCells.forEach((target) => ownedTargets.add(targetCellKey(target)));
    normalized.push({ reference, target_cells: targetCells });
  });
  return normalized;
}

function displayToActualCell(row, column, transposed) {
  return transposed ? { row: column, column: row } : { row, column };
}

function actualToDisplayCell(row, column, transposed) {
  return transposed ? { row: column, column: row } : { row, column };
}

function referenceDatasetName(reference) {
  const parsed = parseInternalDatasetReference(reference);
  return parsed.ok ? parsed.datasetName : "";
}

function referenceCoordinateText(reference) {
  const parsed = parseInternalDatasetReference(reference);
  if (!parsed.ok) return "";
  const formatted = formatInternalDatasetReference(parsed);
  const open = formatted.lastIndexOf("][");
  return open >= 0 ? formatted.slice(open + 2, -1) : "";
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
  return `${link.reference}${link.target_cells.map(targetCellKey).join(";")}`;
}

export function createDatasetInternalLinksController({
  state,
  resolveReferences,
  isReadOnly = () => false,
  isTransposed = () => false,
  isAtLinkedShape = () => true,
  onInventoryChanged = () => {},
  onTargetsClaimed = () => {},
} = {}) {
  let links = [];
  let savedLinks = [];
  let requestGeneration = 0;
  let targetDecorationIndex = null;
  // Target cell key -> the reference failure that left its stored value in
  // place, so a re-render repaints the same cells red without another resolve.
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
  }

  function load(value) {
    abort();
    failuresByTargetKey = new Map();
    links = normalizeDatasetInternalLinks(value);
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
    links = normalizeDatasetInternalLinks(value);
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
      reference: link.reference,
      datasetName: referenceDatasetName(link.reference),
      sourceKind: "internal",
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
    // A blank the mask left inside a linked rectangle still carries that
    // rectangle's edge, so the frame closes across the empty corner.
    const outline = decoration?.isRange ? decoration : index.outlineGaps.get(key);
    cell.classList.toggle("arInternalLinkCell", !!link || !!outline);
    cell.classList.toggle("arInternalLinkErrorCell", !!failure);
    applyDatasetLinkOutlineClasses(cell, outline, "internal");
    if (link) {
      cell.dataset.internalLinkReference = link.reference;
    } else {
      delete cell.dataset.internalLinkReference;
    }
  }

  // The Value and Destination columns are read off the grid on screen. A view
  // at other lengths is not the grid the link points into, so they are left
  // blank rather than quoting a number and a period the link never named.
  function listRecords() {
    const atLinkedShape = !!isAtLinkedShape();
    return links.map((link) => {
      const targets = link.target_cells;
      return {
        id: linkRecordId(link),
        datasetName: referenceDatasetName(link.reference),
        sourceRange: referenceCoordinateText(link.reference),
        reference: link.reference,
        value: atLinkedShape ? targetValuePreview(state?.model, targets, targets.length > 1) : "",
        destination: (atLinkedShape ? describeTargetDestination(state?.model, targets) : "") || "Data",
        affectedCellCount: targets.length,
        readOnly: !!isReadOnly(),
      };
    });
  }

  async function resolveSingleReference(reference, generation) {
    let resp;
    try {
      resp = await resolveReferences([reference]);
    } catch (error) {
      return { ok: false, error: String(error?.message || error || "Dataset link resolve failed.") };
    }
    if (generation !== requestGeneration) return { ok: false, stale: true };
    if (!resp?.ok) {
      return {
        ok: false,
        error: String(resp?.data?.detail || resp?.data?.error || "The dataset reference could not be resolved."),
      };
    }
    const result = Array.isArray(resp.data?.results) ? resp.data.results[0] : null;
    if (!result || !Array.isArray(result.cells)) {
      return { ok: false, error: "The dataset reference could not be resolved." };
    }
    return { ok: true, result };
  }

  async function commitReference({ displayRow, displayColumn, reference } = {}) {
    if (isReadOnly()) return { handled: true, ok: false, error: "This dataset is read-only." };
    const parsed = parseInternalDatasetReference(reference);
    if (!parsed.ok) return { handled: true, ok: false, error: parsed.error };
    const canonical = formatInternalDatasetReference(parsed);

    abort();
    const generation = requestGeneration;
    const resolved = await resolveSingleReference(canonical, generation);
    if (!resolved.ok) return { handled: true, ...resolved };
    const { result } = resolved;
    const rowCount = Number(result.row_count) || 0;
    const columnCount = Number(result.column_count) || 0;
    const targetResult = buildDatasetExternalLinkTargets({
      model: state?.model,
      transposed: !!isTransposed(),
      startRow: displayRow,
      startColumn: displayColumn,
      rowCount,
      columnCount,
    });
    if (!targetResult.ok) {
      return {
        handled: true,
        ok: false,
        error: targetResult.error === "The Excel range does not overlap an editable dataset cell."
          ? "The referenced range does not overlap an editable dataset cell."
          : targetResult.error,
      };
    }
    const rowStart = Number(result.row_start) || 0;
    const columnStart = Number(result.column_start) || 0;
    const targets = [];
    const values = [];
    for (const target of targetResult.targets) {
      const cell = result.cells[target.rowOffset * columnCount + target.columnOffset];
      if (!cell) return { handled: true, ok: false, error: "The referenced range mapping is invalid." };
      targets.push({
        row: target.row,
        column: target.column,
        source_row: rowStart + target.rowOffset,
        source_column: columnStart + target.columnOffset,
      });
      values.push(cell.value ?? null);
    }

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
    links.push({ reference: canonical, target_cells: targets });
    notifyInventoryChanged();
    return {
      handled: true,
      ok: true,
      changedCount,
      affectedCellCount: targets.length,
      reference: canonical,
      message: `Linked ${targets.length} dataset cell${targets.length === 1 ? "" : "s"} to ${referenceDatasetName(canonical) || "another dataset"}.`,
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
        reference: link.reference,
        datasetName: referenceDatasetName(link.reference),
        destination: targetDestinationLabel(state?.model, target),
        error,
      };
      failuresByTargetKey.set(key, failure);
      failures.push(failure);
    });
    return failures;
  }

  /**
   * Re-resolve saved links and apply the current source values. Each link
   * resolves in its own request so one broken reference (a deleted or renamed
   * source dataset) cannot fail the rest of the batch; a link is refreshed
   * whole or not at all.
   */
  async function refreshAll(ids = null, options = {}) {
    const markRefreshedCellsDirty = options?.markRefreshedCellsDirty === true;
    const requestedIds = Array.isArray(ids)
      ? new Set(ids.map((id) => String(id || "")).filter(Boolean))
      : null;
    const scopedLinks = requestedIds
      ? links.filter((link) => requestedIds.has(linkRecordId(link)))
      : links.slice();
    if (!scopedLinks.length || !state?.model) {
      return { linkedCellCount: 0, changedCount: 0, failedCount: 0, failures: [] };
    }
    abort();
    const generation = requestGeneration;
    const resolutions = await Promise.all(
      scopedLinks.map((link) => resolveSingleReference(link.reference, generation)),
    );
    if (generation !== requestGeneration) {
      return { linkedCellCount: 0, changedCount: 0, failedCount: 0, failures: [], stale: true };
    }
    const failures = [];
    let linkedCellCount = 0;
    let changedCount = 0;
    let failedCount = 0;
    scopedLinks.forEach((link, index) => {
      const count = link.target_cells.length;
      linkedCellCount += count;
      const resolution = resolutions[index];
      const result = resolution.ok ? resolution.result : null;
      const rowStart = Number(result?.row_start) || 0;
      const columnStart = Number(result?.column_start) || 0;
      const columnCount = Number(result?.column_count) || 0;
      const validTargets = link.target_cells.every((target) => (
        state.model?.mask?.[target.row]?.[target.column] === true
        && Array.isArray(state.model?.values?.[target.row])
      ));
      const nextValues = [];
      let linkError = resolution.ok
        ? (validTargets ? "" : "The linked dataset cell is no longer part of this dataset.")
        : String(resolution.error || "The dataset reference could not be resolved.");
      if (!linkError && result) {
        for (const target of link.target_cells) {
          const rowOffset = target.source_row - rowStart;
          const columnOffset = target.source_column - columnStart;
          const cell = rowOffset >= 0
            && columnOffset >= 0
            && columnOffset < columnCount
            ? result.cells[rowOffset * columnCount + columnOffset]
            : null;
          if (
            !cell
            || Number(cell.row) !== target.source_row
            || Number(cell.column) !== target.source_column
          ) {
            linkError = "The referenced cells are no longer part of the source dataset.";
            break;
          }
          nextValues.push(cell.value ?? null);
        }
      }
      failures.push(...recordLinkFailures(link, linkError));
      if (linkError) {
        failedCount += count;
        return;
      }
      link.target_cells.forEach((target, targetIndex) => {
        const value = nextValues[targetIndex];
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
    const requestedIds = new Set(
      (Array.isArray(ids) ? ids : [ids]).map((id) => String(id || "")).filter(Boolean),
    );
    const indexes = new Set();
    let affectedCellCount = 0;
    links.forEach((link, index) => {
      if (!requestedIds.has(linkRecordId(link))) return;
      indexes.add(index);
      affectedCellCount += link.target_cells.length;
    });
    if (!indexes.size) return { ok: false, error: "The dataset link is no longer available." };
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
