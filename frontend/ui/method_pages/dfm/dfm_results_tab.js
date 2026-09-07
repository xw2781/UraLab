/*
===============================================================================
DFM Results Tab - results table rendering and CSV export
===============================================================================
*/
import { getDataset } from "/ui/shared/dataset/dataset_api.js";
import { formatCellValue } from "/ui/shared/tabs/data/dataset_grid_view.js?v=20260907c";
import { renderDatasetGridPlaceholder } from "/ui/shared/tabs/data/dataset_grid_placeholder.js?v=20260809a";
import { formatDatasetNumberValue } from "/ui/shared/dataset/dataset_number_format.js";
import { openDatasetNamePicker } from "/ui/shared/components/pickers/dataset_name_picker.js";
import { openContextMenu } from "/ui/shared/components/context_menu/context_menu.js";
import { wireSelectableTable } from "/ui/shared/components/spreadsheet/table_selection.js?v=20260726a";
import {
  state,
  getEffectiveDevLabelsForModel,
  getRatioHeaderLabels,
  getCumulativeFactors,
  getLatestRowValue,
  roundRatio,
  ensureDefaultSummarySelectionForColumns,
  getOriginLabelTextForRatio,
  getResolvedProjectName,
  getResolvedReservingClass,
  escapeCsvCell,
  isResultsTabVisible,
  markDfmDirty,
} from "/ui/method_pages/dfm/dfm_state.js";

let ratioBasisControlsWired = false;
let resultsCopyMenuWired = false;
let ratioBasisOptionsLoadSeq = 0;
let ratioBasisColumnLoadSeq = 0;
let ratioBasisSelectedName = "";
let ratioBasisSelectedFormat = "";
let ratioBasisEmbeddedSnapshot = Boolean(
  new URLSearchParams(globalThis.location?.search || "").get("method_name"),
);
let ratioBasisNumberFormat = "";
let ratioBasisDecimalPlaces = null;
let ratioBasisSourceRevision = "";
let persistedUltimateVector = null;
let ratioBasisProgrammaticUpdate = false;
let ultimateRatioDecimalProgrammaticUpdate = false;
let ratioBasisOptionsRenderedProjectKey = "";
const ratioBasisOptionsByProject = new Map();
const ratioBasisOptionsInFlightByProject = new Map();
const ratioBasisColumnCache = new Map();
let ratioBasisColumnLoadPromise = Promise.resolve();
let ratioBasisColumnState = {
  requestKey: "",
  status: "idle", // idle | loading | ready | error
  datasetName: "",
  dataFormat: "",
  headerText: "",
  valuesByOrigin: new Map(),
  valuesByIndex: [],
  error: "",
};

function wireResultsCopyMenu() {
  if (resultsCopyMenuWired) return;
  resultsCopyMenuWired = true;
  const menu = document.getElementById("ctxMenu");
  if (!menu) return;
  menu.addEventListener("click", async (event) => {
    const btn = event.target?.closest?.(".ctx-item");
    if (!btn) return;
    if (btn.dataset.action === "copy_value" && typeof window.__arcRhoCopyActiveGridSelection === "function") {
      await window.__arcRhoCopyActiveGridSelection();
    }
    menu.style.display = "none";
  });
  document.addEventListener("mousedown", (event) => {
    if (!menu.contains(event.target)) menu.style.display = "none";
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") menu.style.display = "none";
  });
}

function toText(value) {
  return String(value ?? "").trim();
}

function normalizeKey(value) {
  return toText(value).toLowerCase();
}

function getRatioBasisInputEl() {
  return document.getElementById("dfmRatioBasisInput");
}

function getRatioBasisBtnEl() {
  return document.getElementById("dfmRatioBasisBtn");
}

function measureRatioBasisTextWidth(input, text) {
  if (!text) return 0;
  const canvas = measureRatioBasisTextWidth._canvas
    || (measureRatioBasisTextWidth._canvas = document.createElement("canvas"));
  const ctx = canvas.getContext("2d");
  if (!ctx) return text.length * 7;
  const cs = getComputedStyle(input);
  ctx.font = `${cs.fontStyle} ${cs.fontVariant} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
  const measured = Number(ctx.measureText(text).width || 0);
  return Number.isFinite(measured) ? measured : text.length * 7;
}

function syncRatioBasisInputWidth() {
  const input = getRatioBasisInputEl();
  const wrap = input?.closest(".dfmResultsPickerWrap");
  if (!input || !wrap) return;
  const text = toText(input.value) || input.placeholder || "";
  const cs = getComputedStyle(input);
  const chrome = (Number.parseFloat(cs.paddingLeft) || 0)
    + (Number.parseFloat(cs.paddingRight) || 0)
    + (Number.parseFloat(cs.borderLeftWidth) || 0)
    + (Number.parseFloat(cs.borderRightWidth) || 0);
  wrap.style.width = `${Math.ceil(measureRatioBasisTextWidth(input, text) + chrome + 6)}px`;
}

function getRatioBasisStatusEl() {
  return document.getElementById("dfmRatioBasisStatus");
}

function getUltimateRatioDecimalInputEl() {
  return document.getElementById("dfmUltimateRatioDecimalPlacesInput");
}

function setRatioBasisStatus(message = "", tone = "") {
  const el = getRatioBasisStatusEl();
  if (!el) return;
  el.textContent = String(message || "");
  el.classList.remove("is-error", "is-loading");
  if (tone === "error") el.classList.add("is-error");
  if (tone === "loading") el.classList.add("is-loading");
}

function buildRatioBasisHeaderText(datasetName) {
  const name = toText(datasetName);
  return name ? `${name}` : "Ratio Basis";
}

function formatPercentCellValue(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return "";
  const normalized = Math.abs(n) < 0.0000005 ? 0 : n;
  return `${(normalized * 100).toFixed(getResultsUltimateRatioDecimalPlaces())}%`;
}

function getResultsUltimateRatioDecimalPlaces() {
  const input = getUltimateRatioDecimalInputEl();
  const raw = Number.parseInt(String(input?.value ?? "").trim(), 10);
  if (!Number.isFinite(raw)) return 2;
  return Math.max(0, Math.min(6, raw));
}

function normalizeUltimateRatioDecimalInput() {
  const input = getUltimateRatioDecimalInputEl();
  if (!input) return { changed: false, value: 2 };
  const normalized = String(getResultsUltimateRatioDecimalPlaces());
  const changed = String(input.value ?? "") !== normalized;
  if (changed) input.value = normalized;
  return { changed, value: Number.parseInt(normalized, 10) || 2 };
}

function getDatasetTypeColumnIndexes(columns) {
  const indexByName = {};
  for (let i = 0; i < columns.length; i += 1) {
    const key = normalizeKey(columns[i]);
    if (!key || indexByName[key] != null) continue;
    indexByName[key] = i;
  }
  return {
    name: indexByName.name,
    dataFormat: indexByName["data_format"],
    calculated: indexByName.calculated,
  };
}

function getDatasetTypeCell(row, index, fallbackKeys) {
  if (Array.isArray(row)) {
    if (Number.isInteger(index) && index >= 0) return row[index];
    return "";
  }
  if (row && typeof row === "object") {
    for (const key of fallbackKeys) {
      if (Object.prototype.hasOwnProperty.call(row, key)) return row[key];
    }
  }
  return "";
}

function parseCalculatedFlag(value) {
  if (typeof value === "boolean") return value;
  const text = normalizeKey(value);
  return text === "true" || text === "1" || text === "yes" || text === "y";
}

function extractRatioBasisDatasetOptions(data) {
  const columns = Array.isArray(data?.columns) ? data.columns : [];
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  const indexes = getDatasetTypeColumnIndexes(columns);
  const out = [];
  const seen = new Set();

  for (const row of rows) {
    const name = toText(getDatasetTypeCell(row, indexes.name, ["Name", "name"]));
    if (!name) continue;
    const dataFormat = normalizeKey(
      getDatasetTypeCell(row, indexes.dataFormat, ["Data Format", "dataFormat", "data_format"]),
    );
    const calculated = parseCalculatedFlag(
      getDatasetTypeCell(row, indexes.calculated, ["Calculated", "calculated"]),
    );
    const key = normalizeKey(name);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push({ name, dataFormat, calculated });
  }

  out.sort((a, b) => a.name.localeCompare(b.name));
  return out;
}

function markRatioBasisOptionsRendered(projectKey) {
  ratioBasisOptionsRenderedProjectKey = String(projectKey || "");
}

function formatRatioBasisCellValue(value) {
  if (!Number.isFinite(Number(value))) return "";
  if (!ratioBasisNumberFormat) return formatCellValue(value);
  return formatDatasetNumberValue(
    value,
    ratioBasisNumberFormat,
    Number.isFinite(ratioBasisDecimalPlaces) ? ratioBasisDecimalPlaces : 0,
  );
}

function syncResultsSelectionState(table, ranges = []) {
  if (!table) return;
  const selectedRows = new Set();
  const selectedCols = new Set();
  const normalizedRanges = Array.isArray(ranges) ? ranges : [];
  table.querySelectorAll("td[data-r][data-c]").forEach((cell) => {
    const row = Number(cell.dataset.r);
    const col = Number(cell.dataset.c);
    const selected = normalizedRanges.some((range) => (
      row >= range.r0 && row <= range.r1 && col >= range.c0 && col <= range.c1
    ));
    if (selected) {
      cell.setAttribute("aria-selected", "true");
      selectedRows.add(row);
      selectedCols.add(col);
    } else {
      cell.removeAttribute("aria-selected");
    }
  });
  table.querySelectorAll("th.arSpreadsheetSelectedLabel").forEach((header) => {
    header.classList.remove("arSpreadsheetSelectedLabel");
  });
  selectedRows.forEach((row) => {
    table.querySelector(`tbody th[data-r="${row}"]`)?.classList.add("arSpreadsheetSelectedLabel");
  });
  selectedCols.forEach((col) => {
    table.querySelector(`thead th[data-c="${col}"]`)?.classList.add("arSpreadsheetSelectedLabel");
  });
}

async function ensureRatioBasisOptionsForCurrentProject(options = {}) {
  const projectName = getResolvedProjectName();
  const projectKey = normalizeKey(projectName);
  if (!projectKey) {
    markRatioBasisOptionsRendered("");
    return [];
  }

  if (!options?.forceReload && ratioBasisOptionsByProject.has(projectKey)) {
    const cached = ratioBasisOptionsByProject.get(projectKey) || [];
    if (ratioBasisOptionsRenderedProjectKey !== projectKey) {
      markRatioBasisOptionsRendered(projectKey);
    }
    return cached;
  }

  if (!options?.forceReload && ratioBasisOptionsInFlightByProject.has(projectKey)) {
    return ratioBasisOptionsInFlightByProject.get(projectKey);
  }

  const seq = ++ratioBasisOptionsLoadSeq;
  if (toText(getRatioBasisInputEl()?.value)) {
    setRatioBasisStatus("Loading ratio-basis options...", "loading");
  }

  const loadPromise = (async () => {
    const response = await fetch(`/dataset_types?project_name=${encodeURIComponent(projectName)}`);
    if (!response.ok) {
      let detail = "";
      try {
        detail = toText(await response.text());
      } catch {}
      throw new Error(detail || `Failed to load dataset types (${response.status})`);
    }
    const payload = await response.json().catch(() => ({}));
    const rows = extractRatioBasisDatasetOptions(payload?.data || {});
    ratioBasisOptionsByProject.set(projectKey, rows);

    if (seq === ratioBasisOptionsLoadSeq && normalizeKey(getResolvedProjectName()) === projectKey) {
      markRatioBasisOptionsRendered(projectKey);
    }
    return rows;
  })();
  ratioBasisOptionsInFlightByProject.set(projectKey, loadPromise);
  try {
    return await loadPromise;
  } finally {
    if (ratioBasisOptionsInFlightByProject.get(projectKey) === loadPromise) {
      ratioBasisOptionsInFlightByProject.delete(projectKey);
    }
  }
}

function findRatioBasisOption(projectName, datasetName) {
  const projectKey = normalizeKey(projectName);
  const nameKey = normalizeKey(datasetName);
  if (!projectKey || !nameKey) return null;
  const list = ratioBasisOptionsByProject.get(projectKey) || [];
  return list.find((item) => normalizeKey(item.name) === nameKey) || null;
}

function clearRatioBasisColumnState(options = {}) {
  ratioBasisColumnState = {
    requestKey: "",
    status: "idle",
    datasetName: "",
    dataFormat: "",
    headerText: "",
    valuesByOrigin: new Map(),
    valuesByIndex: [],
    error: "",
  };
  if (!options?.keepStatus) setRatioBasisStatus("", "");
}

function applyRatioBasisColumnMetadata(columnState) {
  ratioBasisNumberFormat = toText(columnState?.numberFormat || columnState?.number_format);
  ratioBasisDecimalPlaces = Number.isFinite(Number(columnState?.decimalPlaces ?? columnState?.decimal_places))
    ? Number(columnState?.decimalPlaces ?? columnState?.decimal_places)
    : null;
  ratioBasisSourceRevision = toText(
    columnState?.sourceRevision || columnState?.source_revision || columnState?.revision,
  );
}

function clonePersistedNumber(value) {
  if (value == null || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

/**
 * Hydrates Results directly from the canonical v2 method payload. This path is
 * deliberately disk-free: it does not validate the Ratio Basis against the
 * project index or reload the source dataset.
 */
export function applyPersistedResultsSnapshot(resultsTab = {}) {
  const source = resultsTab && typeof resultsTab === "object" ? resultsTab : {};
  const datasetName = toText(source["ratio_basis_dataset"]);
  const originLabels = Array.isArray(source["ratio_basis_origin_labels"])
    ? source["ratio_basis_origin_labels"].map((label) => String(label ?? ""))
    : [];
  const values = Array.isArray(source["ratio_basis_values"])
    ? source["ratio_basis_values"].map(clonePersistedNumber)
    : [];
  const input = getRatioBasisInputEl();
  ratioBasisEmbeddedSnapshot = true;
  ratioBasisSelectedName = datasetName;
  ratioBasisSelectedFormat = toText(source["ratio_basis_data_format"] || "vector").toLowerCase();
  ratioBasisNumberFormat = toText(source["ratio_basis_number_format"]);
  ratioBasisDecimalPlaces = Number.isFinite(Number(source["ratio_basis_decimal_places"]))
    ? Number(source["ratio_basis_decimal_places"])
    : null;
  ratioBasisSourceRevision = toText(source["ratio_basis_source_revision"]);
  if (input) input.value = datasetName;
  syncRatioBasisInputWidth();

  const valuesByOrigin = new Map();
  originLabels.forEach((label, index) => {
    const value = values[index];
    const key = normalizeOriginKey(label);
    if (key && Number.isFinite(value)) valuesByOrigin.set(key, value);
  });
  ratioBasisColumnState = {
    requestKey: `embedded::${datasetName}::${ratioBasisSourceRevision}`,
    status: datasetName ? "ready" : "idle",
    datasetName,
    dataFormat: ratioBasisSelectedFormat,
    headerText: buildRatioBasisHeaderText(datasetName),
    valuesByOrigin,
    valuesByIndex: values,
    originLabels,
    error: "",
  };
  persistedUltimateVector = Array.isArray(source["ultimate_vector"])
    ? source["ultimate_vector"].map(clonePersistedNumber)
    : null;
  setRatioBasisStatus("", "");
}

export function getResultsRatioBasisSnapshot() {
  const origins = Array.isArray(ratioBasisColumnState.originLabels)
    ? ratioBasisColumnState.originLabels.slice()
    : Array.from(ratioBasisColumnState.valuesByOrigin?.keys?.() || []);
  const values = Array.isArray(ratioBasisColumnState.valuesByIndex)
    ? ratioBasisColumnState.valuesByIndex.slice()
    : [];
  return {
    "ratio_basis_dataset": toText(ratioBasisSelectedName),
    "ratio_basis_data_format": toText(ratioBasisSelectedFormat),
    "ratio_basis_origin_labels": origins,
    "ratio_basis_values": values,
    "ratio_basis_number_format": ratioBasisNumberFormat,
    "ratio_basis_decimal_places": ratioBasisDecimalPlaces,
    "ratio_basis_source_revision": ratioBasisSourceRevision,
  };
}

export function invalidatePersistedResultsDerivations() {
  persistedUltimateVector = null;
}

function getDfmOriginLabels() {
  return Array.isArray(state?.model?.origin_labels)
    ? state.model.origin_labels.map((label) => String(label ?? ""))
    : [];
}

function matchesOriginLabels(labels, origins) {
  if (!Array.isArray(labels) || labels.length !== origins.length) return false;
  return origins.every((label, index) => String(labels[index] ?? "") === label);
}

/**
 * The embedded snapshot stays authoritative only while it still describes the
 * DFM's own origins. Changing Origin Length rebuilds the input triangle on a
 * new origin basis, so the saved column no longer lines up and must be re-read
 * from its source dataset at the current basis; the method contract requires
 * the Ratio Basis labels to equal the DFM origins exactly.
 */
function dropEmbeddedRatioBasisSnapshotOnOriginChange() {
  if (!ratioBasisEmbeddedSnapshot || !ratioBasisSelectedName) return;
  const origins = getDfmOriginLabels();
  if (!origins.length) return;
  if (matchesOriginLabels(ratioBasisColumnState.originLabels, origins)) return;
  ratioBasisEmbeddedSnapshot = false;
  clearRatioBasisColumnState();
}

/**
 * Settles the Ratio Basis column against the DFM's current origins so a save
 * or preview built right after an Origin Length change carries an aligned
 * column instead of the previous basis.
 */
export async function ensureResultsRatioBasisAligned() {
  dropEmbeddedRatioBasisSnapshotOnOriginChange();
  const datasetName = toText(ratioBasisSelectedName);
  if (!datasetName) return { ok: true };
  if (!ratioBasisEmbeddedSnapshot) {
    await ensureRatioBasisOptionsForCurrentProject().catch(() => {});
    // A settled load can queue the next one when the origin basis moved again
    // while this one was in flight; a few passes converge without looping.
    for (let attempt = 0; attempt < 3; attempt += 1) {
      queueRatioBasisColumnLoadIfNeeded();
      if (ratioBasisColumnState.status !== "loading") break;
      await ratioBasisColumnLoadPromise;
    }
  }
  const origins = getDfmOriginLabels();
  const snapshotOrigins = getResultsRatioBasisSnapshot()["ratio_basis_origin_labels"];
  if (origins.length && !matchesOriginLabels(snapshotOrigins, origins)) {
    const detail = toText(ratioBasisColumnState.error)
      || "Reselect or clear it in the Results tab, then save.";
    return {
      ok: false,
      error: `Ratio Basis "${datasetName}" does not line up with the current origins. ${detail}`,
    };
  }
  return { ok: true };
}

function readCurrentOriginLen() {
  const raw = Number.parseInt(document.getElementById("originLenSelect")?.value, 10);
  return Number.isFinite(raw) ? raw : 12;
}

function readCurrentCumulativeFlag() {
  return !!document.getElementById("cumulativeChk")?.checked;
}

function normalizeOriginKey(value) {
  return String(value ?? "").trim();
}

function buildRatioBasisRequestContext() {
  const datasetName = toText(ratioBasisSelectedName);
  if (!datasetName) return null;

  const projectName = getResolvedProjectName();
  const reservingClass = getResolvedReservingClass();
  const projectKey = normalizeKey(projectName);
  if (!projectKey || !reservingClass) return null;

  const option = findRatioBasisOption(projectName, datasetName);
  if (!option && ratioBasisOptionsByProject.has(projectKey)) {
    return null;
  }
  const dataFormat = option?.dataFormat || normalizeKey(ratioBasisSelectedFormat);
  if (!dataFormat) return null;

  const originLen = readCurrentOriginLen();
  const cumulative = readCurrentCumulativeFlag();
  // Ratio-basis extraction only needs origin granularity alignment with Results rows.
  // Use DevelopmentLength = OriginLength to request a full diagonal at that origin basis.
  const devLen = originLen;
  const requestKey = [
    normalizeKey(projectName),
    normalizeKey(reservingClass),
    normalizeKey(datasetName),
    dataFormat,
    String(originLen),
    String(cumulative ? 1 : 0),
  ].join("||");

  return {
    requestKey,
    projectName,
    reservingClass,
    datasetName: option?.name || datasetName,
    dataFormat,
    originLen,
    devLen,
    cumulative,
    headerText: buildRatioBasisHeaderText(option?.name || datasetName),
  };
}

function extractTriangleRowLatestValue(model, rowIndex) {
  const vals = Array.isArray(model?.values) ? model.values : [];
  const mask = Array.isArray(model?.mask) ? model.mask : [];
  const rowVals = Array.isArray(vals[rowIndex]) ? vals[rowIndex] : [];
  const maxCol = rowVals.length - 1;
  if (maxCol < 0) return null;
  const latest = getLatestRowValue(vals, mask, rowIndex, maxCol);
  return latest?.value ?? null;
}

function extractVectorRowValue(model, rowIndex) {
  const vals = Array.isArray(model?.values) ? model.values : [];
  const mask = Array.isArray(model?.mask) ? model.mask : [];
  const rowVals = Array.isArray(vals[rowIndex]) ? vals[rowIndex] : [];
  const rowMask = Array.isArray(mask[rowIndex]) ? mask[rowIndex] : null;
  for (let c = 0; c < rowVals.length; c += 1) {
    if (rowMask && !rowMask[c]) continue;
    const raw = rowVals[c];
    const n = typeof raw === "number" ? raw : Number(raw);
    if (Number.isFinite(n)) return n;
  }
  return null;
}

function extractRatioBasisColumnFromModel(model, ctx) {
  const origins = Array.isArray(model?.origin_labels) ? model.origin_labels : [];
  const valuesByOrigin = new Map();
  const valuesByIndex = [];

  for (let r = 0; r < origins.length; r += 1) {
    const value = ctx.dataFormat === "triangle"
      ? extractTriangleRowLatestValue(model, r)
      : extractVectorRowValue(model, r);
    valuesByIndex.push(Number.isFinite(value) ? value : null);
    const key = normalizeOriginKey(origins[r]);
    if (key && Number.isFinite(value)) valuesByOrigin.set(key, value);
  }

  return { valuesByOrigin, valuesByIndex };
}

function getRatioBasisRowValue(stateLike, originLabel, rowIndex) {
  if (!stateLike || stateLike.status !== "ready") return null;
  const originKey = normalizeOriginKey(originLabel);
  if (originKey && stateLike.valuesByOrigin instanceof Map && stateLike.valuesByOrigin.has(originKey)) {
    return stateLike.valuesByOrigin.get(originKey);
  }
  if (Array.isArray(stateLike.originLabels) && stateLike.originLabels.length) return null;
  if (Array.isArray(stateLike.valuesByIndex) && rowIndex >= 0 && rowIndex < stateLike.valuesByIndex.length) {
    return stateLike.valuesByIndex[rowIndex];
  }
  return null;
}

async function loadRatioBasisColumnForContext(ctx) {
  if (ctx.dataFormat !== "triangle" && ctx.dataFormat !== "vector") {
    throw new Error(`Ratio Basis supports Triangle/Vector only (selected: ${ctx.dataFormat || "unknown"}).`);
  }
  const payload = {
    Path: ctx.reservingClass,
    TriangleName: ctx.datasetName,
    ProjectName: ctx.projectName,
    Cumulative: ctx.cumulative,
    OriginLength: ctx.originLen,
    DevelopmentLength: ctx.devLen,
  };

  const arcrhoResp = await fetch("/arcrho/tri", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const arcrhoData = await arcrhoResp.json().catch(() => ({}));
  if (!arcrhoResp.ok) {
    throw new Error(toText(arcrhoData?.detail || arcrhoData?.error || arcrhoData?.message)
      || `Ratio basis request failed (${arcrhoResp.status}).`);
  }
  if (!arcrhoData?.ok || !toText(arcrhoData?.ds_id)) {
    throw new Error(toText(arcrhoData?.detail || arcrhoData?.error || arcrhoData?.message || arcrhoData?.status)
      || "Ratio basis dataset timed out or CSV was not available.");
  }

  const dsResp = await getDataset(arcrhoData.ds_id, {
    projectName: ctx.projectName,
    originLength: ctx.originLen,
  });
  if (!dsResp.ok) {
    throw new Error(toText(dsResp.data?.detail || dsResp.data?.error || dsResp.data?.message)
      || `Failed to load ratio basis dataset (${dsResp.status}).`);
  }

  const extracted = extractRatioBasisColumnFromModel(dsResp.data, ctx);
  return {
    requestKey: ctx.requestKey,
    status: "ready",
    datasetName: ctx.datasetName,
    dataFormat: ctx.dataFormat,
    headerText: ctx.headerText,
    valuesByOrigin: extracted.valuesByOrigin,
    valuesByIndex: extracted.valuesByIndex,
    originLabels: Array.isArray(dsResp.data?.origin_labels) ? dsResp.data.origin_labels.slice() : [],
    numberFormat: dsResp.data?.number_format,
    decimalPlaces: dsResp.data?.decimal_places,
    sourceRevision: dsResp.data?.source_revision || dsResp.data?.revision || dsResp.data?.sidecar_revision,
    error: "",
  };
}

function queueRatioBasisColumnLoadIfNeeded() {
  const currentProjectKey = normalizeKey(getResolvedProjectName());
  const ctx = buildRatioBasisRequestContext();
  if (!ctx) {
    if (!ratioBasisSelectedName) {
      clearRatioBasisColumnState();
    } else if (currentProjectKey && ratioBasisOptionsByProject.has(currentProjectKey)) {
      clearRatioBasisColumnState({ keepStatus: true });
      setRatioBasisStatus("Selected Ratio Basis dataset is not available in current project.", "error");
    }
    return null;
  }

  if (ratioBasisColumnState.requestKey === ctx.requestKey && ratioBasisColumnState.status === "ready") {
    return ctx;
  }

  const cached = ratioBasisColumnCache.get(ctx.requestKey);
  if (cached) {
    ratioBasisColumnState = { ...cached, status: "ready", error: "" };
    applyRatioBasisColumnMetadata(ratioBasisColumnState);
    setRatioBasisStatus("", "");
    return ctx;
  }

  if (ratioBasisColumnState.requestKey === ctx.requestKey && ratioBasisColumnState.status === "loading") {
    return ctx;
  }

  ratioBasisColumnState = {
    requestKey: ctx.requestKey,
    status: "loading",
    datasetName: ctx.datasetName,
    dataFormat: ctx.dataFormat,
    headerText: ctx.headerText,
    valuesByOrigin: new Map(),
    valuesByIndex: [],
    error: "",
  };
  setRatioBasisStatus(`Loading ${ctx.datasetName}...`, "loading");

  const seq = ++ratioBasisColumnLoadSeq;
  ratioBasisColumnLoadPromise = (async () => {
    try {
      const loaded = await loadRatioBasisColumnForContext(ctx);
      ratioBasisColumnCache.set(ctx.requestKey, loaded);
      if (seq !== ratioBasisColumnLoadSeq) return;
      const latestCtx = buildRatioBasisRequestContext();
      if (!latestCtx || latestCtx.requestKey !== ctx.requestKey) return;
      ratioBasisColumnState = loaded;
      applyRatioBasisColumnMetadata(loaded);
      setRatioBasisStatus("", "");
      if (isResultsTabVisible()) renderResultsTable();
    } catch (err) {
      if (seq !== ratioBasisColumnLoadSeq) return;
      const latestCtx = buildRatioBasisRequestContext();
      if (!latestCtx || latestCtx.requestKey !== ctx.requestKey) return;
      const message = toText(err?.message) || "Failed to load ratio basis dataset.";
      ratioBasisColumnState = {
        requestKey: ctx.requestKey,
        status: "error",
        datasetName: ctx.datasetName,
        dataFormat: ctx.dataFormat,
        headerText: ctx.headerText,
        valuesByOrigin: new Map(),
        valuesByIndex: [],
        error: message,
      };
      setRatioBasisStatus(message, "error");
      if (isResultsTabVisible()) renderResultsTable();
    }
  })();

  return ctx;
}

async function commitRatioBasisSelectionFromInput(options = {}) {
  const input = getRatioBasisInputEl();
  if (!input) return;
  const markDirty = options?.markDirty !== false;
  const shouldRender = options?.render !== false;
  const prevName = ratioBasisSelectedName;
  const prevFormat = ratioBasisSelectedFormat;
  const raw = toText(input.value);
  if (!raw || normalizeKey(raw) === "none") {
    input.value = "";
    syncRatioBasisInputWidth();
    ratioBasisSelectedName = "";
    ratioBasisSelectedFormat = "";
    clearRatioBasisColumnState();
    if (!ratioBasisProgrammaticUpdate && markDirty && (prevName || prevFormat)) {
      markDfmDirty();
    }
    if (shouldRender && isResultsTabVisible()) renderResultsTable();
    return { ok: true, value: "" };
  }

  let datasetOptions = [];
  try {
    datasetOptions = await ensureRatioBasisOptionsForCurrentProject();
  } catch (err) {
    console.error("Failed to load ratio-basis options:", err);
    setRatioBasisStatus(toText(err?.message) || "Failed to load dataset types.", "error");
    return { ok: false, error: toText(err?.message) || "Failed to load dataset types." };
  }

  const selected = (Array.isArray(datasetOptions) ? datasetOptions : [])
    .find((item) => normalizeKey(item.name) === normalizeKey(raw));
  if (!selected) {
    ratioBasisSelectedName = "";
    ratioBasisSelectedFormat = "";
    clearRatioBasisColumnState({ keepStatus: true });
    setRatioBasisStatus("Ratio Basis must match a dataset name from dataset_types.", "error");
    if (shouldRender && isResultsTabVisible()) renderResultsTable();
    return { ok: false, invalid: true };
  }

  input.value = selected.name;
  syncRatioBasisInputWidth();
  ratioBasisSelectedName = selected.name;
  ratioBasisSelectedFormat = selected.dataFormat;
  setRatioBasisStatus("", "");
  const changed = prevName !== ratioBasisSelectedName || prevFormat !== ratioBasisSelectedFormat;
  if (!ratioBasisProgrammaticUpdate && markDirty && changed) {
    markDfmDirty();
  }
  if (shouldRender && isResultsTabVisible()) renderResultsTable();
  return { ok: true, value: ratioBasisSelectedName };
}

export function wireResultsRatioBasisControls() {
  if (ratioBasisControlsWired) return;
  ratioBasisControlsWired = true;

  const input = getRatioBasisInputEl();
  const pickerBtn = getRatioBasisBtnEl();
  const ultRatioDecimalInput = getUltimateRatioDecimalInputEl();
  if (!input && !ultRatioDecimalInput) return;

  if (input) {
    syncRatioBasisInputWidth();
    input.addEventListener("focus", () => {
      ratioBasisEmbeddedSnapshot = false;
      void ensureRatioBasisOptionsForCurrentProject().catch((err) => {
        console.error("Failed to load ratio-basis options:", err);
        if (toText(input.value)) {
          setRatioBasisStatus(toText(err?.message) || "Failed to load dataset types.", "error");
        }
      });
    });

    input.addEventListener("input", () => {
      ratioBasisEmbeddedSnapshot = false;
      syncRatioBasisInputWidth();
      const raw = toText(input.value);
      if (!raw) {
        ratioBasisSelectedName = "";
        ratioBasisSelectedFormat = "";
        clearRatioBasisColumnState();
        if (isResultsTabVisible()) renderResultsTable();
        return;
      }
      void ensureRatioBasisOptionsForCurrentProject().catch(() => {});
    });

    input.addEventListener("change", () => {
      void commitRatioBasisSelectionFromInput();
    });

    input.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      void commitRatioBasisSelectionFromInput();
    });
  }

  if (input && pickerBtn && pickerBtn.dataset.wired !== "1") {
    pickerBtn.dataset.wired = "1";
    pickerBtn.addEventListener("click", async (e) => {
      e.preventDefault();
      e.stopPropagation();
      const projectName = getResolvedProjectName();
      const out = await openDatasetNamePicker({
        projectName,
        initialName: input.value,
        anchorElement: input,
        title: "Select Ratio Basis Dataset",
        // Ratio Basis picker intentionally shows all dataset types; unsupported
        // types are rejected later with a clear Results status message.
        includeCalculated: true,
        setStatus: (msg) => setRatioBasisStatus(String(msg || ""), "error"),
        onError: (err) => {
          console.error("Failed to open Ratio Basis picker:", err);
          setRatioBasisStatus(String(err?.message || err || "Failed to load dataset names."), "error");
        },
        onSelect: (name) => {
          const selected = toText(name);
          if (!selected) return;
          input.value = selected;
          syncRatioBasisInputWidth();
          input.dispatchEvent(new Event("change", { bubbles: true }));
        },
      });
      if (out?.ok) {
        try { input.focus({ preventScroll: true }); } catch { try { input.focus(); } catch {} }
      }
    });
  }

  if (ultRatioDecimalInput && ultRatioDecimalInput.dataset.wired !== "1") {
    ultRatioDecimalInput.dataset.wired = "1";
    let lastCommitted = String(getResultsUltimateRatioDecimalPlaces());
    const apply = () => {
      const normalized = String(getResultsUltimateRatioDecimalPlaces());
      if (ultRatioDecimalInput.value !== normalized) {
        ultRatioDecimalInput.value = normalized;
      }
      const changed = normalized !== lastCommitted;
      if (!changed) return;
      lastCommitted = normalized;
      if (!ultimateRatioDecimalProgrammaticUpdate) {
        markDfmDirty();
      }
      if (isResultsTabVisible()) renderResultsTable();
    };
    ultRatioDecimalInput.addEventListener("change", apply);
    ultRatioDecimalInput.addEventListener("blur", apply);
  }
}

export function getResultsRatioBasisSelection() {
  return toText(ratioBasisSelectedName);
}

export function getResultsUltimateRatioDecimalPlacesSelection() {
  return getResultsUltimateRatioDecimalPlaces();
}

export async function setResultsRatioBasisSelection(value, options = {}) {
  const input = getRatioBasisInputEl();
  if (!input) return { ok: false, error: "ratio basis input not found" };

  const next = toText(value);
  input.value = next;
  syncRatioBasisInputWidth();
  ratioBasisProgrammaticUpdate = true;
  try {
    if (!next) {
      const prevName = ratioBasisSelectedName;
      const prevFormat = ratioBasisSelectedFormat;
      ratioBasisSelectedName = "";
      ratioBasisSelectedFormat = "";
      clearRatioBasisColumnState();
      if (!options?.silent && (prevName || prevFormat)) markDfmDirty();
      if (options?.render !== false && isResultsTabVisible()) renderResultsTable();
      return { ok: true, value: "" };
    }
    return await commitRatioBasisSelectionFromInput({
      markDirty: !options?.silent,
      render: options?.render !== false,
    });
  } finally {
    ratioBasisProgrammaticUpdate = false;
  }
}

export function setResultsUltimateRatioDecimalPlacesSelection(value, options = {}) {
  const input = getUltimateRatioDecimalInputEl();
  if (!input) return { ok: false, error: "ultimate ratio decimal input not found" };
  const prev = getResultsUltimateRatioDecimalPlaces();
  input.value = String(value ?? "");
  ultimateRatioDecimalProgrammaticUpdate = true;
  try {
    const { value: normalized } = normalizeUltimateRatioDecimalInput();
    const changed = normalized !== prev;
    if (changed && !options?.silent) {
      markDfmDirty();
    }
    if ((changed || options?.forceRender) && options?.render !== false && isResultsTabVisible()) {
      renderResultsTable();
    }
    return { ok: true, value: normalized };
  } finally {
    ultimateRatioDecimalProgrammaticUpdate = false;
  }
}

export function buildResultsVector() {
  if (Array.isArray(persistedUltimateVector)) return persistedUltimateVector.slice();
  const model = state.model;
  if (!model || !Array.isArray(model.values) || !Array.isArray(model.mask)) return [];
  const origins = model.origin_labels || [];
  const devs = getEffectiveDevLabelsForModel(model);
  if (!devs.length) return [];
  const cumulative = getCumulativeFactors(model, devs);
  const vals = model.values;
  const mask = model.mask;
  const out = [];
  for (let r = 0; r < origins.length; r++) {
    const maxCol = Math.min(devs.length - 1, (vals?.[r] || []).length - 1);
    const latest = getLatestRowValue(vals, mask, r, maxCol);
    if (latest && Number.isFinite(cumulative[latest.col])) {
      out.push(latest.value * cumulative[latest.col]);
    } else {
      out.push(null);
    }
  }
  return out;
}

export function buildPercentDevelopedVector() {
  // The pattern a dependent method applies: one over the cumulative factor at
  // each origin's own development age. It comes from the selected factors
  // alone, so an origin whose latest observation is zero still develops.
  const model = state.model;
  if (!model || !Array.isArray(model.values) || !Array.isArray(model.mask)) return [];
  const origins = model.origin_labels || [];
  const devs = getEffectiveDevLabelsForModel(model);
  if (!devs.length) return [];
  const cumulative = getCumulativeFactors(model, devs);
  const vals = model.values;
  const mask = model.mask;
  const out = [];
  for (let r = 0; r < origins.length; r++) {
    const maxCol = Math.min(devs.length - 1, (vals?.[r] || []).length - 1);
    const latest = getLatestRowValue(vals, mask, r, maxCol);
    const factor = latest ? cumulative[latest.col] : null;
    out.push(Number.isFinite(factor) && factor !== 0 ? roundRatio(1 / factor, 6) : null);
  }
  return out;
}

export function buildResultsVectorCsv(vector) {
  if (!Array.isArray(vector) || !vector.length) return "";
  return `${vector.map((v) => escapeCsvCell(v == null ? "" : v)).join("\n")}\n`;
}

export function renderResultsTable() {
  const wrap = document.getElementById("resultsWrap");
  if (!wrap) return;
  wrap.innerHTML = "";

  dropEmbeddedRatioBasisSnapshotOnOriginChange();
  if (getRatioBasisInputEl() && !ratioBasisEmbeddedSnapshot) {
    void ensureRatioBasisOptionsForCurrentProject().catch((err) => {
      console.error("Failed to load ratio-basis options:", err);
      if (ratioBasisSelectedName) {
        setRatioBasisStatus(toText(err?.message) || "Failed to load dataset types.", "error");
      }
    });
  }
  const ratioBasisCtx = ratioBasisEmbeddedSnapshot ? null : queueRatioBasisColumnLoadIfNeeded();
  const ratioBasisActive = ratioBasisEmbeddedSnapshot
    ? Boolean(ratioBasisSelectedName && ratioBasisColumnState.status === "ready")
    : Boolean(ratioBasisCtx);
  const ratioBasisHeaderText = ratioBasisColumnState.headerText || ratioBasisCtx?.headerText || "Ratio Basis";
  const ratioBasisStateForRender = ratioBasisEmbeddedSnapshot
    ? ratioBasisColumnState
    : ratioBasisActive && ratioBasisCtx && ratioBasisColumnState.requestKey === ratioBasisCtx.requestKey
      ? ratioBasisColumnState
      : null;

  const model = state.model;
  if (!model || !Array.isArray(model.values) || !Array.isArray(model.mask)) {
    renderDatasetGridPlaceholder(wrap, {
      emptyHint: "Load a dataset in the Data tab to compute results.",
    });
    return;
  }

  const origins = model.origin_labels || [];
  const devs = getEffectiveDevLabelsForModel(model);
  if (!devs.length) {
    wrap.innerHTML = `<div class="small">Not enough data to compute results.</div>`;
    return;
  }

  const ratioLabels = getRatioHeaderLabels(devs);
  const colCount = ratioLabels.length || devs.length;
  ensureDefaultSummarySelectionForColumns(colCount);
  const cumulative = getCumulativeFactors(model, devs);
  const inputTriangleName = String(document.getElementById("triInput")?.value || "").trim();
  const latestHeaderText = inputTriangleName ? `Latest ${inputTriangleName}` : "Latest";

  const table = document.createElement("table");
  table.classList.add("arSpreadsheetTable", "dfmResultsTable");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  const corner = document.createElement("th");
  corner.textContent = getOriginLabelTextForRatio();
  headRow.appendChild(corner);

  const latestHead = document.createElement("th");
  latestHead.textContent = latestHeaderText;
  latestHead.dataset.c = "0";
  headRow.appendChild(latestHead);

  const reserveHead = document.createElement("th");
  reserveHead.textContent = "Reserve";
  reserveHead.dataset.c = "1";
  headRow.appendChild(reserveHead);

  const ultHead = document.createElement("th");
  ultHead.textContent = "Ultimate";
  ultHead.dataset.c = "2";
  headRow.appendChild(ultHead);

  if (ratioBasisActive) {
    const basisHead = document.createElement("th");
    basisHead.textContent = ratioBasisHeaderText;
    basisHead.title = ratioBasisHeaderText;
    basisHead.dataset.c = "3";
    headRow.appendChild(basisHead);

    const ultRatioHead = document.createElement("th");
    ultRatioHead.textContent = "Ultimate Ratio";
    ultRatioHead.dataset.c = "4";
    headRow.appendChild(ultRatioHead);
  }
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  const vals = model.values;
  const mask = model.mask;
  let latestTotal = 0;
  let latestTotalHasValue = false;
  let reserveTotal = 0;
  let reserveTotalHasValue = false;
  let ultimateTotal = 0;
  let ultimateTotalHasValue = false;
  let basisTotal = 0;
  let basisTotalHasValue = false;
  const tagResultCell = (cell, row, col) => {
    if (!cell) return cell;
    cell.dataset.r = String(row);
    cell.dataset.c = String(col);
    return cell;
  };
  for (let r = 0; r < origins.length; r++) {
    const tr = document.createElement("tr");
    const rowHead = document.createElement("th");
    rowHead.textContent = String(origins[r] ?? "");
    rowHead.dataset.r = String(r);
    tr.appendChild(rowHead);

    const latestTd = document.createElement("td");
    const reserveTd = document.createElement("td");
    const ultTd = document.createElement("td");
    const basisTd = ratioBasisActive ? document.createElement("td") : null;
    const ultRatioTd = ratioBasisActive ? document.createElement("td") : null;
    tagResultCell(latestTd, r, 0);
    tagResultCell(reserveTd, r, 1);
    tagResultCell(ultTd, r, 2);
    tagResultCell(basisTd, r, 3);
    tagResultCell(ultRatioTd, r, 4);
    const maxCol = Math.min(devs.length - 1, (vals?.[r] || []).length - 1);
    const latest = getLatestRowValue(vals, mask, r, maxCol);
    const latestValue = latest?.value;
    latestTd.textContent = Number.isFinite(latestValue) ? formatCellValue(latestValue) : "";
    if (Number.isFinite(latestValue)) {
      latestTotal += latestValue;
      latestTotalHasValue = true;
    }

    let ultimateValue = Array.isArray(persistedUltimateVector)
      ? clonePersistedNumber(persistedUltimateVector[r])
      : null;
    if (!Array.isArray(persistedUltimateVector) && latest && Number.isFinite(cumulative[latest.col])) {
      ultimateValue = latest.value * cumulative[latest.col];
    }
    if (Number.isFinite(ultimateValue) && Number.isFinite(latestValue)) {
      const reserveValue = ultimateValue - latestValue;
      reserveTd.textContent = formatCellValue(reserveValue);
      reserveTotal += reserveValue;
      reserveTotalHasValue = true;
    } else {
      reserveTd.textContent = "";
    }
    ultTd.textContent = Number.isFinite(ultimateValue) ? formatCellValue(ultimateValue) : "";
    if (Number.isFinite(ultimateValue)) {
      ultimateTotal += ultimateValue;
      ultimateTotalHasValue = true;
    }

    tr.appendChild(latestTd);
    tr.appendChild(reserveTd);
    tr.appendChild(ultTd);
    if (basisTd) {
      const basisValue = getRatioBasisRowValue(ratioBasisStateForRender, origins[r], r);
      basisTd.textContent = formatRatioBasisCellValue(basisValue);
      if (Number.isFinite(basisValue)) {
        basisTotal += basisValue;
        basisTotalHasValue = true;
      }
      tr.appendChild(basisTd);
      if (ultRatioTd) {
        const ultRatioValue =
          Number.isFinite(ultimateValue) &&
          Number.isFinite(basisValue) &&
          basisValue !== 0
            ? (ultimateValue / basisValue)
            : null;
        ultRatioTd.textContent = formatPercentCellValue(ultRatioValue);
        tr.appendChild(ultRatioTd);
      }
    }
    tbody.appendChild(tr);
  }

  const totalTr = document.createElement("tr");
  totalTr.className = "dfmResultsTotalRow";
  const totalHead = document.createElement("th");
  totalHead.textContent = "Total";
  totalHead.classList.add("rowhdr");
  totalHead.dataset.r = String(origins.length);
  totalTr.appendChild(totalHead);

  const latestTotalTd = document.createElement("td");
  tagResultCell(latestTotalTd, origins.length, 0);
  latestTotalTd.textContent = latestTotalHasValue ? formatCellValue(latestTotal) : "";
  totalTr.appendChild(latestTotalTd);

  const reserveTotalTd = document.createElement("td");
  tagResultCell(reserveTotalTd, origins.length, 1);
  reserveTotalTd.textContent = reserveTotalHasValue ? formatCellValue(reserveTotal) : "";
  totalTr.appendChild(reserveTotalTd);

  const ultimateTotalTd = document.createElement("td");
  tagResultCell(ultimateTotalTd, origins.length, 2);
  ultimateTotalTd.textContent = ultimateTotalHasValue ? formatCellValue(ultimateTotal) : "";
  totalTr.appendChild(ultimateTotalTd);

  if (ratioBasisActive) {
    const basisTotalTd = document.createElement("td");
    tagResultCell(basisTotalTd, origins.length, 3);
    basisTotalTd.textContent = basisTotalHasValue ? formatRatioBasisCellValue(basisTotal) : "";
    totalTr.appendChild(basisTotalTd);

    const totalUltRatioTd = document.createElement("td");
    tagResultCell(totalUltRatioTd, origins.length, 4);
    const totalUltRatioValue =
      ultimateTotalHasValue && basisTotalHasValue && basisTotal !== 0
        ? (ultimateTotal / basisTotal)
        : null;
    totalUltRatioTd.textContent = formatPercentCellValue(totalUltRatioValue);
    totalTr.appendChild(totalUltRatioTd);
  }
  tbody.appendChild(totalTr);

  table.appendChild(tbody);
  wrap.appendChild(table);
  wireResultsCopyMenu();
  const tableHighlight = wireSelectableTable({
    container: wrap,
    selectedClass: "dfmTableHighlight",
    activeClass: "dfmTableActive",
    rowHeaderSelector: "tbody th[data-r]",
    columnHeaderSelector: "thead th[data-c]",
    canStartLabelSelection: true,
    canHandleKeyboardNavigation: true,
    scrollHost: wrap,
    onSelectionChange: ({ ranges }) => {
      syncResultsSelectionState(wrap.querySelector("table.dfmResultsTable"), ranges);
      // Match the shared Data grid: once a table selection starts, keyboard
      // commands belong to the table rather than the control last focused.
      // This lets Escape clear a selected cell or range immediately.
      wrap.focus({ preventScroll: true });
    },
    onContextMenu: (event, cell, api) => {
      event.preventDefault();
      window.__arcRhoCopyActiveGridSelection = api.copySelection;
      const menu = document.getElementById("ctxMenu");
      if (!menu) return;
      openContextMenu(menu, {
        anchorEl: cell,
        clientX: event.clientX,
        clientY: event.clientY,
        offset: 8,
        align: "top-left",
      });
    },
  });
  if (tableHighlight) window.__arcRhoCopyActiveGridSelection = tableHighlight.copySelection;
}
