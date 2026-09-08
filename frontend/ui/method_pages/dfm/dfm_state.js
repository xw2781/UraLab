/*
===============================================================================
DFM State - Shared state variables, constants, and utility functions
used across all DFM tab modules.
===============================================================================
*/
import { state } from "/ui/shared/dataset/dataset_state.js";
import {
  sanitizeDataFolderPart,
  sanitizeFileNamePart,
} from "/ui/shared/utils/filename.js";
import {
  getSummaryConfigKey,
  loadCustomSummaryRows,
} from "/ui/method_pages/dfm/dfm_storage.js";
import {
  curvesTable,
  normalizeCurvesTab,
} from "/ui/method_pages/dfm/dfm_curve_fit.js?v=20260903a";

// =============================================================================
// Dynamic Calc Import
// =============================================================================
const __ratioParams = new URL(import.meta.url).search;
const __ratioCalcUrl = new URL("/ui/method_pages/dfm/dfm_ratio_calc.js", import.meta.url);
__ratioCalcUrl.search = __ratioParams;
const {
  calcRatio,
  ratioNumberOrNull,
  persistedRatioOrNull,
  roundRatio,
  roundHalfUp,
  averageRowReferenceValue,
  formatRatio,
  computeAverageForColumn,
} = await import(__ratioCalcUrl.toString());

export {
  calcRatio,
  ratioNumberOrNull,
  persistedRatioOrNull,
  roundRatio,
  roundHalfUp,
  averageRowReferenceValue,
  formatRatio,
  computeAverageForColumn,
};
export { state };

// =============================================================================
// Runtime Params + Constants
// =============================================================================
const pageParams = new URLSearchParams(window.location.search);
export const ratioSyncParams = pageParams;
export const ratioSyncInst = ratioSyncParams.get("inst") || "default";
export const ratioSyncChannelName = `arcrho-dfm-ratio-sync::${ratioSyncInst}`;
export const ratioSyncSourceId = `dfm_${Math.random().toString(36).slice(2)}_${Date.now()}`;
export const RATIO_SAVE_PATH_KEY = `arcrho_dfm_ratio_save_path_v1::${ratioSyncInst}`;
export const BASE_SUMMARY_ROWS = [
  { id: "volume_all", label: "Volume - all", base: "volume", periods: "all" },
];

// =============================================================================
// Mutable State (exported directly for Set/Map, via getter/setter for primitives)
// =============================================================================
export const ratioStrikeSet = new Set();
export const activeRatioCols = new Set();
export const selectedSummaryByCol = new Map();
export const ratioChartThresholdByCol = new Map();
export const ratioChartLowerThresholdByCol = new Map();
export const ratioChartLeftThresholdByCol = new Map();

let ratioColAllActive = false;
let cachedRootPath = null;
let cachedWorkspacePaths = null;
let dfmIsDirty = false;
let showNaBorders = false;
let currentDfmTab = "details";
let ratioSummaryRaf = null;
let lastSummaryCtxRowId = null;
let ratioChartCol = null;
let ratioChartRaf = null;
let ratioChartWired = false;
let ratioChartPoints = [];
let ratioChartScale = null;
let ratioChartDragActive = false;
let ratioChartDragMoved = false;
let ratioChartHoverLine = null;
let ratioChartDragTarget = null;
let ratioChartHoverTimer = null;
let ratioChartHoverKey = null;
let ratioChartTooltipVisible = false;
let ratioSyncChannel = null;
let ratioSyncMuted = false;
let dfmProgrammaticDepth = 0;

export const summaryRowConfigs = [];
export const summaryRowMap = new Map();

// Getter/setter pairs for primitives
export function getRatioColAllActive() { return ratioColAllActive; }
export function setRatioColAllActive(v) { ratioColAllActive = v; }

export function getDfmIsDirty() { return dfmIsDirty; }

// A save that enqueued an Engine dependent-propagation job records the job id
// here so the dependency-source "cleared" message carries it; Project Instance
// then keeps downstream live previews until the job's terminal status.
let pendingDfmPropagationJobId = "";
export function setPendingDfmPropagationJobId(jobId) {
  pendingDfmPropagationJobId = String(jobId || "").trim();
}
export function consumePendingDfmPropagationJobId() {
  const jobId = pendingDfmPropagationJobId;
  pendingDfmPropagationJobId = "";
  return jobId;
}

export function getShowNaBorders() { return showNaBorders; }
export function setShowNaBorders(v) { showNaBorders = v; }

export function getCurrentDfmTab() { return currentDfmTab; }
export function setCurrentDfmTab(v) { currentDfmTab = v; }

export function getRatioSummaryRaf() { return ratioSummaryRaf; }
export function setRatioSummaryRaf(v) { ratioSummaryRaf = v; }

export function getLastSummaryCtxRowId() { return lastSummaryCtxRowId; }
export function setLastSummaryCtxRowId(v) { lastSummaryCtxRowId = v; }

export function getRatioChartCol() { return ratioChartCol; }
export function setRatioChartCol(v) { ratioChartCol = v; }

export function getRatioChartRaf() { return ratioChartRaf; }
export function setRatioChartRaf(v) { ratioChartRaf = v; }

export function getRatioChartWired() { return ratioChartWired; }
export function setRatioChartWired(v) { ratioChartWired = v; }

export function getRatioChartPoints() { return ratioChartPoints; }
export function setRatioChartPoints(v) { ratioChartPoints = v; }

export function getRatioChartScale() { return ratioChartScale; }
export function setRatioChartScale(v) { ratioChartScale = v; }

export function getRatioChartDragActive() { return ratioChartDragActive; }
export function setRatioChartDragActive(v) { ratioChartDragActive = v; }

export function getRatioChartDragMoved() { return ratioChartDragMoved; }
export function setRatioChartDragMoved(v) { ratioChartDragMoved = v; }

export function getRatioChartHoverLine() { return ratioChartHoverLine; }
export function setRatioChartHoverLine(v) { ratioChartHoverLine = v; }

export function getRatioChartDragTarget() { return ratioChartDragTarget; }
export function setRatioChartDragTarget(v) { ratioChartDragTarget = v; }

export function getRatioChartHoverTimer() { return ratioChartHoverTimer; }
export function setRatioChartHoverTimer(v) { ratioChartHoverTimer = v; }

export function getRatioChartHoverKey() { return ratioChartHoverKey; }
export function setRatioChartHoverKey(v) { ratioChartHoverKey = v; }

export function getRatioChartTooltipVisible() { return ratioChartTooltipVisible; }
export function setRatioChartTooltipVisible(v) { ratioChartTooltipVisible = v; }

export function getRatioSyncChannel() { return ratioSyncChannel; }
export function setRatioSyncChannel(v) { ratioSyncChannel = v; }

export function getRatioSyncMuted() { return ratioSyncMuted; }
export function setRatioSyncMuted(v) { ratioSyncMuted = v; }

// =============================================================================
// Utility Functions
// =============================================================================
export function getDfmInst() {
  const params = new URLSearchParams(window.location.search);
  return params.get("inst") || "";
}

export function isDfmApplyingProgrammatically() {
  return dfmProgrammaticDepth > 0;
}

export async function runDfmProgrammatic(fn) {
  dfmProgrammaticDepth++;
  try {
    return await fn();
  } finally {
    dfmProgrammaticDepth--;
  }
}

export function runDfmProgrammaticSync(fn) {
  dfmProgrammaticDepth++;
  try {
    return fn();
  } finally {
    dfmProgrammaticDepth--;
  }
}

function notifyDfmDirtyState(dirty, options = {}) {
  const nextDirty = !!dirty;
  if (dfmIsDirty === nextDirty && !options?.force) return;
  dfmIsDirty = nextDirty;
  try {
    window.dispatchEvent(new CustomEvent("arcrho:dfm-dirty-state", { detail: { dirty: nextDirty } }));
  } catch {
    // ignore
  }
  const inst = getDfmInst();
  window.parent.postMessage({ type: "arcrho:dfm-dirty", inst, dirty: nextDirty }, "*");
}

export function markDfmDirty() {
  if (dfmProgrammaticDepth > 0) return;
  if (dfmIsDirty) {
    try {
      window.dispatchEvent(new CustomEvent("arcrho:dfm-dirty-state", { detail: { dirty: true } }));
    } catch {
      // ignore
    }
    return;
  }
  notifyDfmDirtyState(true);
}

export function markDfmClean(options = {}) {
  notifyDfmDirtyState(false, options);
}

export function getDfmInputSnapshot() {
  try {
    if (typeof window.ADA_GET_DFM_INPUTS === "function") {
      return window.ADA_GET_DFM_INPUTS();
    }
  } catch {
    // ignore
  }
  const tri = document.getElementById("triInput")?.value?.trim() || "";
  const project = document.getElementById("projectSelect")?.value?.trim() || "";
  const reservingClass = document.getElementById("pathInput")?.value?.trim() || "";
  return {
    resolved: { project, reservingClass, tri },
    display: { project, reservingClass, tri },
    defaults: { projectDefault: false, reservingClassDefault: false },
  };
}

export function getResolvedProjectName() {
  const snap = getDfmInputSnapshot();
  return (snap.resolved?.project || "").trim();
}

export function getResolvedReservingClass() {
  const snap = getDfmInputSnapshot();
  return (snap.resolved?.reservingClass || "").trim();
}

function normalizeWorkspacePathConfig(data) {
  const config = data?.config && typeof data.config === "object" ? data.config : {};
  const paths = config.paths && typeof config.paths === "object" ? config.paths : {};
  const root = String(config.workspace_root || "E:\\ArcRho").trim() || "E:\\ArcRho";
  return {
    workspace_root: root,
    paths: {
      projects_dir: String(paths.projects_dir || "projects").trim() || "projects",
      requests_dir: String(paths.requests_dir || "requests").trim() || "requests",
    },
  };
}

function trimPathSeparators(value) {
  return String(value || "").replace(/^[\\/]+|[\\/]+$/g, "");
}

function trimTrailingPathSeparators(value) {
  return String(value || "").replace(/[\\/]+$/g, "");
}

function isAbsolutePath(value) {
  const text = String(value || "").trim();
  return /^[A-Za-z]:[\\/]/.test(text) || /^\\\\/.test(text);
}

function joinWorkspacePath(...parts) {
  const cleaned = [];
  parts.forEach((part, index) => {
    const text = String(part || "").trim();
    if (!text) return;
    cleaned.push(index === 0 ? trimTrailingPathSeparators(text) : trimPathSeparators(text));
  });
  return cleaned.join("\\");
}

export async function getWorkspacePathsConfig() {
  if (cachedWorkspacePaths) return cachedWorkspacePaths;
  try {
    const res = await fetch("/workspace_paths");
    if (res.ok) {
      const data = await res.json();
      cachedWorkspacePaths = normalizeWorkspacePathConfig(data);
    } else {
      cachedWorkspacePaths = normalizeWorkspacePathConfig(null);
    }
  } catch {
    cachedWorkspacePaths = normalizeWorkspacePathConfig(null);
  }
  cachedRootPath = cachedWorkspacePaths.workspace_root;
  return cachedWorkspacePaths;
}

export async function getRootPath() {
  if (cachedRootPath) return cachedRootPath;
  const config = await getWorkspacePathsConfig();
  cachedRootPath = config.workspace_root;
  return cachedRootPath;
}

export function setCachedRootPath(value) {
  const next = String(value || "").trim();
  cachedRootPath = next || null;
  if (cachedWorkspacePaths) {
    cachedWorkspacePaths = next
      ? { ...cachedWorkspacePaths, workspace_root: next }
      : null;
  }
}

export function getDefaultMethodName() {
  const tri = document.getElementById("triInput")?.value?.trim();
  return tri ? `DFM ${tri}` : "DFM";
}

export function getDfmDecimalPlaces() {
  const el = document.getElementById("decimalPlaces");
  const raw = Number.parseInt(String(el?.value ?? "").trim(), 10);
  if (!Number.isFinite(raw)) return 4;
  return Math.max(0, Math.min(6, raw));
}

export function getHostApi() {
  if (window.ADAHost) return window.ADAHost;
  try {
    let w = window.parent;
    while (w && w !== window) {
      if (w.ADAHost) return w.ADAHost;
      if (w === w.parent) break;
      w = w.parent;
    }
  } catch {}
  return null;
}

export { sanitizeFileNamePart };

export function sanitizeDfmMethodFilePart(value, fallback) {
  return sanitizeDataFolderPart(value, fallback);
}

export function getRatioSaveProjectName() {
  const project = getResolvedProjectName();
  return project ? project : "UnknownProject";
}

export function getRatioSaveSuggestedName(options = {}) {
  const rawMethodName = typeof options.methodName === "string"
    ? options.methodName
    : document.getElementById("dfmMethodName")?.value?.trim();
  const methodName = sanitizeFileNamePart(
    rawMethodName,
    "Name",
  );
  return `DFM@${methodName}.json`;
}

export async function getRatioSaveBaseDir() {
  const workspaceConfig = await getWorkspacePathsConfig();
  const projectsDir = workspaceConfig.paths.projects_dir || "projects";
  const projectsPath = isAbsolutePath(projectsDir)
    ? trimTrailingPathSeparators(projectsDir)
    : joinWorkspacePath(workspaceConfig.workspace_root, projectsDir);
  const project = sanitizeFileNamePart(getRatioSaveProjectName(), "UnknownProject");
  const reservingClass = sanitizeDfmMethodFilePart(
    getResolvedReservingClass() || String(document.getElementById("pathInput")?.value || "").trim(),
    "ReservingClass",
  );
  return joinWorkspacePath(projectsPath, project || "UnknownProject", "data", reservingClass, "methods");
}

export async function buildRatioSavePath(options = {}) {
  const baseDir = await getRatioSaveBaseDir();
  const filename = getRatioSaveSuggestedName(options);
  return `${baseDir}\\${filename}`;
}

export async function getRatioDataDir() {
  const workspaceConfig = await getWorkspacePathsConfig();
  const projectsDir = workspaceConfig.paths.projects_dir || "projects";
  const projectsPath = isAbsolutePath(projectsDir)
    ? trimTrailingPathSeparators(projectsDir)
    : joinWorkspacePath(workspaceConfig.workspace_root, projectsDir);
  const project = sanitizeFileNamePart(getRatioSaveProjectName(), "UnknownProject");
  const reservingClass = sanitizeDfmMethodFilePart(
    getResolvedReservingClass() || String(document.getElementById("pathInput")?.value || "").trim(),
    "ReservingClass",
  );
  return joinWorkspacePath(projectsPath, project || "UnknownProject", "data", reservingClass, "datasets");
}

export function getResultsCsvSuggestedName(options = {}) {
  const datasetNameRaw = typeof options.datasetName === "string"
    ? options.datasetName
    : (String(document.getElementById("dfmMethodName")?.value || "").trim() || getDefaultMethodName());
  const originLen = normalizePositiveIntegerText(options.originLen ?? options.originLength ?? document.getElementById("originLenSelect")?.value);
  const cacheVariantSuffix = originLen ? `@${originLen}` : "";
  return `${sanitizeFileNamePart(datasetNameRaw, "Dataset")}${cacheVariantSuffix}.csv`;
}

export function getInputTriangleCsvSuggestedName(options = {}) {
  const triangleName = typeof options.datasetName === "string"
    ? options.datasetName
    : String(document.getElementById("triInput")?.value || "").trim();
  const originLen = normalizePositiveIntegerText(options.originLen ?? options.originLength ?? document.getElementById("originLenSelect")?.value);
  const devLen = normalizePositiveIntegerText(options.devLen ?? options.developmentLength ?? document.getElementById("devLenSelect")?.value);
  const cacheVariantSuffix = originLen && devLen ? `@${originLen}@${devLen}@cum@dev` : "";
  return `${sanitizeFileNamePart(triangleName, "Dataset")}${cacheVariantSuffix}.csv`;
}

export async function buildInputTriangleCsvPath(options = {}) {
  const dataDir = await getRatioDataDir();
  return `${dataDir}\\${getInputTriangleCsvSuggestedName(options)}`;
}

function normalizePositiveIntegerText(value) {
  const parsed = Number.parseInt(String(value ?? "").trim(), 10);
  return Number.isFinite(parsed) && parsed > 0 ? String(parsed) : "";
}

export function escapeCsvCell(value) {
  const text = value == null ? "" : String(value);
  if (/[",\r\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

export function getEffectiveDevLabelsForModel(model) {
  const devs = Array.isArray(model?.dev_labels) ? model.dev_labels : [];
  const vals = Array.isArray(model?.values) ? model.values : [];
  let maxCols = 0;
  for (const row of vals) {
    if (Array.isArray(row)) maxCols = Math.max(maxCols, row.length);
  }
  if (!maxCols || maxCols >= devs.length) return devs;
  return devs.slice(0, maxCols);
}

export function toLabelNum(value) {
  const s = String(value ?? "").trim();
  const m = s.match(/[-+]?\d*\.?\d+/);
  return m ? m[0] : "";
}

export function getRatioHeaderLabels(devs) {
  const labels = [];
  for (let c = 0; c < devs.length - 1; c++) {
    const left = toLabelNum(devs[c]);
    const right = toLabelNum(devs[c + 1]);
    if (left && right) {
      labels.push(`${left}-${right}`);
    } else {
      labels.push(`${String(devs[c] ?? "")}-${String(devs[c + 1] ?? "")}`);
    }
  }

  if (devs.length) {
    const lastRaw = devs[devs.length - 1];
    const lastNum = toLabelNum(lastRaw);
    const left = (lastNum || String(lastRaw ?? "").trim() || "Ult");
    if (String(left).trim().toLowerCase() === "ult") {
      labels.push("Ult");
    } else {
      labels.push(`${left} - Ult`);
    }
  }

  return labels;
}

export function getOriginLabelTextForRatio() {
  const originLen = Number(document.getElementById("originLenSelect")?.value || 12);
  switch (originLen) {
    case 12: return "Accident Year";
    case 6: return "Accident Half-Year";
    case 3: return "Accident Quarter";
    case 1: return "Accident Month";
    default: return "Accident Period";
  }
}

export function buildSummaryRows() {
  const key = getSummaryConfigKey();
  const savedRows = loadCustomSummaryRows(key);
  const merged = Array.isArray(savedRows) && savedRows.length
    ? savedRows
    : BASE_SUMMARY_ROWS.map((row) => ({ ...row }));
  summaryRowConfigs.splice(0, summaryRowConfigs.length, ...merged);
  summaryRowMap.clear();
  summaryRowConfigs.forEach((row) => summaryRowMap.set(row.id, row));
  return summaryRowConfigs;
}

export function parsePeriodsValue(raw) {
  if (!raw) return "all";
  const txt = String(raw).trim();
  if (!txt || txt.toLowerCase() === "all") return "all";
  const n = Number(txt);
  if (!Number.isFinite(n) || n <= 0) return "all";
  return Math.floor(n);
}

export function parseExcludeValue(raw) {
  if (!raw) return 0;
  const txt = String(raw).trim();
  if (!txt) return 0;
  if (txt.toLowerCase() === "none") return 0;
  const n = Number(txt);
  if (!Number.isFinite(n) || n <= 0) return 0;
  return Math.floor(n);
}

export function buildExcludedSetForColumn(model, col, cfg, baseExcludedSet) {
  const baseSet = baseExcludedSet || new Set();
  const excludeCount = parseExcludeValue(cfg?.exclude);
  if (!excludeCount) return baseSet;
  if (!model || !Array.isArray(model.values) || !Array.isArray(model.mask)) return baseSet;

  const vals = model.values;
  const mask = model.mask;
  const rowCount = Array.isArray(model.origin_labels) ? model.origin_labels.length : vals.length;
  const periodsRaw = cfg?.periods ?? "all";
  const periods = typeof periodsRaw === "string" && periodsRaw.toLowerCase() === "all"
    ? "all"
    : Number(periodsRaw);
  const lookback = Number.isFinite(periods) && periods > 0 ? Math.floor(periods) : null;

  const includeRow = (r) => {
    const hasA = !!(mask[r] && mask[r][col]);
    const hasB = !!(mask[r] && mask[r][col + 1]);
    if (!hasA || !hasB) return null;
    return calcRatio(vals?.[r]?.[col], vals?.[r]?.[col + 1]);
  };

  const candidates = [];
  if (lookback) {
    let picked = 0;
    for (let r = rowCount - 1; r >= 0; r--) {
      if (picked >= lookback) break;
      const ratio = includeRow(r);
      if (!Number.isFinite(ratio)) continue;
      if (baseSet && baseSet.has(`${r},${col}`)) continue;
      picked += 1;
      candidates.push({ r, ratio });
    }
  } else {
    for (let r = 0; r < rowCount; r++) {
      const ratio = includeRow(r);
      if (!Number.isFinite(ratio)) continue;
      if (baseSet && baseSet.has(`${r},${col}`)) continue;
      candidates.push({ r, ratio });
    }
  }

  // ResQ drops pairs of highest and lowest ratios "for as long as the remaining
  // number of ratios is greater than two", so a column down to two ratios keeps
  // both and averages them rather than excluding itself empty. Allowing
  // (n - 1) / 2 pairs is that rule written as a count.
  const n = Math.min(excludeCount, Math.floor((candidates.length - 1) / 2));
  if (n <= 0) return baseSet;

  const sorted = [...candidates].sort((a, b) => a.ratio - b.ratio);
  const merged = new Set(baseSet);
  for (let i = 0; i < n; i++) {
    merged.add(`${sorted[i].r},${col}`);
    merged.add(`${sorted[sorted.length - 1 - i].r},${col}`);
  }
  return merged;
}

export function ensureDefaultSummarySelectionForColumns(colCount) {
  if (!colCount) return;
  const rows = buildSummaryRows();
  const defaultRowId = rows[0]?.id || "";
  if (!defaultRowId) return;
  for (let c = 0; c < colCount; c++) {
    if (!selectedSummaryByCol.has(c)) selectedSummaryByCol.set(c, defaultRowId);
  }
}

export function getSelectedRatioValues(model, devs) {
  const ratioLabels = getRatioHeaderLabels(devs);
  const values = new Array(ratioLabels.length).fill(1);
  if (!ratioLabels.length) return values;

  const rows = buildSummaryRows();
  const defaultRowId = rows[0]?.id || "";

  for (let c = 0; c < ratioLabels.length; c++) {
    const rowId = selectedSummaryByCol.get(c) || defaultRowId;
    const cfg = rowId ? summaryRowMap.get(rowId) : null;
    if (!cfg) {
      values[c] = 1;
      continue;
    }
    if (c >= devs.length - 1) {
      values[c] = getSummaryRowTailFactor(cfg, c);
      continue;
    }
    const averageType = String(cfg.averageType || "").trim().toLowerCase();
    if (averageType === "user_entry") {
      const raw = Array.isArray(cfg.values) ? cfg.values[c] : 1;
      const manual = Number(raw);
      values[c] = Number.isFinite(manual) && manual > 0 ? manual : 1;
      continue;
    }
    const excluded = buildExcludedSetForColumn(model, c, cfg, ratioStrikeSet);
    const summary = computeAverageForColumn(model, c, excluded, cfg, ratioStrikeSet);
    if (summary.totalValid > 0 && summary.totalIncluded === 0) {
      values[c] = 1;
      continue;
    }
    const isVolume = String(cfg.base || "volume").toLowerCase() === "volume";
    const hasValue =
      summary.value !== null &&
      (isVolume ? summary.sumA : summary.totalIncluded > 0);
    values[c] = hasValue ? summary.value : 1;
  }

  return values;
}

// The "- Ult" column is a row's own tail factor, entered rather than averaged
// (ResQ keeps it as the average row's TailFactor). A User Entry or frozen
// benchmark row carries it in its stored values; a computed average row has
// none and stays at 1.
export function summaryRowOwnsTail(cfg) {
  const averageType = String(cfg?.averageType || "").trim().toLowerCase();
  const base = String(cfg?.base || "").trim().toLowerCase();
  return averageType === "user_entry" || base === "benchmark";
}

export function getSummaryRowTailFactor(cfg, col) {
  if (!summaryRowOwnsTail(cfg)) return 1;
  const raw = Array.isArray(cfg.values) ? cfg.values[col] : 1;
  const manual = Number(raw);
  return Number.isFinite(manual) && manual > 0 ? manual : 1;
}

// =============================================================================
// Curves tab state
// =============================================================================
// The person's Curves-tab choices, in the persisted `curves_tab` shape. Null
// means "never touched", which normalizes to the default tab: the Initial
// Selection everywhere, so the factors are exactly the Ratios tab's.
let curvesTabState = null;
let curvesTableCache = null;

export function getCurvesTab() {
  return curvesTabState;
}

export function setCurvesTab(tab) {
  curvesTabState = tab && typeof tab === "object" ? JSON.parse(JSON.stringify(tab)) : null;
  curvesTableCache = null;
}

export function invalidateCurvesTable() {
  curvesTableCache = null;
}

export function getNormalizedCurvesTab(model, devs) {
  const ratioValues = getSelectedRatioValues(model, devs);
  const initial = ratioValues.slice(0, Math.max(0, ratioValues.length - 1));
  return normalizeCurvesTab(curvesTabState, initial.length, initial);
}

// The whole Curves | Data table for the current Ratios selection, cached until
// the ratios or the tab change.
export function getCurvesTable(model, devs) {
  const ratioValues = getSelectedRatioValues(model, devs);
  const initial = ratioValues.slice(0, Math.max(0, ratioValues.length - 1));
  const initialTail = ratioValues.length ? ratioValues[ratioValues.length - 1] : 1;
  const key = JSON.stringify([ratioValues, curvesTabState]);
  if (curvesTableCache && curvesTableCache.key === key) return curvesTableCache.table;
  const table = curvesTable(initial, initialTail, curvesTabState);
  curvesTableCache = { key, table };
  return table;
}

// The factors the ultimates chain: the Curves tab's selected value per period
// and its selected tail, mirroring dfm_contract.selected_development_factors.
export function getSelectedDevelopmentFactors(model, devs) {
  const ratioValues = getSelectedRatioValues(model, devs);
  if (!ratioValues.length) return ratioValues;
  const table = getCurvesTable(model, devs);
  return [...table.selected_values, table.selected_tail];
}

export function getCumulativeFactors(model, devs) {
  const ratioValues = getSelectedDevelopmentFactors(model, devs);
  const cumulative = new Array(ratioValues.length).fill(null);
  let running = null;
  for (let i = ratioValues.length - 1; i >= 0; i--) {
    const v = ratioValues[i];
    if (!Number.isFinite(v)) {
      cumulative[i] = null;
      running = null;
      continue;
    }
    if (i === ratioValues.length - 1) {
      running = v;
    } else if (Number.isFinite(running)) {
      running = v * running;
    } else {
      cumulative[i] = null;
      running = null;
      continue;
    }
    cumulative[i] = running;
  }
  return cumulative;
}

export function getLatestRowValue(vals, mask, rowIndex, maxCol) {
  if (!Array.isArray(vals) || !Array.isArray(mask) || maxCol < 0) return null;
  const rowVals = vals[rowIndex] || [];
  for (let c = maxCol; c >= 0; c--) {
    if (!(mask[rowIndex] && mask[rowIndex][c])) continue;
    const raw = rowVals[c];
    const n = (typeof raw === "number") ? raw : Number(raw);
    if (!Number.isFinite(n)) continue;
    return { value: n, col: c };
  }
  return null;
}

export function isRatiosTabVisible() {
  const ratiosPage = document.getElementById("dfmRatiosPage");
  return !!ratiosPage && ratiosPage.style.display !== "none";
}

export function isResultsTabVisible() {
  const resultsPage = document.getElementById("dfmResultsPage");
  return !!resultsPage && resultsPage.style.display !== "none";
}

export function isCurvesTabVisible() {
  const curvesPage = document.getElementById("dfmCurvesPage");
  return !!curvesPage && curvesPage.style.display !== "none";
}

export function notifyDfmEditState() {
  const enabled = isRatiosTabVisible() && (ratioColAllActive || activeRatioCols.size > 0);
  window.parent.postMessage({ type: "arcrho:dfm-edit-state", enabled }, "*");
}
