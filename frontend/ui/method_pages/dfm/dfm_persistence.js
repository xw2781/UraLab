/*
===============================================================================
DFM Persistence - load/save ratio selections to disk via host API
===============================================================================
*/
import {
  beginDatasetGridLoading,
  endDatasetGridLoading,
} from "/ui/shared/tabs/data/dataset_grid_placeholder.js?v=20260809a";
import {
  state,
  ratioStrikeSet,
  selectedSummaryByCol,
  summaryRowConfigs,
  BASE_SUMMARY_ROWS,
  RATIO_SAVE_PATH_KEY,
  getHostApi,
  buildRatioSavePath,
  buildInputTriangleCsvPath,
  getRatioSaveBaseDir,
  getRatioDataDir,
  getRatioSaveSuggestedName,
  getResultsCsvSuggestedName,
  getCurrentDfmTab,
  buildSummaryRows,
  markDfmClean,
  runDfmProgrammatic,
  isRatiosTabVisible,
  isResultsTabVisible,
  getDfmIsDirty,
  sanitizeFileNamePart,
  getRatioSaveProjectName,
  getResolvedProjectName,
  getResolvedReservingClass,
  setPendingDfmPropagationJobId,
  getDfmInst,
  getDfmDecimalPlaces,
  getEffectiveDevLabelsForModel,
  getRatioHeaderLabels,
  calcRatio,
  ratioNumberOrNull,
  roundRatio,
  computeAverageForColumn,
  buildExcludedSetForColumn,
} from "/ui/method_pages/dfm/dfm_state.js";
import { showMethodSaveReviewWarning } from "/ui/shared/components/message_box/method_save_review_warning.js?v=20260827a";
import { showPageMessageBox } from "/ui/shared/components/message_box/message_box.js?v=20260831a";
import { showExcelLinkFailureAlert } from "/ui/shared/integrations/excel_link_alert.js?v=20260819a";
import { createArcRhoSaveProgress } from "/ui/shared/components/progress_popup/save_progress.js?v=20260831a";
import {
  isEngineUnavailableSaveError,
  trackSavePropagation,
} from "/ui/shared/services/dependent_propagation_job.js?v=20260813e";
import {
  createMethodObjectChangeWatchController,
  showObjectUpdatedAlert,
  wireSamePropagationScopePause,
} from "/ui/shared/services/object_change_watch.js?v=20260820a";
import {
  getSummaryConfigKey,
  saveCustomSummaryRows,
  loadCustomSummaryRows,
  markMethodSaved,
  clearMethodSavedFlag,
} from "/ui/method_pages/dfm/dfm_storage.js";
import {
  buildRatioSelectionPattern,
  applyRatioSelectionPattern,
  buildAverageSelectionPayload,
  applyAverageSelectionFromSaved,
  applyPersistedRatioDerivedSnapshot,
  renderRatioTable,
  queueDfmExternalChangeHighlights,
} from "/ui/method_pages/dfm/dfm_ratios_tab.js?v=20260907b";
import {
  applyPersistedResultsSnapshot,
  ensureResultsRatioBasisAligned,
  renderResultsTable,
  buildResultsVector,
  buildResultsVectorCsv,
  getResultsRatioBasisSelection,
  getResultsRatioBasisSnapshot,
  getResultsUltimateRatioDecimalPlacesSelection,
  setResultsRatioBasisSelection,
  setResultsUltimateRatioDecimalPlacesSelection,
} from "/ui/method_pages/dfm/dfm_results_tab.js?v=20260907a";
import { getDfmNotesText, setDfmNotesText } from "/ui/method_pages/dfm/dfm_notes_tab.js?v=20260714a";
import {
  applyDfmCurvesTabPayload,
  buildDfmCurvesTabPayload,
  renderDfmCurvesTab,
} from "/ui/method_pages/dfm/dfm_curves_tab.js?v=20260907a";
import { getSummaryRowTailFactor } from "/ui/method_pages/dfm/dfm_state.js";
import {
  buildDfmAverageFormulaObject,
  buildDfmSummaryRowsFromAverageFormulaObject,
  buildDfmSummaryRowsFromAverageFormulas,
  getDfmAverageFormulaLabels,
  getDfmAverageFormulaSelectedIndex,
  getDfmAverageFormulaValues,
} from "/ui/method_pages/dfm/dfm_average_formula_rows.js?v=20260513b";
import {
  applyDfmCellNotesPayload,
  buildDfmCellNotesPayload,
} from "/ui/method_pages/dfm/dfm_cell_notes.js";
import {
  recordCurrentDfmObjectSnapshot,
  refreshDfmMethodIndex,
} from "/ui/method_pages/dfm/dfm_startup_state.js";
import {
  hydrateDfmOutputSidecar,
  refreshDfmAuditLog,
  renderDfmAuditLog,
} from "/ui/method_pages/dfm/dfm_audit_log.js?v=20260726a";
import {
  DFM_METHOD_JSON_FORMAT,
  isDfmV2Method,
  loadDfmMethod,
  previewDfmMethod,
  readDfmMethodIdentityFromPage,
  saveDfmMethod,
} from "/ui/method_pages/dfm/dfm_method_api.js?v=20260814b";
import {
  cancelDfmExcelFreshnessCheck,
  checkDfmExcelLinkFreshness,
} from "/ui/method_pages/dfm/dfm_ratios_summary_table.js?v=20260903a";
import { containsDfmDatasetReference } from "/ui/method_pages/dfm/dfm_dataset_reference.js?v=20260811b";
import { resolveDfmDatasetReferencesInFormulas } from "/ui/method_pages/dfm/dfm_dataset_formula.js?v=20260820a";
import { setDfmExcelFreshnessState } from "/ui/method_pages/dfm/dfm_links_tab.js?v=20260901a";
import { refreshDfmDetailsDependencies } from "/ui/method_pages/dfm/dfm_details_dependencies.js?v=20260820b";

let ratioLoadTimer = null;
let ratioLoadPendingReason = "";
let ratioFileWatchTimer = null;
let ratioFileWatchInFlight = false;
let ratioFileWatchPath = "";
let ratioFileWatchRevisionToken = "";
let ratioFileWatchDirtyWarnToken = "";
let lastCleanDfmMethodPayload = null;
let lastCleanDfmNotesText = "";
let normalDfmMethodSavePath = "";
let normalDfmMethodSaveName = "";
let currentDfmOutputDataset = "";
// The DFM page never edits the output dataset's category -- it comes from the
// ResQ dataset type -- but it is an owned field, so dropping it from the built
// payload changes the canonical owned revision. Carry the loaded value through.
let currentDfmOutputCategory = "";
// Open-window change alert (advisory): watch the method JSON + output sidecar
// for rewrites by another user or the dependent-propagation job. Self-saves
// pause the watch and rebase its fingerprint through ensureDfmObjectChangeWatch.
const dfmObjectChangeWatch = createMethodObjectChangeWatchController({
  methodType: "dfm",
  onChange: (attribution) => {
    void showObjectUpdatedAlert({
      showMessageBox: showPageMessageBox,
      attribution,
      isDirty: getDfmIsDirty,
      onBlockedRefresh: () => {
        postDfmStatus(
          "Unsaved DFM changes block the refresh. Save or discard them, then reopen the window.",
          { tone: "warn" },
        );
      },
    });
  },
});
wireSamePropagationScopePause({
  watch: dfmObjectChangeWatch,
  getProject: getResolvedProjectName,
  getReservingClass: getResolvedReservingClass,
});

function ensureDfmObjectChangeWatch(methodName, sidecar = null) {
  dfmObjectChangeWatch.ensure({
    projectName: getResolvedProjectName(),
    reservingClass: getResolvedReservingClass(),
    methodName,
    outputDataset: currentDfmOutputDataset,
    // What this window now has in view; a share read reporting this write, or
    // any earlier one, is not an outside change however late it arrives.
    selfWriteStamp: sidecar?.updated_at,
  });
}
let currentOwnedRevision = "";
let currentDerivedRevision = "";
let currentPublicationRevision = "";
let checkedExcelAppliedRevision = "";
let checkingExcelAppliedRevision = "";
let dfmPreviewTimer = null;
let dfmPreviewGeneration = 0;
let dfmPreviewAbortController = null;
const DFM_INSTANCE_PRESENCE_EVENT = "arcrho:dfm-instance-presence";
const DFM_LOCAL_LOOKUP_DEBUG_STATUS = true; // Temporary debug aid.
const DFM_ANALYSIS_DECIMALS = 6;
const DFM_AVERAGE_FORMULA_DECIMALS = 6;
const DFM_METHOD_FILE_WATCH_INTERVAL_MS = 2000;

function decodeFileNameSegment(value) {
  return String(value || "").replace(/_%([0-9A-Fa-f]{2})_/g, (match, hex) => {
    const code = Number.parseInt(hex, 16);
    return Number.isFinite(code) ? String.fromCharCode(code) : match;
  });
}

function getDfmMethodNameFromPath(path) {
  const filename = String(path || "").split(/[\\/]/).pop() || "";
  const stem = filename.replace(/\.json$/i, "");
  const rawName = stem.startsWith("DFM@") ? stem.slice(4) : stem;
  return decodeFileNameSegment(rawName).trim();
}

function getRatioLoadReasonPriority(reason) {
  const key = String(reason || "").trim().toLowerCase();
  switch (key) {
    case "details-change":
      return 50;
    case "global-changed":
    case "dataset-updated":
      return 40;
    case "init":
      return 30;
    case "tab-activated":
      return 10;
    default:
      return 20;
  }
}

function chooseRatioLoadReason(prevReason, nextReason) {
  const prev = String(prevReason || "").trim();
  const next = String(nextReason || "").trim();
  if (!prev) return next;
  if (!next) return prev;
  return getRatioLoadReasonPriority(next) >= getRatioLoadReasonPriority(prev) ? next : prev;
}

function emitDfmInstancePresence(status) {
  try {
    window.dispatchEvent(new CustomEvent(DFM_INSTANCE_PRESENCE_EVENT, { detail: { status } }));
  } catch {
    // ignore
  }
}

function getTrimmedInputValue(id) {
  return String(document.getElementById(id)?.value || "").trim();
}

function postDfmStatus(text, options = {}) {
  window.parent.postMessage(
    {
      type: "arcrho:status",
      text: String(text || ""),
      ...(options?.tone ? { tone: options.tone } : {}),
    },
    "*",
  );
}

function requestProjectInstanceDatasetTableRefresh() {
  try {
    window.parent?.postMessage({ type: "arcrho:project-instance-refresh-datasets" }, "*");
  } catch {
    // ignore stale parent frames
  }
}

function normalizeDfmIdentityKey(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toLowerCase();
}

async function deleteOldDfmIdentityFiles(oldName, newName) {
  const oldKey = normalizeDfmIdentityKey(oldName);
  const newKey = normalizeDfmIdentityKey(newName);
  if (!oldKey || oldKey === newKey) return { ok: true, skipped: true };
  const projectName = String(getRatioSaveProjectName() || getResolvedProjectName() || "").trim();
  const reservingClass = String(getResolvedReservingClass() || "").trim();
  if (!projectName || !reservingClass) return { ok: false, error: "Missing project or reserving class for old DFM cleanup." };
  try {
    const response = await fetch("/datasets/cached/delete", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        project_name: projectName,
        reserving_class: reservingClass,
        dataset_names: [oldName],
      }),
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok || payload?.ok === false) {
      return { ok: false, error: String(payload?.detail || payload?.error || `HTTP ${response.status}`), data: payload };
    }
    return { ok: true, data: payload };
  } catch (err) {
    return { ok: false, error: String(err?.message || err || "Old DFM cleanup failed.") };
  }
}

function postDfmLookupDebugStatus(text, options = {}) {
  if (!DFM_LOCAL_LOOKUP_DEBUG_STATUS) return;
  const reason = String(options?.reason || "").trim();
  const suffix = reason ? ` [${reason}]` : "";
  postDfmStatus(`Debug: DFM local method lookup ${text}${suffix}`);
}

function getRevisionToken(revision) {
  if (!revision || typeof revision !== "object") return "";
  const path = String(revision.path || "");
  const size = Number.isFinite(Number(revision.size)) ? String(Number(revision.size)) : "";
  const mtimeMs = Number.isFinite(Number(revision.mtimeMs)) ? String(Number(revision.mtimeMs)) : "";
  const hash = String(revision.hash || "");
  return `${path}|${size}|${mtimeMs}|${hash}`;
}

function rememberDfmMethodFileRevision(path, revision) {
  ratioFileWatchPath = String(path || revision?.path || "");
  ratioFileWatchRevisionToken = getRevisionToken(revision);
  ratioFileWatchDirtyWarnToken = "";
}

function rememberNormalDfmMethodSavePath(path) {
  normalDfmMethodSavePath = String(path || "").trim();
  normalDfmMethodSaveName = getDfmMethodNameFromPath(normalDfmMethodSavePath);
}

export async function resolveCurrentDfmMethodSavePath() {
  return normalDfmMethodSavePath || await buildRatioSavePath();
}

function clearDfmMethodFileRevision(path = "") {
  ratioFileWatchPath = String(path || "");
  ratioFileWatchRevisionToken = "";
  ratioFileWatchDirtyWarnToken = "";
}

async function readDfmMethodFileRevision(hostApi, path) {
  if (typeof hostApi?.getFileRevision === "function") {
    return hostApi.getFileRevision({ path });
  }
  if (typeof hostApi?.readJsonFile === "function") {
    return hostApi.readJsonFile({ path });
  }
  return { exists: false };
}

async function refreshDfmMethodFileRevision(path) {
  const hostApi = getHostApi();
  if (!hostApi) return;
  try {
    const result = await readDfmMethodFileRevision(hostApi, path);
    if (result?.exists) {
      rememberDfmMethodFileRevision(path, result.revision);
    } else {
      clearDfmMethodFileRevision(path);
    }
  } catch {
    clearDfmMethodFileRevision(path);
  }
}

function hasRequiredDfmInputs() {
  const project = getResolvedProjectName();
  const reservingClass = getResolvedReservingClass();
  const tri = getTrimmedInputValue("triInput");
  const outputVector = getTrimmedInputValue("dfmOutputVector");
  const methodName = getTrimmedInputValue("dfmMethodName");
  const originLen = getTrimmedInputValue("originLenSelect");
  const devLen = getTrimmedInputValue("devLenSelect");
  return !!(project && reservingClass && tri && outputVector && methodName && originLen && devLen);
}

function hasRequiredDfmLookupInputs() {
  const project = getResolvedProjectName();
  const reservingClass = getResolvedReservingClass();
  const methodName = getTrimmedInputValue("dfmMethodName");
  return !!(project && reservingClass && methodName);
}

function getDfmJsonTab(payload, tabKey) {
  const tab = payload && typeof payload === "object" && !Array.isArray(payload) ? payload[tabKey] : null;
  return tab && typeof tab === "object" && !Array.isArray(tab) ? tab : {};
}

function getDfmDetailsTab(payload) {
  return getDfmJsonTab(payload, "details_tab");
}

function getDfmDataTab(payload) {
  return getDfmJsonTab(payload, "data_tab");
}

function getDfmRatiosTab(payload) {
  return getDfmJsonTab(payload, "ratios_tab");
}

function dfmDatasetFormulaInputs(payload) {
  const formulas = getDfmJsonTab(getDfmRatiosTab(payload), "average_formulas");
  const inputs = Array.isArray(formulas.inputs) ? formulas.inputs : [];
  const settings = getDfmJsonTab(formulas, "custom_average_formula_settings");
  const averageTypes = Array.isArray(settings.average_type) ? settings.average_type : [];
  const out = [];
  inputs.forEach((row, index) => {
    if (String(averageTypes[index] || "").trim().toLowerCase() !== "user_entry") return;
    if (!Array.isArray(row)) return;
    row.forEach((formula) => {
      if (containsDfmDatasetReference(formula)) out.push(String(formula));
    });
  });
  return out;
}

// Dataset-referenced User Entry values are kept fresh by the Engine
// dependent-propagation walk when a referenced dataset is saved, so an opened
// DFM shows its persisted values and stays clean. Resolving the references
// here only warms the session value cache that lets those formulas re-evaluate
// live after ratio edits; it never changes values, dirty state, or files.
function warmDfmDatasetReferenceCache(payload) {
  const datasetFormulas = dfmDatasetFormulaInputs(payload);
  if (!datasetFormulas.length) return;
  resolveDfmDatasetReferencesInFormulas(datasetFormulas).catch(() => {
    // Best-effort: stored values stay in use until a reference resolves
    // through a commit, tooltip, or linked refresh.
  });
}

function getDfmRatioTriangleTab(payload) {
  return getDfmJsonTab(getDfmRatiosTab(payload), "ratio_triangle");
}

function getDfmResultsTab(payload) {
  return getDfmJsonTab(payload, "results_tab");
}

function getSavedInputTriangleValue(payload) {
  const details = getDfmDetailsTab(payload);
  if ("input_triangle" in details) return String(details["input_triangle"] ?? "");
  return null;
}

function normalizeSavedLengthValue(value) {
  const raw = Number.parseInt(String(value ?? "").trim(), 10);
  return Number.isFinite(raw) && raw > 0 ? String(raw) : "";
}

function readSelectedLengthNumber(id, fallback = 12) {
  const raw = Number.parseInt(getTrimmedInputValue(id), 10);
  return Number.isFinite(raw) && raw > 0 ? raw : fallback;
}

function getSavedOriginLengthValue(payload) {
  const details = getDfmDetailsTab(payload);
  if ("origin_length" in details) return normalizeSavedLengthValue(details["origin_length"]);
  return null;
}

function getSavedDevelopmentLengthValue(payload) {
  const details = getDfmDetailsTab(payload);
  if ("development_length" in details) return normalizeSavedLengthValue(details["development_length"]);
  return null;
}

function applySavedSelectValueToUi(id, rawValue) {
  if (rawValue == null) return false;
  const select = document.getElementById(id);
  if (!select) return false;
  const next = String(rawValue ?? "").trim();
  if (!next) return false;
  if (![...select.options].some((opt) => String(opt.value) === next)) {
    const opt = document.createElement("option");
    opt.value = next;
    opt.textContent = next;
    select.appendChild(opt);
  }
  if (String(select.value ?? "") === next) return false;
  select.value = next;
  return true;
}

function getSavedMethodNameValue(payload) {
  const details = getDfmDetailsTab(payload);
  if ("name" in details) return String(details["name"] ?? "");
  return null;
}

function applySavedMethodNameToUi(rawValue) {
  if (rawValue == null) return;
  const input = document.getElementById("dfmMethodName");
  if (!input) return;
  const next = String(rawValue ?? "").trim();
  const prev = String(input.value || "").trim();
  if (next === prev) return;
  input.dataset.programmatic = "1";
  input.value = next;
  // `wireMethodName()` handles title/localStorage sync on input without triggering
  // another local-method lookup.
  input.dispatchEvent(new Event("input", { bubbles: true }));
}

function getSavedOutputTypeValue(payload) {
  const details = getDfmDetailsTab(payload);
  if ("output_type" in details) return String(details["output_type"] ?? "");
  return null;
}

function applySavedOutputTypeToUi(rawValue) {
  if (rawValue == null) return;
  const input = document.getElementById("dfmOutputVector");
  if (!input) return;
  const next = String(rawValue ?? "").trim();
  const prev = String(input.value || "").trim();
  if (next === prev) return;
  input.value = next;
  // Keep the picker module's committed value in sync without opening/revalidating
  // the dropdown during programmatic load.
  input.dispatchEvent(new CustomEvent("arcrho:output-type-selected", { detail: { value: next } }));
}

function applySavedInputTriangleToUi(rawValue) {
  if (rawValue == null) return false;
  const triInput = document.getElementById("triInput");
  if (!triInput) return false;
  const next = String(rawValue ?? "").trim();
  const prev = String(triInput.value || "").trim();
  if (next === prev) return false;
  triInput.value = next;
  return true;
}

function getSavedDecimalPlacesValue(payload) {
  const details = getDfmDetailsTab(payload);
  if ("decimal_places" in details) return details["decimal_places"];
  return null;
}

function applySavedDecimalPlacesToUi(rawValue) {
  if (rawValue == null) return;
  const input = document.getElementById("decimalPlaces");
  if (!input) return;
  const parsed = Number.parseInt(String(rawValue).trim(), 10);
  if (!Number.isFinite(parsed)) return;
  const normalized = String(Math.max(0, Math.min(6, parsed)));
  if (String(input.value ?? "") === normalized) return;
  input.dataset.programmatic = "1";
  input.value = normalized;
  input.dispatchEvent(new Event("change", { bubbles: true }));
}

function getSavedUltimateRatioDecimalPlacesValue(payload) {
  const results = getDfmResultsTab(payload);
  if ("ultimate_ratio_decimal_places" in results) return results["ultimate_ratio_decimal_places"];
  return null;
}

const MONTH_NAME_TO_NUM = new Map([
  ["jan", 1], ["january", 1],
  ["feb", 2], ["february", 2],
  ["mar", 3], ["march", 3],
  ["apr", 4], ["april", 4],
  ["may", 5],
  ["jun", 6], ["june", 6],
  ["jul", 7], ["july", 7],
  ["aug", 8], ["august", 8],
  ["sep", 9], ["sept", 9], ["september", 9],
  ["oct", 10], ["october", 10],
  ["nov", 11], ["november", 11],
  ["dec", 12], ["december", 12],
]);

function parseOriginStartMonth(label, baseLen) {
  const s = String(label || "").trim();
  if (!s) return null;

  if (baseLen === 1) {
    const yyyymm = s.match(/^(\d{4})(\d{2})$/);
    if (yyyymm) {
      const year = Number.parseInt(yyyymm[1], 10);
      const month = Number.parseInt(yyyymm[2], 10);
      if (Number.isFinite(year) && month >= 1 && month <= 12) return { year, month };
    }
    const monYear = s.match(/^([A-Za-z]{3,9})\s+(\d{4})$/);
    if (monYear) {
      const month = MONTH_NAME_TO_NUM.get(monYear[1].toLowerCase());
      const year = Number.parseInt(monYear[2], 10);
      if (month && Number.isFinite(year)) return { year, month };
    }
    return null;
  }

  if (baseLen === 3) {
    const yq = s.match(/^(\d{4})\s*Q([1-4])$/i);
    if (yq) {
      const year = Number.parseInt(yq[1], 10);
      const q = Number.parseInt(yq[2], 10);
      return { year, month: (q - 1) * 3 + 1 };
    }
    const qy = s.match(/^Q([1-4])\s*(\d{4})$/i);
    if (qy) {
      const q = Number.parseInt(qy[1], 10);
      const year = Number.parseInt(qy[2], 10);
      return { year, month: (q - 1) * 3 + 1 };
    }
    return null;
  }

  if (baseLen === 6) {
    const yh = s.match(/^(\d{4})\s*H([1-2])$/i);
    if (yh) {
      const year = Number.parseInt(yh[1], 10);
      const h = Number.parseInt(yh[2], 10);
      return { year, month: (h - 1) * 6 + 1 };
    }
    const hy = s.match(/^H([1-2])\s*(\d{4})$/i);
    if (hy) {
      const h = Number.parseInt(hy[1], 10);
      const year = Number.parseInt(hy[2], 10);
      return { year, month: (h - 1) * 6 + 1 };
    }
    return null;
  }

  if (baseLen === 12) {
    const yearOnly = s.match(/^(\d{4})$/);
    if (yearOnly) {
      const year = Number.parseInt(yearOnly[1], 10);
      if (Number.isFinite(year)) return { year, month: 1 };
    }
    return null;
  }

  return null;
}

function aggregateResultsVectorByLength(vector, originLabels, baseLen, targetLen) {
  if (!Array.isArray(vector) || !vector.length) return [];
  const factor = targetLen / baseLen;
  if (!Number.isFinite(factor) || factor <= 1 || Math.floor(factor) !== factor) return [];

  const labels = Array.isArray(originLabels) ? originLabels : [];
  const canUseLabelBuckets = labels.length === vector.length && (baseLen === 1 || baseLen === 3 || baseLen === 6 || baseLen === 12);
  if (canUseLabelBuckets) {
    const orderedKeys = [];
    const bucketMap = new Map();
    let parseFailed = false;
    for (let i = 0; i < vector.length; i++) {
      const parsed = parseOriginStartMonth(labels[i], baseLen);
      if (!parsed) {
        parseFailed = true;
        break;
      }
      const bucketMonth = Math.floor((parsed.month - 1) / targetLen) * targetLen + 1;
      const key = `${parsed.year}-${bucketMonth}`;
      if (!bucketMap.has(key)) {
        bucketMap.set(key, { sum: 0, hasValue: false });
        orderedKeys.push(key);
      }
      const bucket = bucketMap.get(key);
      const num = Number(vector[i]);
      if (Number.isFinite(num)) {
        bucket.sum += num;
        bucket.hasValue = true;
      }
    }
    if (!parseFailed) {
      return orderedKeys.map((key) => {
        const bucket = bucketMap.get(key);
        return bucket?.hasValue ? bucket.sum : null;
      });
    }
  }

  const out = [];
  for (let i = 0; i < vector.length; i += factor) {
    let sum = 0;
    let hasValue = false;
    const end = Math.min(i + factor, vector.length);
    for (let j = i; j < end; j++) {
      const num = Number(vector[j]);
      if (!Number.isFinite(num)) continue;
      sum += num;
      hasValue = true;
    }
    out.push(hasValue ? sum : null);
  }
  return out;
}

function buildAggregatedResultVariants(resultVector) {
  const baseOriginRaw = Number.parseInt(String(document.getElementById("originLenSelect")?.value || "").trim(), 10);
  const baseLen = Number.isFinite(baseOriginRaw) ? baseOriginRaw : 12;
  const targetLens = [3, 6, 12].filter((len) => len > baseLen && len % baseLen === 0);
  if (!targetLens.length) return [];

  const originLabels = Array.isArray(state?.model?.origin_labels) ? state.model.origin_labels : [];
  const out = [];
  for (const targetLen of targetLens) {
    const vec = aggregateResultsVectorByLength(resultVector, originLabels, baseLen, targetLen);
    if (!vec.length) continue;
    out.push({
      originLen: targetLen,
      devLen: targetLen,
      vector: vec,
    });
  }
  return out;
}

function getSummaryRowsForPersistence(cfgKey) {
  const savedSummaryRows = cfgKey ? loadCustomSummaryRows(cfgKey) : [];
  const sourceRows = savedSummaryRows.length
    ? savedSummaryRows
    : (summaryRowConfigs.length ? summaryRowConfigs : BASE_SUMMARY_ROWS);
  return sourceRows.map((row) => {
    const { id: _id, ...rowWithoutId } = row || {};
    if (!isUserEntrySummaryRow(rowWithoutId)) return { ...rowWithoutId };
    const {
      values: _values,
      inputs: _inputs,
      formulas: _legacyFormulas,
      ...baseRow
    } = rowWithoutId;
    const rowInputs = Array.isArray(_inputs)
      ? _inputs
      : Array.isArray(_legacyFormulas)
        ? _legacyFormulas
        : null;
    const nextRow = { ...baseRow };
    if (Array.isArray(rowInputs) && rowInputs.some((value) => String(value ?? "").trim())) {
      nextRow.inputs = rowInputs.map((value) => String(value ?? "").trim());
    }
    return nextRow;
  });
}

function buildRatioDisplayHeaderLabels(devs) {
  const ratioLabels = getRatioHeaderLabels(devs);
  return ratioLabels.map((label, index) => {
    const text = String(label ?? "");
    if (index === ratioLabels.length - 1) return text || "Ult";
    return text ? `(${index + 1}) ${text}` : `(${index + 1})`;
  });
}

// A cell with no ratio must stay null here: the exclusion pattern writes its 2
// sentinel for the same cell, and a ratio of 0 in its place leaves the two rows
// trimming to different lengths, which strict validation rejects.
function roundAnalysisValue(value) {
  return roundRatio(ratioNumberOrNull(value), DFM_ANALYSIS_DECIMALS);
}

function roundAverageFormulaValue(value) {
  return roundRatio(ratioNumberOrNull(value), DFM_AVERAGE_FORMULA_DECIMALS);
}

function trimTrailingNulls(row) {
  const out = Array.isArray(row) ? row.slice() : [];
  while (out.length && out[out.length - 1] === null) {
    out.pop();
  }
  return out;
}

function normalizeSummaryUserEntryValue(raw) {
  const value = Number(raw);
  return Number.isFinite(value) && value > 0 ? value : 1;
}

function isUserEntrySummaryRow(cfg) {
  return String(cfg?.averageType || "").trim().toLowerCase() === "user_entry";
}

function buildCalculatedRatioTriangleValues() {
  const model = state.model;
  if (!model || !Array.isArray(model.values) || !Array.isArray(model.mask)) return [];
  const values = model.values;
  const mask = model.mask;
  const rowCount = Array.isArray(model.origin_labels) ? model.origin_labels.length : values.length;
  const devs = getEffectiveDevLabelsForModel(model);
  const ratioLabels = getRatioHeaderLabels(devs);
  const out = [];
  for (let r = 0; r < rowCount; r++) {
    const row = [];
    for (let c = 0; c < ratioLabels.length; c++) {
      if (c >= devs.length - 1 || !mask?.[r]?.[c] || !mask?.[r]?.[c + 1]) {
        row.push(null);
        continue;
      }
      row.push(roundAnalysisValue(calcRatio(values?.[r]?.[c], values?.[r]?.[c + 1])));
    }
    out.push(trimTrailingNulls(row));
  }
  return out;
}

function trimMatrixToReferenceRowShape(matrix, reference) {
  if (!Array.isArray(matrix)) return [];
  return matrix.map((row, rowIndex) => {
    const out = Array.isArray(row) ? row.slice() : [];
    const referenceRow = Array.isArray(reference?.[rowIndex]) ? reference[rowIndex] : null;
    return referenceRow ? out.slice(0, referenceRow.length) : out;
  });
}

function getSummaryRowsForValues() {
  return Array.isArray(summaryRowConfigs) && summaryRowConfigs.length
    ? summaryRowConfigs
    : buildSummaryRows();
}

function buildAverageFormulaValues() {
  const model = state.model;
  if (!model || !Array.isArray(model.values) || !Array.isArray(model.mask)) return [];
  const rows = getSummaryRowsForValues();
  const devs = getEffectiveDevLabelsForModel(model);
  const ratioLabels = getRatioHeaderLabels(devs);
  const values = rows.map(() => new Array(ratioLabels.length).fill(null));
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const cfg = rows[rowIndex];
    for (let c = 0; c < ratioLabels.length; c++) {
      if (c >= devs.length - 1) {
        values[rowIndex][c] = roundAverageFormulaValue(getSummaryRowTailFactor(cfg, c));
        continue;
      }
      if (isUserEntrySummaryRow(cfg)) {
        const raw = Array.isArray(cfg.values) ? cfg.values[c] : 1;
        values[rowIndex][c] = roundAverageFormulaValue(normalizeSummaryUserEntryValue(raw));
        continue;
      }
      const excluded = buildExcludedSetForColumn(model, c, cfg, ratioStrikeSet);
      const summary = computeAverageForColumn(model, c, excluded, cfg, ratioStrikeSet);
      if (summary.totalValid > 0 && summary.totalIncluded === 0) {
        values[rowIndex][c] = roundAverageFormulaValue(1);
        continue;
      }
      const isVolume = String(cfg.base || "volume").toLowerCase() === "volume";
      const hasValue =
        summary.value !== null &&
        (isVolume ? summary.sumA : summary.totalIncluded > 0);
      values[rowIndex][c] = roundAverageFormulaValue(hasValue ? summary.value : 1);
    }
  }
  return values.map((row) => trimTrailingNulls(row));
}

function hydrateUserEntryValuesFromAverageFormulaValues(summaryRows, formulas, averageFormulaValues) {
  if (!Array.isArray(summaryRows) || !Array.isArray(formulas) || !Array.isArray(averageFormulaValues)) {
    return summaryRows;
  }
  const formulaIndexByLabel = new Map();
  formulas.forEach((formula, index) => {
    const key = String(formula || "").replace(/\s+/g, " ").trim().toLowerCase();
    if (key && !formulaIndexByLabel.has(key)) formulaIndexByLabel.set(key, index);
  });
  return summaryRows.map((row) => {
    const isFrozenBenchmark = String(row?.base || "").trim().toLowerCase() === "benchmark";
    if ((!isUserEntrySummaryRow(row) && !isFrozenBenchmark) || Array.isArray(row?.values)) return row;
    const labelKey = String(row?.label || row?.id || "").replace(/\s+/g, " ").trim().toLowerCase();
    const rowIndex = formulaIndexByLabel.get(labelKey);
    const valueRow = Number.isInteger(rowIndex) ? averageFormulaValues[rowIndex] : null;
    if (!Array.isArray(valueRow)) return row;
    return {
      ...row,
      values: valueRow.map((value) => normalizeSummaryUserEntryValue(value)),
    };
  });
}

export async function buildDfmMethodPayloadWithPaths(options = {}) {
  let inputTriangleCsvPath = String(options?.inputTriangleCsvPath || "").trim();
  if (!inputTriangleCsvPath) {
    try {
      inputTriangleCsvPath = await buildInputTriangleCsvPath();
    } catch {
      inputTriangleCsvPath = "";
    }
  }
  let ultimateVectorCsvPath = String(options?.ultimateVectorCsvPath || "").trim();
  if (!ultimateVectorCsvPath) {
    try {
      const dataDir = await getRatioDataDir();
      ultimateVectorCsvPath = `${dataDir}\\${getResultsCsvSuggestedName()}`;
    } catch {
      ultimateVectorCsvPath = "";
    }
  }
  return buildDfmMethodPayload({
    ...options,
    inputTriangleCsvPath,
    ultimateVectorCsvPath,
  });
}

function copyExistingFields(source, keys) {
  const out = {};
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(source, key)) {
      out[key] = source[key];
    }
  }
  return out;
}

function copyExistingField(source, sourceKey, target, targetKey = sourceKey) {
  if (Object.prototype.hasOwnProperty.call(source, sourceKey)) {
    target[targetKey] = source[sourceKey];
  }
}

function buildDfmGroupedMethodPayload(methodPayload) {
  const data = methodPayload && typeof methodPayload === "object" ? methodPayload : {};
  const dataTab = {};
  copyExistingField(data, "origin_labels", dataTab);
  copyExistingField(data, "data_development_labels", dataTab, "development_labels");
  copyExistingField(data, "input_data_triangle_values", dataTab);
  copyExistingField(data, "input_data_triangle_mask", dataTab);
  copyExistingField(data, "data_format", dataTab);
  copyExistingField(data, "number_format", dataTab);
  copyExistingField(data, "data_decimal_places", dataTab, "decimal_places");
  copyExistingField(data, "input_source_revision", dataTab, "source_revision");
  const ratiosTab = {};
  const ratioTriangle = {};
  copyExistingField(data, "origin_labels", ratioTriangle);
  copyExistingField(data, "ratio_development_labels", ratioTriangle, "development_labels");
  copyExistingField(data, "ratio_values", ratioTriangle);
  copyExistingField(data, "excluded", ratioTriangle);
  ratiosTab["ratio_triangle"] = ratioTriangle;
  copyExistingField(data, "average_formulas", ratiosTab);
  copyExistingField(data, "cell_notes", ratiosTab);
  const curvesTab = {};
  copyExistingField(data, "curves_tab", curvesTab, "value");
  const grouped = {
    "json_format": DFM_METHOD_JSON_FORMAT,
    "details_tab": copyExistingFields(data, [
      "name",
      "output_type",
      "output_dataset",
      "output_category",
      "input_triangle",
      "origin_length",
      "development_length",
      "decimal_places",
    ]),
    "data_tab": dataTab,
    "ratios_tab": ratiosTab,
    ...("value" in curvesTab ? { "curves_tab": curvesTab.value } : {}),
    "results_tab": copyExistingFields(data, [
      "ratio_basis_dataset",
      "ratio_basis_data_format",
      "ratio_basis_origin_labels",
      "ratio_basis_values",
      "ratio_basis_number_format",
      "ratio_basis_decimal_places",
      "ratio_basis_source_revision",
      "ultimate_ratio_decimal_places",
      "ultimate_vector",
    ]),
    "method_metadata": copyExistingFields(data, [
      "last_modified",
      "data_refreshed",
      "owned_revision",
      "derived_revision",
      "publication_revision",
    ]),
  };
  return grouped;
}

function recordCleanDfmMethodPayload(payload = null) {
  const cleanPayload = payload || buildDfmMethodPayload();
  try {
    lastCleanDfmMethodPayload = JSON.parse(JSON.stringify(cleanPayload));
  } catch {
    lastCleanDfmMethodPayload = cleanPayload;
  }
  lastCleanDfmNotesText = getDfmNotesText();
}

export function recordCurrentDfmCleanState() {
  recordCleanDfmMethodPayload();
  markDfmClean({ force: true });
}

export async function buildDfmAssistantContextPayload(options = {}) {
  const payload = await buildDfmMethodPayloadWithPaths(options);
  // Method Notes live in the output sidecar, so the method payload never
  // carries them; stamp the transient `method metadata.method notes` carrier
  // (same convention as the RPC bridge apply path) so macros and ArcBot read
  // the live, possibly dirty, Notes tab instead of the persisted sidecar.
  if (payload && typeof payload === "object" && !Array.isArray(payload)) {
    const metadata = payload["method_metadata"];
    if (metadata && typeof metadata === "object" && !Array.isArray(metadata)) {
      metadata["method_notes"] = getDfmNotesText();
    } else {
      payload["method_metadata"] = { "method_notes": getDfmNotesText() };
    }
  }
  return payload;
}

async function refreshDfmDatasetAfterDetailsApply(options = {}) {
  if (options.refreshDataset === false) return;
  const refreshFn = window.ADA_DFM_REFRESH_DATASET;
  if (typeof refreshFn !== "function") return;
  try {
    await refreshFn();
  } catch (err) {
    console.warn("Failed to refresh DFM dataset after applying Details fields:", err);
    postDfmStatus("DFM settings were applied, but the Data table refresh failed.", { tone: "warn" });
  }
}

export async function applyDfmMethodPayload(payload, options = {}) {
  return runDfmProgrammatic(() => applyDfmMethodPayloadProgrammatically(payload, options));
}

function cloneJsonValue(value) {
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return value;
  }
}

function mergePlainObject(target, patch) {
  const out = target && typeof target === "object" && !Array.isArray(target)
    ? cloneJsonValue(target)
    : {};
  Object.entries(patch && typeof patch === "object" && !Array.isArray(patch) ? patch : {})
    .forEach(([key, value]) => {
      if (value && typeof value === "object" && !Array.isArray(value)) {
        out[key] = mergePlainObject(out[key], value);
      } else {
        out[key] = cloneJsonValue(value);
      }
    });
  return out;
}

function projectDfmOwnedPatch(payload) {
  const patch = payload && typeof payload === "object" && !Array.isArray(payload) ? payload : {};
  const projected = {};
  const details = getDfmJsonTab(patch, "details_tab");
  if (Object.keys(details).length) projected["details_tab"] = cloneJsonValue(details);
  const ratios = getDfmJsonTab(patch, "ratios_tab");
  const ratioTriangle = getDfmJsonTab(ratios, "ratio_triangle");
  const projectedRatios = {};
  if (Object.prototype.hasOwnProperty.call(ratioTriangle, "excluded")) {
    projectedRatios["ratio_triangle"] = { excluded: cloneJsonValue(ratioTriangle.excluded) };
  }
  for (const key of ["average_formulas", "cell_notes"]) {
    if (Object.prototype.hasOwnProperty.call(ratios, key)) projectedRatios[key] = cloneJsonValue(ratios[key]);
  }
  if (Object.keys(projectedRatios).length) projected["ratios_tab"] = projectedRatios;
  const results = getDfmJsonTab(patch, "results_tab");
  const projectedResults = {};
  for (const key of ["ratio_basis_dataset", "ultimate_ratio_decimal_places"]) {
    if (Object.prototype.hasOwnProperty.call(results, key)) projectedResults[key] = cloneJsonValue(results[key]);
  }
  if (Object.keys(projectedResults).length) projected["results_tab"] = projectedResults;
  const curves = getDfmJsonTab(patch, "curves_tab");
  if (Object.keys(curves).length) projected["curves_tab"] = cloneJsonValue(curves);
  return projected;
}

export async function applyDfmOwnedPatchPayload(payload, options = {}) {
  const merged = isDfmV2Method(payload)
    ? cloneJsonValue(payload)
    : mergePlainObject(buildDfmMethodPayload(), projectDfmOwnedPatch(payload));
  // Method Notes are sidecar-owned and stripped by canonicalization, so read
  // the transient `method metadata.method notes` carrier (macro results, RPC
  // bridge, ArcBot proposals) from the incoming payload before preview.
  const incomingMetadata = getDfmJsonTab(payload, "method_metadata");
  const hasMethodNotes = Object.prototype.hasOwnProperty.call(incomingMetadata, "method_notes");
  try {
    const response = await previewDfmMethod(merged);
    if (!response?.method || !isDfmV2Method(response.method)) {
      throw new Error("Owned DFM patch preview did not return a canonical v2 method.");
    }
    const applied = await applyDfmMethodPayload(response.method, {
      ...options,
      markClean: false,
      reason: options.reason || "owned-patch",
    });
    if (applied?.ok && hasMethodNotes) {
      // Deliver carried Method Notes to the Notes tab; the next normal Save
      // persists them to the output sidecar through the existing notes field.
      setDfmNotesText(String(incomingMetadata["method_notes"] ?? ""));
    }
    return applied;
  } catch (error) {
    return { ok: false, error: String(error?.message || error || "Could not preview DFM owned patch.") };
  }
}

async function applyDfmMethodPayloadProgrammatically(payload, options = {}) {
  const isV2 = isDfmV2Method(payload);
  let datasetInputsChanged = false;
  if (payload && !Array.isArray(payload)) {
    datasetInputsChanged = applySavedSelectValueToUi("originLenSelect", getSavedOriginLengthValue(payload)) || datasetInputsChanged;
    datasetInputsChanged = applySavedSelectValueToUi("devLenSelect", getSavedDevelopmentLengthValue(payload)) || datasetInputsChanged;
  }

  const ratiosTab = getDfmRatiosTab(payload);
  const ratioTriangle = getDfmRatioTriangleTab(payload);
  const resultsTab = getDfmResultsTab(payload);
  if (isV2) {
    const appliedDetails = getDfmDetailsTab(payload);
    currentDfmOutputCategory = String(
      appliedDetails["output_category"]
        ?? appliedDetails["output dataset_category"]
        ?? currentDfmOutputCategory,
    ).trim();
    const dataTab = getDfmDataTab(payload);
    const snapshotResult = window.ADA_DFM_APPLY_DATASET_SNAPSHOT?.({
      origin_labels: dataTab["origin_labels"],
      dev_labels: dataTab["development_labels"],
      values: dataTab["input_data_triangle_values"],
      mask: dataTab["input_data_triangle_mask"],
      data_format: dataTab["data_format"] || "Triangle",
      number_format: dataTab["number_format"],
      decimal_places: dataTab["decimal_places"],
      source_revision: dataTab["source_revision"],
      source_kind: "dfm-v2-snapshot",
    });
    if (snapshotResult?.ok === false) {
      return { ok: false, error: snapshotResult.error || "Could not hydrate the embedded input snapshot." };
    }
    applyPersistedRatioDerivedSnapshot(ratioTriangle);
  }
  const pattern = Array.isArray(payload) ? payload : ratioTriangle.excluded;
  let applied = applyRatioSelectionPattern(pattern);
  if (payload && !Array.isArray(payload)) {
    const cfgKey = getSummaryConfigKey();
    const averageFormulas = ratiosTab["average_formulas"];
    const cellNotes = ratiosTab["cell_notes"];
    const formulas = getDfmAverageFormulaLabels(averageFormulas);
    const matrix = getDfmAverageFormulaSelectedIndex(averageFormulas);
    const averageFormulaValues = getDfmAverageFormulaValues(averageFormulas);
    const averageFormulaRows = buildDfmSummaryRowsFromAverageFormulaObject(averageFormulas);
    const resolvedSummary = buildDfmSummaryRowsFromAverageFormulas(averageFormulaRows, formulas);
    const summaryRows = hydrateUserEntryValuesFromAverageFormulaValues(
      resolvedSummary.rows,
      formulas,
      averageFormulaValues,
    );
    let summaryUpdated = false;

    if (Array.isArray(summaryRows) && cfgKey) {
      saveCustomSummaryRows(cfgKey, summaryRows);
      summaryUpdated = true;
    }
    if (summaryUpdated) buildSummaryRows();
    applyDfmCellNotesPayload(cellNotes);

    const savedMethodName = getSavedMethodNameValue(payload);
    const savedOutputType = getSavedOutputTypeValue(payload);
    const savedInputTriangle = getSavedInputTriangleValue(payload);
    const savedDecimalPlaces = getSavedDecimalPlacesValue(payload);
    const savedUltimateRatioDecimalPlaces = getSavedUltimateRatioDecimalPlacesValue(payload);
    const ratioBasisDataset = resultsTab["ratio_basis_dataset"] ?? "";
    applySavedOutputTypeToUi(savedOutputType);
    datasetInputsChanged = applySavedInputTriangleToUi(savedInputTriangle) || datasetInputsChanged;
    // Apply saved Name after tri-input restore so custom Names win over
    // any default name derived from the selected Output Vector.
    applySavedMethodNameToUi(savedMethodName);
    applySavedDecimalPlacesToUi(savedDecimalPlaces);
    setResultsUltimateRatioDecimalPlacesSelection(savedUltimateRatioDecimalPlaces, { silent: true, render: false });
    if (isV2) {
      applyPersistedResultsSnapshot(resultsTab);
    } else {
      await setResultsRatioBasisSelection(ratioBasisDataset, { silent: true, render: false });
    }
    if (Array.isArray(formulas) && Array.isArray(matrix)) {
      applyAverageSelectionFromSaved(formulas, matrix);
    }
    applyDfmCurvesTabPayload(getDfmJsonTab(payload, "curves_tab"));
  } else {
    applyDfmCellNotesPayload(null);
    applyDfmCurvesTabPayload(null);
    await setResultsRatioBasisSelection("", { silent: true, render: false });
  }

  if (datasetInputsChanged && !isV2) {
    await refreshDfmDatasetAfterDetailsApply(options);
    if (!applied) {
      applied = applyRatioSelectionPattern(pattern);
    }
  }

  if (applied && options.render !== false) {
    renderRatioTable();
    renderDfmCurvesTab();
    renderResultsTable();
  }
  if (applied && options.markClean !== false) {
    recordCleanDfmMethodPayload();
    markMethodSaved();
    markDfmClean({ force: true });
  }
  if (applied && isV2 && options.markClean !== false) {
    const details = getDfmDetailsTab(payload);
    const metadata = getDfmJsonTab(payload, "method_metadata");
    currentDfmOutputDataset = String(details["output_dataset"] || currentDfmOutputDataset || details.name || "").trim();
    currentOwnedRevision = String(metadata["owned_revision"] || currentOwnedRevision || "").trim();
    currentDerivedRevision = String(metadata["derived_revision"] || currentDerivedRevision || "").trim();
    currentPublicationRevision = String(metadata["publication_revision"] || currentPublicationRevision || "").trim();
  }
  if (applied && getCurrentDfmTab() === "audit") {
    void refreshDfmAuditLog();
  }
  return { ok: applied, datasetInputsChanged };
}

function applyDfmAggregateRevisions(response, method) {
  const metadata = getDfmJsonTab(method, "method_metadata");
  currentOwnedRevision = String(response?.owned_revision || metadata["owned_revision"] || "").trim();
  currentDerivedRevision = String(response?.derived_revision || metadata["derived_revision"] || "").trim();
  currentPublicationRevision = String(response?.publication_revision || metadata["publication_revision"] || "").trim();
}

function syncDfmIdentityQuery(method) {
  const details = getDfmDetailsTab(method);
  const methodName = String(details.name || "").trim();
  const outputDataset = String(details["output_dataset"] || "").trim();
  if (!methodName) return;
  if (globalThis.history?.replaceState && globalThis.location?.href) {
    try {
      const url = new URL(globalThis.location.href);
      url.searchParams.set("method_name", methodName);
      if (outputDataset) url.searchParams.set("output_dataset", outputDataset);
      globalThis.history.replaceState(globalThis.history.state, "", url.toString());
    } catch {
      // Identity state remains available in memory when URL replacement is unavailable.
    }
  }
  try {
    const query = new URLSearchParams(globalThis.location?.search || "");
    window.parent?.postMessage?.({
      type: "arcrho:dfm-identity",
      inst: query.get("inst") || "",
      methodName,
      outputDataset,
    }, "*");
  } catch {
    // The aggregate loader still retains both identities locally.
  }
}

function scheduleDfmExcelFreshnessCheck(method) {
  const metadata = getDfmJsonTab(method, "method_metadata");
  const appliedRevision = [
    currentOwnedRevision || metadata["owned_revision"],
    currentDerivedRevision || metadata["derived_revision"],
    currentPublicationRevision || metadata["publication_revision"],
  ]
    .map((value) => String(value || "").trim())
    .filter(Boolean)
    .join("\u001f");
  if (
    !appliedRevision
    || appliedRevision === checkedExcelAppliedRevision
    || appliedRevision === checkingExcelAppliedRevision
  ) return;
  checkingExcelAppliedRevision = appliedRevision;
  cancelDfmExcelFreshnessCheck();
  setTimeout(async () => {
    try {
      if (getDfmIsDirty() || appliedRevision !== checkingExcelAppliedRevision) return;
      const result = await checkDfmExcelLinkFreshness();
      if (result?.aborted || appliedRevision !== checkingExcelAppliedRevision || getDfmIsDirty()) return;
      checkedExcelAppliedRevision = appliedRevision;
      setDfmExcelFreshnessState(result);
      const invalidLinks = Array.isArray(result?.invalidLinks) ? result.invalidLinks : [];
      if (invalidLinks.length) {
        // A reference the workbook can no longer answer is not a freshness
        // note: the stored ratios stay in place, the cells are already red,
        // and the user is told which reference to fix before anything else.
        await showExcelLinkFailureAlert({ failures: invalidLinks, valueNoun: "linked ratio cell" });
        return;
      }
      const staleCount = Number(result?.staleCount || 0);
      const unverifiedCount = Number(result?.unverifiedCount || 0);
      if (staleCount || unverifiedCount) {
        const parts = [];
        if (staleCount) parts.push(`${staleCount} stale`);
        if (unverifiedCount) parts.push(`${unverifiedCount} unverified`);
        postDfmStatus(`Excel links: ${parts.join(", ")}. Stored values remain active; use Links > Refresh to update.`, { tone: "warn" });
      }
    } finally {
      if (checkingExcelAppliedRevision === appliedRevision) checkingExcelAppliedRevision = "";
    }
  }, 0);
}

export async function loadRatioSelectionIfExists(reason) {
  // The method JSON and its input snapshot come from the same network drive as
  // dataset data, so the DFM grids stay in the shared loading state until this
  // read settles either way.
  const gridPlaceholderToken = beginDatasetGridLoading({ message: "Loading DFM method" });
  try {
    return await loadRatioSelectionIfExistsOnce(reason);
  } finally {
    endDatasetGridLoading(gridPlaceholderToken);
  }
}

async function loadRatioSelectionIfExistsOnce(reason) {
  postDfmLookupDebugStatus("triggered", { reason });
  if (!hasRequiredDfmLookupInputs()) {
    postDfmLookupDebugStatus("skipped (waiting for required fields)", { reason });
    emitDfmInstancePresence("incomplete");
    return { ok: false, incomplete: true };
  }
  if (getDfmIsDirty()) {
    postDfmLookupDebugStatus("skipped (dirty)", { reason });
    return { ok: false, dirty: true };
  }

  cancelDfmExcelFreshnessCheck();
  checkingExcelAppliedRevision = "";
  const identity = readDfmMethodIdentityFromPage();
  if (!identity.output_dataset && currentDfmOutputDataset) {
    identity.output_dataset = currentDfmOutputDataset;
  }
  postDfmStatus("Loading DFM method...");
  try {
    const response = await loadDfmMethod(identity);
    const method = response?.method;
    if (!method || !isDfmV2Method(method)) {
      throw new Error("DFM load did not return a canonical v2 method.");
    }
    applyDfmAggregateRevisions(response, method);
    const applied = await applyDfmMethodPayload(method, { reason: reason || "dfm-open" });
    if (!applied?.ok) throw new Error(applied?.error || "The DFM method could not be applied.");
    const details = getDfmDetailsTab(method);
    currentDfmOutputDataset = String(details["output_dataset"] || identity.output_dataset || details.name || "").trim();
    syncDfmIdentityQuery(method);
    hydrateDfmOutputSidecar(response?.sidecar, {
      hydrateNotes: true,
      outputDataset: currentDfmOutputDataset,
    });
    // The raw method-load sidecar carries bare names; the Details rows want the
    // enriched graph, so they are read through the shared Details loader.
    void refreshDfmDetailsDependencies(currentDfmOutputDataset);
    recordCleanDfmMethodPayload(method);
    markDfmClean({ force: true });
    emitDfmInstancePresence("found");
    const sidecarStatus = response?.sidecar?.status;
    const reviewNeeded = Number(sidecarStatus) === 2
      || /review/i.test(String(sidecarStatus || response?.sidecar?.status_label || ""));
    warmDfmDatasetReferenceCache(method);
    postDfmStatus(
      reviewNeeded ? "DFM loaded with Review Needed status." : "Ready",
      reviewNeeded ? { tone: "warn" } : {},
    );
    scheduleDfmExcelFreshnessCheck(method);
    ensureDfmObjectChangeWatch(details.name, response?.sidecar);
    return { ok: true, method, sidecar: response?.sidecar };
  } catch (error) {
    if (Number(error?.status) === 404) {
      emitDfmInstancePresence("missing");
      postDfmStatus("This method object has not been created yet.", { tone: "warn" });
      return { ok: false, missing: true, error: error.message };
    }
    emitDfmInstancePresence("incomplete");
    postDfmStatus(`DFM load failed: ${String(error?.message || error)}`, { tone: "error" });
    return { ok: false, error: String(error?.message || error) };
  }
}

export function scheduleRatioSelectionLoad(reason) {
  ratioLoadPendingReason = chooseRatioLoadReason(ratioLoadPendingReason, reason);
  if (ratioLoadTimer) clearTimeout(ratioLoadTimer);
  ratioLoadTimer = setTimeout(() => {
    const scheduledReason = ratioLoadPendingReason || reason;
    ratioLoadPendingReason = "";
    ratioLoadTimer = null;
    loadRatioSelectionIfExists(scheduledReason);
  }, 120);
}

export async function restoreCleanDfmMethodState() {
  if (lastCleanDfmMethodPayload) {
    const cleanNotes = lastCleanDfmNotesText;
    const result = await applyDfmMethodPayload(lastCleanDfmMethodPayload, { reason: "cancel", markClean: true });
    if (result?.ok) setDfmNotesText(cleanNotes);
    return result;
  }
  return loadRatioSelectionIfExists("cancel");
}

export function buildDfmMethodPayload(options = {}) {
  const devs = getEffectiveDevLabelsForModel(state?.model || {});
  const originLabels = Array.isArray(state?.model?.origin_labels)
    ? state.model.origin_labels.map((label) => String(label ?? ""))
    : [];
  const dataDevelopmentLabels = devs.map((label) => String(label ?? ""));
  const ratioDevelopmentLabels = buildRatioDisplayHeaderLabels(devs);
  const avgSelection = buildAverageSelectionPayload();
  const calculatedRatioTriangleValues = buildCalculatedRatioTriangleValues();
  const pattern = trimMatrixToReferenceRowShape(buildRatioSelectionPattern(), calculatedRatioTriangleValues);
  const averageFormulaValues = buildAverageFormulaValues();
  const cellNotes = buildDfmCellNotesPayload();
  const ratioBasisDataset = getResultsRatioBasisSelection();
  const ratioBasisSnapshot = getResultsRatioBasisSnapshot();
  const outputVector = getTrimmedInputValue("dfmOutputVector");
  const methodName = getTrimmedInputValue("dfmMethodName");
  const queryOutputDataset = new URLSearchParams(globalThis.location?.search || "").get("output_dataset") || "";
  const outputDataset = String(options?.outputDataset || currentDfmOutputDataset || queryOutputDataset || methodName).trim();
  const inputTriangle = getTrimmedInputValue("triInput");
  const originLength = readSelectedLengthNumber("originLenSelect");
  const developmentLength = readSelectedLengthNumber("devLenSelect");
  const decimalPlaces = getDfmDecimalPlaces();
  const ultimateRatioDecimalPlaces = getResultsUltimateRatioDecimalPlacesSelection();
  const cfgKey = getSummaryConfigKey();
  const summaryRows = getSummaryRowsForPersistence(cfgKey);
  const data = {
    excluded: pattern,
    "origin_labels": originLabels,
    "data_development_labels": dataDevelopmentLabels,
    "ratio_development_labels": ratioDevelopmentLabels,
    "input_data_triangle_values": Array.isArray(state?.model?.values)
      ? state.model.values.map((row) => (Array.isArray(row) ? row.slice() : []))
      : [],
    "input_data_triangle_mask": Array.isArray(state?.model?.mask)
      ? state.model.mask.map((row) => (Array.isArray(row) ? row.map(Boolean) : []))
      : [],
    "data_format": String(state?.model?.data_format || "Triangle"),
    "number_format": String(state?.model?.number_format || "Number"),
    "data_decimal_places": Number.isFinite(Number(state?.model?.decimal_places))
      ? Number(state.model.decimal_places)
      : decimalPlaces,
    "input_source_revision": String(state?.model?.source_revision || state?.model?.revision || ""),
    "ratio_values": calculatedRatioTriangleValues,
    "average_formulas": buildDfmAverageFormulaObject(summaryRows, avgSelection.matrix, averageFormulaValues),
    "cell_notes": cellNotes,
    "curves_tab": buildDfmCurvesTabPayload(),
    "ultimate_vector": buildResultsVector(),
    name: methodName,
    "output_type": outputVector,
    "output_dataset": outputDataset,
    // Omitted rather than sent empty when unknown: Save merges owned fields as a
    // patch, so an empty value would clear the category stored on disk.
    ...(currentDfmOutputCategory ? { "output_category": currentDfmOutputCategory } : {}),
    "input_triangle": inputTriangle,
    "origin_length": originLength,
    "development_length": developmentLength,
    "decimal_places": decimalPlaces,
    "ultimate_ratio_decimal_places": ultimateRatioDecimalPlaces,
    "ratio_basis_dataset": ratioBasisDataset,
    ...ratioBasisSnapshot,
    "last_modified": new Date().toISOString(),
    "data_refreshed": String(lastCleanDfmMethodPayload?.["method_metadata"]?.["data_refreshed"] || ""),
    "owned_revision": currentOwnedRevision,
    "derived_revision": currentDerivedRevision,
    "publication_revision": currentPublicationRevision,
  };
  return buildDfmGroupedMethodPayload(data);
}

function normalizeRatioMatrixCellValue(matrix, row, col) {
  const sourceRow = Array.isArray(matrix?.[row]) ? matrix[row] : [];
  if (col >= sourceRow.length) return 2;
  const raw = sourceRow[col];
  if (raw === true) return 1;
  if (raw === false) return 0;
  const num = Number(raw);
  return Number.isFinite(num) ? num : 2;
}

function matrixMaxColumnCount(...matrices) {
  let count = 0;
  matrices.forEach((matrix) => {
    if (!Array.isArray(matrix)) return;
    matrix.forEach((row) => {
      if (Array.isArray(row)) count = Math.max(count, row.length);
    });
  });
  return count;
}

function buildChangedRatioCells(prevPattern, nextPattern) {
  const rows = Math.max(
    Array.isArray(prevPattern) ? prevPattern.length : 0,
    Array.isArray(nextPattern) ? nextPattern.length : 0,
  );
  const cells = [];
  for (let r = 0; r < rows; r++) {
    const cols = matrixMaxColumnCount([prevPattern?.[r]], [nextPattern?.[r]]);
    for (let c = 0; c < cols; c++) {
      if (normalizeRatioMatrixCellValue(prevPattern, r, c) !== normalizeRatioMatrixCellValue(nextPattern, r, c)) {
        cells.push({ r, c });
      }
    }
  }
  return cells;
}

function buildAverageMatrixByLabel(formulas, matrix) {
  const out = new Map();
  const formulaList = Array.isArray(formulas) ? formulas : [];
  formulaList.forEach((formula, row) => {
    const label = String(formula || "").trim();
    if (!label) return;
    const sourceRow = Array.isArray(matrix?.[row]) ? matrix[row] : [];
    out.set(label, sourceRow.map((value) => (Number(value) === 1 ? 1 : 0)));
  });
  return out;
}

function buildChangedAverageCells(prevFormulas, prevMatrix, nextFormulas, nextMatrix) {
  const prevByLabel = buildAverageMatrixByLabel(prevFormulas, prevMatrix);
  const nextByLabel = buildAverageMatrixByLabel(nextFormulas, nextMatrix);
  const labels = new Set([...prevByLabel.keys(), ...nextByLabel.keys()]);
  const cells = [];
  labels.forEach((label) => {
    const prevRow = prevByLabel.get(label) || [];
    const nextRow = nextByLabel.get(label) || [];
    const cols = Math.max(prevRow.length, nextRow.length);
    for (let c = 0; c < cols; c++) {
      if (Number(prevRow[c] || 0) !== Number(nextRow[c] || 0)) cells.push({ label, c });
    }
  });
  return cells;
}

function buildDfmExternalChangedCells(nextPayload) {
  const prevPattern = buildRatioSelectionPattern();
  const nextPattern = getDfmRatioTriangleTab(nextPayload).excluded;
  const prevAverage = buildAverageSelectionPayload();
  const nextAverage = getDfmAverageFormulaSelectedIndex(getDfmRatiosTab(nextPayload)["average_formulas"]);
  return {
    ratioCells: buildChangedRatioCells(prevPattern, nextPattern),
    averageCells: buildChangedAverageCells(
      prevAverage.formulas,
      prevAverage.matrix,
      getDfmAverageFormulaLabels(getDfmRatiosTab(nextPayload)["average_formulas"]),
      nextAverage,
    ),
  };
}

async function checkDfmMethodFileWatch() {
  if (ratioFileWatchInFlight) return;
  if (!hasRequiredDfmLookupInputs()) {
    clearDfmMethodFileRevision();
    return;
  }
  const hostApi = getHostApi();
  if (!hostApi || typeof hostApi.readJsonFile !== "function") return;
  ratioFileWatchInFlight = true;
  try {
    const path = await resolveCurrentDfmMethodSavePath();
    if (path !== ratioFileWatchPath) {
      clearDfmMethodFileRevision(path);
    }
    const revisionResult = await readDfmMethodFileRevision(hostApi, path);
    if (!revisionResult?.exists) {
      if (ratioFileWatchRevisionToken) clearDfmMethodFileRevision(path);
      return;
    }
    const token = getRevisionToken(revisionResult.revision);
    if (!token) return;
    if (!ratioFileWatchRevisionToken) {
      rememberDfmMethodFileRevision(path, revisionResult.revision);
      return;
    }
    if (token === ratioFileWatchRevisionToken) return;

    if (getDfmIsDirty()) {
      if (ratioFileWatchDirtyWarnToken !== token) {
        ratioFileWatchDirtyWarnToken = token;
        postDfmStatus("DFM method JSON changed on disk. Save or reload before closing to avoid overwriting external changes.", { tone: "warn" });
      }
      return;
    }

    const result = await hostApi.readJsonFile({ path });
    if (!result?.exists) {
      clearDfmMethodFileRevision(path);
      return;
    }
    const changedCells = buildDfmExternalChangedCells(result.data);
    const applied = await applyDfmMethodPayload(result.data, { reason: "external-file-change" });
    if (applied.ok) {
      queueDfmExternalChangeHighlights(changedCells);
      if ((changedCells.ratioCells.length || changedCells.averageCells.length) && isRatiosTabVisible()) {
        renderRatioTable();
      }
      rememberDfmMethodFileRevision(path, result.revision || revisionResult.revision);
      postDfmStatus(`Ready: Reloaded external DFM JSON changes from ${path}`);
    }
  } catch (err) {
    console.warn("DFM method file watch failed:", err);
  } finally {
    ratioFileWatchInFlight = false;
  }
}

export function startDfmMethodFileWatcher() {
  // v2 refreshes are published by ArcRho mutations and reloaded through the
  // aggregate endpoint. Out-of-band file edits require explicit Refresh/Repair.
}

export function stopDfmMethodFileWatcher() {
  if (!ratioFileWatchTimer) return;
  clearInterval(ratioFileWatchTimer);
  ratioFileWatchTimer = null;
}

async function runDfmMethodPreview() {
  if (!getDfmIsDirty()) return { ok: true, skipped: true };
  // A Details change can move the origin basis, which invalidates the Ratio
  // Basis column the preview payload carries. Let it settle first so the
  // preview is not rejected for a column the window is already re-reading.
  await ensureResultsRatioBasisAligned();
  dfmPreviewGeneration += 1;
  const generation = dfmPreviewGeneration;
  dfmPreviewAbortController?.abort?.();
  const controller = new AbortController();
  dfmPreviewAbortController = controller;
  try {
    const response = await previewDfmMethod(buildDfmMethodPayload(), { signal: controller.signal });
    if (generation !== dfmPreviewGeneration || controller.signal.aborted) {
      return { ok: false, aborted: true };
    }
    const method = response?.method;
    if (!method || !isDfmV2Method(method)) {
      throw new Error("DFM preview did not return a canonical v2 method.");
    }
    const applied = await applyDfmMethodPayload(method, {
      reason: "owned-state-preview",
      markClean: false,
    });
    return applied?.ok ? { ok: true, method } : applied;
  } catch (error) {
    if (error?.name === "AbortError") return { ok: false, aborted: true };
    postDfmStatus(`DFM preview failed: ${String(error?.message || error)}`, { tone: "warn" });
    return { ok: false, error: String(error?.message || error) };
  } finally {
    if (dfmPreviewAbortController === controller) dfmPreviewAbortController = null;
  }
}

export function scheduleDfmMethodPreview() {
  cancelDfmExcelFreshnessCheck();
  checkingExcelAppliedRevision = "";
  if (dfmPreviewTimer) clearTimeout(dfmPreviewTimer);
  dfmPreviewTimer = setTimeout(() => {
    dfmPreviewTimer = null;
    void runDfmMethodPreview();
  }, 180);
}

export async function flushDfmMethodPreview() {
  // Dirty-state and owned-state events schedule debounced previews that abort
  // an in-flight run by bumping the generation. An explicit flush must not
  // report a save-blocking failure because of that benign race, so it retries
  // until one run finishes without being superseded.
  for (let attempt = 0; attempt < 5; attempt += 1) {
    if (dfmPreviewTimer) {
      clearTimeout(dfmPreviewTimer);
      dfmPreviewTimer = null;
    }
    const result = await runDfmMethodPreview();
    if (!result?.aborted) return result;
  }
  return { ok: false, error: "DFM preview kept restarting while saving. Try saving again." };
}

export function cancelDfmMethodAsyncTasks() {
  if (dfmPreviewTimer) clearTimeout(dfmPreviewTimer);
  dfmPreviewTimer = null;
  dfmPreviewGeneration += 1;
  dfmPreviewAbortController?.abort?.();
  dfmPreviewAbortController = null;
  cancelDfmExcelFreshnessCheck();
}

// A DFM save is a multi-step round trip -- final recalculation, method write,
// then dependent-propagation queueing -- so the window blocks edits behind the
// shared saving animation until it settles. Overlapping saves (save bar plus
// an Excel bridge save) share one popup through its scope counter.
const dfmSaveProgress = createArcRhoSaveProgress({ subject: "DFM Method" });

export async function saveRatioSelectionPattern(forceSaveAs, options = {}) {
  return dfmSaveProgress.run((progress) => runDfmMethodSave(forceSaveAs, options, progress));
}

async function runDfmMethodSave(forceSaveAs, options, progress) {
  // The Ratio Basis column is read on the DFM's origin basis. When Origin
  // Length changed, wait for the re-read before any payload is built and name
  // the field when it still cannot be aligned.
  const ratioBasis = await ensureResultsRatioBasisAligned();
  if (!ratioBasis.ok) {
    postDfmStatus(ratioBasis.error, { tone: "error" });
    return { ok: false, error: ratioBasis.error };
  }
  const preview = await flushDfmMethodPreview();
  if (preview?.ok === false && !preview?.skipped) return preview;

  const previousDetails = getDfmDetailsTab(lastCleanDfmMethodPayload || {});
  const previousMethodName = String(previousDetails.name || "").trim();
  const previousOutputDataset = String(
    previousDetails["output_dataset"] || currentDfmOutputDataset || previousMethodName,
  ).trim();
  const currentMethodName = getTrimmedInputValue("dfmMethodName");
  const identityChanged = Boolean(
    previousMethodName
      && normalizeDfmIdentityKey(previousMethodName) !== normalizeDfmIdentityKey(currentMethodName),
  );
  if (forceSaveAs && !identityChanged) {
    const message = "Save As requires a new unique Name before saving.";
    postDfmStatus(message, { tone: "warn" });
    return { ok: false, error: message };
  }
  const nextOutputDataset = (forceSaveAs || identityChanged)
    ? currentMethodName
    : (previousOutputDataset || currentMethodName);
  const method = buildDfmMethodPayload({ outputDataset: nextOutputDataset });
  const identity = readDfmMethodIdentityFromPage();
  const saveInput = {
    project_name: identity.project_name,
    reserving_class: identity.reserving_class,
    method,
    notes: getDfmNotesText(),
    ...((forceSaveAs || identityChanged) ? {} : {
      expected_owned_revision: currentOwnedRevision,
      expected_derived_revision: currentDerivedRevision,
    }),
  };
  dfmObjectChangeWatch.pause();
  try {
    postDfmStatus("Saving DFM method...");
    progress.writing();
    const response = await saveDfmMethod(saveInput);
    const canonicalMethod = response?.method;
    if (!canonicalMethod || !isDfmV2Method(canonicalMethod)) {
      throw new Error("DFM save did not return a canonical v2 method.");
    }
    // Record the queued propagation job before the clean transition posts the
    // dependency-source cleared message, so the message carries the job id.
    setPendingDfmPropagationJobId(response?.propagation?.job_id);
    applyDfmAggregateRevisions(response, canonicalMethod);
    const applied = await applyDfmMethodPayload(canonicalMethod, { reason: "save", markClean: true });
    if (!applied?.ok) throw new Error(applied?.error || "Saved DFM could not be applied.");
    const details = getDfmDetailsTab(canonicalMethod);
    currentDfmOutputDataset = String(details["output_dataset"] || nextOutputDataset).trim();
    ensureDfmObjectChangeWatch(details.name, response?.sidecar);
    syncDfmIdentityQuery(canonicalMethod);
    hydrateDfmOutputSidecar(response?.sidecar, {
      hydrateNotes: true,
      outputDataset: currentDfmOutputDataset,
    });
    // A save rewrites the graph on both sides, so the Details rows are stale
    // until they are re-read.
    void refreshDfmDetailsDependencies(currentDfmOutputDataset);
    recordCleanDfmMethodPayload(canonicalMethod);
    markMethodSaved();
    markDfmClean({ force: true });
    emitDfmInstancePresence("found");
    requestProjectInstanceDatasetTableRefresh();
    postDfmStatus(`Method saved at ${new Date().toLocaleTimeString()}.`);
    // Hold the saving card open through the dependent walk so the user sees
    // each live update; a null outcome (failed or stalled walk) keeps the
    // window open and leaves the dataset table as the failure surface.
    const propagationOutcome = await trackSavePropagation(response?.propagation, {
      onStatus: (text, statusOptions) => {
        progress.setMessage?.(text, statusOptions);
        postDfmStatus(text, statusOptions);
      },
      onComplete: () => requestProjectInstanceDatasetTableRefresh(),
    });
    scheduleDfmExcelFreshnessCheck(canonicalMethod);
    // The save and its dependent walk are done; drop the spinner before any
    // follow-up dialog.
    progress.finish();
    if (options.showReviewWarning !== false) {
      await showMethodSaveReviewWarning(response, {
        instanceId: getDfmInst(),
        projectName: getResolvedProjectName(),
        reservingClass: getResolvedReservingClass(),
      });
    }
    return {
      ok: true,
      method: canonicalMethod,
      sidecar: response?.sidecar,
      propagationClean: propagationOutcome !== null,
      refreshedDatasets: propagationOutcome?.refreshed_datasets || [],
      linkWarnings: propagationOutcome?.link_warnings || [],
    };
  } catch (error) {
    progress.finish();
    const message = String(error?.message || error || "DFM save failed.");
    if (isEngineUnavailableSaveError(error)) {
      // The save was refused before anything was written; unsaved work stays
      // in this window.
      void showPageMessageBox({ title: "ArcRho Engine Unavailable", message, tone: "warn" });
    }
    postDfmStatus(`Save failed: ${message}`, { tone: "error" });
    return { ok: false, error: message, status: error?.status };
  } finally {
    dfmObjectChangeWatch.resume();
  }
}

export async function saveDfmTemplate() {
  const hostApi = getHostApi();
  if (!hostApi || typeof hostApi.saveJsonFile !== "function") {
    alert("Save requires the desktop app.");
    window.parent.postMessage({ type: "arcrho:status", text: "Save failed: desktop app required." }, "*");
    return;
  }

  const avgSelection = buildAverageSelectionPayload();
  const cfgKey = getSummaryConfigKey();
  const summaryRows = getSummaryRowsForPersistence(cfgKey);

  const data = {
    "payload_format": "arcrho-dfm-owned-patch-v4",
    "details_tab": {
      "origin_length": readSelectedLengthNumber("originLenSelect"),
      "development_length": readSelectedLengthNumber("devLenSelect"),
    },
    "ratios_tab": {
      "average_formulas": buildDfmAverageFormulaObject(summaryRows, avgSelection.matrix),
    },
  };

  const project = sanitizeFileNamePart(getRatioSaveProjectName(), "UnknownProject");
  const rc = sanitizeFileNamePart(getResolvedReservingClass() || "ReservingClass", "ReservingClass");
  const suggestedName = `DFM_Template@${project}@${rc}.arc-dfm`;

  let startDir = "";
  try {
    const dirRes = await fetch("/template/default_dir");
    if (dirRes.ok) {
      const dirData = await dirRes.json();
      startDir = dirData.path || "";
    }
  } catch {}

  const result = await hostApi.saveJsonFile({
    data,
    suggestedName,
    startDir,
    filters: [{ name: "DFM Template", extensions: ["arc-dfm"] }],
  });

  if (result && result.path) {
    const time = new Date().toLocaleTimeString();
    window.parent.postMessage({ type: "arcrho:status", text: `Template saved at ${time}: ${result.path}` }, "*");
  } else if (result && result.error) {
    window.parent.postMessage({ type: "arcrho:status", text: `Template save failed: ${result.error}` }, "*");
  } else {
    window.parent.postMessage({ type: "arcrho:status", text: "Template save canceled." }, "*");
  }
}
