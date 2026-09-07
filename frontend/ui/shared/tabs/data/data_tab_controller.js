// Shared Dataset Data-tab controller used by Dataset Viewer and DFM hosts.

import { state } from "/ui/shared/dataset/dataset_state.js";
import { config } from "/ui/shared/dataset/dataset_config.js";
import { getBerquistShermanContract } from "/ui/shared/dataset/berquist_sherman_contract.js";
import { $, logLine } from "/ui/shared/tabs/data/data_tab_dom.js";
import {
  getDataset,
  getDatasetNumberFormatDefaults,
  loadCachedDataset,
  loadDatasetSidecar,
  patchDataset,
  previewCalculatedDatasetDependents,
  resolveDatasetInternalLinks,
  saveDatasetNotes,
  saveDatasetSidecar,
} from "/ui/shared/dataset/dataset_api.js";
import {
  renderTable,
  setDatasetRenderNumberFormatSettings,
  setDatasetRenderVectorColumnLabel,
} from "/ui/shared/tabs/data/dataset_grid_view.js?v=20260907c";
import {
  beginDatasetGridLoading,
  endDatasetGridLoading,
  renderDatasetGridPlaceholder,
  setDatasetGridEmpty,
} from "/ui/shared/tabs/data/dataset_grid_placeholder.js?v=20260809a";
import {
  redrawDataTabChartSafely as redrawChartSafely,
  renderDataTabChart as renderChart,
} from "/ui/shared/tabs/data/data_tab_chart_port.js";
import {
  requestTabbedPageWindowClose,
  updateTabbedPageSaveControls,
} from "/ui/shared/tabbed_page/tabbed_page.js?v=20260714a";
import { createDatasetDependencyGuard } from "/ui/shared/dataset/dataset_dependency_service.js";
import { createDatasetHeadersService } from "/ui/shared/dataset/dataset_headers_service.js";
import { validateDatasetOriginLabels } from "/ui/shared/dataset/dataset_origin_labels.js";
import { wireDatasetGridInteractions } from "/ui/shared/tabs/data/dataset_grid_interactions.js?v=20260907c";
import { mountDataTabNotes } from "/ui/shared/tabs/data/data_tab_notes_port.js";
import { publishDataTabHostInputs } from "/ui/shared/tabs/data/data_tab_host_port.js";
import { wireDatasetHostBridge } from "/ui/shared/integrations/dataset_host_bridge.js";
import { createDatasetRunController } from "/ui/shared/dataset/dataset_run_controller.js?v=20260906c";
import { hasResultSelectionUpdates } from "/ui/shared/dataset/result_selection_update_report.js?v=20260725b";
import { wireDatasetInputController } from "/ui/shared/tabs/data/data_tab_controls.js?v=20260907a";
import { readDatasetInputQueryValues } from "/ui/shared/tabs/data/data_tab_query_inputs.js";
import {
  applyDecimalPlacesToDatasetNumberFormat,
  clampDatasetDecimalPlaces,
  normalizeDatasetNumberFormat,
} from "/ui/shared/dataset/dataset_number_format.js";
import {
  isDfmDataTabHost,
  isPersistedDfmMethodBootstrap,
} from "/ui/shared/tabs/data/data_tab_context.js";
import { mountDataTabPageHost } from "/ui/shared/tabs/data/data_tab_page_host_port.js";
import {
  appDefaultWindowTab,
  windowTabIds,
} from "/ui/shared/tabs/window_tab_catalog.js?v=20260903a";
import { openProjectNameTreePicker } from "/ui/shared/components/pickers/project_name_tree_picker.js";
import { openDatasetNamePicker } from "/ui/shared/components/pickers/dataset_name_picker.js";
import { getDataTabAuditController } from "/ui/shared/tabs/data/data_tab_audit_port.js";
import { getDataTabCloseConfirm } from "/ui/shared/tabs/data/data_tab_close_port.js";
import { getDataTabLinksController } from "/ui/shared/tabs/data/data_tab_links_port.js";
import { createDatasetExternalLinksController } from "/ui/shared/dataset/dataset_external_links.js?v=20260907b";
import { createDatasetInternalLinksController } from "/ui/shared/dataset/dataset_internal_links.js?v=20260907b";
import { createDatasetFormulaLinksController } from "/ui/shared/dataset/dataset_formula_links.js?v=20260907b";
import {
  loadProjectUserPreferences,
  scheduleProjectUserPreferencesSave,
} from "/ui/shared/services/project_user_preferences.js";
import {
  loadProjectValidValueList,
  loadDatasetValidValueList,
  loadReservingClassValidValueList,
  clearValidValueListCache,
  validateReservingClassPathByTypeNames,
  buildReservingClassPathPartLookup,
  normalizeReservingClassPathByPartLookup,
  normalizeReservingClassPath,
  normalizeReservingClassPathKey,
} from "/ui/shared/services/valid_value_lists.js";
import {
  getLastViewedDatasetInputs,
  setLastViewedDatasetInputs,
  pushBrowsingHistoryEntry,
  normalizeBrowsingHistoryEntry,
} from "/ui/shell/browsing_history.js";
import "/ui/shared/integrations/zoom_bridge.js?v=20260715a";

import { registerDataTabHostController } from "/ui/shared/tabs/data/data_tab_host_controller.js?v=20260906c";
import { registerDataTabDetailsController } from "/ui/shared/tabs/data/data_tab_details_controller.js?v=20260824b";
import { registerDataTabInputsController } from "/ui/shared/tabs/data/data_tab_inputs_controller.js?v=20260906b";
import { registerDataTabPreferencesController } from "/ui/shared/tabs/data/data_tab_preferences_controller.js?v=20260906c";
import { registerDataTabRequestController } from "/ui/shared/tabs/data/data_tab_request_controller.js?v=20260907b";
import { registerDataTabPersistenceController } from "/ui/shared/tabs/data/data_tab_persistence_controller.js?v=20260907e";

const LS_DS_KEY = "arcrho_last_ds_id";
const LS_FORM_KEY = "arcrho_tri_inputs";
const LOCAL_PROJECT_PREFS_ENDPOINT = "/local-project/preferences";
const WF_GLOBAL_CTRL_PREFIX = "arcrho_workflow_global_ctrl_v1::";
const DEFAULT_PROJECT_DISPLAY = "Default Project";
const DEFAULT_PATH_DISPLAY = "Default Path";
const DEFAULT_TOKEN = "__DEFAULT__";
// Records, not distinct datasets: the store keeps one record per dataset per day.
const BROWSING_HISTORY_MAX_ENTRIES = 100;

const qs = new URLSearchParams(window.location.search);
const instanceId = qs.get("inst") || "default";
const isProjectInstanceHost = qs.get("project_instance") === "1";
const isProjectInstanceDraft = qs.get("draft_instance") === "1" || qs.get("draft") === "1";
const isReadOnlyDatasetViewer = qs.get("readonly") === "1";
const temporaryDatasetSessionId = String(qs.get("temporary_session_id") || "").trim();
const isTemporaryDatasetView = qs.get("temporary_view") === "1" && !!temporaryDatasetSessionId;
const isProjectInstanceCachedDatasetOpen = isProjectInstanceHost
  && !isDfmDataTabHost()
  && !isProjectInstanceDraft
  && !isTemporaryDatasetView;
const stepId = instanceId.startsWith("step_") ? instanceId : null;
const scopedKey = (key) => `${key}::${instanceId}`;
const workflowId = qs.get("wf") || "";

const runtime = {
  state,
  config,
  $,
  logLine,
  getBerquistShermanContract,
  getDataset,
  getDatasetNumberFormatDefaults,
  loadCachedDataset,
  loadDatasetSidecar,
  patchDataset,
  previewCalculatedDatasetDependents,
  resolveDatasetInternalLinks,
  saveDatasetNotes,
  saveDatasetSidecar,
  renderTable,
  setDatasetRenderNumberFormatSettings,
  setDatasetRenderVectorColumnLabel,
  redrawChartSafely,
  renderChart,
  requestTabbedPageWindowClose,
  updateTabbedPageSaveControls,
  createDatasetDependencyGuard,
  createDatasetHeadersService,
  validateDatasetOriginLabels,
  wireDatasetGridInteractions,
  mountDataTabNotes,
  publishDataTabHostInputs,
  wireDatasetHostBridge,
  createDatasetRunController,
  hasResultSelectionUpdates,
  wireDatasetInputController,
  readDatasetInputQueryValues,
  applyDecimalPlacesToDatasetNumberFormat,
  clampDatasetDecimalPlaces,
  normalizeDatasetNumberFormat,
  isDfmDataTabHost,
  isPersistedDfmMethodBootstrap,
  getDataTabAuditController,
  getDataTabCloseConfirm,
  getDataTabLinksController,
  createDatasetExternalLinksController,
  createDatasetInternalLinksController,
  createDatasetFormulaLinksController,
  loadProjectUserPreferences,
  scheduleProjectUserPreferencesSave,
  loadProjectValidValueList,
  loadDatasetValidValueList,
  loadReservingClassValidValueList,
  clearValidValueListCache,
  validateReservingClassPathByTypeNames,
  buildReservingClassPathPartLookup,
  normalizeReservingClassPathByPartLookup,
  normalizeReservingClassPath,
  normalizeReservingClassPathKey,
  getLastViewedDatasetInputs,
  setLastViewedDatasetInputs,
  pushBrowsingHistoryEntry,
  normalizeBrowsingHistoryEntry,
  qs,
  instanceId,
  isProjectInstanceHost,
  isProjectInstanceDraft,
  isReadOnlyDatasetViewer,
  temporaryDatasetSessionId,
  isTemporaryDatasetView,
  isProjectInstanceCachedDatasetOpen,
  stepId,
  scopedKey,
  workflowId,
  LS_DS_KEY,
  LS_FORM_KEY,
  LOCAL_PROJECT_PREFS_ENDPOINT,
  WF_GLOBAL_CTRL_PREFIX,
  DEFAULT_PROJECT_DISPLAY,
  DEFAULT_PATH_DISPLAY,
  DEFAULT_TOKEN,
  BROWSING_HISTORY_MAX_ENTRIES,
  DATASET_VIEWER_TAB_IDS: windowTabIds("dataset"),
  DATASET_VIEWER_APP_DEFAULT_TAB: appDefaultWindowTab("dataset"),
  activeDependencyPreviewKey: "",
  allDatasetTypes: [],
  allProjects: [],
  currentDatasetPrecedents: [],
  currentDatasetSidecarDataFormat: "",
  currentDatasetSidecarSourceKind: "",
  currentDatasetStoredDevelopmentLength: 0,
  currentDatasetStoredOriginLength: 0,
  datasetDependencyGuard: null,
  datasetExternalLinks: null,
  datasetInternalLinks: null,
  datasetFormulaLinks: null,
  datasetHeadersService: null,
  datasetInstanceNameConflict: false,
  datasetInstanceNameConflictMessage: "",
  datasetRunController: null,
  datasetSaveInFlight: false,
  isSidecarReadOnlyDataset: false,
  lastProjectSelection: "",
  savedProjectInstanceDraftName: "",
};

registerDataTabInputsController(runtime);
registerDataTabPreferencesController(runtime);
registerDataTabRequestController(runtime);
registerDataTabDetailsController(runtime);
registerDataTabPersistenceController(runtime);
registerDataTabHostController(runtime);

let datasetGridInteractions = null;
let eventsWired = false;
let bootPromise = null;

export function getDatasetExternalLinkRecords() {
  return runtime.datasetExternalLinks.listRecords();
}

export function getDatasetExternalLinkCellInfo(displayRow, displayColumn) {
  return runtime.datasetExternalLinks.getCellLinkInfo(displayRow, displayColumn);
}

export async function breakDatasetExternalLinks(ids) {
  const result = runtime.datasetExternalLinks.breakLinks(ids);
  if (!result.ok) return result;
  renderTable();
  runtime.notifyDatasetUpdated({ publishPreview: false });
  runtime.setStatus(result.message || "Links broken. Current dataset values are now hard-coded.");
  return result;
}

export async function breakDatasetExternalLink(id) {
  return breakDatasetExternalLinks([id]);
}

export async function refreshDatasetExternalLinkRecords(ids) {
  return runtime.refreshDatasetExternalLinks({ ids });
}

export function getDatasetInternalLinkRecords() {
  return runtime.datasetInternalLinks.listRecords();
}

export async function breakDatasetInternalLinks(ids) {
  const result = runtime.datasetInternalLinks.breakLinks(ids);
  if (!result.ok) return result;
  renderTable();
  runtime.notifyDatasetUpdated({ publishPreview: false });
  runtime.setStatus(result.message || "Links broken. Current dataset values are now hard-coded.");
  return result;
}

export async function refreshDatasetInternalLinkRecords(ids) {
  return runtime.refreshDatasetInternalLinks({ ids });
}

export function getDatasetFormulaLinkRecords() {
  return runtime.datasetFormulaLinks.listRecords();
}

export async function breakDatasetFormulaLinks(ids) {
  const result = runtime.datasetFormulaLinks.breakLinks(ids);
  if (!result.ok) return result;
  renderTable();
  runtime.notifyDatasetUpdated({ publishPreview: false });
  runtime.setStatus(result.message || "Links broken. Current dataset values are now hard-coded.");
  return result;
}

export async function refreshDatasetFormulaLinkRecords(ids) {
  return runtime.refreshDatasetFormulaLinks({ ids });
}

// The Links tab shows every kind of link in one table; these three route a
// mixed selection back to the controller that owns each record.
const LINK_KIND_HANDLERS = {
  excel: {
    refresh: refreshDatasetExternalLinkRecords,
    break: (ids) => runtime.datasetExternalLinks.breakLinks(ids),
  },
  internal: {
    refresh: refreshDatasetInternalLinkRecords,
    break: (ids) => runtime.datasetInternalLinks.breakLinks(ids),
  },
  formula: {
    refresh: refreshDatasetFormulaLinkRecords,
    break: (ids) => runtime.datasetFormulaLinks.breakLinks(ids),
  },
};

export function getDatasetLinkRecords() {
  return [
    ...getDatasetExternalLinkRecords().map((record) => ({ ...record, sourceKind: "excel" })),
    ...getDatasetInternalLinkRecords(),
    ...getDatasetFormulaLinkRecords(),
  ];
}

function groupLinkRecordIdsByKind(records) {
  const groups = new Map();
  (Array.isArray(records) ? records : []).forEach((record) => {
    const kind = LINK_KIND_HANDLERS[record?.sourceKind] ? record.sourceKind : "excel";
    if (!record?.id) return;
    if (!groups.has(kind)) groups.set(kind, []);
    groups.get(kind).push(record.id);
  });
  return groups;
}

export async function refreshDatasetLinkRecords(records) {
  const results = [];
  for (const [kind, ids] of groupLinkRecordIdsByKind(records)) {
    results.push(await LINK_KIND_HANDLERS[kind].refresh(ids));
  }
  const failures = results.flatMap((result) => result?.failures || []);
  return {
    linkedCellCount: results.reduce((sum, result) => sum + (Number(result?.linkedCellCount) || 0), 0),
    changedCount: results.reduce((sum, result) => sum + (Number(result?.changedCount) || 0), 0),
    failedCount: results.reduce((sum, result) => sum + (Number(result?.failedCount) || 0), 0),
    failures,
    error: results.map((result) => result?.error).filter(Boolean).join(" "),
  };
}

export async function breakDatasetLinks(records) {
  let affectedCellCount = 0;
  let brokenLinkCount = 0;
  for (const [kind, ids] of groupLinkRecordIdsByKind(records)) {
    const result = LINK_KIND_HANDLERS[kind].break(ids);
    if (!result.ok) return result;
    affectedCellCount += Number(result.affectedCellCount) || 0;
    brokenLinkCount += ids.length;
  }
  if (!brokenLinkCount) return { ok: false, error: "No links were selected." };
  renderTable();
  runtime.notifyDatasetUpdated({ publishPreview: false });
  const message = `${brokenLinkCount === 1 ? "Link" : `${brokenLinkCount} links`} broken. Current dataset values are now hard-coded.`;
  runtime.setStatus(message);
  return { ok: true, affectedCellCount, message };
}

async function openProjectNameTreeForDataset(targetInput) {
  const initialProject = runtime.getResolvedProjectValue() || targetInput?.value || "";
  await openProjectNameTreePicker({
    initialProject,
    anchorElement: targetInput || null,
    title: "Select a Project",
    setStatus: runtime.setStatus,
    onError: (err) => {
      console.error("Failed to load project tree:", err);
      runtime.setStatus("Error loading project tree.");
    },
    onSelect: async (projectName) => {
      const selected = String(projectName || "").trim();
      if (!selected || !targetInput) return;
      runtime.setInputDefaultBound(targetInput, false);
      targetInput.value = selected;
      runtime.showProjectDropdown(false);
      runtime.setStatus("Loading dataset...");
      await runtime.handleProjectSelection(selected, { strict: true, showMessage: true });
    },
  });
}

async function openDatasetNameTreeForDataset(targetInput) {
  await openDatasetNamePicker({
    projectName: runtime.getResolvedProjectValue(),
    initialName: targetInput?.value || "",
    anchorElement: targetInput || null,
    title: "Select a Dataset Type",
    setStatus: runtime.setStatus,
    onError: (err) => {
      console.error("Failed to load dataset type tree:", err);
      runtime.setStatus("Error loading dataset types.");
    },
    onSelect: (datasetName) => {
      const selected = String(datasetName || "").trim();
      if (!selected || !targetInput) return;
      targetInput.value = selected;
      runtime.showDatasetDropdown(false);
      const knownName = runtime.ensureDatasetTypeOption(selected) || selected;
      void runtime.handleDatasetSelection(knownName, { strict: true });
    },
  });
}

// The sentence a linked cell shows in place of its formula while the window is
// off the lengths its links were read at, or "" whenever there is nothing to
// explain — the display is on those lengths, or this dataset has no links.
function offLinkedShapeLinkNote() {
  if (runtime.datasetDisplayIsAtLinkedShape()) return "";
  const hasLinks = linkControllersHaveLinks();
  return hasLinks ? runtime.datasetOffLinkedShapeLinkHint() : "";
}

function linkControllersHaveLinks() {
  return !!(
    runtime.datasetExternalLinks?.hasLinks()
    || runtime.datasetInternalLinks?.hasLinks()
    || runtime.datasetFormulaLinks?.hasLinks()
  );
}

function wireGridInteractions() {
  if (datasetGridInteractions) return;
  const dfmLinksRefused = () => Promise.resolve({
    handled: true,
    ok: false,
    error: "Enter external Excel links in DFM Ratios User Entry cells.",
  });
  const mapLinkCells = (cells) => (Array.isArray(cells) ? cells : []).map((cell) => ({
    row: Number(cell?.row ?? cell?.r),
    column: Number(cell?.column ?? cell?.c),
  }));
  const linkControllers = () => [
    runtime.datasetExternalLinks,
    runtime.datasetInternalLinks,
    runtime.datasetFormulaLinks,
  ];
  datasetGridInteractions = wireDatasetGridInteractions({
    state,
    renderTable,
    isReadOnly: runtime.isDatasetReadOnly,
    readOnlyMessage: runtime.getDatasetReadOnlyMessage,
    setStatus: runtime.setStatus,
    notifyDatasetUpdated: runtime.notifyDatasetUpdated,
    refreshDatasetSettingsDirty: runtime.refreshDatasetSettingsDirty,
    commitExternalReference: (request) => (
      isDfmDataTabHost()
        ? dfmLinksRefused()
        : runtime.datasetExternalLinks.commitReference(request)
    ),
    commitInternalReference: (request) => (
      isDfmDataTabHost()
        ? dfmLinksRefused()
        : runtime.datasetInternalLinks.commitReference(request)
    ),
    commitFormulaReference: (request) => (
      isDfmDataTabHost()
        ? dfmLinksRefused()
        : runtime.datasetFormulaLinks.commitReference(request)
    ),
    cancelExternalReference: () => linkControllers().forEach((controller) => controller.abort()),
    hardCodeExternalLinkCells: (cells) => {
      const mapped = mapLinkCells(cells);
      return linkControllers().reduce(
        (count, controller) => count + controller.hardCodeTargetCells(mapped),
        0,
      );
    },
    decorateExternalLinkCell: (cell, displayRow, displayColumn) => {
      linkControllers().forEach((controller) => controller.decorateCell(cell, displayRow, displayColumn));
    },
    // A cell holds at most one link (Excel, ArcRho, or formula), enforced on
    // commit and save, so the first answer wins here. While the window shows
    // the dataset at other lengths than its links were read at, no cell on
    // screen is one a link names, so every cell of a linked dataset answers
    // with the note that says which length to put back instead.
    getExternalLinkCellInfo: (displayRow, displayColumn) => {
      const note = offLinkedShapeLinkNote();
      if (note) return { note, anchorDisplayRow: displayRow, anchorDisplayColumn: displayColumn };
      return runtime.datasetFormulaLinks.getCellLinkInfo(displayRow, displayColumn)
        || runtime.datasetInternalLinks.getCellLinkInfo(displayRow, displayColumn)
        || runtime.datasetExternalLinks.getCellLinkInfo(displayRow, displayColumn);
    },
    beginReferencePick: () => runtime.publishDatasetReferencePickBegin?.(),
    endReferencePick: () => runtime.publishDatasetReferencePickEnd?.(),
    publishReferencePick: (range) => runtime.publishDatasetReferencePick?.(range),
  });
  runtime.applyDatasetReferencePick = datasetGridInteractions.applyDatasetReferencePick;
}

function applyGridSelectionFromState() {
  datasetGridInteractions?.applySelectionFromState?.();
}

Object.assign(runtime, {
  openProjectNameTreeForDataset,
  openDatasetNameTreeForDataset,
  wireGridInteractions,
  applyGridSelectionFromState,
});

function wireEvents() {
  if (eventsWired) return;
  eventsWired = true;
  wireDatasetInputController({
    ...runtime,
    state,
    $,
    openProjectNameTreeForDataset,
    openDatasetNameTreeForDataset,
    wireDatasetHostBridge,
    wireGridInteractions,
  });
  runtime.wireDatasetInstanceNameInput();
  runtime.wireDatasetSaveControls();
}

async function bootDatasetDataTabOnce() {
  // Boot owns the grid placeholder until a load, a run, or an explicit empty
  // state takes over, so the first paint of a Client PC window shows the grid
  // that is on its way rather than an empty-looking one.
  const gridPlaceholderToken = beginDatasetGridLoading();
  const gridHost = document.getElementById("tableWrap");
  // The grid host is mounted before boot runs, so the skeleton is the window's
  // first paint of that area instead of a blank panel or a stale empty state.
  if (gridHost && !state.model) renderDatasetGridPlaceholder(gridHost);
  try {
    await bootDatasetDataTabSteps();
  } finally {
    endDatasetGridLoading(gridPlaceholderToken);
  }
}

async function bootDatasetDataTabSteps() {
  runtime.wireDataTabHostLifecycle();
  runtime.wireDataTabInputLifecycle();
  runtime.wireDataTabPersistenceLifecycle();
  runtime.initializeDatasetId();
  setDatasetRenderVectorColumnLabel(isProjectInstanceHost ? qs.get("vector_column_label") : "");
  runtime.wireNotesEditor();
  runtime.fillLenDropdowns();

  const persistedDfmBootstrap = isPersistedDfmMethodBootstrap();
  try {
    if (persistedDfmBootstrap || isProjectInstanceCachedDatasetOpen || isTemporaryDatasetView) {
      // Temporary views carry complete inputs in the URL and cannot save, so
      // they skip the dropdown/preference/sidecar boot chain like cached
      // Project Instance opens; the authoritative run validation reloads the
      // lists it needs on demand.
      runtime.applyTriInputsFromQueryParams();
      // Skipping that chain also skips the only step that resolves a number
      // format, and the grid formats from the toolbar controls when the run
      // paints it. Resolve the Dataset Type default first so the first paint is
      // already formatted instead of waiting for the next input change.
      if (isTemporaryDatasetView) await runtime.applyTemporaryNumberFormatDefaults();
    } else {
      await runtime.loadProjectsDropdown();
      runtime.applyWorkflowDefaultsIfNew();
      await runtime.restoreTriInputsFromStorage();
      runtime.applyTriInputsFromQueryParams();
      const projectResult = runtime.validateAndNormalizeProjectInput({ strict: true, showMessage: false });
      if (projectResult.ok) {
        runtime.lastProjectSelection = projectResult.value;
        if (!isDfmDataTabHost()) runtime.saveLastDatasetViewerProjectToAppData(projectResult.value);
        await Promise.all([
          runtime.refreshDatasetTypesForProject(projectResult.value),
          runtime.refreshReservingClassPathsForProject(projectResult.value),
        ]);
      } else {
        await Promise.all([
          runtime.refreshDatasetTypesForProject(""),
          runtime.refreshReservingClassPathsForProject(""),
        ]);
      }
      await runtime.validateAndNormalizeReservingClassInput(
        runtime.getResolvedProjectValue(),
        { strict: true, showMessage: false },
      );
      runtime.validateAndNormalizeDatasetInput({ strict: true, showMessage: false });
      await runtime.syncSidecarForCurrentDataset({ applyLengths: !isProjectInstanceDraft });
      await runtime.refreshDatasetInstanceNameConflict();
    }
    runtime.enforceDevLenRule({ source: "origin" });

    mountDataTabPageHost({
      initialTab: runtime.getDatasetInitialTab(),
      onDetailsActivated: () => requestAnimationFrame(runtime.resizeDetailFormulaInput),
      onChartActivated: () => {
        requestAnimationFrame(() => requestAnimationFrame(redrawChartSafely));
      },
      wireDataTabTopBarToggle: runtime.wireDatasetDataTabTopBarToggle,
    });

    wireEvents();

    const { project, path, tri } = runtime.getTriInputs();
    if (persistedDfmBootstrap) {
      runtime.setStatus("Loading DFM method...");
    } else if (project && path && tri) {
      // The loading popup is reserved for clear-cache rebuilds, which show it
      // immediately, and for runs still pending after the run controller's
      // short delay (the cache-miss engine path). Cached opens render the
      // grid without a spinner; the status line still reports progress.
      if (isProjectInstanceCachedDatasetOpen) {
        await runtime.loadProjectInstanceCachedDataset();
      } else if (isProjectInstanceDraft) {
        await runtime.refreshProjectInstanceDraftModel();
      } else {
        runtime.scheduleAutoRun(0);
      }
    } else if (isDfmDataTabHost()) {
      setDatasetGridEmpty({
        title: "Waiting For DFM Inputs",
        hint: "This table fills in once the method has a project, reserving class, and dataset.",
      });
      runtime.setStatus("Waiting for DFM inputs...");
    } else {
      await runtime.loadDataset();
    }
  } catch (err) {
    runtime.hideDatasetLoadingPopup();
    throw err;
  }
}

export function bootDatasetDataTab() {
  if (!bootPromise) bootPromise = bootDatasetDataTabOnce();
  return bootPromise;
}
