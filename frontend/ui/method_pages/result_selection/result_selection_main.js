import { fetchProjectDatasetTypeItems } from "/ui/shared/dataset/dataset_types_source.js";
import {
  ensureDatasetOriginLabels,
  formatDatasetOriginLabel,
  getDatasetOriginLabelText,
  validateDatasetOriginLabels,
} from "/ui/shared/dataset/dataset_origin_labels.js";
import { openDatasetNamePicker } from "/ui/shared/components/pickers/dataset_name_picker.js";
import { sanitizeDataFolderPart, sanitizeFileNamePart } from "/ui/shared/utils/filename.js";
import {
  applyTabbedPageSaveBar,
  createTabbedPage,
  requestTabbedPageWindowClose,
  updateTabbedPageSaveControls,
} from "/ui/shared/tabbed_page/tabbed_page.js?v=20260816a";
import { wireTabPopoutWindows } from "/ui/shared/tabbed_page/tab_popout_window.js?v=20260722a";
import { mountNotesTab } from "/ui/shared/tabs/notes/notes_tab.js?v=20260714a";
import { syncDetailsLabelWidth } from "/ui/shared/tabs/details/details_form_layout.js?v=20260820b";
import { createDetailsDependenciesController } from "/ui/shared/tabs/details/details_dependencies.js?v=20260820b";
import { createPageCloseConfirm } from "/ui/shared/components/close_confirm/close_confirm.js";
import { showMethodSaveReviewWarning } from "/ui/shared/components/message_box/method_save_review_warning.js?v=20260827a";
import { showPageMessageBox } from "/ui/shared/components/message_box/message_box.js?v=20260831a";
import { createArcRhoSaveProgress, showSavedDependentsNotice } from "/ui/shared/components/progress_popup/save_progress.js?v=20260831a";
import {
  isEngineUnavailableSaveError,
  trackSavePropagation,
} from "/ui/shared/services/dependent_propagation_job.js?v=20260813e";
import {
  createMethodObjectChangeWatchController,
  showObjectUpdatedAlert,
  wireSamePropagationScopePause,
} from "/ui/shared/services/object_change_watch.js?v=20260820a";
import { createSpreadsheetTableController } from "/ui/shared/components/spreadsheet/spreadsheet_table.js?v=20260712c";
import { createAuditLogView } from "/ui/shared/tabs/audit_log/audit_log_view.js?v=20260714c";
import {
  formatSidecarAuditEventDate,
  normalizeSidecarAuditEntries,
} from "/ui/shared/tabs/audit_log/sidecar_audit_entries.js?v=20260714c";
import {
  buildResultSelectionChartSeries,
  createResultSelectionChart,
} from "/ui/method_pages/result_selection/result_selection_chart.js?v=20260724a";
import { readProjectInstanceDatasetSnapshot } from "/ui/shared/dataset/project_instance_dataset_snapshot.js?v=20260725a";
import {
  resultSelectionUpdateContexts,
  resultSelectionUpdateNames,
} from "/ui/shared/dataset/result_selection_update_report.js?v=20260725b";
import {
  buildResultSelectionMethodPayload,
  normalizeRatioBasisValueSets,
  ratioBasisValuesForName,
  upsertRatioBasisValueSet,
} from "/ui/method_pages/result_selection/result_selection_json_contract.js?v=20260725a";
import {
  RESULT_SELECTION_TAB_DEFS,
  windowTabIds,
} from "/ui/shared/tabs/window_tab_catalog.js?v=20260903a";

const MAX_RATIO_BASIS_COUNT = 3;
const DEFAULT_ORIGIN_LENGTH = 12;
const VALID_ORIGIN_LENGTHS = [12, 6, 3, 1];
const RS_TAB_DEFS = RESULT_SELECTION_TAB_DEFS;
const ALLOWED_RS_TABS = windowTabIds("result_selection");
const FALLBACK_ORIGIN_LABEL_COUNTS = {
  12: 10,
  6: 20,
  3: 40,
  1: 120,
};
const METHOD_COL_DEFAULT_WIDTHS = {
  origin: 90,
  source: 90,
  weight: 68,
  ultimate: 100,
  reserve: 100,
  ratio: 100,
  spacer: 14,
};
const METHOD_COL_MIN_WIDTHS = {
  origin: 58,
  source: 52,
  weight: 68,
  ultimate: 70,
  reserve: 70,
  ratio: 70,
  spacer: 14,
};
const METHOD_COL_MAX_WIDTH = 320;

function text(value) {
  return String(value ?? "").trim();
}

function norm(value) {
  return text(value).replace(/\s+/g, " ").toLowerCase();
}

const params = new URLSearchParams(window.location.search);
const inst = params.get("inst") || `rs_${Date.now()}`;
let programmatic = false;
let isDirty = false;
let cleanSnapshot = "";
let datasetTypeItems = [];
let cachedRows = [];
const closeConfirm = createPageCloseConfirm({ subject: "Result Selection" });
let rsTabSystem = null;

const state = {
  project: text(params.get("project")),
  reservingClass: text(params.get("class") || params.get("path")),
  outputCategory: text(params.get("category")),
  sources: [],
  ratioBasisValues: [],
  ratioBasisValueSets: [],
  outputValues: [],
  activeRatioBasisName: "",
  originLabels: [],
  originLabelsKey: "",
  sidecarOriginLength: null,
  sidecarOriginLabels: [],
  methodColumnWidths: {},
  showEffectiveWeights: false,
  methodHighlight: null,
  methodHighlights: [],
  methodHighlightDragging: false,
  ratioBasisDragIndex: null,
  resultsHighlight: null,
  resultsHighlightDragging: false,
  weightEditSession: null,
  sourceReloadSeq: 0,
  datasetCatalogLoaded: false,
  datasetCatalogPromise: null,
  methodRevision: "",
  loadBlocked: false,
  needsReview: false,
  dependencyPreviewPublished: false,
  hasDependencyPreview: false,
  dependencyPreviews: new Map(),
  persistedRefreshSeq: 0,
  persistedRefreshTimer: null,
  persistedRefreshReason: "",
  dependencyRefreshPromise: null,
  initialLoadPending: true,
  pendingDependencyClearMessages: [],
  persistedDependencyNames: new Set(),
  dependencyRestorePending: false,
  dependencyRestoreError: "",
  dependencyEventSeq: 0,
  persistedMutationSeq: 0,
  persistedMutationInFlight: 0,
  ultimateOverrides: [],
  activeTab: ALLOWED_RS_TABS.has(norm(params.get("tab"))) ? norm(params.get("tab")) : "details",
};

const rsNotesController = mountNotesTab({
  container: document.getElementById("rsNotesMount"),
  ariaLabel: "Result Selection notes",
  onChange: () => ctx.markDirty(),
  onStatus: (message) => ctx.postStatus(message),
});

const els = {
  tabBar: document.getElementById("rsTabBar"),
  nameInput: document.getElementById("rsNameInput"),
  outputTypeInput: document.getElementById("rsOutputTypeInput"),
  outputTypeBtn: document.getElementById("rsOutputTypeBtn"),
  originLengthInput: document.getElementById("rsOriginLengthInput"),
  originLengthDropdown: document.getElementById("rsOriginLengthDropdown"),
  originLengthButton: document.getElementById("rsOriginLengthButton"),
  originLengthLabel: document.getElementById("rsOriginLengthLabel"),
  originLengthMenu: document.getElementById("rsOriginLengthMenu"),
  ratioBasisInputs: [
    document.getElementById("rsRatioBasisInput"),
    document.getElementById("rsRatioBasisInput2"),
    document.getElementById("rsRatioBasisInput3"),
  ],
  ratioBasisList: document.getElementById("rsRatioBasisList"),
  ratioBasisAddButton: document.getElementById("rsRatioBasisAddBtn"),
  ratioBasisPicker: document.querySelector(".rsRatioBasisPicker"),
  ratioBasisContextMenu: document.getElementById("rsRatioBasisContextMenu"),
  showRatiosPctInput: document.getElementById("rsShowRatiosPctInput"),
  statisticDecimalsInput: document.getElementById("rsStatisticDecimalsInput"),
  methodStatisticDecimalsInput: document.getElementById("rsMethodStatisticDecimalsInput"),
  methodStatisticDecimalsUp: document.getElementById("rsMethodStatisticDecimalsUp"),
  methodStatisticDecimalsDown: document.getElementById("rsMethodStatisticDecimalsDown"),
  showWeightsInput: document.getElementById("rsShowWeightsInput"),
  weightDisplayDropdown: document.getElementById("rsWeightDisplayDropdown"),
  weightDisplayButton: document.getElementById("rsWeightDisplayButton"),
  weightDisplayLabel: document.getElementById("rsWeightDisplayLabel"),
  weightDisplayMenu: document.getElementById("rsWeightDisplayMenu"),
  activeRatioBasisDropdown: document.getElementById("rsActiveRatioBasisDropdown"),
  activeRatioBasisButton: document.getElementById("rsActiveRatioBasisButton"),
  activeRatioBasisLabel: document.getElementById("rsActiveRatioBasisLabel"),
  activeRatioBasisMenu: document.getElementById("rsActiveRatioBasisMenu"),
  methodGrid: document.getElementById("rsMethodGrid"),
  chartCanvas: document.getElementById("rsChartCanvas"),
  chartLegendList: document.getElementById("rsChartLegendList"),
  chartLegendCount: document.getElementById("rsChartLegendCount"),
  chartEmpty: document.getElementById("rsChartEmpty"),
  chartTooltip: document.getElementById("rsChartTooltip"),
  resultsGrid: document.getElementById("rsResultsGrid"),
  auditLogMount: document.getElementById("rsAuditLogMount"),
  saveBar: document.querySelector(".rsSaveBar"),
  saveBtn: document.getElementById("rsSaveBtn"),
  cancelBtn: document.getElementById("rsCancelBtn"),
  notesInput: rsNotesController.elements.input,
  cellContextMenu: document.getElementById("rsCellContextMenu"),
  sourceContextMenu: document.getElementById("rsSourceContextMenu"),
};

const auditLogView = createAuditLogView({
  container: els.auditLogMount,
  ariaLabel: "Result Selection audit log",
  emptyDescription: "Result Selection saves will appear here after the first save.",
  normalizeEntries: normalizeSidecarAuditEntries,
  formatEventDate: formatSidecarAuditEventDate,
});

// Open-window change alert (advisory): fires once when another user or the
// dependent-propagation job rewrites this method while it is open.
const rsObjectChangeWatch = createMethodObjectChangeWatchController({
  methodType: "result_selection",
  onChange: (attribution) => {
    void showObjectUpdatedAlert({
      showMessageBox: showPageMessageBox,
      attribution,
      isDirty: () => isDirty,
      onBlockedRefresh: () => {
        postStatus("Unsaved changes block the refresh. Save or discard them, then reopen the window.", "warn");
      },
    });
  },
});
wireSamePropagationScopePause({
  watch: rsObjectChangeWatch,
  getProject: () => state.project,
  getReservingClass: () => state.reservingClass,
});

// A Result Selection save is a multi-step round trip -- origin-label refresh,
// method write, then dependent-propagation queueing -- so the window blocks
// edits behind the shared saving animation until it settles. Overlapping saves
// (save bar plus a Sync dialog save) share one popup through its scope counter.
const rsSaveProgress = createArcRhoSaveProgress({ subject: "Result Selection" });

const ctx = {
  fetchProjectDatasetTypeItems,
  ensureDatasetOriginLabels,
  formatDatasetOriginLabel,
  getDatasetOriginLabelText,
  validateDatasetOriginLabels,
  openDatasetNamePicker,
  sanitizeDataFolderPart,
  sanitizeFileNamePart,
  applyTabbedPageSaveBar,
  createTabbedPage,
  requestTabbedPageWindowClose,
  updateTabbedPageSaveControls,
  wireTabPopoutWindows,
  notesController: rsNotesController,
  createSpreadsheetTableController,
  showMethodSaveReviewWarning,
  showPageMessageBox,
  isEngineUnavailableSaveError,
  trackSavePropagation,
  rsObjectChangeWatch,
  rsSaveProgress,
  showSavedDependentsNotice,
  readProjectInstanceDatasetSnapshot,
  resultSelectionUpdateContexts,
  resultSelectionUpdateNames,
  buildResultSelectionMethodPayload,
  normalizeRatioBasisValueSets,
  ratioBasisValuesForName,
  upsertRatioBasisValueSet,
  MAX_RATIO_BASIS_COUNT,
  DEFAULT_ORIGIN_LENGTH,
  VALID_ORIGIN_LENGTHS,
  RS_TAB_DEFS,
  ALLOWED_RS_TABS,
  FALLBACK_ORIGIN_LABEL_COUNTS,
  METHOD_COL_DEFAULT_WIDTHS,
  METHOD_COL_MIN_WIDTHS,
  METHOD_COL_MAX_WIDTH,
  params,
  inst,
  state,
  els,
  auditLogView,
  text,
  norm,
  closeConfirm,
};

Object.defineProperties(ctx, {
  programmatic: {
    get: () => programmatic,
    set: (value) => { programmatic = value; },
  },
  isDirty: {
    get: () => isDirty,
    set: (value) => { isDirty = value; },
  },
  cleanSnapshot: {
    get: () => cleanSnapshot,
    set: (value) => { cleanSnapshot = value; },
  },
  datasetTypeItems: {
    get: () => datasetTypeItems,
    set: (value) => { datasetTypeItems = value; },
  },
  cachedRows: {
    get: () => cachedRows,
    set: (value) => { cachedRows = value; },
  },
  rsTabSystem: {
    get: () => rsTabSystem,
    set: (value) => { rsTabSystem = value; },
  },
});

function installResultSelectionPart(name, installer) {
  if (typeof installer !== "function") {
    throw new Error(`Result Selection ${name} module failed to load.`);
  }
  Object.assign(ctx, installer(ctx));
}

const resultSelectionParts = window.ResultSelectionParts || {};
installResultSelectionPart("ui", resultSelectionParts.installUi);
installResultSelectionPart("data", resultSelectionParts.installData);
installResultSelectionPart("grids", resultSelectionParts.installGrids);
installResultSelectionPart("model", resultSelectionParts.installModel);

const rsChart = createResultSelectionChart({
  canvas: els.chartCanvas,
  legendList: els.chartLegendList,
  legendCount: els.chartLegendCount,
  emptyState: els.chartEmpty,
  tooltip: els.chartTooltip,
});

ctx.renderResultSelectionChart = function renderResultSelectionChart() {
  if (!rsChart) return;
  const rowCount = ctx.getRowCount();
  const sourceIndexes = ctx.orderedSourceEntries().map((entry) => entry.index);
  rsChart.render({
    originLabels: Array.from({ length: rowCount }, (_, index) => ctx.originLabel(index)),
    series: buildResultSelectionChartSeries({
      sources: state.sources,
      sourceIndexes,
      selectedUltimateValues: ctx.selectedUltimateVector(),
      selectedUltimateLabel: "Selected Ultimate",
      rowCount,
    }),
    decimalPlaces: ctx.getDetails().statisticDecimalPlaces,
  });
};
ctx.refreshResultSelectionChart = () => rsChart?.refresh();

syncDetailsLabelWidth({
  root: "#rsDetailsPage",
  labelSelector: ".arDetailsLabel",
});

// The UI layer runs inside a `with (ctx)` block and cannot import, so the shared
// Details dependency controller is built here and reached through `ctx`.
const detailsDependencies = createDetailsDependenciesController({
  precedentsList: "rsPrecedentsList",
  dependentsList: "rsDependentsList",
  // The graph is keyed by the dataset this method publishes, not the method.
  getIdentity: () => ({
    projectName: state.project,
    reservingClass: state.reservingClass,
    datasetName: ctx.getDetails().name,
  }),
  instanceId: inst,
  isProjectInstanceHost: window.parent !== window,
  setStatus: (message) => ctx.postStatus(message),
});
ctx.refreshDetailsDependencies = () => detailsDependencies.refresh().catch(() => null);

// The window is held blank until the opening tab is rendered; see
// ui/shared/tabbed_page/initial_tab_paint.js.
ctx.init().catch((err) => {
  console.error("Result Selection initialization failed:", err);
  ctx.postStatus(`Result Selection failed to initialize: ${err?.message || err}`, "error");
}).finally(() => window.arcrhoRevealPage?.());
