import { ensureDatasetOriginLabels } from "/ui/shared/dataset/dataset_origin_labels.js";
import { createDatasetHeadersService } from "/ui/shared/dataset/dataset_headers_service.js";
import { openDatasetNamePicker } from "/ui/shared/components/pickers/dataset_name_picker.js";
import { sanitizeFileNamePart } from "/ui/shared/utils/filename.js";
import {
  applyTabbedPageSaveBar,
  createTabbedPage,
  requestTabbedPageWindowClose,
  updateTabbedPageSaveControls,
} from "/ui/shared/tabbed_page/tabbed_page.js?v=20260816a";
import { mountNotesTab } from "/ui/shared/tabs/notes/notes_tab.js?v=20260714a";
import { syncDetailsLabelWidth, syncDetailsSections } from "/ui/shared/tabs/details/details_form_layout.js?v=20260820b";
import { createDetailsDependenciesController } from "/ui/shared/tabs/details/details_dependencies.js?v=20260820b";
import { applyHostFixedDetailsFields } from "/ui/shared/tabs/details/details_host_fields.js?v=20260820b";
import { createAuditLogView } from "/ui/shared/tabs/audit_log/audit_log_view.js?v=20260714c";
import {
  formatSidecarAuditEventDate,
  normalizeSidecarAuditEntries,
} from "/ui/shared/tabs/audit_log/sidecar_audit_entries.js?v=20260714c";
import { createPageCloseConfirm } from "/ui/shared/components/close_confirm/close_confirm.js";
import { wireNumberFormatField } from "/ui/shared/components/pickers/number_format_field.js?v=20260817a";
import { showMethodSaveReviewWarning } from "/ui/shared/components/message_box/method_save_review_warning.js?v=20260827a";
import { createArcRhoSaveProgress, showSavedDependentsNotice } from "/ui/shared/components/progress_popup/save_progress.js?v=20260831a";
import { trackSavePropagation } from "/ui/shared/services/dependent_propagation_job.js?v=20260813e";
import {
  getBerquistShermanContract,
  normalizeBerquistShermanVariant,
} from "/ui/shared/dataset/berquist_sherman_contract.js";
import {
  DEFAULT_DATASET_NUMBER_FORMAT,
  applyDecimalPlacesToDatasetNumberFormat,
  clampDatasetDecimalPlaces,
  formatDatasetNumberValue,
  getDatasetNumberFormatDecimalPlaces,
  normalizeDatasetNumberFormat,
} from "/ui/shared/dataset/dataset_number_format.js";
import {
  DEFAULT_LOESS_SPAN,
  normalizeAnnualTriangle,
  normalizeLoessSpan,
} from "./calculation_helpers.js";
import { calculateSettlementRate } from "./settlement_rate_calculation.js";
import { calculateCaseReserveAdequacy } from "./case_reserve_adequacy_calculation.js";
import { readProjectInstanceDatasetSnapshot } from "/ui/shared/dataset/project_instance_dataset_snapshot.js?v=20260725a";
import {
  BERQUIST_SHERMAN_TAB_DEFS,
  windowTabIds,
} from "/ui/shared/tabs/window_tab_catalog.js?v=20260903a";
import {
  readExcelCellsBatch,
  validateExcelLinksBatch,
} from "/ui/shared/integrations/excel_api.js?v=20260819a";
import { showExcelLinkFailureAlert } from "/ui/shared/integrations/excel_link_alert.js?v=20260819a";
import { normalizeExcelReferenceAddressCase } from "/ui/shared/integrations/excel_reference.js?v=20260715a";
import {
  buildUserValueLinkRecords,
  collectUserValueExcelLinkGroups,
  evaluateUserValueFormula,
  excelReadItemsForInputs,
  excelSourceKey,
  expandUserValueRangeEntry,
  isUserValueFormula,
  normalizeUserValueInputs,
  parseUserValueNumber,
  userValueExcelReferences,
} from "./berquist_sherman_user_values.js";
import { createBerquistShermanLinksTab } from "./berquist_sherman_links_tab.js";
import { createBerquistShermanCellSelection } from "./berquist_sherman_cell_selection.js";

const ANNUAL_PERIOD_LENGTH = 12;
const TABS = BERQUIST_SHERMAN_TAB_DEFS;
const ALLOWED_TABS = windowTabIds("berquist_sherman");
const params = new URLSearchParams(window.location.search || "");
const variant = normalizeBerquistShermanVariant(params.get("variant")) || "sr";
const contract = getBerquistShermanContract(variant);
const inst = text(params.get("inst")) || `bs_${Date.now()}`;

const ROLE_DEFINITIONS = Object.freeze({
  sr: Object.freeze([
    { key: "paid_claims", label: "Paid Claims", format: "Triangle", inputId: "bsSrPaidInput" },
    { key: "closed_claim_numbers", label: "Closed Claim Counts", format: "Triangle", inputId: "bsSrClosedInput" },
    { key: "ultimate_claim_numbers", label: "Ultimate Claim Counts", format: "Vector", inputId: "bsSrUltimateInput" },
  ]),
  // ResQ lists the CRA sources as Paid, Incurred, Reported, Closed.
  cra: Object.freeze([
    { key: "paid_claims", label: "Paid Claims", format: "Triangle", inputId: "bsCraPaidInput" },
    { key: "incurred_claims", label: "Incurred Claims", format: "Triangle", inputId: "bsCraIncurredInput" },
    { key: "reported_claim_numbers", label: "Reported Claim Counts", format: "Triangle", inputId: "bsCraReportedInput" },
    { key: "closed_claim_numbers", label: "Closed Claim Counts", format: "Triangle", inputId: "bsCraClosedInput" },
  ]),
});

const VIEW_DEFINITIONS = Object.freeze({
  // Settlement Rate mirrors the ResQ Method sub-tab order and captions. The
  // Pairs, All, and Loess estimates are shown per origin period inside the
  // Adjusted Paid Claims selection grid.
  sr: Object.freeze([
    { key: "paidClaims", label: "Paid Claims", caption: "Paid Claims:" },
    { key: "numbersClosed", label: "Numbers Closed", caption: "Numbers Closed:" },
    { key: "proportionSettled", label: "Proportion Settled", caption: "Proportion Settled:" },
    { key: "selectedClaimNumbers", label: "Selected Numbers Closed", caption: "Selected Numbers Closed:" },
    {
      key: "output",
      label: "Adjusted Paid Claims",
      caption: "Paid Claims Adjusted to Constant Proportions Settled",
    },
  ]),
  // Case Reserve Adequacy mirrors the same ResQ Method sub-tab order and
  // captions. "Avg. Selections" stacks the two ResQ selection grids.
  cra: Object.freeze([
    { key: "reportedClaimNumbers", label: "Reported", caption: "Reported Claim Counts:" },
    { key: "closedClaimNumbers", label: "Closed", caption: "Closed Claim Counts:" },
    { key: "openClaimNumbers", label: "Open", caption: "Number of Open Claims (Reported - Closed):" },
    { key: "paidClaims", label: "Paid Claims", caption: "Paid Claims:" },
    { key: "incurredClaims", label: "Incurred Claims", caption: "Incurred Claims:" },
    { key: "caseReserves", label: "Case Reserves", caption: "Case Reserves (Incurred - Paid):" },
    {
      key: "averageCaseReserves",
      label: "Avg. Case Reserves",
      caption: "Average Case Reserves (Case Reserves / Open):",
    },
    {
      key: "averagePaidClaims",
      label: "Avg. Paid",
      caption: "Average Paid Claims (Incremental Paid / Incremental Closed):",
    },
    {
      key: "avgSelections",
      label: "Avg. Selections",
      caption: "Average Inflation:",
      secondaryCaption: "Current Average Case Reserves:",
    },
    {
      key: "adjustedAverageCaseReserves",
      label: "Adj. Avg. Case Reserves",
      caption: "Adjusted Average Case Reserves:",
    },
    {
      key: "output",
      label: "Adj. Incurred",
      caption: "Adjusted Incurred (Paid + Adjusted Average Case Reserve x Open):",
    },
  ]),
});

// Views that show a source triangle rather than a calculated result.
const SOURCE_VIEW_ROLES = Object.freeze({
  paidClaims: "paid_claims",
  numbersClosed: "closed_claim_numbers",
  incurredClaims: "incurred_claims",
  reportedClaimNumbers: "reported_claim_numbers",
  closedClaimNumbers: "closed_claim_numbers",
});

// ResQ "Average Inflation" rows. The two group captions carry no values, so the
// grid reads as two estimator pairs above the user and selected bands.
const INFLATION_GRID_ROWS = Object.freeze([
  { group: "Case Reserves" },
  { method: "case_column", label: "Column", key: "caseInflationByColumn" },
  { method: "case_all", label: "All", key: "caseInflationOverall", constant: true },
  { group: "Paid" },
  { method: "paid_column", label: "Column", key: "paidInflationByColumn" },
  { method: "paid_all", label: "All", key: "paidInflationOverall", constant: true },
  { method: "user", label: "User Value", user: true },
  { selected: true, label: "Selected", key: "selectedInflation" },
]);

// ResQ "Current Average Case Reserves" rows.
const AVERAGE_GRID_ROWS = Object.freeze([
  { method: "latest", label: "Latest", key: "latestAverageCaseReserves" },
  { method: "monotone", label: "Monotone", key: "monotoneAverageCaseReserves" },
  { method: "loess", label: "Loess", key: "loessAverageCaseReserves", loess: true },
  { method: "user", label: "User Value", user: true },
  { selected: true, label: "Selected", key: "selectedAverageCaseReserves" },
]);

function defaultNumberFormat() {
  return {
    number_format: DEFAULT_DATASET_NUMBER_FORMAT,
    decimal_places: getDatasetNumberFormatDecimalPlaces(DEFAULT_DATASET_NUMBER_FORMAT),
  };
}

// One shape for every recorded format, so the method JSON, the output sidecar,
// and the Details controls cannot disagree about how a format is written down.
function normalizeNumberFormatEntry(value, fallback = defaultNumberFormat()) {
  const source = value && typeof value === "object" ? value : {};
  const raw = source.number_format ?? source.numberFormat;
  const pattern = normalizeDatasetNumberFormat(raw, fallback.number_format);
  const places = clampDatasetDecimalPlaces(
    source.decimal_places ?? source.decimalPlaces,
    getDatasetNumberFormatDecimalPlaces(pattern),
  );
  return {
    number_format: applyDecimalPlacesToDatasetNumberFormat(pattern, places),
    decimal_places: places,
  };
}

function sameNumberFormat(left, right) {
  return left?.number_format === right?.number_format
    && left?.decimal_places === right?.decimal_places;
}

const state = {
  project: text(params.get("project")),
  reservingClass: text(params.get("class") || params.get("path")),
  cachedRows: [],
  sourceNames: {},
  sourceValues: {},
  sourcePayloads: {},
  originLabels: [],
  developmentLabels: [],
  sidecarOriginLabels: [],
  // `derived` is the Details-tab format for every calculated triangle and is
  // also the output dataset's own format. `sources` mirrors each input
  // dataset instance's format so the input views display exactly as the
  // Dataset Viewer does; it is a recorded copy, refreshed from the live
  // sidecars whenever the sources load.
  numberFormats: { derived: defaultNumberFormat(), sources: {} },
  result: null,
  currentView: VIEW_DEFINITIONS[variant][0].key,
  selectedProportionSettled: [],
  selectedProportionIsDefault: [],
  selectedAdjustment: [],
  loessSpan: DEFAULT_LOESS_SPAN,
  inflationSelection: [],
  userInflation: [],
  averageCaseReserveSelection: [],
  userAverageCaseReserves: [],
  // The text behind each User Value: "" for a typed number, otherwise the
  // formula or Excel reference whose evaluated number sits in the vector above.
  userInflationInputs: [],
  userAverageCaseReserveInputs: [],
  // The User Value cell that owns keyboard input, as `{ scope, devIndex }`.
  activeUserCell: null,
  // Excel references the workbook refused on the last check or refresh, keyed
  // by user cell, so the cell stays red across re-renders until it is fixed.
  excelLinkErrors: new Map(),
};

let cleanSnapshot = "";
let isDirty = false;
let programmatic = false;
let tabbedPage = null;
let bsLinksTab = null;
let sidecarLoadSequence = 0;
let outputPreviewTimer = 0;
let lastOutputPreviewMessage = null;
let derivedNumberFormatSeeded = false;
const activeDependencyPreviews = new Map();
const closeConfirm = createPageCloseConfirm({ subject: contract.displayLabel });
const headerState = { headerLabels: [], devHeaderLabels: [] };
const headersService = createDatasetHeadersService({
  state: headerState,
  setStatus: (message) => {
    if (text(message)) postStatus(message);
  },
});

const els = {
  projectInput: document.getElementById("bsProjectInput"),
  classInput: document.getElementById("bsClassInput"),
  nameInput: document.getElementById("bsNameInput"),
  methodTypeInput: document.getElementById("bsMethodTypeInput"),
  outputTypeInput: document.getElementById("bsOutputTypeInput"),
  outputTypeBtn: document.getElementById("bsOutputTypeBtn"),
  originLengthInput: document.getElementById("bsOriginLengthInput"),
  developmentLengthInput: document.getElementById("bsDevelopmentLengthInput"),
  numberFormatInput: document.getElementById("bsNumberFormatInput"),
  numberFormatField: document.getElementById("bsNumberFormatField"),
  numberFormatBtn: document.getElementById("bsNumberFormatBtn"),
  numberFormatMenu: document.getElementById("bsNumberFormatMenu"),
  decimalPlacesInput: document.getElementById("bsDecimalPlacesInput"),
  decimalPlacesUp: document.getElementById("bsDecimalPlacesUp"),
  decimalPlacesDown: document.getElementById("bsDecimalPlacesDown"),
  srInputs: document.getElementById("bsSrInputs"),
  craInputs: document.getElementById("bsCraInputs"),
  viewButtons: document.getElementById("bsViewButtons"),
  methodCaption: document.getElementById("bsMethodCaption"),
  methodCaptionText: document.getElementById("bsMethodCaptionText"),
  loessSpanField: document.getElementById("bsLoessSpanField"),
  loessSpanInput: document.getElementById("bsLoessSpanInput"),
  loessSpanUp: document.getElementById("bsLoessSpanUp"),
  loessSpanDown: document.getElementById("bsLoessSpanDown"),
  methodMessage: document.getElementById("bsMethodMessage"),
  methodHead: document.getElementById("bsMethodHead"),
  methodBody: document.getElementById("bsMethodBody"),
  secondaryPane: document.getElementById("bsSecondaryPane"),
  secondaryCaption: document.getElementById("bsSecondaryCaption"),
  secondaryCaptionText: document.getElementById("bsSecondaryCaptionText"),
  secondaryHead: document.getElementById("bsSecondaryHead"),
  secondaryBody: document.getElementById("bsSecondaryBody"),
  methodTable: document.getElementById("bsMethodTable"),
  methodTableWrap: document.getElementById("bsMethodTableWrap"),
  secondaryTable: document.getElementById("bsSecondaryTable"),
  secondaryTableWrap: document.getElementById("bsSecondaryTableWrap"),
  cellContextMenu: document.getElementById("bsCellContextMenu"),
  selectLeadingDiagonalItem: document.getElementById("bsSelectLeadingDiagonalItem"),
  auditLogMount: document.getElementById("bsAuditLogMount"),
  saveBtn: document.getElementById("bsSaveBtn"),
  cancelBtn: document.getElementById("bsCancelBtn"),
};

const notesController = mountNotesTab({
  container: document.getElementById("bsNotesMount"),
  ariaLabel: `${contract.displayLabel} notes`,
  onChange: () => markDirty(),
  onStatus: postStatus,
});

const auditLogView = createAuditLogView({
  container: els.auditLogMount,
  ariaLabel: `${contract.displayLabel} audit log`,
  emptyDescription: "Method saves will appear here after the first save.",
  normalizeEntries: normalizeSidecarAuditEntries,
  formatEventDate: formatSidecarAuditEventDate,
});

// Every grid tags its value cells for the shared spreadsheet selection, so a
// range can be picked and copied from any calculation view the way the
// Dataset Viewer and Result Selection grids allow.
const cellSelection = createBerquistShermanCellSelection({
  tables: [
    { key: "primary", table: els.methodTable, scrollHost: els.methodTableWrap },
    { key: "secondary", table: els.secondaryTable, scrollHost: els.secondaryTableWrap },
  ],
  contextMenu: els.cellContextMenu,
  onContextAction: (action) => {
    if (action === "select_leading_diagonal") selectLeadingDiagonal();
  },
  onCopied: () => postStatus(`Copied selected ${contract.displayLabel} values.`),
});

function text(value) {
  return String(value ?? "").trim();
}

function norm(value) {
  return text(value).replace(/\s+/g, " ").toLowerCase();
}

function numberOrNull(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function csvBaseName(value) {
  return text(value).split(/[\\/]/).pop();
}

function roleDefinitions() {
  return ROLE_DEFINITIONS[variant];
}

function viewDefinitions() {
  return VIEW_DEFINITIONS[variant];
}

function roleByKey(key) {
  return roleDefinitions().find((role) => role.key === key) || null;
}

function getDetails() {
  return {
    name: text(els.nameInput?.value),
    outputType: text(els.outputTypeInput?.value),
  };
}

function withProgrammatic(callback) {
  programmatic = true;
  try {
    return callback();
  } finally {
    programmatic = false;
  }
}

function postStatus(message, tone = "") {
  const statusText = text(message);
  try {
    window.parent?.postMessage(
      { type: "arcrho:status", text: statusText, ...(tone ? { tone } : {}) },
      "*",
    );
  } catch {}
}

function postDirty(dirty, force = false) {
  const next = !!dirty;
  if (!force && next === isDirty) return;
  isDirty = next;
  updateTabbedPageSaveControls({
    saveButton: els.saveBtn,
    cancelButton: els.cancelBtn,
    dirty: next,
  });
  try {
    window.parent?.postMessage({ type: "arcrho:dataset-dirty", inst, dirty: next }, "*");
  } catch {}
  if (next) scheduleOutputDependencyPreview("dirty");
  else clearOutputDependencyPreview("clean");
}

function configSnapshot() {
  const details = getDetails();
  const selection = variant === "sr"
    ? {
        selectedProportionSettled: state.selectedProportionSettled,
        selectedProportionIsDefault: state.selectedProportionIsDefault,
        selectedAdjustment: state.selectedAdjustment,
        loessSpan: state.loessSpan,
      }
    : {
        inflationSelection: state.inflationSelection,
        userInflation: state.userInflation,
        userInflationInputs: state.userInflationInputs,
        averageCaseReserveSelection: state.averageCaseReserveSelection,
        userAverageCaseReserves: state.userAverageCaseReserves,
        userAverageCaseReserveInputs: state.userAverageCaseReserveInputs,
        loessSpan: state.loessSpan,
      };
  return JSON.stringify({
    name: details.name,
    outputType: details.outputType,
    sourceNames: roleDefinitions().map((role) => [role.key, text(state.sourceNames[role.key])]),
    selection,
    // Only the format the user owns counts as an edit; the recorded source
    // formats follow their datasets and sync without making the page dirty.
    derivedNumberFormat: derivedNumberFormat(),
    notes: notesController.getValue(),
  });
}

function markDirty() {
  if (programmatic) return;
  const dirty = configSnapshot() !== cleanSnapshot;
  postDirty(dirty);
  if (dirty) scheduleOutputDependencyPreview("dirty");
}

function markClean() {
  cleanSnapshot = configSnapshot();
  notesController.markClean();
  postDirty(false, true);
}

function syncTitle() {
  const name = getDetails().name;
  document.title = name ? `${name} - ${contract.displayLabel}` : contract.displayLabel;
}

function setMethodMessage(message, tone = "") {
  if (!els.methodMessage) return;
  els.methodMessage.textContent = text(message);
  els.methodMessage.classList.toggle("error", tone === "error");
}

function latestDiagonalValues(matrix) {
  return (Array.isArray(matrix) ? matrix : []).map((row) => {
    const cells = Array.isArray(row) ? row : [];
    for (let index = cells.length - 1; index >= 0; index -= 1) {
      const value = numberOrNull(cells[index]);
      if (value !== null) return value;
    }
    return null;
  });
}

function outputDependencyMessage(type, reason = "") {
  const details = getDetails();
  const output = state.result?.output;
  const message = {
    type,
    inst,
    project: state.project,
    reservingClass: state.reservingClass,
    datasetName: details.name,
    datasetTypeName: details.outputType || details.name,
    names: [details.name, details.outputType].map(text).filter(Boolean),
    methodType: contract.methodType,
    sourceKind: contract.sourceKind,
    dataFormat: "Triangle",
    originLength: ANNUAL_PERIOD_LENGTH,
    developmentLength: ANNUAL_PERIOD_LENGTH,
    reason,
  };
  if (type === "arcrho:dependency-source-preview" && Array.isArray(output)) {
    message.values = latestDiagonalValues(output);
    message.matrixValues = cloneMatrix(output);
    message.mask = output.map((row) => (
      Array.isArray(row) ? row.map((value) => numberOrNull(value) !== null) : []
    ));
    message.originLabels = state.originLabels.map(String);
    message.developmentLabels = Array.from(
      { length: matrixDevelopmentCount() },
      (_, index) => getDevelopmentLabel(index),
    );
  }
  return message;
}

function outputPreviewIdentity(message) {
  return [
    norm(message?.project),
    normalizedReservingClassPath(message?.reservingClass),
    ...(Array.isArray(message?.names) ? message.names.map(norm).sort() : []),
  ].join("\u001f");
}

function postOutputDependencyPreview(reason = "dirty") {
  const output = state.result?.output;
  if (!Array.isArray(output) || !output.length) return;
  const message = outputDependencyMessage("arcrho:dependency-source-preview", reason);
  if (!message.names.length) return;
  if (
    lastOutputPreviewMessage
    && outputPreviewIdentity(lastOutputPreviewMessage) !== outputPreviewIdentity(message)
  ) {
    try {
      window.parent?.postMessage({
        ...lastOutputPreviewMessage,
        type: "arcrho:dependency-source-cleared",
        reason: "identity-changed",
      }, "*");
    } catch {}
  }
  lastOutputPreviewMessage = message;
  try {
    window.parent?.postMessage(message, "*");
  } catch {}
}

function scheduleOutputDependencyPreview(reason = "dirty", force = false) {
  window.clearTimeout(outputPreviewTimer);
  outputPreviewTimer = window.setTimeout(() => {
    outputPreviewTimer = 0;
    if (force || isDirty) postOutputDependencyPreview(reason);
  }, 120);
}

let pendingBsPropagationJobId = "";

function clearOutputDependencyPreview(reason = "clean") {
  window.clearTimeout(outputPreviewTimer);
  outputPreviewTimer = 0;
  const message = lastOutputPreviewMessage || outputDependencyMessage(
    "arcrho:dependency-source-cleared",
    reason,
  );
  lastOutputPreviewMessage = null;
  // Set only by a save that enqueued an Engine propagation job; Project
  // Instance defers the downstream preview clear until the job's terminal
  // status so dependents never snap back to stale values.
  const propagationJobId = pendingBsPropagationJobId;
  pendingBsPropagationJobId = "";
  if (!message.names?.length) return;
  try {
    window.parent?.postMessage({
      ...message,
      type: "arcrho:dependency-source-cleared",
      reason,
      propagationJobId,
    }, "*");
  } catch {}
}

function normalizeMatrix(rawValues) {
  const rows = Array.isArray(rawValues) ? rawValues : [];
  return rows.map((rawRow) => {
    const row = (Array.isArray(rawRow) ? rawRow : [rawRow]).map(numberOrNull);
    while (row.length && row[row.length - 1] === null) row.pop();
    return row;
  });
}

function normalizeVector(rawValues) {
  const rows = Array.isArray(rawValues) ? rawValues : [];
  return rows.map((rawRow) => {
    if (!Array.isArray(rawRow)) return numberOrNull(rawRow);
    for (const value of rawRow) {
      const number = numberOrNull(value);
      if (number !== null) return number;
    }
    return null;
  });
}

function normalizeStringVector(values) {
  return Array.isArray(values) ? values.map((value) => text(value) || null) : [];
}

function normalizeNumberVector(values) {
  return Array.isArray(values) ? values.map(numberOrNull) : [];
}

function normalizeBooleanVector(values) {
  return Array.isArray(values) ? values.map((value) => !!value) : [];
}

function normalizeSelectionMatrix(values) {
  return Array.isArray(values)
    ? values.map((row) => Array.isArray(row) ? row.map((value) => text(value) || null) : [])
    : [];
}

function matrixRowCount() {
  for (const role of roleDefinitions()) {
    if (role.format !== "Triangle") continue;
    const matrix = state.sourceValues[role.key];
    if (Array.isArray(matrix) && matrix.length) return matrix.length;
  }
  return 0;
}

function matrixDevelopmentCount() {
  let count = state.developmentLabels.length;
  for (const role of roleDefinitions()) {
    if (role.format !== "Triangle") continue;
    for (const row of state.sourceValues[role.key] || []) {
      count = Math.max(count, Array.isArray(row) ? row.length : 0);
    }
  }
  return count;
}

function getOriginLabel(index) {
  return text(state.originLabels[index]) || String(index + 1);
}

function getDevelopmentLabel(index) {
  return text(state.developmentLabels[index]) || `Age ${index + 1}`;
}

function ensureSelectionConfig() {
  const rowCount = matrixRowCount();
  const developmentCount = matrixDevelopmentCount();
  if (!rowCount || !developmentCount) return;

  if (variant === "sr") {
    if (state.selectedProportionSettled.length !== developmentCount) {
      const closed = state.sourceValues.closed_claim_numbers || [];
      const ultimate = state.sourceValues.ultimate_claim_numbers || [];
      state.selectedProportionSettled = Array.from({ length: developmentCount }, (_, devIndex) => {
        const originIndex = rowCount - devIndex - 1;
        const numerator = numberOrNull(closed[originIndex]?.[devIndex]);
        const denominator = numberOrNull(ultimate[originIndex]);
        return numerator !== null && denominator !== null && denominator !== 0
          ? numerator / denominator
          : 0;
      });
    }
    if (state.selectedProportionIsDefault.length !== developmentCount) {
      state.selectedProportionIsDefault = Array(developmentCount).fill(true);
    }
    const paid = state.sourceValues.paid_claims || [];
    const selectionMatches = state.selectedAdjustment.length === paid.length
      && paid.every((row, rowIndex) => (
        Array.isArray(state.selectedAdjustment[rowIndex])
        && state.selectedAdjustment[rowIndex].length >= row.length
      ));
    if (!selectionMatches) {
      state.selectedAdjustment = paid.map((row) => (
        row.map(() => row.length > 1 ? "pairs" : "unadjusted")
      ));
    }
    return;
  }

  if (state.inflationSelection.length !== developmentCount) {
    state.inflationSelection = Array(developmentCount).fill("paid_all");
  }
  // A User Value cell starts blank; the calculation reads a blank as zero.
  if (state.userInflation.length !== developmentCount) {
    state.userInflation = Array(developmentCount).fill(null);
  }
  if (state.userInflationInputs.length !== developmentCount) {
    state.userInflationInputs = normalizeUserValueInputs(state.userInflationInputs, developmentCount);
  }
  if (state.averageCaseReserveSelection.length !== developmentCount) {
    state.averageCaseReserveSelection = Array(developmentCount).fill("latest");
  }
  if (state.userAverageCaseReserves.length !== developmentCount) {
    state.userAverageCaseReserves = Array(developmentCount).fill(null);
  }
  if (state.userAverageCaseReserveInputs.length !== developmentCount) {
    state.userAverageCaseReserveInputs = normalizeUserValueInputs(
      state.userAverageCaseReserveInputs,
      developmentCount,
    );
  }
}

function validateSourceShape() {
  const missing = roleDefinitions()
    .filter((role) => !text(state.sourceNames[role.key]))
    .map((role) => role.label);
  if (missing.length) throw new Error(`Select ${missing.join(", ")}.`);

  const triangleRoles = roleDefinitions().filter((role) => role.format === "Triangle");
  const firstMatrix = state.sourceValues[triangleRoles[0]?.key];
  if (!Array.isArray(firstMatrix) || !firstMatrix.length) {
    throw new Error("The selected annual triangle is empty.");
  }
  const rowLengths = firstMatrix.map((row) => Array.isArray(row) ? row.length : 0);
  if (!rowLengths.some((count) => count > 0)) throw new Error("The selected annual triangle is empty.");

  for (const role of triangleRoles.slice(1)) {
    const matrix = state.sourceValues[role.key];
    if (!Array.isArray(matrix) || matrix.length !== firstMatrix.length) {
      throw new Error(`${role.label} does not match the selected annual triangle shape.`);
    }
    const matches = matrix.every((row, index) => (
      (Array.isArray(row) ? row.length : 0) === rowLengths[index]
    ));
    if (!matches) throw new Error(`${role.label} does not match the selected annual triangle shape.`);
  }

  for (const role of roleDefinitions().filter((item) => item.format === "Vector")) {
    const vector = state.sourceValues[role.key];
    if (!Array.isArray(vector) || vector.length !== firstMatrix.length) {
      throw new Error(`${role.label} must have one value for each origin period.`);
    }
  }

  const developmentCount = Math.max(...rowLengths, 0);
  if (state.originLabels.length !== firstMatrix.length) {
    throw new Error("Annual origin labels do not match the selected source rows.");
  }
  if (state.developmentLabels.length < developmentCount) {
    throw new Error(
      "Annual development labels are unavailable. Check the project's development dates.",
    );
  }
  for (const role of roleDefinitions()) {
    const payload = state.sourcePayloads[role.key] || {};
    const originLabels = payload.origin_labels || payload.originLabels;
    if (
      Array.isArray(originLabels)
      && originLabels.length
      && (
        originLabels.length !== state.originLabels.length
        || originLabels.some((label, index) => text(label) !== text(state.originLabels[index]))
      )
    ) {
      throw new Error(`${role.label} uses different origin periods.`);
    }
    if (role.format !== "Triangle") continue;
    const developmentLabels = payload.development_labels
      || payload.developmentLabels
      || payload.dev_labels;
    if (
      Array.isArray(developmentLabels)
      && developmentLabels.length
      && (
        developmentLabels.length < developmentCount
        || state.developmentLabels.slice(0, developmentCount).some(
          (label, index) => text(label) !== text(developmentLabels[index]),
        )
      )
    ) {
      throw new Error(`${role.label} uses different development periods.`);
    }
  }
}

function calculateCurrent() {
  try {
    validateSourceShape();
    ensureSelectionConfig();
    state.result = variant === "sr"
      ? calculateSettlementRate({
          paidClaims: state.sourceValues.paid_claims,
          closedClaimNumbers: state.sourceValues.closed_claim_numbers,
          ultimateClaimNumbers: state.sourceValues.ultimate_claim_numbers,
          selectedProportionSettled: state.selectedProportionSettled,
          selectedProportionIsDefault: state.selectedProportionIsDefault,
          selectedAdjustment: state.selectedAdjustment,
          loessSpan: state.loessSpan,
        })
      : calculateCaseReserveAdequacy({
          reportedClaimNumbers: state.sourceValues.reported_claim_numbers,
          closedClaimNumbers: state.sourceValues.closed_claim_numbers,
          incurredClaims: state.sourceValues.incurred_claims,
          paidClaims: state.sourceValues.paid_claims,
          // The stored COL exclusion flags are non-contributing and deferred from the MVP.
          avgCaseReserveExclusions: [],
          avgPaidClaimsExclusions: [],
          inflationSelection: state.inflationSelection,
          userInflation: state.userInflation,
          averageCaseReserveSelection: state.averageCaseReserveSelection,
          userAverageCaseReserves: state.userAverageCaseReserves,
          loessSpan: state.loessSpan,
        });
    // A good calculation reports nothing, as in ResQ: the message row only
    // carries failures, and clearing it drops any error from a prior attempt.
    setMethodMessage("");
  } catch (error) {
    state.result = null;
    if (lastOutputPreviewMessage) clearOutputDependencyPreview("invalid");
    setMethodMessage(text(error?.message || error), "error");
  }
  renderMethodTable();
  if (isDirty) scheduleOutputDependencyPreview("dirty");
}

function matrixForCurrentView() {
  const roleKey = SOURCE_VIEW_ROLES[state.currentView];
  if (roleKey) return state.sourceValues[roleKey] || [];
  const value = state.result?.[state.currentView];
  return Array.isArray(value) ? value : [];
}

function currentViewDefinition() {
  return viewDefinitions().find((item) => item.key === state.currentView) || viewDefinitions()[0];
}

// A rate is not currency: at the Details format's zero or two decimals an
// inflation estimate or a proportion settled would collapse to the same figure
// in every column, so those two grids keep ResQ's four-decimal presentation.
const RATE_DECIMAL_PLACES = 4;

function formatPercentValue(value) {
  const number = numberOrNull(value);
  if (number === null) return "";
  return `${(number * 100).toLocaleString(undefined, {
    minimumFractionDigits: RATE_DECIMAL_PLACES,
    maximumFractionDigits: RATE_DECIMAL_PLACES,
  })}%`;
}

function parsePercentInput(rawText) {
  const numeric = Number.parseFloat(text(rawText).replace(/%/gu, "").replace(/,/gu, ""));
  return Number.isFinite(numeric) ? numeric / 100 : null;
}

function derivedNumberFormat() {
  return state.numberFormats.derived;
}

// An input view shows its own dataset instance exactly as the Dataset Viewer
// does; everything the method calculates uses the Details-tab format.
function roleNumberFormat(roleKey) {
  return state.numberFormats.sources[roleKey] || derivedNumberFormat();
}

function viewNumberFormat(viewKey = state.currentView) {
  const roleKey = SOURCE_VIEW_ROLES[viewKey];
  return roleKey ? roleNumberFormat(roleKey) : derivedNumberFormat();
}

// A blank is not a zero. `formatDatasetNumberValue` coerces its argument with
// `Number(...)`, which turns a missing cell into 0, so an absent value has to be
// resolved to an empty string here instead of being handed to the formatter.
function formatCellValue(value, format) {
  const number = numberOrNull(value);
  if (number === null) return "";
  return formatDatasetNumberValue(
    number,
    format.number_format,
    format.decimal_places,
  );
}

function formatRateValue(value) {
  const number = numberOrNull(value);
  if (number === null) return "";
  return number.toLocaleString(undefined, {
    minimumFractionDigits: RATE_DECIMAL_PLACES,
    maximumFractionDigits: RATE_DECIMAL_PLACES,
  });
}

function syncNumberFormatControls() {
  const derived = derivedNumberFormat();
  if (els.numberFormatInput && document.activeElement !== els.numberFormatInput) {
    els.numberFormatInput.value = derived.number_format;
  }
  if (els.decimalPlacesInput && document.activeElement !== els.decimalPlacesInput) {
    els.decimalPlacesInput.value = String(derived.decimal_places);
  }
}

// The Details pair behaves like the Dataset Viewer's: typing a pattern re-reads
// its decimal places, and stepping the places rewrites the pattern.
function applyDerivedNumberFormat(entry, { fromDecimalPlaces = false } = {}) {
  const current = derivedNumberFormat();
  const pattern = normalizeDatasetNumberFormat(
    entry.number_format ?? current.number_format,
    current.number_format,
  );
  const places = fromDecimalPlaces
    ? clampDatasetDecimalPlaces(entry.decimal_places, current.decimal_places)
    : getDatasetNumberFormatDecimalPlaces(pattern);
  const next = normalizeNumberFormatEntry({ number_format: pattern, decimal_places: places });
  if (sameNumberFormat(next, current)) {
    syncNumberFormatControls();
    return;
  }
  derivedNumberFormatSeeded = true;
  state.numberFormats.derived = next;
  syncNumberFormatControls();
  renderMethodTable();
  markDirty();
}

function syncLoessSpanControls() {
  if (els.loessSpanInput && document.activeElement !== els.loessSpanInput) {
    els.loessSpanInput.value = String(state.loessSpan);
  }
}

// The caption rows mirror the ResQ tab headers, including the Loess Span
// spinner beside the caption of whichever grid owns the loess estimator.
function syncMethodChrome() {
  const view = currentViewDefinition();
  const caption = text(view?.caption);
  const secondaryCaption = text(view?.secondaryCaption);
  const loessCaption = variant === "sr"
    ? (state.currentView === "output" ? els.methodCaption : null)
    : (state.currentView === "avgSelections" ? els.secondaryCaption : null);
  if (els.methodCaptionText) els.methodCaptionText.textContent = caption;
  if (els.secondaryCaptionText) els.secondaryCaptionText.textContent = secondaryCaption;
  if (els.secondaryPane) els.secondaryPane.hidden = !secondaryCaption;
  if (els.loessSpanField) {
    els.loessSpanField.hidden = !loessCaption;
    if (loessCaption && els.loessSpanField.parentElement !== loessCaption) {
      loessCaption.appendChild(els.loessSpanField);
    }
  }
  if (els.methodCaption) {
    els.methodCaption.hidden = !caption && loessCaption !== els.methodCaption;
  }
  syncLoessSpanControls();
}

function applyLoessSpan(value) {
  const span = normalizeLoessSpan(value);
  if (span === state.loessSpan) {
    syncLoessSpanControls();
    return;
  }
  state.loessSpan = span;
  syncLoessSpanControls();
  recalculateAfterSelectionEdit();
}

function renderViewButtons() {
  if (!els.viewButtons) return;
  els.viewButtons.replaceChildren();
  for (const view of viewDefinitions()) {
    const button = document.createElement("button");
    button.type = "button";
    const active = view.key === state.currentView;
    button.className = `bsViewTab tabbedPageTab${active ? " active" : ""}`;
    button.dataset.view = view.key;
    button.textContent = view.label;
    button.setAttribute("role", "tab");
    button.setAttribute("aria-selected", active ? "true" : "false");
    button.tabIndex = active ? 0 : -1;
    button.addEventListener("click", () => {
      state.currentView = view.key;
      cellSelection.clearAll();
      renderViewButtons();
      renderMethodTable();
    });
    els.viewButtons.appendChild(button);
  }
  syncMethodChrome();
}

function buildMethodHeaderRow(cornerLabel, developmentCount) {
  const headerRow = document.createElement("tr");
  const corner = document.createElement("th");
  corner.scope = "col";
  corner.textContent = cornerLabel;
  headerRow.appendChild(corner);
  for (let devIndex = 0; devIndex < developmentCount; devIndex += 1) {
    const header = document.createElement("th");
    header.scope = "col";
    header.dataset.c = String(devIndex);
    header.textContent = getDevelopmentLabel(devIndex);
    headerRow.appendChild(header);
  }
  return headerRow;
}

// A value cell carries its grid position and its raw figure for the shared
// selection and copy; a row label carries the position alone, so a click on it
// selects the row.
function tagValueCell(cell, rowIndex, colIndex, rawValue) {
  cell.dataset.r = String(rowIndex);
  cell.dataset.c = String(colIndex);
  cell.dataset.copyValue = numberOrNull(rawValue) === null ? "" : String(rawValue);
  cell.setAttribute("aria-selected", "false");
  return cell;
}

function tagRowLabel(cell, rowIndex) {
  cell.dataset.r = String(rowIndex);
  return cell;
}

// Ingestion applies the Dataset Viewer mask and the annual staircase, so a
// triangle row is exactly as long as the cells it owns. Everything past that
// length is the unavailable lower-right area: it carries no value and no grid,
// the way the Dataset Viewer blanks the same region.
function populatedRowLength(matrix, rowIndex) {
  const row = matrix?.[rowIndex];
  return Array.isArray(row) ? row.length : 0;
}

function maskedCell(rowIndex, devIndex) {
  const cell = document.createElement("td");
  cell.className = "bsMaskedCell";
  // Part of the grid for selection, so a drag to the triangle's corner takes
  // the whole triangle; it copies as an empty entry.
  return tagValueCell(cell, rowIndex, devIndex, null);
}

function adjustmentGridRows() {
  const span = state.result?.loessSpan ?? state.loessSpan;
  return [
    { method: "unadjusted", label: "Unadjusted :", matrix: state.sourceValues.paid_claims || [] },
    { method: "pairs", label: "Pairs :", matrix: state.result?.pairsAdjustment || [] },
    { method: "all", label: "All :", matrix: state.result?.allAdjustment || [] },
    { method: "loess", label: `Loess (${span}) :`, matrix: state.result?.loessAdjustment || [] },
    { method: "selected", label: "Selected :", matrix: state.result?.adjustedPaidClaims || [] },
  ];
}

function selectedAdjustmentFor(rowIndex, devIndex) {
  const effective = state.result?.selectedAdjustment;
  return norm(effective?.[rowIndex]?.[devIndex] ?? state.selectedAdjustment?.[rowIndex]?.[devIndex]);
}

// ResQ-style Adjusted Paid Claims grid: five estimator rows per origin period.
// Click a value to select that cell, click a row label for the whole origin,
// Ctrl+click a row label for every origin and development period.
function renderAdjustedPaidGrid() {
  const format = derivedNumberFormat();
  const paid = state.sourceValues.paid_claims || [];
  const developmentCount = matrixDevelopmentCount();
  els.methodHead.replaceChildren(buildMethodHeaderRow("Accident Year", developmentCount));

  const body = document.createDocumentFragment();
  const gridRows = adjustmentGridRows();
  // The estimator rows number consecutively across origins for the selection;
  // an origin's caption row is not a row of values.
  let gridRowIndex = 0;
  for (let rowIndex = 0; rowIndex < paid.length; rowIndex += 1) {
    const populatedCount = populatedRowLength(paid, rowIndex);
    const yearRow = document.createElement("tr");
    yearRow.className = "bsAdjYearRow";
    const yearLabel = document.createElement("td");
    yearLabel.textContent = getOriginLabel(rowIndex);
    yearRow.appendChild(yearLabel);
    const yearFill = document.createElement("td");
    yearFill.colSpan = Math.max(developmentCount, 1);
    yearRow.appendChild(yearFill);
    body.appendChild(yearRow);

    for (const gridRow of gridRows) {
      const isSelectedRow = gridRow.method === "selected";
      const rowElement = document.createElement("tr");
      const label = tagRowLabel(document.createElement("td"), gridRowIndex);
      label.className = `bsAdjRowLabel${isSelectedRow ? " bsAdjSelectedRowLabel" : ""}`;
      label.textContent = gridRow.label;
      if (!isSelectedRow) {
        label.classList.add("bsAdjPick");
        label.dataset.adjustMethod = gridRow.method;
        label.dataset.adjustOrigin = String(rowIndex);
        label.tabIndex = 0;
        label.title = "Click selects this estimator for the origin period. Ctrl+Click selects it everywhere.";
      }
      rowElement.appendChild(label);
      for (let devIndex = 0; devIndex < developmentCount; devIndex += 1) {
        if (devIndex >= populatedCount) {
          rowElement.appendChild(maskedCell(gridRowIndex, devIndex));
          continue;
        }
        const cell = document.createElement("td");
        const rawValue = gridRow.matrix[rowIndex]?.[devIndex];
        tagValueCell(cell, gridRowIndex, devIndex, rawValue);
        cell.textContent = formatCellValue(rawValue, format);
        cell.title = numberOrNull(rawValue) === null ? "" : String(rawValue);
        if (isSelectedRow) {
          cell.className = "bsAdjSelectedValue";
        } else {
          cell.className = "bsAdjCell bsAdjPick";
          cell.dataset.adjustMethod = gridRow.method;
          cell.dataset.adjustOrigin = String(rowIndex);
          cell.dataset.adjustDev = String(devIndex);
          cell.tabIndex = 0;
          if (selectedAdjustmentFor(rowIndex, devIndex) === gridRow.method) {
            cell.classList.add("bsAdjSelectedSource");
          }
        }
        rowElement.appendChild(cell);
      }
      body.appendChild(rowElement);
      gridRowIndex += 1;
    }
  }
  els.methodBody.replaceChildren(body);
}

function applyAdjustmentSelection(method, originIndex, devIndex, applyToAllOrigins) {
  const paid = state.sourceValues.paid_claims || [];
  if (applyToAllOrigins) {
    state.selectedAdjustment = paid.map((row) => (Array.isArray(row) ? row : []).map(() => method));
  } else if (devIndex === null) {
    state.selectedAdjustment[originIndex] = (paid[originIndex] || []).map(() => method);
  } else {
    const row = paid[originIndex] || [];
    const previous = Array.isArray(state.selectedAdjustment[originIndex])
      ? state.selectedAdjustment[originIndex]
      : [];
    if (previous.length < row.length) {
      state.selectedAdjustment[originIndex] = row.map((_, index) => previous[index] ?? null);
    }
    state.selectedAdjustment[originIndex][devIndex] = method;
  }
  recalculateAfterSelectionEdit();
}

function handleAdjustmentGridEvent(event) {
  if (variant !== "sr" || state.currentView !== "output") return;
  const target = event.target instanceof Element
    ? event.target.closest("[data-adjust-method]")
    : null;
  if (!target) return;
  if (event.type === "keydown" && event.key !== "Enter" && event.key !== " ") return;
  event.preventDefault();
  const method = text(target.dataset.adjustMethod);
  const originIndex = Number.parseInt(target.dataset.adjustOrigin || "", 10);
  if (!method || !Number.isInteger(originIndex)) return;
  const devRaw = Number.parseInt(target.dataset.adjustDev || "", 10);
  applyAdjustmentSelection(
    method,
    originIndex,
    Number.isInteger(devRaw) ? devRaw : null,
    (event.ctrlKey || event.metaKey) && !Number.isInteger(devRaw),
  );
}

// The green highlight marks the proportion cell currently feeding each
// development period: the leading diagonal by default, or the clicked value.
function proportionHighlightRowIndex(proportionMatrix, devIndex) {
  const isDefault = state.selectedProportionIsDefault[devIndex] === true;
  const selected = numberOrNull(state.result?.selectedProportionSettled?.[devIndex]);
  for (let rowIndex = proportionMatrix.length - 1; rowIndex >= 0; rowIndex -= 1) {
    const value = numberOrNull(proportionMatrix[rowIndex]?.[devIndex]);
    if (value === null) continue;
    if (isDefault) return rowIndex;
    if (selected !== null && Math.abs(value - selected) <= Math.max(1e-12, Math.abs(selected) * 1e-9)) {
      return rowIndex;
    }
  }
  return -1;
}

// ResQ-style Proportion Settled tab: click a triangle value to select it for
// that development period, type into the Selected row for custom values, and
// right-click for Select Leading Diagonal.
function renderProportionSettledGrid() {
  const proportionMatrix = Array.isArray(state.result?.proportionSettled)
    ? state.result.proportionSettled
    : [];
  const developmentCount = matrixDevelopmentCount();
  els.methodHead.replaceChildren(buildMethodHeaderRow("Accident Year", developmentCount));

  const body = document.createDocumentFragment();
  const highlightByColumn = Array.from(
    { length: developmentCount },
    (_, devIndex) => proportionHighlightRowIndex(proportionMatrix, devIndex),
  );
  for (let rowIndex = 0; rowIndex < proportionMatrix.length; rowIndex += 1) {
    const rowElement = document.createElement("tr");
    const origin = tagRowLabel(document.createElement("td"), rowIndex);
    origin.textContent = getOriginLabel(rowIndex);
    rowElement.appendChild(origin);
    const populatedCount = populatedRowLength(proportionMatrix, rowIndex);
    for (let devIndex = 0; devIndex < developmentCount; devIndex += 1) {
      if (devIndex >= populatedCount) {
        rowElement.appendChild(maskedCell(rowIndex, devIndex));
        continue;
      }
      const cell = document.createElement("td");
      const rawValue = numberOrNull(proportionMatrix[rowIndex]?.[devIndex]);
      tagValueCell(cell, rowIndex, devIndex, rawValue);
      if (rawValue === null) {
        rowElement.appendChild(cell);
        continue;
      }
      cell.className = "bsPropCell";
      cell.textContent = formatPercentValue(rawValue);
      cell.title = String(rawValue);
      cell.dataset.propDev = String(devIndex);
      cell.dataset.propValue = String(rawValue);
      cell.tabIndex = 0;
      if (highlightByColumn[devIndex] === rowIndex) cell.classList.add("bsPropSelectedSource");
      rowElement.appendChild(cell);
    }
    body.appendChild(rowElement);
  }

  const spacerRow = document.createElement("tr");
  spacerRow.className = "bsPropSpacerRow";
  for (let index = 0; index <= developmentCount; index += 1) {
    spacerRow.appendChild(document.createElement("td"));
  }
  body.appendChild(spacerRow);

  // The Selected row is the row after the triangle for the selection, so a
  // drag down from the triangle takes the selections with it.
  const selectedRow = document.createElement("tr");
  selectedRow.className = "bsPropSelectedRow";
  const selectedLabel = tagRowLabel(document.createElement("td"), proportionMatrix.length);
  selectedLabel.textContent = "Selected";
  selectedRow.appendChild(selectedLabel);
  const selectedValues = Array.isArray(state.result?.selectedProportionSettled)
    ? state.result.selectedProportionSettled
    : state.selectedProportionSettled;
  for (let devIndex = 0; devIndex < developmentCount; devIndex += 1) {
    const cell = tagValueCell(document.createElement("td"), proportionMatrix.length, devIndex, selectedValues[devIndex]);
    const input = document.createElement("input");
    input.type = "text";
    input.className = "bsPropSelectedInput";
    input.value = formatPercentValue(selectedValues[devIndex]);
    input.setAttribute(
      "aria-label",
      `${getDevelopmentLabel(devIndex)} selected proportion settled`,
    );
    input.dataset.propDev = String(devIndex);
    cell.appendChild(input);
    selectedRow.appendChild(cell);
  }
  body.appendChild(selectedRow);
  els.methodBody.replaceChildren(body);
}

function applyProportionSelection(devIndex, value) {
  const parsed = numberOrNull(value);
  if (parsed === null || !Number.isInteger(devIndex) || devIndex < 0) return;
  ensureSelectionConfig();
  state.selectedProportionSettled[devIndex] = parsed;
  state.selectedProportionIsDefault[devIndex] = false;
  recalculateAfterSelectionEdit();
}

function selectLeadingDiagonal() {
  state.selectedProportionIsDefault = Array(matrixDevelopmentCount()).fill(true);
  recalculateAfterSelectionEdit();
}

function handleProportionGridEvent(event) {
  if (variant !== "sr" || state.currentView !== "proportionSettled") return;
  if (event.target instanceof HTMLInputElement && event.target.classList.contains("bsPropSelectedInput")) {
    const input = event.target;
    const commit = () => {
      const devIndex = Number.parseInt(input.dataset.propDev || "", 10);
      const parsed = parsePercentInput(input.value);
      if (parsed === null) {
        input.value = formatPercentValue(state.result?.selectedProportionSettled?.[devIndex]);
        return;
      }
      applyProportionSelection(devIndex, parsed);
    };
    if (event.type === "change") commit();
    if (event.type === "keydown" && event.key === "Enter") input.blur();
    return;
  }
  if (event.type === "keydown" && event.key !== "Enter" && event.key !== " ") return;
  const cell = event.target instanceof Element ? event.target.closest("td.bsPropCell") : null;
  if (!cell) return;
  event.preventDefault();
  applyProportionSelection(
    Number.parseInt(cell.dataset.propDev || "", 10),
    numberOrNull(cell.dataset.propValue),
  );
}

const SELECTION_SCOPES = Object.freeze({
  inflation: Object.freeze({
    rows: INFLATION_GRID_ROWS,
    selectionKey: "inflationSelection",
    userKey: "userInflation",
    inputsKey: "userInflationInputs",
    defaultMethod: "paid_all",
    label: "Average Inflation",
    ariaSuffix: "inflation",
  }),
  average: Object.freeze({
    rows: AVERAGE_GRID_ROWS,
    selectionKey: "averageCaseReserveSelection",
    userKey: "userAverageCaseReserves",
    inputsKey: "userAverageCaseReserveInputs",
    defaultMethod: "latest",
    label: "Current Average Case Reserves",
    ariaSuffix: "average case reserve",
  }),
});

function selectionRowValue(row, devIndex, userValues) {
  if (row.user) return userValues[devIndex];
  const source = state.result?.[row.key];
  if (row.constant) return numberOrNull(source);
  return Array.isArray(source) ? source[devIndex] : null;
}

function userCellKey(scope, devIndex) {
  return `${scope}:${devIndex}`;
}

function userCellHost(scope) {
  return scope === "inflation" ? els.methodBody : els.secondaryBody;
}

function findUserCell(scope, devIndex) {
  return userCellHost(scope)?.querySelector(`td.bsSelUserCell[data-select-dev="${devIndex}"]`) || null;
}

// ResQ-style CRA selection grid: one estimator per row, one development period
// per column. Click a value to select that estimator for the development
// period, click a row label to select it for every development period, and type
// into the User Value row to enter and select your own figure.
function renderColumnSelectionGrid(scope, head, body, formatValue) {
  const config = SELECTION_SCOPES[scope];
  const developmentCount = matrixDevelopmentCount();
  // The calculation normalizes the selection, so the grid echoes what actually
  // fed the result rather than the raw state it was built from. The User Value
  // row shows the raw state instead: a blank cell must read blank, not as the
  // zero the calculation stands in for it.
  const selection = state.result?.[config.selectionKey] || state[config.selectionKey];
  const userValues = state[config.userKey];
  const userInputs = state[config.inputsKey];
  head.replaceChildren(buildMethodHeaderRow("", developmentCount));

  const fragment = document.createDocumentFragment();
  // The value rows number consecutively for the selection; a grouping row is
  // not a row of values.
  let gridRowIndex = 0;
  for (const row of config.rows) {
    const rowElement = document.createElement("tr");
    if (row.group) {
      rowElement.className = "bsSelGroupRow";
      const groupLabel = document.createElement("td");
      groupLabel.textContent = row.group;
      rowElement.appendChild(groupLabel);
      const fill = document.createElement("td");
      fill.colSpan = Math.max(developmentCount, 1);
      rowElement.appendChild(fill);
      fragment.appendChild(rowElement);
      continue;
    }

    const label = tagRowLabel(document.createElement("td"), gridRowIndex);
    label.textContent = row.loess
      ? `${row.label} (${state.result?.loessSpan ?? state.loessSpan})`
      : row.label;
    if (row.selected) {
      label.className = "bsSelSelectedRowLabel";
    } else {
      label.className = "bsSelPick";
      label.dataset.selectScope = scope;
      label.dataset.selectMethod = row.method;
      label.tabIndex = 0;
      label.title = "Click selects this estimator for every development period.";
    }
    rowElement.appendChild(label);

    for (let devIndex = 0; devIndex < developmentCount; devIndex += 1) {
      const age = getDevelopmentLabel(devIndex);
      const cell = document.createElement("td");
      const rawValue = selectionRowValue(row, devIndex, userValues);
      tagValueCell(cell, gridRowIndex, devIndex, rawValue);
      cell.title = numberOrNull(rawValue) === null ? "" : String(rawValue);
      if (row.selected) {
        cell.className = "bsSelSelectedValue";
        cell.textContent = formatValue(rawValue);
        rowElement.appendChild(cell);
        continue;
      }
      cell.className = row.user ? "bsSelUserCell bsSelPick" : "bsSelCell bsSelPick";
      cell.dataset.selectScope = scope;
      cell.dataset.selectMethod = row.method;
      cell.dataset.selectDev = String(devIndex);
      if (norm(selection[devIndex]) === row.method) {
        cell.classList.add("bsSelSelectedSource");
      }
      cell.tabIndex = 0;
      if (row.user) {
        const input = text(userInputs[devIndex]);
        cell.textContent = numberOrNull(rawValue) === null ? "" : formatValue(rawValue);
        cell.setAttribute("aria-label", `${age} user ${config.ariaSuffix}`);
        if (input) {
          cell.classList.add(userValueExcelReferences(input).length ? "bsSelExcelLinked" : "bsSelFormulaCell");
          cell.title = input;
          const failure = state.excelLinkErrors.get(userCellKey(scope, devIndex));
          if (failure) {
            cell.classList.add("bsSelExcelLinkError");
            cell.title = `${input}\n${failure}`;
          }
        }
      } else {
        cell.textContent = formatValue(rawValue);
      }
      rowElement.appendChild(cell);
    }
    fragment.appendChild(rowElement);
    gridRowIndex += 1;
  }
  body.replaceChildren(fragment);
}

function renderSelectionGrids() {
  if (!state.result) return;
  const derived = derivedNumberFormat();
  renderColumnSelectionGrid("inflation", els.methodHead, els.methodBody, formatRateValue);
  if (els.secondaryHead && els.secondaryBody) {
    renderColumnSelectionGrid(
      "average",
      els.secondaryHead,
      els.secondaryBody,
      (value) => formatCellValue(value, derived),
    );
  }
}

function applySelectionChoice(scope, method, devIndex) {
  const config = SELECTION_SCOPES[scope];
  if (!config || !method) return;
  ensureSelectionConfig();
  const selection = state[config.selectionKey];
  if (devIndex === null) {
    for (let index = 0; index < matrixDevelopmentCount(); index += 1) selection[index] = method;
  } else {
    selection[devIndex] = method;
  }
  recalculateAfterSelectionEdit();
}

/*
The User Value row follows the Result Selection entry UX rather than a row of
permanent inputs: a click activates a cell, typing a digit, a sign, `=`, or a
quote starts an edit seeded with that key, Enter or F2 or a double-click opens
the current text, Enter commits, Escape abandons, Tab commits and moves on, the
arrow keys move the active cell, and Delete or Backspace clears it. A committed
figure is entered and selected for its column; clearing a cell whose column
selected it falls back to the grid's default estimator. Excel references and
formulas are committed through the same path, so a pasted link behaves like a
typed number that happens to be read from a workbook.
*/
let userCellEditor = null;

// The active User Value cell is the anchor of the shared grid selection, so it
// wears the shared dashed outline and Ctrl+C copies it with the range.
function setActiveUserCell(scope, devIndex, { focus = true } = {}) {
  const cell = findUserCell(scope, devIndex);
  if (!cell) {
    state.activeUserCell = null;
    return;
  }
  state.activeUserCell = { scope, devIndex };
  cellSelection.selectCell(cell);
  if (focus) cell.focus();
}

function openUserCellEditor(scope, devIndex, seedText = null) {
  const config = SELECTION_SCOPES[scope];
  if (!config) return;
  if (userCellEditor) {
    if (userCellEditor.scope === scope && userCellEditor.devIndex === devIndex) return;
    cancelUserCellEditor();
  }
  const cell = findUserCell(scope, devIndex);
  if (!cell) return;
  ensureSelectionConfig();
  const stored = text(state[config.inputsKey][devIndex]);
  const value = state[config.userKey][devIndex];
  const input = document.createElement("input");
  input.type = "text";
  input.className = "bsSelUserInput";
  input.autocomplete = "off";
  input.spellcheck = false;
  input.value = seedText ?? (stored || (numberOrNull(value) === null ? "" : String(value)));
  input.setAttribute("aria-label", `${getDevelopmentLabel(devIndex)} user ${config.ariaSuffix}`);
  const editor = { scope, devIndex, cell, input, busy: false };
  input.addEventListener("keydown", (event) => {
    if (editor.busy) {
      event.preventDefault();
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      void commitUserCellEditor(editor);
    } else if (event.key === "Escape") {
      event.preventDefault();
      cancelUserCellEditor();
      setActiveUserCell(scope, devIndex);
    } else if (event.key === "Tab") {
      event.preventDefault();
      void commitUserCellEditor(editor, { moveBy: event.shiftKey ? -1 : 1 });
    }
    event.stopPropagation();
  });
  input.addEventListener("input", () => input.classList.remove("isInvalid"));
  // Leaving the editor commits; an entry the page refuses on the way out is
  // dropped with its reason in the status bar rather than trapping focus.
  input.addEventListener("blur", () => {
    if (userCellEditor !== editor || editor.busy) return;
    void commitUserCellEditor(editor, { fromBlur: true });
  });
  cell.classList.add("isEditing");
  cell.replaceChildren(input);
  userCellEditor = editor;
  state.activeUserCell = { scope, devIndex };
  input.focus();
  if (seedText === null) input.select();
}

function cancelUserCellEditor() {
  const editor = userCellEditor;
  if (!editor) return;
  userCellEditor = null;
  renderMethodTable();
}

async function commitUserCellEditor(editor, { moveBy = 0, fromBlur = false } = {}) {
  if (!editor || userCellEditor !== editor || editor.busy) return false;
  const { scope, devIndex, input } = editor;
  editor.busy = true;
  input.readOnly = true;
  input.setAttribute("aria-busy", "true");
  try {
    await applyUserValueEntry(scope, devIndex, input.value);
  } catch (error) {
    const message = text(error?.message || error);
    if (fromBlur) {
      if (userCellEditor === editor) cancelUserCellEditor();
      restoreActiveUserCellFocus();
      postStatus(message, "error");
      return false;
    }
    editor.busy = false;
    input.readOnly = false;
    input.removeAttribute("aria-busy");
    input.classList.add("isInvalid");
    postStatus(message, "error");
    input.focus();
    return false;
  }
  // The grid re-rendered under the commit, so the editor is already gone.
  if (userCellEditor === editor) userCellEditor = null;
  if (fromBlur) {
    // Focus already went where the user sent it - often another User Value
    // cell, which the click activated before this commit finished - so only
    // the anchor is restored, and focus only when the re-render dropped it.
    restoreActiveUserCellFocus();
    return true;
  }
  const nextIndex = Math.min(Math.max(devIndex + moveBy, 0), matrixDevelopmentCount() - 1);
  setActiveUserCell(scope, nextIndex);
  return true;
}

function restoreActiveUserCellFocus() {
  const active = state.activeUserCell;
  if (!active) return;
  const focusLost = !document.activeElement || document.activeElement === document.body;
  setActiveUserCell(active.scope, active.devIndex, { focus: focusLost });
}

function setUserValue(scope, devIndex, value, input) {
  const config = SELECTION_SCOPES[scope];
  state[config.userKey][devIndex] = value;
  state[config.inputsKey][devIndex] = text(input);
  state.excelLinkErrors.delete(userCellKey(scope, devIndex));
  const selection = state[config.selectionKey];
  if (value === null) {
    if (selection[devIndex] === "user") selection[devIndex] = config.defaultMethod;
  } else {
    selection[devIndex] = "user";
  }
}

// One workbook read for every distinct cell the given inputs reference; a cell
// the workbook refused is absent from the map, so evaluation names it.
async function readExcelValuesForInputs(inputs) {
  const items = excelReadItemsForInputs(inputs);
  const values = new Map();
  const failures = new Map();
  if (!items.length) return { values, failures };
  const response = await readExcelCellsBatch(
    items.map(({ book_path, sheet, cell }) => ({ book_path, sheet, cell })),
  );
  if (!response?.ok) throw new Error(text(response?.error) || "Excel could not be read.");
  const results = Array.isArray(response.results) ? response.results : [];
  items.forEach((item, index) => {
    const result = results[index];
    if (result?.ok && numberOrNull(result.value) !== null) values.set(item.key, Number(result.value));
    else failures.set(item.key, text(result?.error) || "The linked cell has no numeric value.");
  });
  return { values, failures };
}

// The workbook's own reason for refusing a cell, named with the cell, so an
// entry that fails on a missing sheet says so rather than "no numeric value".
function refusedReferenceError(input, failures) {
  const reference = userValueExcelReferences(input)
    .find((candidate) => failures.has(excelSourceKey(candidate)));
  if (!reference) return null;
  const reason = failures.get(excelSourceKey(reference));
  return Object.assign(
    new Error(`${reference.filename} ${reference.sheet}!${reference.cell}: ${reason}`),
    { reference, reason },
  );
}

async function applyUserValueEntry(scope, devIndex, rawText) {
  const config = SELECTION_SCOPES[scope];
  if (!config || !Number.isInteger(devIndex) || devIndex < 0) return;
  ensureSelectionConfig();
  const raw = normalizeExcelReferenceAddressCase(text(rawText));
  const entries = expandUserValueRangeEntry(raw, devIndex, matrixDevelopmentCount())
    || [{ column: devIndex, input: raw }];
  const formulas = entries.filter((entry) => isUserValueFormula(entry.input));
  const { values, failures } = formulas.length
    ? await readExcelValuesForInputs(formulas.map((entry) => entry.input))
    : { values: new Map(), failures: new Map() };
  const updates = entries.map((entry) => {
    if (isUserValueFormula(entry.input)) {
      const refused = refusedReferenceError(entry.input, failures);
      if (refused) throw refused;
      return { column: entry.column, input: entry.input, value: evaluateUserValueFormula(entry.input, values) };
    }
    const parsed = parseUserValueNumber(entry.input);
    if (!parsed.ok) throw new Error("Enter a number, a formula, or an Excel cell reference.");
    return { column: entry.column, input: "", value: parsed.value };
  });
  for (const update of updates) setUserValue(scope, update.column, update.value, update.input);
  recalculateAfterSelectionEdit();
  void bsLinksTab?.refresh();
}

function clearUserValue(scope, devIndex) {
  ensureSelectionConfig();
  setUserValue(scope, devIndex, null, "");
  recalculateAfterSelectionEdit();
  setActiveUserCell(scope, devIndex);
  void bsLinksTab?.refresh();
}

const USER_CELL_EDIT_START_KEYS = /^[0-9.\-+='"[]$/u;

function handleUserCellKeydown(event, scope, devIndex) {
  if (event.ctrlKey || event.metaKey || event.altKey) return;
  const config = SELECTION_SCOPES[scope];
  if (event.key === "ArrowLeft" || event.key === "ArrowRight") {
    event.preventDefault();
    const nextIndex = devIndex + (event.key === "ArrowLeft" ? -1 : 1);
    if (nextIndex >= 0 && nextIndex < matrixDevelopmentCount()) setActiveUserCell(scope, nextIndex);
  } else if (event.key === "Enter" || event.key === "F2") {
    event.preventDefault();
    openUserCellEditor(scope, devIndex);
  } else if (event.key === " ") {
    // Space selects the figure the cell already holds, as it does on the
    // estimator rows.
    event.preventDefault();
    if (numberOrNull(state[config.userKey][devIndex]) !== null) applySelectionChoice(scope, "user", devIndex);
    setActiveUserCell(scope, devIndex);
  } else if (event.key === "Delete" || event.key === "Backspace") {
    event.preventDefault();
    clearUserValue(scope, devIndex);
  } else if (event.key === "Escape") {
    event.preventDefault();
    state.activeUserCell = null;
    cellSelection.clearAll();
  } else if (event.key.length === 1 && USER_CELL_EDIT_START_KEYS.test(event.key)) {
    event.preventDefault();
    openUserCellEditor(scope, devIndex, event.key);
  }
}

function handleSelectionGridEvent(event) {
  if (variant !== "cra" || state.currentView !== "avgSelections") return;
  const target = event.target instanceof Element ? event.target : null;
  if (target?.closest("input.bsSelUserInput")) return;
  const pick = target?.closest("[data-select-method]") || null;
  if (!pick) return;
  const scope = text(pick.dataset.selectScope);
  const devRaw = Number.parseInt(pick.dataset.selectDev || "", 10);
  const devIndex = Number.isInteger(devRaw) ? devRaw : null;
  const config = SELECTION_SCOPES[scope];
  if (pick.classList.contains("bsSelUserCell") && devIndex !== null) {
    if (event.type === "keydown") {
      handleUserCellKeydown(event, scope, devIndex);
    } else if (event.type === "dblclick") {
      event.preventDefault();
      openUserCellEditor(scope, devIndex);
    } else if (event.type === "click") {
      // A click activates the cell; a cell that already holds a figure is also
      // selected for its column, as clicking any other estimator's value is.
      if (userCellEditor) return;
      if (numberOrNull(state[config.userKey][devIndex]) !== null
        && norm(state[config.selectionKey][devIndex]) !== "user") {
        applySelectionChoice(scope, "user", devIndex);
      }
      setActiveUserCell(scope, devIndex);
    }
    return;
  }
  if (event.type !== "click" && event.type !== "keydown") return;
  if (event.type === "keydown" && event.key !== "Enter" && event.key !== " ") return;
  event.preventDefault();
  applySelectionChoice(scope, text(pick.dataset.selectMethod), devIndex);
}

// Ctrl+V on an active User Value cell fills the row from that column with the
// first line of the clipboard, one tab-separated entry per column; each entry
// is committed exactly as if it had been typed, so a pasted Excel reference or
// range links the cells it lands on.
function handleUserCellPaste(event) {
  if (!state.activeUserCell || userCellEditor) return;
  if (variant !== "cra" || state.currentView !== "avgSelections") return;
  const active = document.activeElement;
  if (!(active instanceof Element) || !active.classList.contains("bsSelUserCell")) return;
  const clipboard = text(event.clipboardData?.getData("text"));
  if (!clipboard) return;
  event.preventDefault();
  const { scope, devIndex } = state.activeUserCell;
  const entries = clipboard.split(/\r?\n/u)[0].split("\t");
  void (async () => {
    try {
      for (let offset = 0; offset < entries.length; offset += 1) {
        const column = devIndex + offset;
        if (column >= matrixDevelopmentCount()) break;
        await applyUserValueEntry(scope, column, entries[offset]);
      }
      setActiveUserCell(scope, devIndex);
      postStatus(`Pasted ${entries.length === 1 ? "a user value" : `${entries.length} user values`}.`);
    } catch (error) {
      setActiveUserCell(scope, devIndex);
      postStatus(text(error?.message || error), "error");
    }
  })();
}

function userValueGrids() {
  const columnLabels = Array.from({ length: matrixDevelopmentCount() }, (_, index) => getDevelopmentLabel(index));
  return Object.entries(SELECTION_SCOPES).map(([scope, config]) => ({
    scope,
    label: config.label,
    inputs: state[config.inputsKey],
    columnLabels,
  }));
}

function userValueLinkRecords() {
  return variant === "cra" ? buildUserValueLinkRecords(userValueGrids()) : [];
}

// The User Value cells whose formulas read any of the given workbook cells;
// every linked cell when no source is named.
function userValueLinkTargets(sourceIds = null) {
  const wanted = Array.isArray(sourceIds) && sourceIds.length ? new Set(sourceIds) : null;
  const targets = new Map();
  for (const group of collectUserValueExcelLinkGroups(userValueGrids()).values()) {
    if (wanted && !wanted.has(group.id)) continue;
    for (const consumer of group.consumers) {
      const key = userCellKey(consumer.scope, consumer.column);
      if (!targets.has(key)) targets.set(key, { scope: consumer.scope, devIndex: consumer.column });
    }
  }
  return Array.from(targets.values());
}

function userValueFailure(target, reference, error) {
  const config = SELECTION_SCOPES[target.scope];
  return {
    workbookPath: reference?.bookPath || "",
    worksheet: reference?.sheet || "",
    sourceCell: reference?.cell || "",
    destination: `${config.label} / ${getDevelopmentLabel(target.devIndex)}`,
    error: text(error),
  };
}

// Re-reads every linked workbook cell and re-evaluates the formulas that use
// it. A value that moved is entered like a typed one and makes the page dirty;
// a reference the workbook refuses keeps its stored value, turns red, and is
// reported in the shared Excel link message box.
async function refreshUserValueExcelLinks(sourceIds = null) {
  const targets = userValueLinkTargets(sourceIds);
  if (!targets.length) return { ok: true, linkedCellCount: 0, changedCount: 0, failedCount: 0 };
  const inputs = targets.map(({ scope, devIndex }) => state[SELECTION_SCOPES[scope].inputsKey][devIndex]);
  const { values, failures } = await readExcelValuesForInputs(inputs);
  const failed = [];
  let changedCount = 0;
  targets.forEach((target, index) => {
    const config = SELECTION_SCOPES[target.scope];
    const input = inputs[index];
    const key = userCellKey(target.scope, target.devIndex);
    try {
      const refused = refusedReferenceError(input, failures);
      if (refused) throw refused;
      const value = evaluateUserValueFormula(input, values);
      state.excelLinkErrors.delete(key);
      if (state[config.userKey][target.devIndex] !== value) {
        setUserValue(target.scope, target.devIndex, value, input);
        changedCount += 1;
      }
    } catch (error) {
      const reason = text(error?.reason || error?.message || error);
      state.excelLinkErrors.set(key, reason);
      failed.push(userValueFailure(target, error?.reference, reason));
    }
  });
  if (changedCount) recalculateAfterSelectionEdit();
  else renderMethodTable();
  if (failed.length) {
    await showExcelLinkFailureAlert({ failures: failed, valueNoun: "linked user value" });
  } else {
    bsLinksTab?.setFreshness(null);
  }
  return {
    ok: true,
    linkedCellCount: targets.length,
    changedCount,
    failedCount: failed.length,
    message: `${changedCount} of ${targets.length} linked user value${targets.length === 1 ? "" : "s"} changed.`,
  };
}

// Replaces each linked cell's formula with the number it currently shows.
function breakUserValueExcelLinks(sourceIds = null) {
  const targets = userValueLinkTargets(sourceIds);
  for (const { scope, devIndex } of targets) {
    const config = SELECTION_SCOPES[scope];
    setUserValue(scope, devIndex, state[config.userKey][devIndex], "");
  }
  if (targets.length) recalculateAfterSelectionEdit();
  return {
    ok: true,
    affectedCellCount: targets.length,
    message: `${targets.length} linked user value${targets.length === 1 ? "" : "s"} hard-coded.`,
  };
}

// Opening a method checks its saved Excel links once without changing any
// value: a reference the workbook refuses is broken, a readable value that no
// longer matches the stored one is stale, and an unreadable one is unverified.
// The Links tab carries the tally; a broken link also opens the shared alert.
async function checkUserValueExcelLinkFreshness() {
  if (variant !== "cra" || !bsLinksTab) return;
  const targets = userValueLinkTargets();
  if (!targets.length) {
    bsLinksTab.setFreshness(null);
    return;
  }
  const inputs = targets.map(({ scope, devIndex }) => state[SELECTION_SCOPES[scope].inputsKey][devIndex]);
  const items = excelReadItemsForInputs(inputs);
  let response = null;
  try {
    response = await validateExcelLinksBatch(items.map(({ book_path, sheet, cell }) => ({ book_path, sheet, cell })));
  } catch (error) {
    postStatus(`Excel links could not be checked: ${text(error?.message || error)}`, "warn");
    return;
  }
  if (!response?.ok) {
    postStatus(`Excel links could not be checked: ${text(response?.error) || "Excel could not be read."}`, "warn");
    return;
  }
  const results = Array.isArray(response.results) ? response.results : [];
  const values = new Map();
  const failures = new Map();
  items.forEach((item, index) => {
    const result = results[index];
    if (!result?.ok) failures.set(item.key, text(result?.error) || "The linked cell could not be read.");
    else if (numberOrNull(result.value) !== null) values.set(item.key, Number(result.value));
  });
  const invalidLinks = [];
  let staleCount = 0;
  let unverifiedCount = 0;
  targets.forEach((target, index) => {
    const config = SELECTION_SCOPES[target.scope];
    const input = inputs[index];
    const references = userValueExcelReferences(input);
    const refused = references.find((reference) => failures.has(excelSourceKey(reference)));
    if (refused) {
      const error = failures.get(excelSourceKey(refused));
      state.excelLinkErrors.set(userCellKey(target.scope, target.devIndex), error);
      invalidLinks.push(userValueFailure(target, refused, error));
      return;
    }
    if (references.some((reference) => !values.has(excelSourceKey(reference)))) {
      unverifiedCount += 1;
      return;
    }
    let value = null;
    try {
      value = evaluateUserValueFormula(input, values);
    } catch {
      unverifiedCount += 1;
      return;
    }
    const stored = numberOrNull(state[config.userKey][target.devIndex]);
    if (stored === null || Math.abs(stored - value) > 5e-7) staleCount += 1;
  });
  if (invalidLinks.length) renderMethodTable();
  bsLinksTab.setFreshness({ staleCount, unverifiedCount, invalidCount: invalidLinks.length });
  if (invalidLinks.length) {
    await showExcelLinkFailureAlert({ failures: invalidLinks, valueNoun: "linked user value" });
  } else if (staleCount || unverifiedCount) {
    postStatus("Saved Excel values may be out of date. Refresh them from the Links tab.", "warn");
  }
}

function renderMethodTable() {
  if (!els.methodHead || !els.methodBody) return;
  renderMethodGrids();
  if (els.selectLeadingDiagonalItem) {
    els.selectLeadingDiagonalItem.hidden = !(variant === "sr" && state.currentView === "proportionSettled");
  }
  // A re-render replaces every cell, so the selection is painted back onto the
  // new ones from its row and column positions.
  cellSelection.applyDom();
}

function renderMethodGrids() {
  syncMethodChrome();
  els.secondaryHead?.replaceChildren();
  els.secondaryBody?.replaceChildren();
  if (variant === "sr" && state.currentView === "output" && state.result) {
    renderAdjustedPaidGrid();
    return;
  }
  if (variant === "sr" && state.currentView === "proportionSettled" && state.result) {
    renderProportionSettledGrid();
    return;
  }
  if (variant === "cra" && state.currentView === "avgSelections") {
    els.methodHead.replaceChildren();
    els.methodBody.replaceChildren();
    renderSelectionGrids();
    return;
  }
  const matrix = matrixForCurrentView();
  const format = viewNumberFormat();
  const developmentCount = Math.max(
    matrixDevelopmentCount(),
    ...matrix.map((row) => Array.isArray(row) ? row.length : 0),
    0,
  );
  const showUltimateColumn = variant === "sr" && state.currentView === "numbersClosed";

  const headerRow = buildMethodHeaderRow("Accident Year", developmentCount);
  if (showUltimateColumn) {
    const ultimateHeader = document.createElement("th");
    ultimateHeader.scope = "col";
    ultimateHeader.dataset.c = String(developmentCount);
    ultimateHeader.textContent = "Ultimate";
    headerRow.appendChild(ultimateHeader);
  }
  els.methodHead.replaceChildren(headerRow);

  const body = document.createDocumentFragment();
  const rowCount = Math.max(matrixRowCount(), matrix.length);
  const ultimateValues = state.sourceValues.ultimate_claim_numbers || [];
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const rowElement = document.createElement("tr");
    const origin = tagRowLabel(document.createElement("td"), rowIndex);
    origin.textContent = getOriginLabel(rowIndex);
    rowElement.appendChild(origin);
    const populatedCount = populatedRowLength(matrix, rowIndex);
    for (let devIndex = 0; devIndex < developmentCount; devIndex += 1) {
      if (devIndex >= populatedCount) {
        rowElement.appendChild(maskedCell(rowIndex, devIndex));
        continue;
      }
      const cell = document.createElement("td");
      const rawValue = matrix[rowIndex]?.[devIndex];
      tagValueCell(cell, rowIndex, devIndex, rawValue);
      cell.textContent = formatCellValue(rawValue, format);
      cell.title = numberOrNull(rawValue) === null ? "" : String(rawValue);
      rowElement.appendChild(cell);
    }
    if (showUltimateColumn) {
      const ultimateCell = document.createElement("td");
      const rawValue = ultimateValues[rowIndex];
      tagValueCell(ultimateCell, rowIndex, developmentCount, rawValue);
      ultimateCell.textContent = formatCellValue(rawValue, roleNumberFormat("ultimate_claim_numbers"));
      ultimateCell.title = numberOrNull(rawValue) === null ? "" : String(rawValue);
      rowElement.appendChild(ultimateCell);
    }
    body.appendChild(rowElement);
  }
  els.methodBody.replaceChildren(body);
}

function recalculateAfterSelectionEdit() {
  calculateCurrent();
  markDirty();
}

function normalizeCachedRow(row) {
  const rawName = text(row?.datasetName || row?.dataset_name || row?.name || row?.datasetTypeName || row?.dataset_type);
  const name = rawName;
  return {
    ...row,
    name,
    datasetName: name,
    datasetType: text(row?.datasetTypeName || row?.dataset_type || row?.datasetType || name),
    dataFormat: text(row?.dataFormat || row?.data_format || row?.meta?.dataFormat),
    methodType: text(row?.methodType || row?.method_type || row?.meta?.methodType),
    sourceKind: text(row?.sourceKind || row?.source_kind || row?.meta?.sourceKind),
    csvFile: text(row?.csvFile || row?.csv_file || row?.meta?.csvFile || csvBaseName(row?.path)),
    originLength: Number(row?.meta?.originLength ?? row?.originLength ?? row?.origin_length) || 0,
    developmentLength: Number(row?.meta?.developmentLength ?? row?.developmentLength ?? row?.development_length) || 0,
  };
}

async function loadCachedRows(force = false) {
  if (!state.project || !state.reservingClass) return [];
  const sharedPayload = !force && params.get("project_instance") === "1"
    ? readProjectInstanceDatasetSnapshot(state.project, state.reservingClass)
    : null;
  if (sharedPayload) {
    state.cachedRows = sharedPayload.files.map(normalizeCachedRow).filter((row) => row.name);
    return state.cachedRows;
  }
  const query = new URLSearchParams({
    project_name: state.project,
    reserving_class: state.reservingClass,
  });
  if (force) query.set("refresh", "1");
  const response = await fetch(`/datasets/cached?${query.toString()}`, { cache: "no-store" });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || payload?.ok === false) {
    throw new Error(payload?.detail || payload?.error || `Dataset cache failed (${response.status}).`);
  }
  const rows = Array.isArray(payload?.files)
    ? payload.files
    : Array.isArray(payload?.rows)
      ? payload.rows
      : Array.isArray(payload?.items)
        ? payload.items
        : Array.isArray(payload?.datasets)
          ? payload.datasets
          : [];
  state.cachedRows = rows.map(normalizeCachedRow).filter((row) => row.name);
  return state.cachedRows;
}

function cachedRecordByName(name) {
  const key = norm(name);
  return state.cachedRows.find((row) => (
    norm(row.name) === key || norm(row.datasetType) === key
  )) || null;
}

async function loadDatasetPayload(datasetName) {
  const name = text(datasetName);
  if (!state.project || !state.reservingClass || !name) {
    throw new Error("Missing project, reserving class, or dataset name.");
  }
  const record = cachedRecordByName(name);
  if (
    record
    && (
      (record.originLength && record.originLength !== ANNUAL_PERIOD_LENGTH)
      || (record.developmentLength && record.developmentLength !== ANNUAL_PERIOD_LENGTH)
    )
  ) {
    throw new Error(`${name} is not an annual dataset.`);
  }
  const body = {
    project_name: state.project,
    reserving_class: state.reservingClass,
    dataset_name: name,
    origin_length: ANNUAL_PERIOD_LENGTH,
    development_length: ANNUAL_PERIOD_LENGTH,
    cumulative: true,
    calendar: false,
    // A hand-entered source may be held at finer periods than the annual grid
    // it is shown at, and B&S wants the grid: asking for the display shape
    // rolls that file up on the way out, so a source stored monthly reads as
    // the annual triangle the Dataset window shows rather than as its store.
    at_display_shape: true,
  };
  if (record?.csvFile) body.csv_file = record.csvFile;
  const response = await fetch("/dataset/cache/load", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || payload?.ok === false) {
    throw new Error(payload?.detail || payload?.error || `Dataset load failed (${response.status}).`);
  }
  const originLength = Number(payload?.origin_length) || ANNUAL_PERIOD_LENGTH;
  const developmentLength = Number(payload?.development_length) || ANNUAL_PERIOD_LENGTH;
  if (originLength !== ANNUAL_PERIOD_LENGTH || developmentLength !== ANNUAL_PERIOD_LENGTH) {
    throw new Error(`${name} is not an annual dataset.`);
  }
  return payload;
}

function applyPayloadToRole(role, payload) {
  const payloadFormat = text(payload?.data_format || payload?.dataFormat);
  if (norm(payloadFormat) !== norm(role.format)) {
    throw new Error(
      `${role.label} must be an annual ${role.format.toLowerCase()} dataset.`,
    );
  }
  const values = role.format === "Vector"
    ? normalizeVector(payload?.values)
    : normalizeAnnualTriangle(payload?.values, payload?.mask);
  state.sourceValues[role.key] = values;
  state.sourcePayloads[role.key] = payload;
  if (role.format === "Triangle") {
    const originLabels = payload?.origin_labels || payload?.originLabels;
    const developmentLabels = payload?.development_labels || payload?.developmentLabels || payload?.dev_labels;
    if (Array.isArray(originLabels) && originLabels.length === values.length) {
      if (!state.originLabels.length || state.originLabels.length !== values.length) {
        state.originLabels = originLabels.map(String);
      }
    }
    const developmentCount = Math.max(0, ...values.map((row) => row.length));
    if (Array.isArray(developmentLabels) && developmentLabels.length >= developmentCount) {
      if (!state.developmentLabels.length || state.developmentLabels.length < developmentCount) {
        state.developmentLabels = developmentLabels.slice(0, developmentCount).map(String);
      }
    }
  }
}

async function refreshOriginLabels() {
  const expectedCount = matrixRowCount();
  if (!expectedCount || state.originLabels.length === expectedCount) return;
  if (state.sidecarOriginLabels.length === expectedCount) {
    state.originLabels = state.sidecarOriginLabels.slice();
    return;
  }
  try {
    state.originLabels = await ensureDatasetOriginLabels(
      state.project,
      ANNUAL_PERIOD_LENGTH,
      { expectedCount, requireMatchingPeriod: true },
    );
  } catch {
    state.originLabels = [];
  }
}

async function refreshDevelopmentLabels() {
  const expectedCount = matrixDevelopmentCount();
  if (!expectedCount || state.developmentLabels.length >= expectedCount) return;
  const labels = await headersService.ensureDevHeadersForProject(state.project);
  state.developmentLabels = Array.isArray(labels)
    ? labels.slice(0, expectedCount).map(String)
    : [];
}

async function refreshSourceRoles(roles = roleDefinitions(), { refreshCache = false } = {}) {
  if (refreshCache) await loadCachedRows(true).catch(() => {});
  const selectedRoles = roles.filter((role) => text(state.sourceNames[role.key]));
  const loaded = await Promise.all(selectedRoles.map(async (role) => ({
    role,
    payload: await loadDatasetPayload(state.sourceNames[role.key]),
  })));
  for (const item of loaded) applyPayloadToRole(item.role, item.payload);
  reapplyActiveDependencyPreviews();
  await refreshOriginLabels();
  await refreshDevelopmentLabels();
  calculateCurrent();
  await syncRecordedSourceNumberFormats();
}

// A source's number format is display state, not a result input: when the user
// restyles a source dataset the method's recorded copy catches up on its own,
// without making the page dirty, bumping the method's last-modified stamp, or
// touching the output's review status.
async function syncRecordedSourceNumberFormats() {
  let changed = false;
  for (const role of roleDefinitions()) {
    const payload = state.sourcePayloads[role.key];
    if (!payload || !text(payload.number_format)) continue;
    const live = normalizeNumberFormatEntry(payload);
    if (sameNumberFormat(state.numberFormats.sources[role.key], live)) continue;
    state.numberFormats.sources[role.key] = live;
    changed = true;
  }
  if (!changed) return;
  renderMethodTable();
  await rewriteRecordedNumberFormats();
}

async function rewriteRecordedNumberFormats() {
  const details = getDetails();
  if (!state.project || !state.reservingClass || !details.name) return;
  try {
    const existing = await requestMethodAndSidecar(details.name);
    const data = existing?.method;
    if (!existing?.exists || !data || typeof data !== "object") return;
    const methodTab = data.method_tab;
    if (!methodTab || typeof methodTab !== "object") return;
    const record = buildNumberFormatsRecord();
    if (JSON.stringify(methodTab.number_formats ?? null) === JSON.stringify(record)) return;
    // The stored payload is rewritten in place, so `method_metadata.last_modified`
    // and every other saved field stay exactly as the last real save left them.
    await requestMethodSave({
      method: { ...data, method_tab: { ...methodTab, number_formats: record } },
    });
  } catch (error) {
    postStatus(`Could not sync the recorded number formats: ${text(error?.message || error)}`, "warn");
  }
}

function normalizedReservingClassPath(value) {
  return text(value).replace(/\\+/g, "\\").toLowerCase();
}

function dependencyMessageMatchesContext(message = {}) {
  if (text(message.inst) && text(message.inst) === inst) return false;
  const project = text(message.project || message.project_name);
  if (project && norm(project) !== norm(state.project)) return false;
  const reservingClass = text(message.reservingClass || message.reserving_class);
  return !reservingClass
    || normalizedReservingClassPath(reservingClass) === normalizedReservingClassPath(state.reservingClass);
}

function dependencyMessageNames(message = {}) {
  return new Set([
    ...(Array.isArray(message.names) ? message.names : []),
    message.datasetName,
    message.datasetTypeName,
    message.name,
  ].map(norm).filter(Boolean));
}

function dependencyRolesMatchingMessage(message = {}) {
  if (!dependencyMessageMatchesContext(message)) return [];
  const names = dependencyMessageNames(message);
  if (!names.size) return [];
  return roleDefinitions().filter((role) => names.has(norm(state.sourceNames[role.key])));
}

function dependencyValuesForRole(message, role) {
  const matrixValues = Array.isArray(message.matrixValues)
    ? message.matrixValues
    : Array.isArray(message.matrix_values)
      ? message.matrix_values
      : message.values;
  return role.format === "Vector"
    ? normalizeVector(matrixValues)
    : normalizeAnnualTriangle(matrixValues, message.mask);
}

function assertDependencyPreviewCompatible(message, role, values) {
  const dataFormat = text(message.dataFormat || message.data_format);
  if (dataFormat && norm(dataFormat) !== norm(role.format)) {
    throw new Error(`${role.label} preview has the wrong data format.`);
  }
  const originLabels = message.originLabels || message.origin_labels;
  if (
    Array.isArray(originLabels)
    && originLabels.length
    && (
      originLabels.length !== state.originLabels.length
      || originLabels.some((label, index) => text(label) !== text(state.originLabels[index]))
    )
  ) {
    throw new Error(`${role.label} preview uses different origin periods.`);
  }
  if (role.format === "Triangle") {
    const developmentLabels = message.developmentLabels || message.development_labels;
    const developmentCount = Math.max(0, ...values.map((row) => row.length));
    if (
      Array.isArray(developmentLabels)
      && developmentLabels.length
      && state.developmentLabels.slice(0, developmentCount).some(
        (label, index) => text(label) !== text(developmentLabels[index]),
      )
    ) {
      throw new Error(`${role.label} preview uses different development periods.`);
    }
  }
}

function applyDependencyValues(role, values) {
  state.sourceValues[role.key] = role.format === "Vector"
    ? normalizeVector(values)
    : normalizeAnnualTriangle(values);
}

function reapplyActiveDependencyPreviews() {
  for (const [roleKey, preview] of Array.from(activeDependencyPreviews.entries())) {
    const role = roleByKey(roleKey);
    if (!role || norm(state.sourceNames[role.key]) !== norm(preview.sourceName)) {
      activeDependencyPreviews.delete(roleKey);
      continue;
    }
    applyDependencyValues(role, preview.values);
  }
}

function applyDependencySourcePreview(message = {}) {
  const roles = dependencyRolesMatchingMessage(message);
  if (!roles.length) return false;
  let changed = false;
  for (const role of roles) {
    const values = dependencyValuesForRole(message, role);
    if (!values.length) continue;
    assertDependencyPreviewCompatible(message, role, values);
    activeDependencyPreviews.set(role.key, {
      sourceName: state.sourceNames[role.key],
      values,
    });
    applyDependencyValues(role, values);
    changed = true;
  }
  if (changed) {
    calculateCurrent();
    scheduleOutputDependencyPreview("upstream-preview", true);
  }
  return changed;
}

async function clearDependencySourcePreview(message = {}) {
  const roles = dependencyRolesMatchingMessage(message)
    .filter((role) => activeDependencyPreviews.has(role.key));
  if (!roles.length) return false;
  for (const role of roles) activeDependencyPreviews.delete(role.key);
  const reason = norm(message.reason);
  await refreshSourceRoles(roles, { refreshCache: reason === "save" || reason === "clean" });
  if (isDirty || activeDependencyPreviews.size > 0) {
    scheduleOutputDependencyPreview("dirty", true);
  } else {
    clearOutputDependencyPreview("upstream-cleared");
  }
  return true;
}

function syncSourceInputs() {
  for (const role of roleDefinitions()) {
    const input = document.getElementById(role.inputId);
    if (input) input.value = text(state.sourceNames[role.key]);
  }
}

async function openPicker(roleKey, anchorElement) {
  await loadCachedRows().catch(() => {});
  const role = roleKey === "output" ? null : roleByKey(roleKey);
  const format = role?.format || "Triangle";
  await openDatasetNamePicker({
    projectName: state.project,
    initialName: role ? state.sourceNames[role.key] : getDetails().outputType,
    anchorElement: anchorElement instanceof Element ? anchorElement : null,
    title: role ? `Select ${role.label}` : "Select Output Type",
    allowedDataFormats: [format],
    includeCalculated: true,
    emptyMessage: `No annual ${format.toLowerCase()} datasets found.`,
    itemFilter: (item) => {
      if (!role) return true;
      const record = cachedRecordByName(item?.name);
      if (!record || norm(record.dataFormat) !== norm(format)) return false;
      return (!record.originLength || record.originLength === ANNUAL_PERIOD_LENGTH)
        && (!record.developmentLength || record.developmentLength === ANNUAL_PERIOD_LENGTH);
    },
    setStatus: (message) => {
      if (text(message)) postStatus(message, "warn");
    },
    onError: (error) => {
      console.error(error);
      postStatus(`Dataset picker failed: ${text(error?.message || error)}`, "error");
    },
    onSelect: async (name) => {
      const selected = text(name);
      if (!selected) return;
      if (!role) {
        withProgrammatic(() => {
          els.outputTypeInput.value = selected;
          if (!text(els.nameInput.value)) els.nameInput.value = selected;
        });
        syncTitle();
        markDirty();
        return;
      }
      const priorName = state.sourceNames[role.key];
      const hadPriorValues = Object.prototype.hasOwnProperty.call(state.sourceValues, role.key);
      const priorValues = state.sourceValues[role.key];
      const hadPriorPayload = Object.prototype.hasOwnProperty.call(state.sourcePayloads, role.key);
      const priorPayload = state.sourcePayloads[role.key];
      const priorPreview = activeDependencyPreviews.get(role.key);
      state.sourceNames[role.key] = selected;
      activeDependencyPreviews.delete(role.key);
      syncSourceInputs();
      try {
        await refreshSourceRoles([role]);
        postStatus(`${role.label} selected: ${selected}`);
      } catch (error) {
        state.sourceNames[role.key] = priorName;
        if (hadPriorValues) state.sourceValues[role.key] = priorValues;
        else delete state.sourceValues[role.key];
        if (hadPriorPayload) state.sourcePayloads[role.key] = priorPayload;
        else delete state.sourcePayloads[role.key];
        if (priorPreview) activeDependencyPreviews.set(role.key, priorPreview);
        else activeDependencyPreviews.delete(role.key);
        syncSourceInputs();
        calculateCurrent();
        setMethodMessage(text(error?.message || error), "error");
        postStatus(`Source load failed: ${text(error?.message || error)}`, "error");
      }
      markDirty();
    },
  });
}

function getCsvFilename() {
  const name = sanitizeFileNamePart(getDetails().name || contract.methodType, "Dataset");
  return `${name}@${ANNUAL_PERIOD_LENGTH}@${ANNUAL_PERIOD_LENGTH}@cum@dev.csv`;
}

function cloneMatrix(values) {
  return Array.isArray(values)
    ? values.map((row) => Array.isArray(row) ? row.slice() : [])
    : [];
}

// Every source role is present in the record, so a reader never has to guess
// which input a missing entry belonged to.
function buildNumberFormatsRecord() {
  const sources = {};
  for (const role of roleDefinitions()) {
    sources[role.key] = { ...roleNumberFormat(role.key) };
  }
  return { derived: { ...derivedNumberFormat() }, sources };
}

function applyNumberFormatsRecord(value) {
  const record = value && typeof value === "object" ? value : {};
  if (record.derived) derivedNumberFormatSeeded = true;
  state.numberFormats.derived = normalizeNumberFormatEntry(
    record.derived,
    derivedNumberFormat(),
  );
  const sources = record.sources && typeof record.sources === "object" ? record.sources : {};
  for (const role of roleDefinitions()) {
    if (sources[role.key] === undefined) continue;
    state.numberFormats.sources[role.key] = normalizeNumberFormatEntry(sources[role.key]);
  }
}

function buildMethodPayload() {
  const details = getDetails();
  const methodTab = {
    origin_labels: state.originLabels.map(String),
    development_labels: Array.from(
      { length: matrixDevelopmentCount() },
      (_, index) => getDevelopmentLabel(index),
    ),
  };
  for (const role of roleDefinitions()) methodTab[role.key] = text(state.sourceNames[role.key]);
  if (variant === "sr") {
    methodTab.selected_proportion_settled = state.selectedProportionSettled.slice();
    methodTab.selected_proportion_is_default = state.selectedProportionIsDefault.slice();
    methodTab.selected_adjustment = cloneMatrix(state.selectedAdjustment);
    methodTab.loess_span = state.loessSpan;
  } else {
    methodTab.inflation_selection = state.inflationSelection.slice();
    methodTab.user_inflation = state.userInflation.slice();
    // The formula text travels beside the numbers only when a cell holds one,
    // so a method with typed figures alone stays byte-identical to the ResQ
    // migration's, which never writes formulas.
    if (state.userInflationInputs.some(Boolean)) {
      methodTab.user_inflation_inputs = state.userInflationInputs.slice();
    }
    methodTab.average_case_reserve_selection = state.averageCaseReserveSelection.slice();
    methodTab.user_average_case_reserves = state.userAverageCaseReserves.slice();
    if (state.userAverageCaseReserveInputs.some(Boolean)) {
      methodTab.user_average_case_reserve_inputs = state.userAverageCaseReserveInputs.slice();
    }
    methodTab.loess_span = state.loessSpan;
  }
  methodTab.number_formats = buildNumberFormatsRecord();
  return {
    json_format: contract.jsonFormat,
    details_tab: {
      name: details.name,
      method_type: contract.methodType,
      output_type: details.outputType,
      origin_length: ANNUAL_PERIOD_LENGTH,
      development_length: ANNUAL_PERIOD_LENGTH,
    },
    method_tab: methodTab,
    method_metadata: {
      method_type: contract.methodType,
      source_kind: contract.sourceKind,
      last_modified: new Date().toISOString(),
    },
  };
}

function applyMethodPayload(payload) {
  const data = payload && typeof payload === "object" ? payload : {};
  const details = data.details_tab || {};
  const method = data.method_tab || {};
  withProgrammatic(() => {
    els.nameInput.value = text(details.name || els.nameInput.value);
    els.outputTypeInput.value = text(details.output_type || els.outputTypeInput.value);
  });
  for (const role of roleDefinitions()) {
    state.sourceNames[role.key] = text(method[role.key] || state.sourceNames[role.key]);
  }
  state.originLabels = Array.isArray(method.origin_labels) ? method.origin_labels.map(String) : [];
  state.developmentLabels = Array.isArray(method.development_labels)
    ? method.development_labels.map(String)
    : [];
  state.selectedProportionSettled = normalizeNumberVector(method.selected_proportion_settled);
  state.selectedProportionIsDefault = normalizeBooleanVector(method.selected_proportion_is_default);
  state.selectedAdjustment = normalizeSelectionMatrix(method.selected_adjustment);
  state.loessSpan = normalizeLoessSpan(method.loess_span);
  state.inflationSelection = normalizeStringVector(method.inflation_selection);
  state.userInflation = normalizeNumberVector(method.user_inflation);
  state.userInflationInputs = normalizeUserValueInputs(
    method.user_inflation_inputs,
    state.userInflation.length,
  );
  state.averageCaseReserveSelection = normalizeStringVector(method.average_case_reserve_selection);
  state.userAverageCaseReserves = normalizeNumberVector(method.user_average_case_reserves);
  state.userAverageCaseReserveInputs = normalizeUserValueInputs(
    method.user_average_case_reserve_inputs,
    state.userAverageCaseReserves.length,
  );
  state.excelLinkErrors.clear();
  applyNumberFormatsRecord(method.number_formats);
  syncSourceInputs();
  syncTitle();
  syncLoessSpanControls();
  syncNumberFormatControls();
}

// The page open is one registered workspace read. `/berquist-sherman/load`
// returns the method JSON and the output sidecar together, so a Client PC pays
// a single workspace visit — served by the ArcRho Gateway when it is available
// — instead of reading the sidecar and then the method file one after the
// other. Reading the method through the host API could never use that
// transport, because a host-API file read does not enter the app server.
async function requestMethodAndSidecar(methodName) {
  const response = await fetch("/berquist-sherman/load", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      project_name: state.project,
      reserving_class: state.reservingClass,
      method_type: contract.methodType,
      method_name: methodName,
    }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || payload?.ok === false) {
    throw new Error(
      payload?.detail || payload?.error || `${contract.displayLabel} load failed (${response.status}).`,
    );
  }
  return payload;
}

// The save is the write half of the same pairing: `/berquist-sherman/save`
// writes the method JSON, the output CSV, and the output sidecar in one
// Engine-hosted call, so the on-disk text is produced by the app server's
// canonical writer next to the data rather than by the renderer across the
// share. The page still owns the payloads it sends. `sidecar` is omitted by
// the in-place number-format rewrite, which publishes nothing.
async function requestMethodSave({ method, csv_file = null, output_csv = null, sidecar = null }) {
  const response = await fetch("/berquist-sherman/save", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      project_name: state.project,
      reserving_class: state.reservingClass,
      method_type: contract.methodType,
      method_name: getDetails().name,
      method,
      csv_file,
      output_csv,
      sidecar,
    }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || payload?.ok === false) {
    throw new Error(
      payload?.detail || payload?.error || `${contract.displayLabel} save failed (${response.status}).`,
    );
  }
  return payload;
}

// Both halves are optional: a method that has never been saved has neither, and
// the page then opens fresh from its Project Instance arguments. The sidecar is
// applied first and the method second, so a saved Details number format still
// wins over the output dataset's own.
async function openMethodPage() {
  const details = getDetails();
  if (!state.project || !state.reservingClass || !details.name) {
    auditLogView.clear();
    return false;
  }
  const requestSequence = ++sidecarLoadSequence;
  auditLogView.setLoading();
  let payload = null;
  try {
    payload = await requestMethodAndSidecar(details.name);
  } catch (error) {
    if (requestSequence !== sidecarLoadSequence) return false;
    auditLogView.setError(`Could not load the audit log. ${text(error?.message || error)}`);
    postStatus(
      `Could not load ${contract.displayLabel}: ${text(error?.message || error)}`,
      "warn",
    );
    return false;
  }
  if (requestSequence !== sidecarLoadSequence) return false;
  applySidecarPayload(payload?.sidecar);
  if (params.get("fresh") === "1" || !payload?.method) return false;
  applyMethodPayload(payload.method);
  postStatus(`Loaded ${contract.displayLabel}: ${getDetails().name}`);
  return true;
}

function getPrecedentNames() {
  const seen = new Set();
  const names = [];
  for (const role of roleDefinitions()) {
    const name = text(state.sourceNames[role.key]);
    const key = norm(name);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    names.push(name);
  }
  return names;
}

let detailsDependencies = null;

function getDetailsDependencies() {
  if (!detailsDependencies) {
    detailsDependencies = createDetailsDependenciesController({
      precedentsList: "bsPrecedentsList",
      dependentsList: "bsDependentsList",
      getIdentity: () => ({
        projectName: state.project,
        reservingClass: state.reservingClass,
        datasetName: getDetails().name,
      }),
      getBerquistShermanContract,
      instanceId: inst,
      isProjectInstanceHost: window.parent !== window,
      setStatus: (message) => postStatus(message),
    });
  }
  return detailsDependencies;
}

// One place applies a loaded sidecar, whether it arrived with the page open or
// from a later Audit-tab refresh.
function applySidecarPayload(sidecar) {
  state.sidecarOriginLabels = Array.isArray(sidecar?.origin_labels)
    ? sidecar.origin_labels.map(String)
    : [];
  // The output dataset's own format seeds the Details controls once; after
  // that the method JSON and the user's edits own it, so reloading the
  // sidecar for the Audit tab cannot revert an unsaved format change.
  if (!derivedNumberFormatSeeded && text(sidecar?.number_format)) {
    derivedNumberFormatSeeded = true;
    state.numberFormats.derived = normalizeNumberFormatEntry(sidecar);
    syncNumberFormatControls();
  }
  notesController.setValue(text(sidecar?.notes), { markClean: true });
  auditLogView.render(sidecar?.audit_log);
  // The Details graph rows come from the same payload, so they cost no extra read.
  getDetailsDependencies().apply(sidecar);
  return sidecar || null;
}

async function loadSidecar() {
  const requestSequence = ++sidecarLoadSequence;
  const details = getDetails();
  if (!state.project || !state.reservingClass || !details.name) {
    auditLogView.clear();
    getDetailsDependencies().clear();
    return null;
  }
  auditLogView.setLoading();
  try {
    // The registered page-open read carries the sidecar, so an Audit-tab
    // refresh runs on the server host beside the data. Asking for the sidecar
    // on its own instead reaches the workspace from here and opens the sidecar,
    // the project's dataset-type rows, the reserving-class index, and one more
    // file for every precedent and dependent — each its own round trip.
    const payload = await requestMethodAndSidecar(details.name);
    if (requestSequence !== sidecarLoadSequence) return null;
    return applySidecarPayload(payload?.sidecar);
  } catch (error) {
    if (requestSequence !== sidecarLoadSequence) return null;
    auditLogView.setError(`Could not load the audit log. ${text(error?.message || error)}`);
    getDetailsDependencies().clear();
    return null;
  }
}

function buildSidecarSaveBody(csvFile) {
  const details = getDetails();
  return {
    project_name: state.project,
    reserving_class: state.reservingClass,
    dataset_name: details.name,
    dataset_type: details.outputType || details.name,
    instance_name: details.name,
    source_kind: contract.sourceKind,
    method_type: contract.methodType,
    status: 0,
    data_format: "Triangle",
    origin_length: ANNUAL_PERIOD_LENGTH,
    development_length: ANNUAL_PERIOD_LENGTH,
    cumulative: true,
    transposed: false,
    calendar: false,
    number_format: derivedNumberFormat().number_format,
    decimal_places: derivedNumberFormat().decimal_places,
    origin_labels: state.originLabels.map(String),
    csv_file: csvFile,
    notes: notesController.getValue(),
    precedents: getPrecedentNames(),
  };
}

// The saved output sidecar comes back inside the save response — the audit log
// the save just appended to, the graph rows it just rewrote, and the notes it
// just stored. Applying it here is what keeps the save to one visit: loading
// the sidecar back would re-read what this payload already holds.
function applySavedOutputSidecar(payload) {
  sidecarLoadSequence += 1;
  pendingBsPropagationJobId = String(payload?.calculated_updates?.job_id || "").trim();
  applySidecarPayload(payload);
  return payload;
}

function matrixCsv(values) {
  const matrix = normalizeMatrix(values);
  const width = Math.max(matrixDevelopmentCount(), ...matrix.map((row) => row.length), 0);
  return `${matrix.map((row) => (
    Array.from({ length: width }, (_, index) => {
      const number = numberOrNull(row[index]);
      return number === null ? "" : String(number);
    }).join(",")
  )).join("\n")}\n`;
}

function blockSaveForActiveSourcePreviews() {
  if (activeDependencyPreviews.size === 0) return false;
  postStatus(
    `${contract.displayLabel} cannot save while a source has unsaved preview changes. Save or discard the source first.`,
    "error",
  );
  return true;
}

// The window blocks edits behind the shared saving animation for the whole
// round trip: source refresh, method and CSV writes, then the sidecar write
// that queues dependent updates.
const bsSaveProgress = createArcRhoSaveProgress({ subject: contract.displayLabel });

async function saveMethod() {
  return bsSaveProgress.run((progress) => runBerquistShermanSave(progress));
}

async function runBerquistShermanSave(progress) {
  const details = getDetails();
  if (!details.name || !details.outputType) {
    postStatus(`${contract.displayLabel} save requires Name and Output Type.`, "error");
    return { ok: false };
  }
  const missing = roleDefinitions().filter((role) => !text(state.sourceNames[role.key]));
  if (missing.length) {
    postStatus(`${contract.displayLabel} save requires ${missing.map((role) => role.label).join(", ")}.`, "error");
    return { ok: false };
  }
  if (blockSaveForActiveSourcePreviews()) return { ok: false };
  await refreshSourceRoles();
  if (blockSaveForActiveSourcePreviews()) return { ok: false };
  const output = state.result?.output;
  if (!Array.isArray(output) || !output.some((row) => row.some((value) => numberOrNull(value) !== null))) {
    postStatus(`${contract.displayLabel} output is blank. Check the selected sources.`, "error");
    return { ok: false };
  }
  const csvFile = getCsvFilename();
  const sidecarBody = buildSidecarSaveBody(csvFile);
  progress.writing();
  // One Engine-hosted call writes the method JSON, the output CSV, and the
  // output sidecar next to the data, so the on-disk text comes from the one
  // canonical writer (`arcrho_api.io`) and the whole save costs a single visit
  // to the workspace. It comes back with the dependent walk queued, or already
  // run inline on the Engine, which is what `trackSavePropagation` sorts out.
  const sidecar = applySavedOutputSidecar(await requestMethodSave({
    method: buildMethodPayload(),
    csv_file: csvFile,
    output_csv: matrixCsv(output),
    sidecar: sidecarBody,
  }));
  // Only the dataset table is refetched; the saved sidecar was applied above.
  await loadCachedRows(true).catch(() => {});
  markClean();
  try {
    window.parent?.postMessage({
      type: "arcrho:project-instance-refresh-datasets",
      inst,
      savedDatasetName: details.name,
      variant,
    }, "*");
  } catch {}
  postStatus(`${contract.displayLabel} saved: ${details.name}`);
  // Hold the saving card open through the dependent walk so the user sees
  // each live update; a null outcome (failed or stalled walk) keeps the
  // window open and leaves the dataset table as the failure surface.
  const propagationOutcome = await trackSavePropagation(sidecar?.calculated_updates, {
    onStatus: (message, statusOptions) => {
      progress.setMessage?.(message, statusOptions);
      postStatus(message, statusOptions?.tone === "warn" ? "warn" : "");
    },
    onComplete: () => {
      try {
        window.parent?.postMessage({ type: "arcrho:project-instance-refresh-datasets" }, "*");
      } catch {}
    },
  });
  // The save and its dependent walk are done; drop the spinner before the
  // review dialog.
  progress.finish();
  await showMethodSaveReviewWarning(sidecar, {
    instanceId: inst,
    projectName: state.project,
    reservingClass: state.reservingClass,
  });
  // The written paths come back from the save route, which is the only place
  // that resolves them now that the renderer no longer builds workspace paths.
  return {
    ok: true,
    path: text(sidecar?.method_path),
    csvPath: text(sidecar?.output_csv_path),
    propagationClean: propagationOutcome !== null,
    refreshedDatasets: propagationOutcome?.refreshed_datasets || [],
    linkWarnings: propagationOutcome?.link_warnings || [],
  };
}

function requestConfirmedClose() {
  clearOutputDependencyPreview("close");
  requestTabbedPageWindowClose({
    messageType: "arcrho:dataset-close-confirmed",
    inst,
  });
}

async function closeOrConfirm() {
  if (isDirty) {
    const discard = await closeConfirm.confirm({ reason: "close" });
    if (!discard) return;
  }
  requestConfirmedClose();
}

function initTabbedPage() {
  const initialTab = text(params.get("tab") || params.get("initial_tab"));
  tabbedPage = createTabbedPage(document.getElementById("bsTabbedPage"), {
    tabs: TABS,
    cssPrefix: "bs",
    initialTab: ALLOWED_TABS.has(initialTab) ? initialTab : "details",
    onTabChange: (tabId) => {
      if (tabId === "audit") void loadSidecar();
      if (tabId === "links") void bsLinksTab?.refresh();
      try {
        window.parent?.postMessage({
          type: "arcrho:berquist-sherman-tab-changed",
          inst,
          tab: tabId,
          variant,
        }, "*");
      } catch {}
    },
  });
  applyTabbedPageSaveBar(document.getElementById("bsSaveBar"));
}

function wireMethodGridControls() {
  els.methodBody?.addEventListener("click", handleAdjustmentGridEvent);
  els.methodBody?.addEventListener("keydown", handleAdjustmentGridEvent);
  els.methodBody?.addEventListener("click", handleProportionGridEvent);
  els.methodBody?.addEventListener("keydown", handleProportionGridEvent);
  els.methodBody?.addEventListener("change", handleProportionGridEvent);
  for (const body of [els.methodBody, els.secondaryBody]) {
    body?.addEventListener("click", handleSelectionGridEvent);
    body?.addEventListener("dblclick", handleSelectionGridEvent);
    body?.addEventListener("keydown", handleSelectionGridEvent);
  }
  document.addEventListener("paste", handleUserCellPaste);
  els.loessSpanInput?.addEventListener("change", () => applyLoessSpan(els.loessSpanInput.value));
  els.loessSpanInput?.addEventListener("keydown", (event) => {
    if (event.key === "Enter") applyLoessSpan(els.loessSpanInput.value);
  });
  els.loessSpanUp?.addEventListener("click", () => applyLoessSpan(state.loessSpan + 1));
  els.loessSpanDown?.addEventListener("click", () => applyLoessSpan(state.loessSpan - 1));
}

function wireNumberFormatControls() {
  const commitFormat = () => applyDerivedNumberFormat({
    number_format: els.numberFormatInput?.value,
  });
  els.numberFormatInput?.addEventListener("change", commitFormat);
  els.numberFormatInput?.addEventListener("keydown", (event) => {
    if (event.key === "Enter") els.numberFormatInput.blur();
  });
  wireNumberFormatField({
    input: els.numberFormatInput,
    field: els.numberFormatField,
    toggle: els.numberFormatBtn,
    menu: els.numberFormatMenu,
    onApply: (preset) => applyDerivedNumberFormat({ number_format: preset }),
  });
  const stepDecimalPlaces = (step) => applyDerivedNumberFormat(
    { decimal_places: derivedNumberFormat().decimal_places + step },
    { fromDecimalPlaces: true },
  );
  els.decimalPlacesInput?.addEventListener("change", () => applyDerivedNumberFormat(
    { decimal_places: els.decimalPlacesInput.value },
    { fromDecimalPlaces: true },
  ));
  els.decimalPlacesInput?.addEventListener("keydown", (event) => {
    if (event.key === "Enter") els.decimalPlacesInput.blur();
  });
  els.decimalPlacesUp?.addEventListener("click", () => stepDecimalPlaces(1));
  els.decimalPlacesDown?.addEventListener("click", () => stepDecimalPlaces(-1));
}

// Every grid pane owns a scroll wrapper, so each one gets its own activity
// state rather than sharing the first wrapper's.
function wireTableScrollbarActivity() {
  for (const host of document.querySelectorAll(".bsTableWrap")) {
    wireScrollbarActivityFor(host);
  }
}

function wireScrollbarActivityFor(host) {
  let idleTimer = null;
  const syncScrollbarHover = (event) => {
    const rect = host.getBoundingClientRect();
    const verticalScrollbarWidth = Math.max(0, host.offsetWidth - host.clientWidth);
    const horizontalScrollbarHeight = Math.max(0, host.offsetHeight - host.clientHeight);
    const nearVerticalScrollbar = host.scrollHeight > host.clientHeight
      && verticalScrollbarWidth > 0
      && event.clientX >= rect.right - Math.max(verticalScrollbarWidth, 16);
    const nearHorizontalScrollbar = host.scrollWidth > host.clientWidth
      && horizontalScrollbarHeight > 0
      && event.clientY >= rect.bottom - Math.max(horizontalScrollbarHeight, 16);
    host.classList.toggle("isScrollbarHover", nearVerticalScrollbar || nearHorizontalScrollbar);
  };
  host.addEventListener("scroll", () => {
    host.classList.add("isScrolling");
    if (idleTimer) clearTimeout(idleTimer);
    idleTimer = setTimeout(() => host.classList.remove("isScrolling"), 550);
  }, { passive: true });
  host.addEventListener("pointermove", syncScrollbarHover, { passive: true });
  host.addEventListener("pointerleave", () => host.classList.remove("isScrollbarHover"), { passive: true });
}

function wireInputs() {
  for (const input of [els.nameInput, els.outputTypeInput]) {
    input?.addEventListener("input", () => {
      syncTitle();
      markDirty();
    });
  }
  wireMethodGridControls();
  wireNumberFormatControls();
  wireTableScrollbarActivity();
  els.outputTypeBtn?.addEventListener("click", () => void openPicker("output", els.outputTypeBtn));
  for (const button of document.querySelectorAll("button[data-picker-role]")) {
    button.addEventListener("click", () => void openPicker(button.dataset.pickerRole, button));
  }
  els.saveBtn?.addEventListener("click", async () => {
    try {
      const saved = await saveMethod();
      // A save keeps the window open; only Cancel and a confirmed dirty close
      // dismiss it.
      if (saved?.ok && saved?.propagationClean) {
        await showSavedDependentsNotice(saved.refreshedDatasets, { linkWarnings: saved.linkWarnings });
      }
    } catch (error) {
      console.error(error);
      postStatus(`Save failed: ${text(error?.message || error)}`, "error");
    }
  });
  els.cancelBtn?.addEventListener("click", () => void closeOrConfirm());
  window.__arcrho_request_close = () => {
    if (!isDirty) return false;
    if (closeConfirm.isOpen) return true;
    void closeOrConfirm();
    return true;
  };
  window.__arcrho_consume_close_shortcut = window.__arcrho_request_close;
}

function wireMessages() {
  window.addEventListener("pagehide", () => clearOutputDependencyPreview("close"));
  window.addEventListener("message", async (event) => {
    const message = event?.data && typeof event.data === "object" ? event.data : {};
    if (message.type === "arcrho:dataset-save") {
      try {
        const saved = await saveMethod();
        if (saved?.ok && saved?.propagationClean) {
          await showSavedDependentsNotice(saved.refreshedDatasets, { linkWarnings: saved.linkWarnings });
        }
      } catch (error) {
        postStatus(`Save failed: ${text(error?.message || error)}`, "error");
      }
      return;
    }
    if (message.type === "arcrho:dependency-source-preview") {
      try {
        applyDependencySourcePreview(message);
      } catch (error) {
        setMethodMessage(text(error?.message || error), "error");
        postStatus(`Source preview rejected: ${text(error?.message || error)}`, "error");
      }
      return;
    }
    if (message.type === "arcrho:dependency-source-cleared") {
      try {
        await clearDependencySourcePreview(message);
      } catch (error) {
        postStatus(`Source preview reload failed: ${text(error?.message || error)}`, "error");
      }
      return;
    }
    if (message.type === "arcrho:close-active-tab" || message.type === "arcrho:dataset-close-request") {
      await closeOrConfirm();
    }
  });
}

async function init() {
  applyHostFixedDetailsFields({ root: "#bsDetailsPage" });
  syncDetailsLabelWidth({
    root: "#bsDetailsPage",
    labelSelector: ".arDetailsLabel",
  });
  withProgrammatic(() => {
    els.projectInput.value = state.project;
    els.classInput.value = state.reservingClass;
    els.nameInput.value = text(params.get("name") || params.get("dataset"));
    els.methodTypeInput.value = contract.displayLabel;
    els.outputTypeInput.value = text(
      params.get("output_type") || params.get("dataset_type") || params.get("datasetType"),
    );
    els.originLengthInput.value = String(ANNUAL_PERIOD_LENGTH);
    els.developmentLengthInput.value = String(ANNUAL_PERIOD_LENGTH);
  });
  els.srInputs.hidden = variant !== "sr";
  els.craInputs.hidden = variant !== "cra";
  // The section divider is a `+` rule, so the hidden source section has to be
  // resolved before the tab paints or the visible one keeps its leading rule.
  syncDetailsSections("#bsDetailsPage");
  const initialTriangle = text(params.get("input_triangle"));
  if (initialTriangle) state.sourceNames.paid_claims = initialTriangle;
  syncSourceInputs();
  syncTitle();
  initTabbedPage();
  renderViewButtons();
  wireInputs();
  wireMessages();
  bsLinksTab = createBerquistShermanLinksTab({
    container: document.getElementById("bsLinksMount"),
    displayLabel: contract.displayLabel,
    getRecords: userValueLinkRecords,
    onRefresh: refreshUserValueExcelLinks,
    onBreak: breakUserValueExcelLinks,
    onStatus: postStatus,
  });

  // The dataset index and the page open are independent workspace reads, and
  // both must land before the sources can load, so they travel together rather
  // than one after the other.
  const [, loaded] = await Promise.all([
    loadCachedRows().catch((error) => {
      postStatus(`Dataset cache unavailable: ${text(error?.message || error)}`, "warn");
    }),
    openMethodPage(),
  ]);
  try {
    await refreshSourceRoles();
  } catch (error) {
    setMethodMessage(text(error?.message || error), "error");
  }
  if (loaded) markClean();
  else {
    cleanSnapshot = "";
    markDirty();
  }
  void bsLinksTab?.refresh();
  // The saved Excel values stay active either way; the check only reports.
  void checkUserValueExcelLinkFreshness();
}

// The window is held blank until the opening tab is rendered; see
// ui/shared/tabbed_page/initial_tab_paint.js.
void init().finally(() => window.arcrhoRevealPage?.());
