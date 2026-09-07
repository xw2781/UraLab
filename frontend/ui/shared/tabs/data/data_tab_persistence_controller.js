// Owns sidecar, settings, notes, external-link, dirty, save, and close lifecycles.
import { notifyDataTabDurableDatasetState, withDataTabDatasetMutation } from "/ui/shared/tabs/data/data_tab_change_watch_port.js?v=20260806a";
import { buildDatasetSaveStatus } from "/ui/shared/tabs/data/data_tab_propagation_report.js?v=20260830a";
import { createTemporaryDatasetFormat } from "/ui/shared/tabs/data/data_tab_temporary_format.js?v=20260805a";
import { createDatasetDirtyState } from "/ui/shared/tabs/data/data_tab_dirty_state.js?v=20260830a";
import { showExcelLinkFailureAlert } from "/ui/shared/integrations/excel_link_alert.js?v=20260819a";
import { showPageMessageBox } from "/ui/shared/components/message_box/message_box.js?v=20260831a";
import { createArcRhoSaveProgress, showSavedDependentsNotice } from "/ui/shared/components/progress_popup/save_progress.js?v=20260831a";
import { trackSavePropagation } from "/ui/shared/services/dependent_propagation_job.js?v=20260813e";
export function registerDataTabPersistenceController(runtime) {
  const { state, config, instanceId, isProjectInstanceDraft, isReadOnlyDatasetViewer, isTemporaryDatasetView } = runtime;
  if (typeof state.showSubtotal !== "boolean") state.showSubtotal = true;
  const defer = (name) => (...args) => runtime[name](...args);
  const { getResolvedProjectValue, getResolvedReservingClassValue, getDatasetInstanceNameValue, normalizeDatasetInstanceKey, getTriInputs, getProjectInstanceDraftDataFormat, getDatasetDecimalPlacesValue, getDatasetSyncedNumberFormatValue, isDfmDataTabHost, clampDatasetDecimalPlaces, normalizeDatasetNumberFormat, applyDecimalPlacesToDatasetNumberFormat, updateTabbedPageSaveControls, setDatasetRenderNumberFormatSettings, renderTable, notifyDatasetUpdated, getDatasetNumberFormatDefaults, getDataTabLinksController, loadDatasetSidecar, renderDatasetAuditLog, getDatasetAuditLog, normalizeDatasetDependencyEntries, renderDetailFormula, getDatasetTypeFormulaByName, renderDatasetPrecedents, renderDatasetDependents, saveTriInputsToStorage, setDatasetDecimalPlacesValue, setDatasetNumberFormatValue, refreshLenDropdowns, validateDatasetOriginLabels, refreshDatasetInstanceNameConflict, saveDatasetSidecar, saveLastDsId, handleCalculationUpdates, invalidateCachedDatasetInstances, clearDatasetDependencyPreview, requestProjectInstanceDatasetTableRefresh, setStatus, requestTabbedPageWindowClose, isInputDefaultBound, loadWorkflowDefaults, saveDatasetNotes, publishDataTabHostInputs, mountDataTabNotes, ensureHeadersForProject, ensureDevHeadersForProject, scheduleAutoRun, applyGridSelectionFromState, setLenSelectValue, getDataTabCloseConfirm, createDatasetExternalLinksController, createDatasetInternalLinksController, createDatasetFormulaLinksController, resolveDatasetInternalLinks } = new Proxy({}, { get: (_target, name) => defer(name) });
  const normalizeProjectText = defer("normalizeProjectText");
  const renderChart = defer("renderChart");
  const isDatasetReadOnly = defer("isDatasetReadOnly");
  const getDatasetRunDataFormat = defer("getDatasetRunDataFormat");
  const setLenSelectLock = defer("setLenSelectLock");
  const setLenSelectStoredLength = defer("setLenSelectStoredLength");
  const setLenSelectDisplayLength = defer("setLenSelectDisplayLength");
  let notesContextKey = "", notesContextPayload = null, notesDirty = false, lastSavedNotesText = "", datasetNotesController = null, datasetSettingsDirty = false, sidecarContextKey = "", sidecarContextPayload = null, lastSavedDatasetSettings = null, sidecarSyncNonce = 0, datasetExternalLinksLoaded = false, datasetCloseConfirm = null, hostInputsPublished = false;
  let datasetExcelLinkCheckAbortController = null;
  let storedDevelopmentChoice = 0, storedDevelopmentChoiceDisplay = 0;
  // Whether the open dataset's file holds a value. See savedDatasetHoldsNoValue.
  let savedDatasetIsEmpty = true;
  // The lengths a cleared hand-entered dataset was reshaped to, or null while
  // the window still shows the shape of the dataset's own file. See
  // releaseStoredShape.
  let releasedLengths = null;
  const datasetExcelLinkCheckedKeys = new Set();
  const {
    loadTemporaryNumberFormatSettings,
    resolveTemporaryDatasetSettings,
    applyTemporaryNumberFormatDefaults,
    applyTemporaryNumberFormatSettings,
  } = createTemporaryDatasetFormat({
    isTemporaryDatasetView,
    state,
    getDatasetNumberFormatDefaults,
    getCurrentDatasetSettings: (...args) => getCurrentDatasetSettings(...args),
    normalizeDatasetSettings: (...args) => normalizeDatasetSettings(...args),
    buildDatasetSidecarContextPayload: (...args) => buildDatasetSidecarContextPayload(...args),
    hasDatasetSidecarContext: (...args) => hasDatasetSidecarContext(...args),
    getDatasetSyncedNumberFormatValue,
    setDatasetDecimalPlacesValue,
    setDatasetNumberFormatValue,
    renderTable,
    notifyDatasetUpdated,
    applyGridSelectionFromState,
  });
  const {
    normalizeDatasetModeText,
    sourceKindIsReadOnly,
    currentDatasetIsManualTriangleOrVector,
    hasManualInputGridChanges,
    hasUnsavedDatasetChanges,
    isUnsavedProjectInstanceDraft,
    shouldPersistManualInputGridValues,
    hasPendingDatasetSaveWork,
    isDraftGridUnavailable,
  } = createDatasetDirtyState({
    state,
    isProjectInstanceDraft,
    isReadOnlyDatasetViewer,
    isTemporaryDatasetView,
    isDfmDataTabHost,
    getProjectInstanceDraftDataFormat,
    getDatasetInstanceNameValue,
    normalizeDatasetInstanceKey,
    getSavedProjectInstanceDraftName: () => runtime.savedProjectInstanceDraftName,
    getDatasetSidecarSourceKind: () => runtime.currentDatasetSidecarSourceKind,
    getDatasetSidecarDataFormat: () => runtime.currentDatasetSidecarDataFormat,
    getDatasetExternalLinks: () => runtime.datasetExternalLinks,
    getDatasetInternalLinks: () => runtime.datasetInternalLinks,
    getDatasetFormulaLinks: () => runtime.datasetFormulaLinks,
    isSettingsDirty: () => datasetSettingsDirty,
    isNotesDirty: () => notesDirty,
  });
  const linksControllerIsReadOnly = () => (
    isDatasetReadOnly()
    || isDfmDataTabHost()
    || !currentDatasetIsManualTriangleOrVector()
    // A link entered or broken here would be filed against the cells of a
    // grid the saved links were never written against.
    || !datasetDisplayIsAtLinkedShape()
  );
  const linksControllerIsTransposed = () => document.getElementById("transposedChk")?.checked === true;
  const notifyLinksInventoryChanged = () => {
    getDataTabLinksController()?.refresh?.();
    updateDatasetSaveUi();
  };
  // One cell holds at most one link, so the controller that takes a cell over
  // releases it from the other two.
  const linkControllerNames = ["datasetExternalLinks", "datasetInternalLinks", "datasetFormulaLinks"];
  const releaseClaimedCells = (owner) => (cells) => {
    linkControllerNames.forEach((name) => {
      if (name !== owner) runtime[name]?.hardCodeTargetCells(cells);
    });
  };
  const resolveReferences = (references) => resolveDatasetInternalLinks({
    project_name: getResolvedProjectValue(),
    reserving_class: getResolvedReservingClassValue(),
    references,
  });
  runtime.datasetExternalLinks = createDatasetExternalLinksController({
    state,
    isReadOnly: linksControllerIsReadOnly,
    isTransposed: linksControllerIsTransposed,
    isAtLinkedShape: datasetDisplayIsAtLinkedShape,
    onInventoryChanged: notifyLinksInventoryChanged,
    onTargetsClaimed: releaseClaimedCells("datasetExternalLinks"),
  });
  runtime.datasetInternalLinks = createDatasetInternalLinksController({
    state,
    resolveReferences,
    isReadOnly: linksControllerIsReadOnly,
    isTransposed: linksControllerIsTransposed,
    isAtLinkedShape: datasetDisplayIsAtLinkedShape,
    onInventoryChanged: notifyLinksInventoryChanged,
    onTargetsClaimed: releaseClaimedCells("datasetInternalLinks"),
  });
  runtime.datasetFormulaLinks = createDatasetFormulaLinksController({
    state,
    resolveReferences,
    isReadOnly: linksControllerIsReadOnly,
    isTransposed: linksControllerIsTransposed,
    isAtLinkedShape: datasetDisplayIsAtLinkedShape,
    onInventoryChanged: notifyLinksInventoryChanged,
    onTargetsClaimed: releaseClaimedCells("datasetFormulaLinks"),
  });
  const forEachLinksController = (apply) => linkControllerNames.forEach((name) => apply(runtime[name]));
  function buildDatasetSidecarContextPayload() {
    return {
      project_name: getResolvedProjectValue(),
      reserving_class: getResolvedReservingClassValue(),
      dataset_name: getDatasetInstanceNameValue(),
      dataset_type: (document.getElementById("triInput")?.value || "").trim(),
      instance_name: getDatasetInstanceNameValue(),
    };
  }

  function hasDatasetSidecarContext(payload) {
    return !!(
      String(payload?.project_name || "").trim()
      && String(payload?.reserving_class || "").trim()
      && String(payload?.dataset_name || "").trim()
    );
  }

  function buildDatasetSidecarContextKey(payload) {
    if (!hasDatasetSidecarContext(payload)) return "";
    return `${payload.project_name}\u001f${payload.reserving_class}\u001f${payload.dataset_type || ""}\u001f${payload.dataset_name}`;
  }

  function getCurrentDatasetSettings() {
    const triInputs = getTriInputs();
    return {
      dataset_type: triInputs.tri,
      instance_name: triInputs.instanceName || triInputs.tri,
      data_format: isProjectInstanceDraft ? getProjectInstanceDraftDataFormat() : undefined,
      origin_length: triInputs.originLen,
      development_length: triInputs.devLen,
      cumulative: !!triInputs.cumulative,
      transposed: !!triInputs.transposed,
      calendar: !!triInputs.calendar,
      show_subtotal: state.showSubtotal !== false,
      decimal_places: getDatasetDecimalPlacesValue(),
      number_format: getDatasetSyncedNumberFormatValue(),
    };
  }

  function getManualInputDatasetValuePayload() {
    if (!shouldPersistManualInputGridValues() || !state.model) return {};
    const values = Array.isArray(state.model.values)
      ? state.model.values.map((row) => (
        Array.isArray(row)
          ? row.map((value) => {
            if (value == null || value === "") return null;
            const numeric = Number(value);
            return Number.isFinite(numeric) ? numeric : null;
          })
          : []
      ))
      : null;
    const mask = Array.isArray(state.model.mask)
      ? state.model.mask.map((row) => (Array.isArray(row) ? row.map(Boolean) : []))
      : null;
    if (!Array.isArray(values) || !values.length) return {};
    return {
      source_kind: "input",
      data_format: runtime.currentDatasetSidecarDataFormat || state.model?.data_format || getProjectInstanceDraftDataFormat(),
      origin_labels: Array.isArray(state.model.origin_labels) ? state.model.origin_labels.map(String) : undefined,
      values,
      mask,
    };
  }

  function getDatasetExternalLinksPayload() {
    if (
      isDfmDataTabHost()
      || !datasetExternalLinksLoaded
      || !currentDatasetIsManualTriangleOrVector()
    ) return {};
    return { external_links: runtime.datasetExternalLinks.serialize() };
  }

  function getDatasetInternalLinksPayload() {
    if (
      isDfmDataTabHost()
      || !datasetExternalLinksLoaded
      || !currentDatasetIsManualTriangleOrVector()
    ) return {};
    return { internal_links: runtime.datasetInternalLinks.serialize() };
  }

  function getDatasetFormulaLinksPayload() {
    if (
      isDfmDataTabHost()
      || !datasetExternalLinksLoaded
      || !currentDatasetIsManualTriangleOrVector()
    ) return {};
    return { formula_links: runtime.datasetFormulaLinks.serialize() };
  }

  function normalizeDatasetSettings(source = {}) {
    const origin = Number(source.origin_length ?? source.originLen);
    const development = Number(source.development_length ?? source.devLen);
    const numberFormat = source.number_format ?? source.numberFormat ?? source.num_format;
    const decimalPlaces = source.decimal_places ?? source.decimalPlaces;
    const normalizedDecimalPlaces = clampDatasetDecimalPlaces(decimalPlaces);
    return {
      dataset_type: String(source.dataset_type ?? source.datasetType ?? source.tri ?? "").trim(),
      instance_name: String(source.instance_name ?? source.instanceName ?? source.dataset_name ?? source.datasetName ?? "").trim(),
      origin_length: Number.isFinite(origin) && origin > 0 ? Math.trunc(origin) : 12,
      development_length: Number.isFinite(development) && development > 0 ? Math.trunc(development) : 12,
      cumulative: typeof source.cumulative === "boolean" ? source.cumulative : true,
      transposed: typeof source.transposed === "boolean" ? source.transposed : false,
      calendar: typeof source.calendar === "boolean" ? source.calendar : false,
      show_subtotal: typeof source.show_subtotal === "boolean" ? source.show_subtotal : true,
      decimal_places: normalizedDecimalPlaces,
      number_format: applyDecimalPlacesToDatasetNumberFormat(
        normalizeDatasetNumberFormat(numberFormat),
        normalizedDecimalPlaces,
      ),
    };
  }

  function sameDatasetSettings(a, b) {
    const left = normalizeDatasetSettings(a || {});
    const right = normalizeDatasetSettings(b || {});
    return (
      left.origin_length === right.origin_length
      && left.development_length === right.development_length
      && left.cumulative === right.cumulative
      && left.transposed === right.transposed
      && left.calendar === right.calendar
      && left.show_subtotal === right.show_subtotal
      && left.decimal_places === right.decimal_places
      && left.number_format === right.number_format
      && normalizeProjectText(left.dataset_type) === normalizeProjectText(right.dataset_type)
      && normalizeProjectText(left.instance_name) === normalizeProjectText(right.instance_name)
    );
  }

  function datasetValuesAreAllZero() {
    const values = Array.isArray(state.model?.values) ? state.model.values : [];
    const mask = Array.isArray(state.model?.mask) ? state.model.mask : [];
    for (let r = 0; r < values.length; r += 1) {
      const row = Array.isArray(values[r]) ? values[r] : [];
      for (let c = 0; c < row.length; c += 1) {
        if (Array.isArray(mask[r]) && mask[r][c] === false) continue;
        const raw = row[c];
        if (raw == null || raw === "") continue;
        const value = Number(raw);
        if (!Number.isFinite(value) || Math.abs(value) > 1e-12) return false;
      }
    }
    return true;
  }

  // Whether the dataset's own file holds a value, which is a different question
  // from whether the grid on screen does. An edit writes straight into
  // `state.model.values`, so once anything is dirty the grid can no longer be
  // asked; the answer taken while it still matched the file stands until the
  // next load or save replaces it.
  function savedDatasetHoldsNoValue() {
    if (!(state.dirty?.size > 0)) savedDatasetIsEmpty = datasetValuesAreAllZero();
    return savedDatasetIsEmpty;
  }

  // The period the open dataset's own file is held at, as the sidecar records
  // it. Zero means it is not known yet, which is the state before a sidecar has
  // loaded and for a draft that has never been saved.
  // Both the sidecar load and the sidecar save answer with the stored pair, so
  // one reader keeps the window's copy of it in step with either.
  function applyStoredLengthsFromResponse(payload) {
    const source = payload && typeof payload === "object" ? payload : {};
    runtime.currentDatasetStoredOriginLength = Number(source.stored_origin_length) || 0;
    runtime.currentDatasetStoredDevelopmentLength = Number(source.stored_development_length) || 0;
    runtime.currentDatasetLinkedOriginLength = Number(source.linked_origin_length) || 0;
    runtime.currentDatasetLinkedDevelopmentLength = Number(source.linked_development_length) || 0;
    // Whatever the sidecar now says is the answer, so any store the user had
    // asked for and not yet saved is spent, and so is a clear-and-reshape.
    storedDevelopmentChoice = 0;
    storedDevelopmentChoiceDisplay = 0;
    releasedLengths = null;
    // A save has just written the grid to the file, so the grid answers for it
    // again even though its cells are still marked as edits.
    savedDatasetIsEmpty = datasetValuesAreAllZero();
  }

  function getStoredLengthPair() {
    const origin = Number(runtime.currentDatasetStoredOriginLength);
    const development = Number(runtime.currentDatasetStoredDevelopmentLength);
    return {
      origin_length: Number.isFinite(origin) && origin > 0 ? Math.trunc(origin) : 0,
      development_length: Number.isFinite(development) && development > 0 ? Math.trunc(development) : 0,
    };
  }

  // A hand-entered dataset that still holds nothing has no stored period worth
  // keeping: the next save fixes it at whatever the length controls read, so
  // the whole ladder stays open and the readout says so. A dataset cleared and
  // then reshaped stays in that state once values are entered again, because
  // the shape those values sit at is the one the next save stores.
  // ResQ fixes the store on the first value a triangle is *saved* with, not on
  // the first value typed into it, so values waiting to be saved -- a pasted
  // 10x10 over a store the user has just lowered -- leave the choice standing
  // and travel to the server with it.
  function storedLengthIsPending() {
    return currentDatasetIsManualTriangleOrVector()
      && (savedDatasetHoldsNoValue() || releasedLengths !== null);
  }

  // Once every value of a hand-entered dataset is 0 and a length control
  // moves, the file's stored period no longer binds the window: the grid is
  // rebuilt empty at the new lengths instead of reloaded from the file, and
  // the save says the old values were cleared, so the server writes a new
  // file at the shape the entered values have.
  function releaseStoredShape() {
    releasedLengths = getCurrentLengthControlValues();
  }

  // ResQ lets an empty triangle be stored finer than it is shown, and moves the
  // store with the display until a value is saved. The period the user asked
  // for is remembered against the display length it was asked for at, so a
  // later display change resets the store to the new display the way ResQ does.
  // The two values themselves are declared with the rest of this controller's
  // state, because the sidecar reader above clears them.
  function chooseStoredDevelopmentLength(value) {
    const requested = Number(value);
    storedDevelopmentChoice = Number.isFinite(requested) && requested > 0 ? Math.trunc(requested) : 0;
    storedDevelopmentChoiceDisplay = getCurrentLengthControlValues().development_length;
  }

  function getStoredDevelopmentLengthChoice() {
    const display = getCurrentLengthControlValues().development_length;
    if (
      storedDevelopmentChoice > 0
      && storedDevelopmentChoiceDisplay === display
      && display % storedDevelopmentChoice === 0
    ) return storedDevelopmentChoice;
    // The sidecar's own store still stands while the display is the one it was
    // saved at; past that the store follows the display.
    const recorded = getStoredLengthPair().development_length;
    const savedDisplay = Number(lastSavedDatasetSettings?.development_length) || 0;
    if (recorded > 0 && savedDisplay === display && display % recorded === 0) return recorded;
    return display;
  }

  // The shape the file will be held at once this save lands: the sidecar's own
  // stored pair, except while a hand-entered dataset is still empty, when the
  // length controls and the development `Stored at` decide it.
  function getStoredLengthControlPair() {
    if (!storedLengthIsPending()) return getStoredLengthPair();
    return {
      origin_length: getCurrentLengthControlValues().origin_length,
      development_length: getStoredDevelopmentLengthChoice(),
    };
  }

  // The one case a save states the store is a still-empty hand-entered
  // triangle; everywhere else the sidecar keeps the period it already records,
  // and a vector's store follows its own length control as ResQ's does.
  function storedDevelopmentLengthForSave() {
    if (isDfmDataTabHost() || !storedLengthIsPending() || currentDatasetIsVector()) return 0;
    return getStoredDevelopmentLengthChoice();
  }

  // Asking for a finer store is a change Save has to carry even when nothing
  // else on the tab moved.
  function storedDevelopmentLengthIsDirty() {
    const requested = storedDevelopmentLengthForSave();
    if (!requested) return false;
    const recorded = getStoredLengthPair().development_length;
    return recorded > 0 && requested !== recorded;
  }

  function currentDatasetIsVector() {
    return normalizeDatasetModeText(getDatasetRunDataFormat()) === "vector";
  }

  function getManualDatasetLengthBaseline() {
    const stored = getStoredLengthPair();
    if (stored.origin_length > 0 && stored.development_length > 0) return stored;
    const settings = lastSavedDatasetSettings;
    if (!settings) {
      return {
        origin_length: 12,
        development_length: 12,
      };
    }
    return {
      origin_length: Number(settings.origin_length) || 12,
      development_length: Number(settings.development_length) || 12,
    };
  }

  function getCurrentLengthControlValues() {
    const origin = Number.parseInt(document.getElementById("originLenSelect")?.value || "", 10);
    const dev = Number.parseInt(document.getElementById("devLenSelect")?.value || "", 10);
    return {
      origin_length: Number.isFinite(origin) ? origin : 12,
      development_length: Number.isFinite(dev) ? dev : 12,
    };
  }

  // The two axes part company on a coarser display. A coarse origin row is the
  // calendar diagonal of several finer rows and has no single cell to write
  // back to, so ResQ refuses it and so does ArcRho. A coarse development column
  // does have one: the stored cell at that column's own age, which is where
  // ResQ puts the value, so the grid stays editable there.
  function datasetOriginDisplayIsCoarserThanStored() {
    if (storedLengthIsPending()) return false;
    const stored = getStoredLengthPair();
    const current = getCurrentLengthControlValues();
    return stored.origin_length > 0 && current.origin_length > stored.origin_length;
  }

  // Read against the period the next save will write at, not the one the
  // sidecar records, so a dataset that is still empty and has just been told to
  // store its figures finer than it shows them says so before that first save.
  function datasetDevelopmentDisplayIsCoarserThanStored() {
    if (currentDatasetIsVector()) return false;
    const stored = getStoredLengthControlPair();
    const current = getCurrentLengthControlValues();
    return stored.development_length > 0 && current.development_length > stored.development_length;
  }

  // The lengths the open dataset's links were written at, as the sidecar
  // records them. A link names a cell of the grid that was on screen when it
  // was written, so that grid -- not the period the file is held at, nor the
  // display the dataset has since been saved at -- is the one it still points
  // into. A triangle kept monthly under a yearly display carries yearly links,
  // and they stay live at the yearly view. Links entered here and not yet
  // saved belong to the display the sidecar was last loaded or saved with.
  function datasetLinkedDisplayLengths() {
    const linked = Number(runtime.currentDatasetLinkedOriginLength);
    const settings = linked > 0
      ? {
        origin_length: linked,
        development_length: Number(runtime.currentDatasetLinkedDevelopmentLength),
      }
      : lastSavedDatasetSettings;
    const origin = Number(settings?.origin_length);
    const development = Number(settings?.development_length);
    if (!Number.isFinite(origin) || origin <= 0) return null;
    return {
      origin_length: Math.trunc(origin),
      development_length: Number.isFinite(development) && development > 0 ? Math.trunc(development) : 0,
    };
  }

  // Only at that display does a saved link name a square the grid has: every
  // cell of a different view stands for other cells entirely, so a link read,
  // checked, painted, or refreshed there would land on the wrong one. The whole
  // link inventory therefore stands still until the lengths come back, while
  // the display itself may move and be saved: the sidecar keeps the linked
  // lengths apart from the display ones. A dataset that holds nothing, or no
  // link, has nothing worth protecting and is never held still.
  function datasetDisplayIsAtLinkedShape() {
    if (storedLengthIsPending()) return true;
    if (!linkControllerNames.some((name) => runtime[name].hasLinks())) return true;
    const linked = datasetLinkedDisplayLengths();
    if (!linked) return true;
    const current = getCurrentLengthControlValues();
    if (current.origin_length !== linked.origin_length) return false;
    // A vector has one length and no development control to compare.
    if (currentDatasetIsVector()) return true;
    return !linked.development_length || current.development_length === linked.development_length;
  }

  // One sentence for a linked dataset being viewed at other lengths than its
  // links were written at: which length to put back, named the way the control
  // beside it is labelled.
  function datasetOffLinkedShapeLinkHint() {
    if (datasetDisplayIsAtLinkedShape()) return "";
    const linked = datasetLinkedDisplayLengths();
    if (!linked) return "";
    if (currentDatasetIsVector()) {
      return `This dataset's cells are linked. Set the length to ${linked.origin_length} to view or edit the formula.`;
    }
    const current = getCurrentLengthControlValues();
    const lengths = [];
    if (current.origin_length !== linked.origin_length) lengths.push(`Origin Length to ${linked.origin_length}`);
    if (linked.development_length && current.development_length !== linked.development_length) {
      lengths.push(`Development Length to ${linked.development_length}`);
    }
    if (!lengths.length) return "";
    return `This dataset's cells are linked. Set ${lengths.join(" and ")} to view or edit the formula.`;
  }

  function datasetCoarserViewMessage() {
    const stored = getStoredLengthPair();
    if (currentDatasetIsVector()) {
      return `Values can be entered only at the stored period (Period ${stored.origin_length}). Set the length back to edit.`;
    }
    return `Values can be entered only at the stored origin period (Origin ${stored.origin_length}). Set the origin length back to edit.`;
  }

  // Editing a coarse development view rewrites the whole stored triangle, so
  // the status line says so in one sentence for as long as that view is up.
  function datasetCoarseDevelopmentNote() {
    if (!datasetDevelopmentDisplayIsCoarserThanStored()) return "";
    const stored = getStoredLengthControlPair();
    return `Saving here writes each value into the stored period (Development ${stored.development_length}) at its own column age and clears the stored periods between.`;
  }

  // The finest period a length control may offer. A dataset whose file still
  // holds nothing has the whole ladder, except that values entered and not yet
  // saved are held at the shape they were entered at until they are -- which is
  // what validateManualDatasetLengthChange refuses -- so the ladder must not
  // offer a length that refusal would bounce.
  function manualDatasetLadderFloor() {
    if (!storedLengthIsPending()) return getStoredLengthPair();
    if (datasetValuesAreAllZero()) return { origin_length: 0, development_length: 0 };
    return releasedLengths || getManualDatasetLengthBaseline();
  }

  function applyStoredLengthChoices() {
    const stored = manualDatasetLadderFloor();
    setLenSelectStoredLength("originLenSelect", stored.origin_length);
    setLenSelectStoredLength("devLenSelect", stored.development_length);
  }

  const STORED_ORIGIN_LOCK_REASON = "The origin period is fixed by Origin Length while the dataset is empty.";
  const STORED_DEVELOPMENT_LOCK_REASON = "Stored at can be changed only while the dataset is empty.";
  const VECTOR_NO_DEVELOPMENT_REASON = "A vector has no development periods.";

  // A `Stored at` control sits beside its length and offers the periods that
  // divide it. It is dimmed rather than removed when it cannot be changed, so
  // the strip keeps its shape and the period the file is held at is always on
  // screen.
  function applyStoredLenControl(selectId, { displayLength, value, enabled, reason, displayValue = "" }) {
    setLenSelectDisplayLength(selectId, displayLength);
    const shown = String(Number(value) > 0 ? Math.trunc(Number(value)) : displayLength);
    setLenSelectValue(selectId, shown);
    const select = document.getElementById(selectId);
    if (select) {
      select.disabled = !enabled;
      select.title = enabled ? "" : reason;
    }
    setLenSelectLock(selectId, { locked: !enabled, displayValue: displayValue || shown, reason });
  }

  // The two `Stored at` controls carry the period the file is really held at.
  // The origin one is never editable: as in ResQ, the Origin Length control
  // fixes the origin store while the dataset is empty. The development one is
  // live only while the dataset holds no value, which is the only time ResQ
  // allows the store to move. The DFM Data tab has neither control: it does
  // not save the dataset sidecar.
  function updateStoredLengthControls() {
    if (isDfmDataTabHost()) return;
    const pending = storedLengthIsPending();
    const vector = currentDatasetIsVector();
    const display = getCurrentLengthControlValues();
    const stored = getStoredLengthControlPair();

    applyStoredLenControl("originStoredLenSelect", {
      displayLength: display.origin_length,
      value: stored.origin_length,
      enabled: false,
      reason: STORED_ORIGIN_LOCK_REASON,
    });
    applyStoredLenControl("devStoredLenSelect", {
      displayLength: display.development_length,
      value: stored.development_length,
      enabled: pending && !vector,
      reason: vector ? VECTOR_NO_DEVELOPMENT_REASON : STORED_DEVELOPMENT_LOCK_REASON,
      // A vector has no development dimension, so its store reads 0 beside the
      // 0 its Development Length already shows.
      displayValue: vector ? "0" : "",
    });
  }

  function validateManualDatasetLengthChange() {
    if (!currentDatasetIsManualTriangleOrVector()) return true;
    if (datasetValuesAreAllZero()) return true;
    const current = getCurrentLengthControlValues();
    if (releasedLengths) {
      // Values entered since the clear have no file yet, so there is nothing
      // to show them at another period from until they are saved.
      if (current.origin_length === releasedLengths.origin_length && current.development_length === releasedLengths.development_length) {
        return true;
      }
      restoreLengthControls(releasedLengths);
      setStatus(`The values entered since the data was cleared are not saved yet, so the lengths stay at Origin ${releasedLengths.origin_length}, Development ${releasedLengths.development_length}. Save them, or set all values to 0, before changing the lengths.`);
      return false;
    }
    const baseline = getManualDatasetLengthBaseline();
    if (current.origin_length >= baseline.origin_length && current.development_length >= baseline.development_length) {
      return true;
    }
    restoreLengthControls(baseline);
    setStatus(`Manual input datasets with non-zero values cannot be shown below the period their values are stored at (Origin ${baseline.origin_length}, Development ${baseline.development_length}). Set all values to 0 before changing to a lower level.`);
    return false;
  }

  function restoreLengthControls(lengths) {
    setLenSelectValue("originLenSelect", String(lengths.origin_length));
    setLenSelectValue("devLenSelect", String(lengths.development_length));
    refreshLenDropdowns();
  }

  function updateManualDatasetModeControls() {
    const locked = currentDatasetIsManualTriangleOrVector();
    const message = "Manual input Triangle/Vector datasets keep their cumulative and development/calendar mode fixed.";
    const cumulativeChk = document.getElementById("cumulativeChk");
    if (cumulativeChk) {
      cumulativeChk.disabled = locked;
      cumulativeChk.title = locked ? message : "";
    }
    document.querySelectorAll('input[name="timeMode"]').forEach((input) => {
      input.disabled = locked;
      input.title = locked ? message : "";
    });
  }

  // A vector has one column of values and no development dimension, so its
  // Development Length is shown as a fixed 0 the user cannot open. The stored
  // length is deliberately left alone: the vector request sends only
  // PeriodLength, so nothing reads it, and rewriting it would make an untouched
  // dataset look edited.
  function updateVectorDevelopmentLengthControl() {
    setLenSelectLock("devLenSelect", {
      locked: currentDatasetIsVector(),
      displayValue: "0",
      reason: VECTOR_NO_DEVELOPMENT_REASON,
    });
  }

  function restoreManualDatasetModeControls() {
    const settings = normalizeDatasetSettings(lastSavedDatasetSettings || getCurrentDatasetSettings());
    const cumulativeChk = document.getElementById("cumulativeChk");
    if (cumulativeChk) cumulativeChk.checked = settings.cumulative;
    const mode = settings.calendar ? "calendar" : "development";
    const modeInput = document.querySelector(`input[name="timeMode"][value="${mode}"]`);
    if (modeInput) modeInput.checked = true;
    updateManualDatasetModeControls();
  }

  function notifyDatasetDirtyState() {
    const dirty = hasUnsavedDatasetChanges();
    try {
      window.parent?.postMessage({
        type: "arcrho:dataset-dirty",
        inst: instanceId,
        dirty,
      }, "*");
    } catch {}
  }

  function updateDatasetSaveUi() {
    const bar = document.getElementById("datasetSaveBar");
    const saveBtn = document.getElementById("datasetSaveBtn");
    const cancelBtn = document.getElementById("datasetCancelBtn");
    const runBtn = document.getElementById("runArcRhoTriBtn");
    const clearBtn = document.getElementById("clearCacheReloadBtn");
    const hasContext = hasDatasetSidecarContext(sidecarContextPayload) || hasNotesContext(notesContextPayload);
    const dirty = hasPendingDatasetSaveWork();
    if (bar) bar.hidden = !hasContext || isTemporaryDatasetView;
    updateTabbedPageSaveControls({
      saveButton: saveBtn,
      cancelButton: cancelBtn,
      dirty,
      saving: runtime.datasetSaveInFlight,
      saveBlocked: isTemporaryDatasetView || runtime.datasetInstanceNameConflict || !hasContext || isDraftGridUnavailable(),
      cancelBlocked: isTemporaryDatasetView || !hasContext,
    });
    for (const button of [runBtn, clearBtn]) {
      if (!button) continue;
      if (runtime.datasetInstanceNameConflict) {
        if (button.dataset.duplicateNameBlocked !== "1") {
          button.dataset.originalTitle = button.title || "";
        }
        button.dataset.duplicateNameBlocked = "1";
        button.disabled = true;
        button.title = runtime.datasetInstanceNameConflictMessage || "Dataset instance name already exists.";
      } else if (button.dataset.duplicateNameBlocked === "1") {
        button.disabled = false;
        button.title = button.dataset.originalTitle || "";
        delete button.dataset.duplicateNameBlocked;
      }
    }
    updateManualDatasetModeControls();
    updateVectorDevelopmentLengthControl();
    applyStoredLengthChoices();
    updateStoredLengthControls();
    notifyDatasetDirtyState();
  }

  function refreshDatasetSettingsDirty() {
    if (isTemporaryDatasetView) {
      datasetSettingsDirty = false;
      updateDatasetSaveUi();
      return;
    }
    if (isDfmDataTabHost()) {
      datasetSettingsDirty = false;
      updateDatasetSaveUi();
      return;
    }
    datasetSettingsDirty = !!lastSavedDatasetSettings
      && (
        !sameDatasetSettings(getCurrentDatasetSettings(), lastSavedDatasetSettings)
        || storedDevelopmentLengthIsDirty()
      );
    updateDatasetSaveUi();
  }

  function applyDatasetSettingsToControls(settings = {}) {
    const normalized = normalizeDatasetSettings(settings);
    // The offered lengths follow the stored period, so they have to be in place
    // before the saved display shape is written into the control.
    applyStoredLengthChoices();
    setLenSelectValue("originLenSelect", String(normalized.origin_length));
    setLenSelectValue("devLenSelect", String(normalized.development_length));
    const cumulativeChk = document.getElementById("cumulativeChk");
    if (cumulativeChk) cumulativeChk.checked = normalized.cumulative;
    const transposedChk = document.getElementById("transposedChk");
    if (transposedChk) transposedChk.checked = normalized.transposed;
    state.showSubtotal = normalized.show_subtotal;
    const mode = normalized.calendar ? "calendar" : "development";
    const modeInput = document.querySelector(`input[name="timeMode"][value="${mode}"]`);
    if (modeInput) modeInput.checked = true;
    setDatasetDecimalPlacesValue(normalized.decimal_places);
    setDatasetNumberFormatValue(normalized.number_format);
    refreshLenDropdowns();
    updateVectorDevelopmentLengthControl();
    updateStoredLengthControls();
  }

  function invalidateDatasetContextLoads() {
    sidecarSyncNonce += 1;
    datasetExcelLinkCheckAbortController?.abort();
    datasetExcelLinkCheckAbortController = null;
    forEachLinksController((controller) => controller.abort());
  }

  async function reportDatasetExcelLinkFailures(failures, options = {}) {
    const isCurrent = typeof options?.isCurrent === "function" ? options.isCurrent : () => true;
    // Repaint first: the alert names the broken references, and the grid behind
    // it must already show those cells red, because the red marking is what
    // survives the dismissal and points at the cells still to fix.
    renderTable();
    applyGridSelectionFromState();
    getDataTabLinksController()?.refresh?.();
    if (!isCurrent()) return;
    await showExcelLinkFailureAlert({
      failures,
      unnamedCount: options?.unnamedCount ?? 0,
      reason: options?.reason ?? "",
      valueNoun: "linked dataset cell",
    });
  }

  function scheduleDatasetExcelLinkCheck({ contextKey, isCurrent }) {
    if (
      !contextKey
      || datasetExcelLinkCheckedKeys.has(contextKey)
      || !datasetExternalLinksLoaded
    ) return;
    window.setTimeout(async () => {
      if (!isCurrent()) return;
      // A view at other lengths holds none of the cells the saved links name,
      // so every one of them would report itself broken. The shape is read here
      // rather than when the check was scheduled, because the length controls
      // settle after the sidecar answers; and nothing is recorded against the
      // key while it is skipped, so the check still runs the first time the
      // window comes back to the lengths the links were read at.
      if (!datasetDisplayIsAtLinkedShape()) return;
      datasetExcelLinkCheckedKeys.add(contextKey);
      datasetExcelLinkCheckAbortController?.abort();
      const abortController = new AbortController();
      datasetExcelLinkCheckAbortController = abortController;
      // One server pass answers both questions this dataset has about its
      // links: is every saved reference still readable, and is any workbook
      // newer than the values stored here.
      const result = await runtime.datasetExternalLinks.validateLinks(
        state.fileMtime,
        { signal: abortController.signal },
      );
      if (datasetExcelLinkCheckAbortController === abortController) {
        datasetExcelLinkCheckAbortController = null;
      }
      if (!isCurrent() || result?.aborted || result?.stale) return;
      if (!result?.ok) {
        setStatus("Excel links could not be verified.");
        return;
      }
      if (result.failures.length) {
        // A reference that no longer resolves is the answer the user has to act
        // on, so it replaces the newer-workbook prompt rather than queueing
        // behind it: refreshing from a workbook whose reference is broken
        // cannot succeed anyway.
        await reportDatasetExcelLinkFailures(result.failures, { isCurrent });
        return;
      }
      if (!result.newerWorkbookCount) return;
      const workbookNames = result.newerWorkbooks
        .map(({ path }) => String(path || "").split(/[\\/]/).pop())
        .filter(Boolean);
      const workbookSummary = workbookNames.length === 1
        ? `The linked workbook ${workbookNames[0]} is newer`
        : `${workbookNames.length} linked workbooks are newer`;
      const choice = await showPageMessageBox({
        title: "Linked Excel File Updated",
        tone: "warning",
        message: `${workbookSummary} than the values stored in this ArcRho dataset. Keep the stored values, or refresh from Excel. Refreshed values remain unsaved until you select Save.`,
        actions: [{ id: "refresh", label: "Refresh from Excel" }],
        okLabel: "Keep Current Values",
        balancedActions: true,
      });
      if (choice === "refresh" && isCurrent()) {
        await refreshDatasetExternalLinks({
          isCurrent,
          markRefreshedCellsDirty: true,
        });
      }
    }, 0);
  }

  // A reload of the same dataset -- the window came back, or only the view it
  // is shown at moved -- reads Excel again only when Excel has something new to
  // say. The dataset's own CSV already holds the figures the last refresh
  // brought over, so the workbooks are stated first and read only where one has
  // been saved since that file. A workbook that cannot be stated counts as
  // unchanged: the stored figures stand rather than being blanked by a drive
  // that happened to be away.
  async function refreshDatasetExternalLinksIfWorkbooksChanged(options = {}) {
    const isCurrent = typeof options?.isCurrent === "function" ? options.isCurrent : () => true;
    if (
      isDfmDataTabHost()
      || !datasetExternalLinksLoaded
      || !datasetDisplayIsAtLinkedShape()
      || !state.model
      || !currentDatasetIsManualTriangleOrVector()
      || !isCurrent()
    ) {
      return { linkedCellCount: 0, changedCount: 0, failedCount: 0 };
    }
    const changed = await runtime.datasetExternalLinks.findNewerWorkbooks(state.fileMtime);
    if (!isCurrent()) return { linkedCellCount: 0, changedCount: 0, failedCount: 0 };
    if (!changed?.ok || !changed.newerWorkbooks?.length) {
      return { linkedCellCount: 0, changedCount: 0, failedCount: 0 };
    }
    return refreshDatasetExternalLinks(options);
  }

  async function refreshDatasetExternalLinks(options = {}) {
    const isCurrent = typeof options?.isCurrent === "function" ? options.isCurrent : () => true;
    if (
      isDfmDataTabHost()
      || !datasetExternalLinksLoaded
      || !datasetDisplayIsAtLinkedShape()
      || !state.model
      || !currentDatasetIsManualTriangleOrVector()
      || !isCurrent()
    ) {
      return { linkedCellCount: 0, changedCount: 0, failedCount: 0 };
    }
    const hadLinkFailures = runtime.datasetExternalLinks.getLinkFailures().length > 0;
    const result = await runtime.datasetExternalLinks.refreshAll(
      options?.ids ?? null,
      { markRefreshedCellsDirty: options?.markRefreshedCellsDirty === true },
    );
    if (!isCurrent() || result?.stale || result?.aborted) return result;
    const failures = Array.isArray(result.failures) ? result.failures : [];
    if (result.changedCount > 0) {
      renderTable();
      notifyDatasetUpdated();
      applyGridSelectionFromState();
    } else if (hadLinkFailures) {
      // A repaired reference reads back the value already stored, so nothing
      // changed - but the red marking it clears still has to leave the grid.
      renderTable();
      applyGridSelectionFromState();
    }
    getDataTabLinksController()?.refresh?.();
    updateDatasetSaveUi();
    // A refresh the user asked for that did not do what they asked always says
    // so in the window, never only in a status line the next action overwrites:
    // with the references to fix when it has them, and with whatever reason it
    // does have when the batch itself did not come back.
    const unnamedCount = failures.length ? 0 : Number(result.failedCount) || 0;
    if (failures.length || unnamedCount) {
      window.setTimeout(() => {
        if (isCurrent()) {
          reportDatasetExcelLinkFailures(failures, {
            isCurrent,
            unnamedCount,
            reason: result.error,
          });
        }
      }, 0);
    } else if (result.changedCount > 0) {
      window.setTimeout(() => {
        if (isCurrent()) {
          setStatus(`Excel refresh updated ${result.changedCount} linked dataset cell${result.changedCount === 1 ? "" : "s"}.`);
        }
      }, 0);
    }
    return result;
  }

  // ArcRho dataset links and formula links refresh the same way: the
  // controller re-resolves and applies, and the grid repaints for any value
  // or broken-link marking that moved.
  async function refreshDatasetLinksOf(controller, noun, options = {}) {
    const isCurrent = typeof options?.isCurrent === "function" ? options.isCurrent : () => true;
    if (
      isDfmDataTabHost()
      || !datasetExternalLinksLoaded
      || !datasetDisplayIsAtLinkedShape()
      || !state.model
      || !currentDatasetIsManualTriangleOrVector()
      || !isCurrent()
    ) {
      return { linkedCellCount: 0, changedCount: 0, failedCount: 0 };
    }
    const hadLinkFailures = controller.getLinkFailures().length > 0;
    const result = await controller.refreshAll(
      options?.ids ?? null,
      { markRefreshedCellsDirty: options?.markRefreshedCellsDirty === true },
    );
    if (!isCurrent() || result?.stale || result?.aborted) return result;
    if (result.changedCount > 0) {
      renderTable();
      notifyDatasetUpdated();
      applyGridSelectionFromState();
    } else if (hadLinkFailures || result.failedCount > 0) {
      // Repaint so a repaired reference loses its red marking, and a newly
      // broken one gains it, even when no value moved.
      renderTable();
      applyGridSelectionFromState();
    }
    getDataTabLinksController()?.refresh?.();
    updateDatasetSaveUi();
    if (result.changedCount > 0 && !result.failedCount) {
      window.setTimeout(() => {
        if (isCurrent()) {
          setStatus(`${noun} refresh updated ${result.changedCount} linked cell${result.changedCount === 1 ? "" : "s"}.`);
        }
      }, 0);
    }
    return result;
  }

  function refreshDatasetInternalLinks(options = {}) {
    return refreshDatasetLinksOf(runtime.datasetInternalLinks, "Dataset link", options);
  }

  function refreshDatasetFormulaLinks(options = {}) {
    return refreshDatasetLinksOf(runtime.datasetFormulaLinks, "Formula", options);
  }

  async function syncSidecarForCurrentDataset(options = {}) {
    const isCurrent = typeof options?.isCurrent === "function" ? options.isCurrent : () => true;
    if (!isCurrent()) return false;
    const context = buildDatasetSidecarContextPayload();
    const key = buildDatasetSidecarContextKey(context);
    sidecarContextPayload = hasDatasetSidecarContext(context) ? context : null;
    sidecarContextKey = key;
    if (!key) {
      if (isDfmDataTabHost()) setDatasetRenderNumberFormatSettings(null);
      runtime.isSidecarReadOnlyDataset = false;
      applyStoredLengthsFromResponse(null);
      runtime.currentDatasetSidecarSourceKind = "";
      runtime.currentDatasetSidecarDataFormat = "";
      runtime.currentDatasetPrecedents = [];
      datasetExternalLinksLoaded = false;
      forEachLinksController((controller) => controller.clear());
      lastSavedDatasetSettings = null;
      datasetSettingsDirty = false;
      renderDatasetAuditLog([]);
      renderDetailFormula("", runtime.currentDatasetPrecedents);
      renderDatasetPrecedents([]);
      renderDatasetDependents([]);
      updateDatasetSaveUi();
      return false;
    }

    const nonce = ++sidecarSyncNonce;
    getDatasetAuditLog()?.setLoading();
    let resp;
    try {
      resp = options?.sidecarData
        ? { ok: true, data: options.sidecarData }
        : await loadDatasetSidecar(context);
    } catch (error) {
      if (!isCurrent()) return false;
      if (nonce === sidecarSyncNonce) {
        getDatasetAuditLog()?.setError(error?.message || "Unable to load the audit log.");
        datasetExternalLinksLoaded = false;
        forEachLinksController((controller) => controller.clear());
      }
      throw error;
    }
    if (nonce !== sidecarSyncNonce || !isCurrent()) return false;
    if (!resp.ok) {
      if (isDfmDataTabHost()) setDatasetRenderNumberFormatSettings(null);
      setStatus(`Dataset settings load failed: ${resp?.data?.detail || "Unknown error."}`);
      applyStoredLengthsFromResponse(null);
      runtime.currentDatasetSidecarSourceKind = isProjectInstanceDraft ? "input" : "";
      runtime.currentDatasetSidecarDataFormat = isProjectInstanceDraft ? getProjectInstanceDraftDataFormat() : "";
      runtime.currentDatasetPrecedents = [];
      datasetExternalLinksLoaded = false;
      forEachLinksController((controller) => controller.clear());
      lastSavedDatasetSettings = normalizeDatasetSettings(getCurrentDatasetSettings());
      datasetSettingsDirty = false;
      getDatasetAuditLog()?.setError(resp?.data?.detail || "Unable to load the audit log.");
      renderDetailFormula(getDatasetTypeFormulaByName(document.getElementById("triInput")?.value || ""), runtime.currentDatasetPrecedents);
      renderDatasetPrecedents([]);
      renderDatasetDependents([]);
      updateDatasetSaveUi();
      return false;
    }

    const data = resp.data || {};
    state.sidecarUpdatedAt = data.exists ? String(data.updated_at || "") : "";
    const notesSynced = await syncNotesForCurrentDataset({
      isCurrent,
      forceReload: options?.forceReload === true,
      notes: data.exists ? String(data.notes ?? "") : "",
    });
    if (!isCurrent() || notesSynced === false) return false;
    applyStoredLengthsFromResponse(data.exists ? data : null);
    runtime.currentDatasetSidecarSourceKind = data.exists ? String(data.source_kind || "") : (isProjectInstanceDraft ? "input" : "");
    runtime.currentDatasetSidecarDataFormat = data.exists ? String(data.data_format || "") : (isProjectInstanceDraft ? getProjectInstanceDraftDataFormat() : "");
    runtime.currentDatasetPrecedents = data.exists ? normalizeDatasetDependencyEntries(data.precedents) : [];
    datasetExternalLinksLoaded = !isDfmDataTabHost() && currentDatasetIsManualTriangleOrVector();
    runtime.datasetExternalLinks.load(
      datasetExternalLinksLoaded && data.exists ? data.external_links : [],
    );
    runtime.datasetInternalLinks.load(
      datasetExternalLinksLoaded && data.exists ? data.internal_links : [],
    );
    runtime.datasetFormulaLinks.load(
      datasetExternalLinksLoaded && data.exists ? data.formula_links : [],
    );
    if (data.exists) scheduleDatasetExcelLinkCheck({ contextKey: key, isCurrent });
    if (isProjectInstanceDraft && data.exists && !String(data.csv_file || "").trim()) {
      runtime.savedProjectInstanceDraftName = String(data.dataset_name || context.dataset_name || "").trim();
    }
    renderDatasetAuditLog(data.exists ? data.audit_log : []);
    renderDetailFormula(
      data.exists
        ? (String(data.formula || "").trim() || getDatasetTypeFormulaByName(data.dataset_type || context.dataset_name || ""))
        : getDatasetTypeFormulaByName(document.getElementById("triInput")?.value || ""),
      runtime.currentDatasetPrecedents,
    );
    renderDatasetPrecedents(runtime.currentDatasetPrecedents);
    renderDatasetDependents(data.exists ? data.dependents : []);
    runtime.isSidecarReadOnlyDataset = !!data.exists && sourceKindIsReadOnly(runtime.currentDatasetSidecarSourceKind);
    const patchSaveBtn = document.getElementById("saveBtn");
    if (patchSaveBtn && !isReadOnlyDatasetViewer) {
      patchSaveBtn.disabled = runtime.isSidecarReadOnlyDataset;
      patchSaveBtn.title = runtime.isSidecarReadOnlyDataset ? "Calculated datasets are read-only." : "";
    }
    let settings;
    if (data.exists) {
      settings = normalizeDatasetSettings(data);
    } else if (isTemporaryDatasetView) {
      settings = await resolveTemporaryDatasetSettings(context);
      if (!isCurrent()) return false;
    } else {
      settings = normalizeDatasetSettings(getCurrentDatasetSettings());
    }
    if (isDfmDataTabHost()) {
      setDatasetRenderNumberFormatSettings(data.exists ? settings : null);
    }
    lastSavedDatasetSettings = settings;
    if (options?.forceReload === true) {
      await refreshDatasetExternalLinksIfWorkbooksChanged({ isCurrent });
      if (!isCurrent()) return false;
    }
    if (options?.applyLengths !== false && data.exists) {
      applyDatasetSettingsToControls(settings);
      saveTriInputsToStorage();
      datasetSettingsDirty = false;
      updateDatasetSaveUi();
      return true;
    }
    applyTemporaryNumberFormatSettings(settings);
    refreshDatasetSettingsDirty();
    return true;
  }

  // The saving animation is created per controller instance so the Dataset
  // window and a method page hosting this Data tab keep separate popups.
  const datasetSaveProgress = createArcRhoSaveProgress({ subject: "Dataset", noun: "dataset" });

  async function saveDatasetSidecarForCurrentContext(progress = null) {
    if (isTemporaryDatasetView) {
      return { ok: false, error: "Temporary view does not save permanent dataset sidecars." };
    }
    if (isProjectInstanceDraft) {
      const originResult = validateDatasetOriginLabels(state.model?.origin_labels, {
        originLen: getTriInputs().originLen,
        requireMatchingPeriod: true,
      });
      if (!originResult.ok) {
        return {
          ok: false,
          error: `Dataset draft cannot be saved: ${originResult.error}. Set a valid Origin Start Date in Project Settings, then try again.`,
        };
      }
    }
    if (await refreshDatasetInstanceNameConflict()) {
      return { ok: false, error: runtime.datasetInstanceNameConflictMessage || "Dataset instance name already exists." };
    }
    const context = buildDatasetSidecarContextPayload();
    if (!hasDatasetSidecarContext(context)) {
      return { ok: false, error: "Project, Reserving Class, and Dataset Type are required." };
    }
    const settings = getCurrentDatasetSettings();
    const payload = {
      ...context,
      ...settings,
      // The period the CSV is written at. Stated only while a hand-entered
      // dataset is still empty, which is the one time it can move.
      stored_development_length: storedDevelopmentLengthForSave() || null,
      // The file's old values were set to 0 and the shape moved since, so the
      // server writes a new file at this shape instead of holding the old one.
      ...(releasedLengths ? { stored_values_cleared: true } : {}),
      notes: String(getNotesEditorElements().input?.value ?? ""),
      ...getManualInputDatasetValuePayload(),
      ...getDatasetExternalLinksPayload(),
      ...getDatasetInternalLinksPayload(),
      ...getDatasetFormulaLinksPayload(),
    };
    progress?.writing();
    const resp = await withDataTabDatasetMutation({ source: "sidecar-save" }, () => saveDatasetSidecar(payload));
    if (!resp.ok) {
      return { ok: false, error: resp?.data?.detail || "Failed to save dataset settings." };
    }
    sidecarSyncNonce += 1;
    sidecarContextPayload = context;
    sidecarContextKey = buildDatasetSidecarContextKey(context);
    notesContextPayload = { ...context };
    notesContextKey = buildNotesContextKey(notesContextPayload);
    applyNotesInputValue(String(resp.data?.notes ?? ""));
    lastSavedDatasetSettings = normalizeDatasetSettings(settings);
    applyStoredLengthsFromResponse(resp.data);
    runtime.currentDatasetSidecarSourceKind = String(resp.data?.source_kind || (isProjectInstanceDraft ? "input" : runtime.currentDatasetSidecarSourceKind) || "");
    runtime.currentDatasetSidecarDataFormat = String(resp.data?.data_format || settings.data_format || runtime.currentDatasetSidecarDataFormat || "");
    runtime.currentDatasetPrecedents = normalizeDatasetDependencyEntries(resp.data?.precedents);
    if (datasetExternalLinksLoaded) {
      runtime.datasetExternalLinks.markClean(resp.data?.external_links ?? runtime.datasetExternalLinks.serialize());
      runtime.datasetInternalLinks.markClean(resp.data?.internal_links ?? runtime.datasetInternalLinks.serialize());
      runtime.datasetFormulaLinks.markClean(resp.data?.formula_links ?? runtime.datasetFormulaLinks.serialize());
    }
    if (isProjectInstanceDraft) {
      runtime.savedProjectInstanceDraftName = context.dataset_name;
    }
    if (hasManualInputGridChanges()) {
      state.dirty.clear();
    }
    if (state.model && currentDatasetIsManualTriangleOrVector()) {
      state.model.source_kind = runtime.currentDatasetSidecarSourceKind;
      state.model.data_format = runtime.currentDatasetSidecarDataFormat;
    }
    if (resp.data?.ds_id) {
      config.DS_ID = String(resp.data.ds_id);
      saveLastDsId(config.DS_ID);
    }
    if (resp.data?.file_mtime !== undefined && resp.data?.file_mtime !== null) {
      state.fileMtime = resp.data.file_mtime;
    }
    renderDatasetAuditLog(resp.data?.audit_log);
    renderDetailFormula(
      String(resp.data?.formula || "").trim() || getDatasetTypeFormulaByName(settings.dataset_type),
      runtime.currentDatasetPrecedents,
    );
    renderDatasetPrecedents(runtime.currentDatasetPrecedents);
    renderDatasetDependents(resp.data?.dependents);
    invalidateCachedDatasetInstances();
    datasetSettingsDirty = false;
    updateDatasetSaveUi();
    clearDatasetDependencyPreview("save");
    // Engine-hosted saves return with the dependent walk already finished;
    // a null outcome (walk failures) keeps the window open and leaves the
    // dataset table as the failure surface.
    const propagationOutcome = await trackSavePropagation(resp.data?.calculated_updates, {
      onStatus: (message, statusOptions) => {
        progress?.setMessage?.(message, statusOptions);
        setStatus(message);
      },
      onComplete: () => requestProjectInstanceDatasetTableRefresh(),
    });
    // The write and its dependent walk are done; drop the spinner before the
    // "Saved" notice that follows the save command.
    progress?.finish();
    handleCalculationUpdates(resp.data?.calculated_updates, "Dataset settings save");
    state.sidecarUpdatedAt = String(resp.data?.updated_at || state.sidecarUpdatedAt || "");
    notifyDataTabDurableDatasetState({ source: "sidecar-save" });
    return {
      ok: true,
      data: resp.data,
      propagationClean: propagationOutcome !== null,
      refreshedDatasets: propagationOutcome?.refreshed_datasets || [],
      linkWarnings: propagationOutcome?.link_warnings || [],
    };
  }
  async function saveDatasetChanges(options = {}) {
    if (isTemporaryDatasetView) {
      return { ok: false, error: "Temporary view is read-only and cannot save permanent dataset changes." };
    }
    if (runtime.datasetSaveInFlight) return { ok: false, error: "Save already in progress." };
    return datasetSaveProgress.run((progress) => runDatasetSave(options, progress));
  }

  async function runDatasetSave(options, progress) {
    forEachLinksController((controller) => controller.abort());
    runtime.datasetSaveInFlight = true;
    updateDatasetSaveUi();
    void getDataTabLinksController()?.refresh?.();
    let saveStatus = buildDatasetSaveStatus();
    // A save with nothing dirty writes nothing and enqueues no walk, so it
    // counts as clean for the close-on-save decision.
    let propagationClean = true;
    let refreshedDatasets = [];
    let linkWarnings = [];
    try {
      if (datasetSettingsDirty || hasManualInputGridChanges() || linkControllerNames.some((name) => runtime[name].isDirty()) || notesDirty || isUnsavedProjectInstanceDraft()) {
        const sidecarResult = await saveDatasetSidecarForCurrentContext(progress);
        if (!sidecarResult.ok) return sidecarResult;
        saveStatus = buildDatasetSaveStatus(sidecarResult.data);
        propagationClean = sidecarResult.propagationClean !== false;
        refreshedDatasets = sidecarResult.refreshedDatasets || [];
        linkWarnings = sidecarResult.linkWarnings || [];
      }
      updateDatasetSaveUi();
      if (!options?.silentStatus) setStatus(saveStatus.text, saveStatus.tone);
      requestProjectInstanceDatasetTableRefresh();
      return { ok: true, propagationClean, refreshedDatasets, linkWarnings };
    } finally {
      runtime.datasetSaveInFlight = false;
      updateDatasetSaveUi();
      void getDataTabLinksController()?.refresh?.();
    }
  }

  async function discardDatasetChanges(options = {}) {
    const reload = options?.reload !== false;
    forEachLinksController((controller) => controller.restoreSaved());
    // A clear-and-reshape is discarded with the rest, so the file's own shape
    // binds the length controls again.
    releasedLengths = null;
    if (lastSavedDatasetSettings) {
      applyDatasetSettingsToControls(lastSavedDatasetSettings);
      saveTriInputsToStorage();
      if (reload) {
        const project = getResolvedProjectValue();
        if (project) {
          await ensureHeadersForProject(project);
          await ensureDevHeadersForProject(project);
        }
        renderTable();
        notifyDatasetUpdated();
        renderChart();
        setStatus("Loading dataset...");
        scheduleAutoRun(0);
      }
    }
    if (notesDirty) applyNotesInputValue(lastSavedNotesText);
    clearDatasetDependencyPreview("cancel");
    state.dirty.clear();
    datasetSettingsDirty = false;
    updateDatasetSaveUi();
  }

  async function confirmCancelDatasetChanges(reason = "close") {
    if (!datasetCloseConfirm) datasetCloseConfirm = getDataTabCloseConfirm();
    if (!hasUnsavedDatasetChanges()) return true;
    if (!datasetCloseConfirm) return false;
    const discard = await datasetCloseConfirm.confirm({ reason });
    if (!discard) return false;
    await discardDatasetChanges({ reload: reason !== "close" });
    return true;
  }

  function requestConfirmedDatasetClose() {
    clearDatasetDependencyPreview("close-discard");
    requestTabbedPageWindowClose({
      messageType: "arcrho:dataset-close-confirmed",
      inst: instanceId,
    });
  }

  function wireDatasetSaveControls() {
    if (!datasetCloseConfirm) datasetCloseConfirm = getDataTabCloseConfirm();
    document.getElementById("datasetSaveBtn")?.addEventListener("click", async () => {
      await handleDatasetSaveCommand();
    });
    document.getElementById("datasetCancelBtn")?.addEventListener("click", async () => {
      const ok = await confirmCancelDatasetChanges("close");
      if (ok) requestConfirmedDatasetClose();
    });
    window.__arcrho_request_close = () => {
      if (!hasUnsavedDatasetChanges()) return false;
      if (datasetCloseConfirm?.isOpen) return true;
      void (async () => {
        const ok = await confirmCancelDatasetChanges("close");
        if (ok) requestConfirmedDatasetClose();
      })();
      return true;
    };
    window.__arcrho_consume_close_shortcut = window.__arcrho_request_close;
    window.addEventListener("beforeunload", (event) => {
      if (!hasUnsavedDatasetChanges()) return;
      event.preventDefault();
      event.returnValue = "";
    });
    updateDatasetSaveUi();
  }

  async function handleDatasetSaveCommand() {
    const result = await saveDatasetChanges();
    // A cancelled dependent-update confirmation is the user's own answer, not
    // a failure: the window stays open with the edit intact and no alarm.
    if (result.cancelled) setStatus(result.error || "Save cancelled; nothing was changed.");
    else if (!result.ok) setStatus(`Dataset save failed: ${result.error || "Unknown error."}`);
    // A save never closes the window; the user keeps working in place. After a
    // clean dependent walk the notice names what the walk refreshed, except for
    // the Data tab hosted inside a DFM window, which reports through DFM's save.
    if (result.ok && result.propagationClean && !isDfmDataTabHost()) {
      await showSavedDependentsNotice(result.refreshedDatasets, { linkWarnings: result.linkWarnings });
    }
    return result;
  }

  function getDisplayProjectValue() {
    return (document.getElementById("projectSelect")?.value || "").trim();
  }

  function getDisplayReservingClassValue() {
    return (document.getElementById("pathInput")?.value || "").trim();
  }

  function getDisplayTriValue() {
    return (document.getElementById("triInput")?.value || "").trim();
  }

  function getRawProjectValueForNotes() {
    const input = document.getElementById("projectSelect");
    if (isInputDefaultBound(input)) {
      const defaults = loadWorkflowDefaults();
      return typeof defaults?.project === "string" ? defaults.project : "";
    }
    return String(input?.value ?? "");
  }

  function getRawReservingClassValueForNotes() {
    const input = document.getElementById("pathInput");
    if (isInputDefaultBound(input)) {
      const defaults = loadWorkflowDefaults();
      return typeof defaults?.reservingClass === "string" ? defaults.reservingClass : "";
    }
    return String(input?.value ?? "");
  }

  function getRawDatasetNameValueForNotes() {
    const input = document.getElementById("dsDetailName") || document.getElementById("triInput");
    return String(input?.value ?? "");
  }

  function buildNotesContextPayload() {
    return {
      project_name: getRawProjectValueForNotes(),
      reserving_class: getRawReservingClassValueForNotes(),
      dataset_name: getRawDatasetNameValueForNotes(),
    };
  }

  function hasNotesContext(payload) {
    if (!payload || typeof payload !== "object") return false;
    const projectName = String(payload.project_name ?? "");
    const reservingClass = String(payload.reserving_class ?? "");
    const datasetName = String(payload.dataset_name ?? "");
    return !!projectName.trim() && !!reservingClass.trim() && !!datasetName.trim();
  }

  function buildNotesContextKey(payload) {
    if (!hasNotesContext(payload)) return "";
    return `${payload.project_name}\u001f${payload.reserving_class}\u001f${payload.dataset_name}`;
  }

  function getNotesErrorMessage(resp, fallback) {
    const detail = resp?.data?.detail;
    if (typeof detail === "string" && detail.trim()) return detail.trim();
    const error = resp?.data?.error;
    if (typeof error === "string" && error.trim()) return error.trim();
    if (typeof fallback === "string" && fallback.trim()) return fallback.trim();
    return "Unknown error.";
  }

  function getNotesEditorElements() {
    return {
      input: datasetNotesController?.elements?.input || null,
      saveState: document.getElementById("dsNotesSaveState"),
    };
  }

  function updateNotesSaveUi() {
    const { saveState } = getNotesEditorElements();
    const hasContext = !!notesContextKey && hasNotesContext(notesContextPayload);

    if (!saveState) return;
    saveState.classList.remove("is-dirty", "is-clean", "is-hidden");
    if (isTemporaryDatasetView) {
      saveState.textContent = "Read-only in temporary view";
      saveState.classList.add("is-clean");
      updateDatasetSaveUi();
      return;
    }
    if (!hasContext) {
      saveState.textContent = "No dataset context";
      updateDatasetSaveUi();
      return;
    }
    if (notesDirty) {
      saveState.textContent = "Unsaved changes";
      saveState.classList.add("is-dirty");
      updateDatasetSaveUi();
      return;
    }
    saveState.textContent = "";
    saveState.classList.add("is-hidden");
    updateDatasetSaveUi();
  }

  function applyNotesInputValue(text) {
    const nextText = String(text ?? "");
    lastSavedNotesText = nextText;
    notesDirty = false;
    datasetNotesController?.setValue(nextText, { markClean: true });
    updateNotesSaveUi();
    updateDatasetSaveUi();
  }

  async function saveNotesForPayload(payload, options = {}) {
    if (isTemporaryDatasetView) {
      return { ok: false, error: "Temporary view is read-only and cannot save notes." };
    }
    const silentStatus = !!options?.silentStatus;
    const isCurrent = typeof options?.isCurrent === "function" ? options.isCurrent : () => true;
    if (!isCurrent()) return { ok: false, stale: true };
    if (!hasNotesContext(payload)) {
      updateNotesSaveUi();
      return { ok: false, error: "Project, Reserving Class, and Dataset Type are required." };
    }

    const { input } = getNotesEditorElements();
    const notesText = String(input?.value ?? "");
    const req = {
      project_name: payload.project_name,
      reserving_class: payload.reserving_class,
      dataset_name: payload.dataset_name,
      notes: notesText,
    };
    const resp = await saveDatasetNotes(req);
    if (!isCurrent()) return { ok: false, stale: true };
    if (!resp.ok) {
      return { ok: false, error: getNotesErrorMessage(resp, "Failed to save notes.") };
    }

    notesContextPayload = {
      project_name: req.project_name,
      reserving_class: req.reserving_class,
      dataset_name: req.dataset_name,
    };
    notesContextKey = buildNotesContextKey(notesContextPayload);
    lastSavedNotesText = notesText;
    datasetNotesController?.markClean(notesText);
    notesDirty = datasetNotesController?.isDirty()
      ?? (String(input?.value ?? "") !== notesText);
    updateNotesSaveUi();
    if (!silentStatus && !notesDirty) setStatus("Notes saved.");
    return { ok: true, data: resp.data, dirty: notesDirty };
  }

  async function saveNotesForCurrentContext(options = {}) {
    return saveNotesForPayload(notesContextPayload, options);
  }

  async function syncNotesForCurrentDataset(options = {}) {
    const isCurrent = typeof options?.isCurrent === "function" ? options.isCurrent : () => true;
    const forceReload = options?.forceReload === true;
    if (!isCurrent()) return false;
    const nextPayload = buildNotesContextPayload();
    const nextKey = buildNotesContextKey(nextPayload);
    if (nextKey === notesContextKey && notesDirty) {
      notesContextPayload = hasNotesContext(nextPayload) ? nextPayload : null;
      updateNotesSaveUi();
      return true;
    }
    if (nextKey === notesContextKey && !forceReload) {
      notesContextPayload = hasNotesContext(nextPayload) ? nextPayload : null;
      updateNotesSaveUi();
      return true;
    }

    if (notesContextKey && notesDirty) {
      const shouldSave = window.confirm(
        "You have unsaved Notes. Click OK to save before switching notes, or Cancel to discard unsaved changes.",
      );
      if (shouldSave) {
        const saveResult = await saveNotesForCurrentContext({ silentStatus: true, isCurrent });
        if (!isCurrent()) return false;
        if (saveResult.stale) return false;
        if (!saveResult.ok) {
          setStatus(`Notes save failed: ${saveResult.error || "Unknown error."}`);
          updateNotesSaveUi();
          return false;
        }
        if (saveResult.dirty) {
          setStatus("Notes changed while saving. Save the latest notes before switching datasets.");
          updateNotesSaveUi();
          return false;
        }
      } else {
        notesDirty = false;
      }
    }

    notesContextPayload = hasNotesContext(nextPayload) ? nextPayload : null;
    notesContextKey = nextKey;
    updateNotesSaveUi();
    if (!nextKey) {
      applyNotesInputValue("");
      return true;
    }

    applyNotesInputValue(String(options?.notes ?? ""));
    return true;
  }

  function wireDataTabPersistenceLifecycle() {
    if (hostInputsPublished) return;
    hostInputsPublished = true;
    publishDataTabHostInputs({
      getResolvedProjectValue,
      getResolvedReservingClassValue,
      getDisplayProjectValue,
      getDisplayReservingClassValue,
      getDisplayTriValue,
      isInputDefaultBound,
    });
  }
  function wireNotesEditor() {
    if (datasetNotesController && !datasetNotesController.destroyed) return datasetNotesController;
    const container = document.getElementById("datasetNotesMount");
    if (!container) return null;
    datasetNotesController = mountDataTabNotes({
      container,
      setNotesDirty: (value) => {
        notesDirty = !!value;
      },
      updateNotesSaveUi,
      setStatus,
    });
    if (isTemporaryDatasetView) {
      const { input, styleControls } = datasetNotesController?.elements || {};
      if (input) {
        input.readOnly = true;
        input.setAttribute("aria-readonly", "true");
        input.title = "Notes are read-only in temporary view.";
      }
      for (const control of Object.values(styleControls || {})) {
        if (control) control.disabled = true;
      }
    }
    return datasetNotesController;
  }


  Object.assign(runtime, {
    wireDataTabPersistenceLifecycle,
    buildDatasetSidecarContextPayload, hasDatasetSidecarContext,
    buildDatasetSidecarContextKey, getCurrentDatasetSettings,
    getManualInputDatasetValuePayload, getDatasetExternalLinksPayload,
    getDatasetInternalLinksPayload, getDatasetFormulaLinksPayload,
    normalizeDatasetSettings, sameDatasetSettings,
    hasManualInputGridChanges, hasUnsavedDatasetChanges,
    isUnsavedProjectInstanceDraft, shouldPersistManualInputGridValues,
    hasPendingDatasetSaveWork, isDraftGridUnavailable,
    normalizeDatasetModeText,
    sourceKindIsReadOnly,
    currentDatasetIsManualTriangleOrVector,
    datasetValuesAreAllZero,
    getManualDatasetLengthBaseline,
    getCurrentLengthControlValues,
    getStoredLengthPair,
    storedLengthIsPending,
    releaseStoredShape,
    chooseStoredDevelopmentLength,
    getStoredDevelopmentLengthChoice,
    getStoredLengthControlPair,
    storedDevelopmentLengthForSave,
    applyStoredLengthChoices,
    updateStoredLengthControls,
    datasetOriginDisplayIsCoarserThanStored,
    datasetDevelopmentDisplayIsCoarserThanStored,
    datasetDisplayIsAtLinkedShape,
    datasetOffLinkedShapeLinkHint,
    datasetCoarserViewMessage,
    datasetCoarseDevelopmentNote,
    validateManualDatasetLengthChange,
    updateManualDatasetModeControls,
    updateVectorDevelopmentLengthControl,
    restoreManualDatasetModeControls,
    notifyDatasetDirtyState,
    updateDatasetSaveUi,
    refreshDatasetSettingsDirty,
    applyDatasetSettingsToControls,
    applyTemporaryNumberFormatDefaults,
    resolveTemporaryDatasetSettings,
    loadTemporaryNumberFormatSettings,
    invalidateDatasetContextLoads,
    refreshDatasetExternalLinks,
    refreshDatasetInternalLinks,
    refreshDatasetFormulaLinks,
    syncSidecarForCurrentDataset,
    saveDatasetSidecarForCurrentContext,
    saveDatasetChanges,
    discardDatasetChanges,
    confirmCancelDatasetChanges,
    requestConfirmedDatasetClose,
    wireDatasetSaveControls,
    handleDatasetSaveCommand,
    getDisplayProjectValue,
    getDisplayReservingClassValue,
    getDisplayTriValue,
    getRawProjectValueForNotes,
    getRawReservingClassValueForNotes,
    getRawDatasetNameValueForNotes,
    buildNotesContextPayload,
    hasNotesContext,
    buildNotesContextKey,
    getNotesErrorMessage,
    getNotesEditorElements,
    updateNotesSaveUi,
    applyNotesInputValue,
    saveNotesForPayload,
    saveNotesForCurrentContext,
    syncNotesForCurrentDataset,
    wireNotesEditor,
  });
}
