// Owns host messaging, dependency previews, status, and calculation update reporting.

import {
  collectDatasetPropagationFailures,
  datasetPropagationFailureStep,
} from "/ui/shared/tabs/data/data_tab_propagation_report.js?v=20260830a";
import {
  beginDatasetGridLoading,
  endDatasetGridLoading,
  setDatasetGridError,
} from "/ui/shared/tabs/data/dataset_grid_placeholder.js?v=20260809a";

export function registerDataTabHostController(runtime) {
  const { state, config, $, instanceId, stepId, workflowId, WF_GLOBAL_CTRL_PREFIX } = runtime;
  const defer = (name) => (...args) => runtime[name](...args);
  const { updateDatasetSaveUi, getDatasetInstanceNameValue, getResolvedProjectValue, getResolvedReservingClassValue, hasManualInputGridChanges, previewCalculatedDatasetDependents, normalizeDatasetModeText, normalizeReservingClassPath, hasUnsavedDatasetChanges, validateDatasetOriginLabels, getTriInputs, renderTable, renderChart, loadDataset, handleDatasetSaveCommand, handleWorkflowGlobalChange, clearValidValueListCache, logLine, loadCachedDataset, saveLastDsId, syncSidecarForCurrentDataset, applyGridSelectionFromState, recordDatasetBrowsingHistory, saveTriInputsToStorage, setDatasetRenderNumberFormatSettings, hasResultSelectionUpdates, createDatasetHeadersService, createDatasetRunController, validateTriInputsBeforeRun, buildTriRequestPayload, buildVecRequestPayload, getDatasetRunDataFormat, invalidateDatasetContextLoads, isDatasetReadOnly, getDatasetReadOnlyMessage, datasetCoarseDevelopmentNote, getDataset, patchDataset, isDfmDataTabHost } = new Proxy({}, { get: (_target, name) => defer(name) });
  const FONT_STORAGE_KEY = "arcrho_app_font";
  const FORCE_REBUILD_KEY = "arcrho_force_rebuild_enabled";
  const CALCULATED_DATASETS_UPDATED_MESSAGE = "arcrho:calculated-datasets-updated";
  let calculatedDatasetRefreshInFlight = false;
  let calculatedDependencyPreviewTimer = null;
  let calculatedDependencyPreviewSeq = 0;
  let hostLifecycleWired = false;
  const activeCalculatedDependencyPreviewTargets = new Map();

  function buildFontStack(font) {
    const raw = String(font || "").trim();
    if (!raw) return "";
    if (raw.includes(",")) return raw;
    const primary = /\s/.test(raw) ? `"${raw.replace(/\"/g, "")}"` : raw;
    return `${primary}, "Segoe UI", "SegoeUI", Tahoma, Arial, sans-serif`;
  }

  function applyAppFont(font) {
    const stack = buildFontStack(font);
    if (!stack) return;
    const root = document.documentElement;
    if (root) root.style.setProperty("--app-font", stack);
    if (document.body) document.body.style.fontFamily = stack;
  }

  function loadAppFontFromStorage() {
    try {
      const raw = localStorage.getItem(FONT_STORAGE_KEY);
      if (raw && typeof raw === "string") return raw;
    } catch {}
    return "";
  }

  function isForceRebuildEnabled() {
    try {
      return localStorage.getItem(FORCE_REBUILD_KEY) === "1";
    } catch {
      return false;
    }
  }

  function notifyDatasetUpdated(options = {}) {
    window.dispatchEvent(new CustomEvent("arcrho:dataset-updated"));
    updateDatasetSaveUi();
    if (options?.publishPreview !== false) publishDatasetDependencyPreview();
  }

  function requestProjectInstanceDatasetTableRefresh() {
    try {
      window.parent?.postMessage({ type: "arcrho:project-instance-refresh-datasets" }, "*");
    } catch {
      // ignore stale parent frames
    }
  }

  function numberOrNull(value) {
    if (value === null || value === undefined || value === "") return null;
    const numeric = Number(value);
    return Number.isFinite(numeric) ? numeric : null;
  }

  function latestDiagonalValues(values, mask) {
    const rows = Array.isArray(values) ? values : [];
    return rows.map((row, r) => {
      if (!Array.isArray(row)) return null;
      for (let c = row.length - 1; c >= 0; c -= 1) {
        if (Array.isArray(mask?.[r]) && mask[r][c] === false) continue;
        const value = numberOrNull(row[c]);
        if (value !== null) return value;
      }
      return null;
    });
  }

  function vectorValues(values) {
    const rows = Array.isArray(values) ? values : [];
    return rows.map((row) => numberOrNull(Array.isArray(row) ? row[0] : row));
  }

  function cloneDatasetMatrixValues(values) {
    return Array.isArray(values)
      ? values.map((row) => (Array.isArray(row) ? row.map(numberOrNull) : []))
      : [];
  }

  function cloneDatasetMask(mask) {
    return Array.isArray(mask)
      ? mask.map((row) => (Array.isArray(row) ? row.map(Boolean) : []))
      : [];
  }

  function datasetDependencySourceValues() {
    const values = Array.isArray(state.model?.values) ? state.model.values : [];
    const mask = Array.isArray(state.model?.mask) ? state.model.mask : [];
    const format = normalizeDatasetModeText(runtime.currentDatasetSidecarDataFormat || state.model?.data_format || "");
    return format === "triangle" ? latestDiagonalValues(values, mask) : vectorValues(values);
  }

  function buildDatasetDependencySourceMessage(type, reason = "") {
    const datasetName = getDatasetInstanceNameValue() || document.getElementById("triInput")?.value || "";
    const datasetTypeName = document.getElementById("triInput")?.value || datasetName;
    const payload = {
      type,
      inst: instanceId,
      project: getResolvedProjectValue(),
      reservingClass: getResolvedReservingClassValue(),
      datasetName,
      datasetTypeName,
      names: [datasetName, datasetTypeName].map((value) => String(value || "").trim()).filter(Boolean),
      methodType: state.model?.method_type || "",
      sourceKind: runtime.currentDatasetSidecarSourceKind || state.model?.source_kind || "",
      dataFormat: runtime.currentDatasetSidecarDataFormat || state.model?.data_format || "",
      reason,
    };
    if (type === "arcrho:dependency-source-preview") {
      payload.values = datasetDependencySourceValues();
      payload.matrixValues = cloneDatasetMatrixValues(state.model?.values);
      payload.mask = cloneDatasetMask(state.model?.mask);
      payload.originLabels = Array.isArray(state.model?.origin_labels) ? state.model.origin_labels.map(String) : [];
      payload.developmentLabels = Array.isArray(state.model?.dev_labels) ? state.model.dev_labels.map(String) : [];
      payload.originLength = payload.originLabels.length;
      payload.developmentLength = payload.developmentLabels.length;
    }
    return payload;
  }

  function postDatasetDependencySourceMessage(type, reason = "") {
    const message = buildDatasetDependencySourceMessage(type, reason);
    if (!message.names.length) return;
    try {
      window.parent?.postMessage(message, "*");
    } catch {}
  }

  function publishDatasetDependencyPreview() {
    if (!hasManualInputGridChanges()) return;
    postDatasetDependencySourceMessage("arcrho:dependency-source-preview", "dirty");
    scheduleCalculatedDependencyPreview();
  }

  function clearDatasetDependencyPreview(reason = "") {
    if (calculatedDependencyPreviewTimer != null) {
      window.clearTimeout(calculatedDependencyPreviewTimer);
      calculatedDependencyPreviewTimer = null;
    }
    calculatedDependencyPreviewSeq += 1;
    postDatasetDependencySourceMessage("arcrho:dependency-source-cleared", reason || "clean");
    clearCalculatedDependencyPreviewTargets(reason || "clean");
  }

  function postCalculatedDependencyPreviewTarget(step, reason = "calculated-preview") {
    const datasetName = String(step?.dataset_name || step?.dataset_type_name || "").trim();
    if (!datasetName) return "";
    const message = {
      type: "arcrho:dependency-source-preview",
      inst: instanceId,
      project: getResolvedProjectValue(),
      reservingClass: getResolvedReservingClassValue(),
      datasetName,
      datasetTypeName: String(step?.dataset_type_name || datasetName).trim(),
      names: [datasetName, step?.dataset_type_name].map((value) => String(value || "").trim()).filter(Boolean),
      methodType: "Calculated Dataset",
      sourceKind: "calculated_preview",
      dataFormat: String(step?.data_format || step?.dataFormat || "").trim(),
      reason,
      values: Array.isArray(step?.values) ? step.values : [],
      matrixValues: Array.isArray(step?.matrix_values) ? step.matrix_values : (Array.isArray(step?.matrixValues) ? step.matrixValues : []),
      mask: Array.isArray(step?.mask) ? step.mask : [],
      originLabels: Array.isArray(step?.origin_labels) ? step.origin_labels.map(String) : [],
      developmentLabels: Array.isArray(step?.development_labels) ? step.development_labels.map(String) : [],
    };
    if (!message.names.length || !message.matrixValues.length) return "";
    const key = dependencyMessageSourceKey(message);
    activeCalculatedDependencyPreviewTargets.set(key, message);
    try {
      window.parent?.postMessage(message, "*");
    } catch {}
    return key;
  }

  function clearCalculatedDependencyPreviewTargets(reason = "clean", keepKeys = new Set()) {
    for (const [key, message] of Array.from(activeCalculatedDependencyPreviewTargets.entries())) {
      if (keepKeys?.has?.(key)) continue;
      activeCalculatedDependencyPreviewTargets.delete(key);
      try {
        window.parent?.postMessage({
          ...message,
          type: "arcrho:dependency-source-cleared",
          reason,
        }, "*");
      } catch {}
    }
  }

  function scheduleCalculatedDependencyPreview() {
    if (calculatedDependencyPreviewTimer != null) {
      window.clearTimeout(calculatedDependencyPreviewTimer);
    }
    calculatedDependencyPreviewTimer = window.setTimeout(() => {
      calculatedDependencyPreviewTimer = null;
      void publishCalculatedDependencyPreview();
    }, 120);
  }

  async function publishCalculatedDependencyPreview() {
    if (!hasManualInputGridChanges()) {
      clearCalculatedDependencyPreviewTargets("clean");
      return;
    }
    const seq = ++calculatedDependencyPreviewSeq;
    const sourceMessage = buildDatasetDependencySourceMessage("arcrho:dependency-source-preview", "dirty");
    if (!sourceMessage.names.length || !Array.isArray(sourceMessage.matrixValues) || !sourceMessage.matrixValues.length) return;
    const result = await previewCalculatedDatasetDependents({
      project_name: sourceMessage.project,
      reserving_class: sourceMessage.reservingClass,
      changed_dataset_name: sourceMessage.datasetName,
      changed_dataset_type_name: sourceMessage.datasetTypeName,
      values: sourceMessage.matrixValues,
      mask: sourceMessage.mask,
      origin_labels: sourceMessage.originLabels,
      development_labels: sourceMessage.developmentLabels,
    }).catch(() => null);
    if (seq !== calculatedDependencyPreviewSeq || !hasManualInputGridChanges()) return;
    const steps = Array.isArray(result?.data?.steps) ? result.data.steps : [];
    const keepKeys = new Set();
    for (const step of steps) {
      if (!step?.ok) continue;
      const key = postCalculatedDependencyPreviewTarget(step);
      if (key) keepKeys.add(key);
    }
    clearCalculatedDependencyPreviewTargets("preview-stale", keepKeys);
  }

  function dependencyMessageSourceKey(message = {}) {
    const names = [
      ...(Array.isArray(message.names) ? message.names : []),
      message.datasetName,
      message.datasetTypeName,
      message.name,
    ]
      .map(normalizeDatasetMatchText)
      .filter(Boolean)
      .sort();
    return [
      normalizeDatasetMatchText(message.inst),
      normalizeDatasetMatchText(message.project),
      normalizeReservingClassPath(message.reservingClass || message.reserving_class || ""),
      names.join("|"),
    ].join("\u001f");
  }

  function dependencyMessageNames(message = {}) {
    return new Set([
      ...(Array.isArray(message.names) ? message.names : []),
      message.datasetName,
      message.datasetTypeName,
      message.name,
    ].map(normalizeDatasetMatchText).filter(Boolean));
  }

  function dependencyMessageMatchesCurrentContext(message = {}) {
    if (!message || typeof message !== "object") return false;
    if (String(message.inst || "") && String(message.inst || "") === String(instanceId || "")) return false;
    const project = String(message.project || message.project_name || "").trim();
    if (project && normalizeDatasetMatchText(project) !== normalizeDatasetMatchText(getResolvedProjectValue())) {
      return false;
    }
    const reservingClass = String(message.reservingClass || message.reserving_class || "").trim();
    if (reservingClass) {
      const left = normalizeDatasetMatchText(normalizeReservingClassPath(reservingClass));
      const right = normalizeDatasetMatchText(normalizeReservingClassPath(getResolvedReservingClassValue()));
      if (left && right && left !== right) return false;
    }
    const names = dependencyMessageNames(message);
    if (!names.size) return false;
    const currentNames = collectCurrentDatasetNamesForMatch();
    for (const name of currentNames) {
      if (names.has(name)) return true;
    }
    return false;
  }

  function previewMatrixFromDependencyMessage(message = {}) {
    const matrix = Array.isArray(message.matrixValues)
      ? message.matrixValues
      : (Array.isArray(message.values) ? message.values.map((value) => [value]) : []);
    return matrix
      .filter((row) => Array.isArray(row))
      .map((row) => row.map(numberOrNull));
  }

  function labelsFromDependencyMessage(message = {}, key, fallback = []) {
    const values = Array.isArray(message[key]) ? message[key] : [];
    const labels = values.map((value) => String(value ?? "").trim()).filter(Boolean);
    return labels.length ? labels : (Array.isArray(fallback) ? fallback.map(String) : []);
  }

  function buildDependencyPreviewMask(values, sourceMask) {
    if (Array.isArray(sourceMask) && sourceMask.length) {
      return values.map((row, r) => row.map((_, c) => !!sourceMask?.[r]?.[c]));
    }
    return values.map((row) => row.map(() => true));
  }

  function applyDependencySourcePreview(message = {}) {
    if (!dependencyMessageMatchesCurrentContext(message)) return false;
    if (hasUnsavedDatasetChanges()) {
      setStatus("A live source preview is available. Save or discard local edits before applying it.");
      return false;
    }
    const values = previewMatrixFromDependencyMessage(message);
    if (!values.length) return false;
    const currentModel = state.model || {};
    const originLabelCandidates = Array.isArray(message.originLabels) && message.originLabels.length
      ? message.originLabels
      : currentModel.origin_labels;
    const originResult = validateDatasetOriginLabels(originLabelCandidates, {
      originLen: getTriInputs().originLen,
      expectedCount: values.length,
      requireMatchingPeriod: true,
    });
    if (!originResult.ok) {
      setStatus(
        `Cannot apply live source preview: ${originResult.error}. `
        + "Reload the dataset after correcting Origin Start Date in Project Settings.",
      );
      return false;
    }
    const originLabels = originResult.labels;
    const developmentLabels = labelsFromDependencyMessage(
      message,
      "developmentLabels",
      Array.isArray(currentModel.dev_labels) && currentModel.dev_labels.length
        ? currentModel.dev_labels
        : ["1"],
    );
    state.model = {
      ...currentModel,
      origin_labels: originLabels,
      dev_labels: developmentLabels.length ? developmentLabels : ["1"],
      values,
      mask: buildDependencyPreviewMask(values, message.mask),
      data_format: message.dataFormat || currentModel.data_format || runtime.currentDatasetSidecarDataFormat || "",
      source_kind: message.sourceKind || currentModel.source_kind || runtime.currentDatasetSidecarSourceKind || "",
    };
    runtime.activeDependencyPreviewKey = dependencyMessageSourceKey(message);
    renderTable();
    renderChart();
    window.dispatchEvent(new CustomEvent("arcrho:dataset-updated", {
      detail: { preview: true, source: message },
    }));
    return true;
  }

  async function clearDependencySourcePreview(message = {}) {
    if (!runtime.activeDependencyPreviewKey) return false;
    if (!dependencyMessageMatchesCurrentContext(message)) return false;
    const sourceKey = dependencyMessageSourceKey(message);
    if (sourceKey !== runtime.activeDependencyPreviewKey) return false;
    runtime.activeDependencyPreviewKey = "";
    try {
      await loadDataset();
    } catch (err) {
      setStatus(`Dataset preview reload failed: ${String(err?.message || err)}`);
    }
    return true;
  }

  function normalizeDatasetMatchText(value) {
    return String(value || "").trim().toLowerCase();
  }

  // ---- Cross-window dataset reference picking ------------------------------
  // The window editing a formula announces the pick; every other durable
  // dataset window of the same project and reserving class answers clicks and
  // drags on its grid with the picked rectangle until the pick ends.

  function publishDatasetReferencePickBegin() {
    try {
      window.parent?.postMessage({
        type: "arcrho:dataset-reference-pick-begin",
        inst: instanceId,
        project: getResolvedProjectValue(),
        reservingClass: getResolvedReservingClassValue(),
        datasetName: getDatasetInstanceNameValue(),
      }, "*");
    } catch {}
  }

  function publishDatasetReferencePickEnd() {
    try {
      window.parent?.postMessage({
        type: "arcrho:dataset-reference-pick-end",
        inst: instanceId,
      }, "*");
    } catch {}
  }

  function publishDatasetReferencePick(range = {}) {
    const toInst = String(state.referencePickRequester || "");
    const datasetName = getDatasetInstanceNameValue();
    if (!toInst || !datasetName) return;
    try {
      window.parent?.postMessage({
        type: "arcrho:dataset-reference-pick",
        inst: instanceId,
        toInst,
        datasetName,
        dataFormat: runtime.currentDatasetSidecarDataFormat || state.model?.data_format || "",
        rowStart: range.rowStart,
        rowEnd: range.rowEnd,
        colStart: range.colStart,
        colEnd: range.colEnd,
        final: range.final === true,
      }, "*");
    } catch {}
  }

  function handleDatasetReferencePickBegin(msg = {}) {
    if (!msg || String(msg.inst || "") === String(instanceId || "")) return;
    if (isDfmDataTabHost() || runtime.isProjectInstanceDraft || runtime.isTemporaryDatasetView) return;
    if (!getDatasetInstanceNameValue()) return;
    // References resolve within one reserving class, so only its windows pick.
    if (normalizeDatasetMatchText(msg.project) !== normalizeDatasetMatchText(getResolvedProjectValue())) return;
    const left = normalizeDatasetMatchText(normalizeReservingClassPath(msg.reservingClass || ""));
    const right = normalizeDatasetMatchText(normalizeReservingClassPath(getResolvedReservingClassValue()));
    if (!left || !right || left !== right) return;
    state.referencePickRequester = String(msg.inst);
    setStatus(`Select cells here to insert them into the ${String(msg.datasetName || "dataset").trim() || "dataset"} formula.`);
    // Repaint so the grid picks up the pointer and dashed-range treatment it
    // wears only while it is answering someone else's formula.
    runtime.applyGridSelectionFromState?.();
  }

  function handleDatasetReferencePickEnd(msg = {}) {
    if (String(state.referencePickRequester || "") !== String(msg?.inst || "")) return;
    state.referencePickRequester = "";
    setStatus("");
    runtime.applyGridSelectionFromState?.();
  }

  function collectCurrentDatasetNamesForMatch() {
    return new Set([
      normalizeDatasetMatchText(getDatasetInstanceNameValue()),
      normalizeDatasetMatchText(document.getElementById("triInput")?.value || ""),
    ].filter(Boolean));
  }

  function isCalculationStepUpdated(step) {
    return !!step?.ok || String(step?.status || "").toLowerCase() === "updated";
  }

  function calculationContextMatches(report, step = {}) {
    const reportProject = String(step?.project_name || report?.project_name || "").trim();
    const reportPath = String(step?.reserving_class || report?.reserving_class || "").trim();
    if (reportProject && normalizeDatasetMatchText(reportProject) !== normalizeDatasetMatchText(getResolvedProjectValue())) {
      return false;
    }
    if (reportPath && normalizeReservingClassPath(reportPath) !== normalizeReservingClassPath(getResolvedReservingClassValue())) {
      return false;
    }
    return true;
  }

  function calculationStepMatchesCurrentDataset(step) {
    const currentNames = collectCurrentDatasetNamesForMatch();
    if (!currentNames.size) return false;
    return [
      step?.dataset_type_name,
      step?.dataset_name,
      step?.instance_name,
    ].some((value) => currentNames.has(normalizeDatasetMatchText(value)));
  }

  function calculationReportTargetsCurrentDataset(report) {
    if (!report || typeof report !== "object") return false;
    const steps = collectCalculationSteps(report);
    if (steps.some((step) => isCalculationStepUpdated(step) && calculationContextMatches(report, step) && calculationStepMatchesCurrentDataset(step))) {
      return true;
    }
    if (!calculationContextMatches(report)) return false;
    const currentNames = collectCurrentDatasetNamesForMatch();
    return Array.isArray(report.targets) && report.targets.some((target) => currentNames.has(normalizeDatasetMatchText(target)));
  }

  async function handleCalculatedDatasetsUpdatedMessage(report) {
    if (calculatedDatasetRefreshInFlight || !calculationReportTargetsCurrentDataset(report)) return;
    if (hasUnsavedDatasetChanges()) {
      setStatus("This dataset was recalculated on disk. Save or discard local edits before reloading.");
      return;
    }
    calculatedDatasetRefreshInFlight = true;
    try {
      setStatus("Upstream formula change refreshed this dataset. Reloading...");
      const result = await loadDataset();
      if (!result?.ok) return;
      setStatus("Dataset refreshed after upstream recalculation.");
    } catch (err) {
      const message = String(err?.message || err || "Dataset refresh failed.");
      setStatus(`Dataset refresh failed: ${message}`);
    } finally {
      calculatedDatasetRefreshInFlight = false;
    }
  }

  function handleHostMessage(e) {
    if (e?.data?.type === "arcrho:dataset-save") {
      void handleDatasetSaveCommand();
      return;
    }
    if (e?.data?.type === "arcrho:set-app-font") {
      applyAppFont(e.data.font);
    }
    if (e?.data?.type === CALCULATED_DATASETS_UPDATED_MESSAGE) {
      void handleCalculatedDatasetsUpdatedMessage(e.data.report || null);
      return;
    }
    if (e?.data?.type === "arcrho:dataset-reference-pick-begin") {
      handleDatasetReferencePickBegin(e.data);
      return;
    }
    if (e?.data?.type === "arcrho:dataset-reference-pick-end") {
      handleDatasetReferencePickEnd(e.data);
      return;
    }
    if (e?.data?.type === "arcrho:dataset-reference-pick") {
      if (String(e.data.toInst || "") === String(instanceId || "")) {
        runtime.applyDatasetReferencePick?.(e.data);
      }
      return;
    }
    if (e?.data?.type === "arcrho:dependency-source-preview") {
      applyDependencySourcePreview(e.data);
      return;
    }
    if (e?.data?.type === "arcrho:dependency-source-cleared") {
      void clearDependencySourcePreview(e.data);
      return;
    }
    if (e?.data?.type === "arcrho:workflow-global-changed") {
      handleWorkflowGlobalChange(e.data.globalControl);
    }
    if (e?.data?.type === "arcrho:force-rebuild-toggle") {
      try {
        localStorage.setItem(FORCE_REBUILD_KEY, e?.data?.enabled ? "1" : "0");
      } catch {
        // ignore
      }
      return;
    }
    if (e?.data?.type === "arcrho:server-connection-updated") {
      clearValidValueListCache();
      logLine("Server connection updated.");
    }
  }

  function handleHostStorage(e) {
    if (!workflowId) return;
    if (e.key === `${WF_GLOBAL_CTRL_PREFIX}${workflowId}`) {
      try {
        const frameEl = window.frameElement;
        if (frameEl && frameEl.offsetParent === null) return;
      } catch {
        // ignore
      }
      handleWorkflowGlobalChange();
    }
  }

  function handleHostMouseDown() {
    window.parent.postMessage({ type: "arcrho:close-shell-menus" }, "*");
  }

  function requestCloseActiveTab() {
    window.parent.postMessage({ type: "arcrho:close-active-tab" }, "*");
  }

  function handleHostKeyDown(e) {
    const key = (e.key || "").toLowerCase();
    if (e.altKey && key === "w") {
      e.preventDefault();
      e.stopPropagation();
      requestCloseActiveTab();
      return;
    }
    if (e.ctrlKey && key === "q") {
      e.preventDefault();
      e.stopPropagation();
      window.parent.postMessage({ type: "arcrho:hotkey", action: "app_shutdown" }, "*");
      return;
    }
    if (e.ctrlKey) {
      if (key === "s") {
        e.preventDefault();
        e.stopPropagation();
        const action = e.shiftKey ? "file_save_as" : "file_save";
        window.parent.postMessage({ type: "arcrho:hotkey", action }, "*");
        return;
      }
      if (key === "o") {
        e.preventDefault();
        e.stopPropagation();
        window.parent.postMessage({ type: "arcrho:hotkey", action: "file_import" }, "*");
        return;
      }
      if (key === "p") {
        e.preventDefault();
        e.stopPropagation();
        window.parent.postMessage({ type: "arcrho:hotkey", action: "file_print" }, "*");
        return;
      }
      if (e.shiftKey && key === "f") {
        e.preventDefault();
        e.stopPropagation();
        window.parent.postMessage({ type: "arcrho:hotkey", action: "view_toggle_nav" }, "*");
        return;
      }
    }
    if (e.altKey && key === "r" && e.ctrlKey) {
      e.preventDefault();
      e.stopPropagation();
      window.parent.postMessage({ type: "arcrho:hotkey", action: "file_restart" }, "*");
      return;
    }
  }

  function wireDataTabHostLifecycle() {
    if (hostLifecycleWired) return;
    hostLifecycleWired = true;
    window.ArcRhoZoomBridge?.wirePageZoomBridge();
    applyAppFont(loadAppFontFromStorage());
    window.addEventListener("message", handleHostMessage);
    window.addEventListener("storage", handleHostStorage);
    window.addEventListener("mousedown", handleHostMouseDown, { capture: true });
    window.addEventListener("keydown", handleHostKeyDown, { capture: true });
    if (isDfmDataTabHost()) {
      window.ADA_DFM_REFRESH_DATASET = refreshDfmDatasetForCurrentInputs;
      window.ADA_DFM_APPLY_DATASET_SNAPSHOT = applyDfmDatasetSnapshot;
    }
  }

  function scheduleAutoRun(delayMs = 150) {
    return runtime.datasetRunController.scheduleAutoRun(delayMs);
  }

  function bindAutoRunOnEnter(el) {
    return runtime.datasetRunController.bindAutoRunOnEnter(el);
  }

  function runArcRhoTri(opts = {}) {
    return runtime.datasetRunController.runArcRhoTri(opts);
  }

  async function loadProjectInstanceCachedDataset() {
    const { project, path, tri, instanceName, originLen, devLen, cumulative, calendar } = getTriInputs();
    const datasetName = instanceName || tri;
    if (!project || !path || !datasetName) return { ok: false, skipped: true };
    const gridPlaceholderToken = beginDatasetGridLoading({ message: `Loading "${datasetName}"` });
    try {
      return await readProjectInstanceCachedDataset({
        project, path, datasetName, originLen, devLen, cumulative, calendar,
      });
    } finally {
      endDatasetGridLoading(gridPlaceholderToken);
    }
  }

  async function readProjectInstanceCachedDataset(context) {
    const { project, path, datasetName, originLen, devLen, cumulative, calendar } = context;
    setStatus(`Loading ${datasetName}...`);
    const response = await loadCachedDataset({
      project_name: project,
      reserving_class: path,
      dataset_name: datasetName,
      origin_length: originLen,
      development_length: devLen,
      cumulative,
      calendar,
      // The window opens at the shape the sidecar saved, so a hand-entered
      // dataset stored finer than it is shown arrives already rolled up.
      at_display_shape: true,
    });
    if (!response.ok || response.data?.ok === false) {
      const message = String(response.data?.detail || response.data?.error || `Dataset cache load failed (${response.status}).`);
      setDatasetGridError(message);
      setStatus(message);
      return { ok: false, status: response.status, data: response.data, message };
    }
    const data = response.data || {};
    config.DS_ID = String(data.id || "");
    if (config.DS_ID) saveLastDsId(config.DS_ID);
    state.dirty.clear();
    state.model = data;
    state.fileMtime = data.mtime;
    state.headerLabels = Array.isArray(data.origin_labels) ? data.origin_labels.map(String) : [];
    state.devHeaderLabels = Array.isArray(data.dev_labels) ? data.dev_labels.map(String) : [];
    const sidecarSynced = await syncSidecarForCurrentDataset({
      applyLengths: true,
      sidecarData: data,
    });
    if (sidecarSynced === false) return { ok: false, contextSyncFailed: true };
    renderTable();
    renderChart();
    notifyDatasetUpdated();
    applyGridSelectionFromState();
    updateCurrentTabTitle();
    recordDatasetBrowsingHistory({ project, path, tri: datasetName });
    const meta = document.getElementById("dsMeta");
    if (meta) meta.textContent = `id=${data.id} | origins=${state.headerLabels.length} | dev=${state.devHeaderLabels.length} | mtime=${data.mtime}`;
    // The same sentence the run path shows while a coarser development view
    // is up, since a dataset opened this way can arrive at one.
    setStatus(datasetCoarseDevelopmentNote() || [path, datasetName].filter(Boolean).join(" | ") || "Ready");
    return { ok: true, data };
  }

  async function refreshDfmDatasetForCurrentInputs() {
    if (!isDfmDataTabHost()) return null;
    saveTriInputsToStorage();
    setStatus("Loading dataset...");
    return runArcRhoTri({ showValidationMessage: false });
  }

  function applyDfmDatasetSnapshot(snapshot = {}) {
    if (!isDfmDataTabHost()) return { ok: false, error: "DFM Data-tab host is not active." };
    const originLabels = Array.isArray(snapshot.origin_labels)
      ? snapshot.origin_labels.map((label) => String(label ?? ""))
      : [];
    const developmentLabels = Array.isArray(snapshot.dev_labels)
      ? snapshot.dev_labels.map((label) => String(label ?? ""))
      : [];
    const values = Array.isArray(snapshot.values)
      ? snapshot.values.map((row) => (Array.isArray(row) ? row.slice() : []))
      : [];
    const mask = Array.isArray(snapshot.mask)
      ? snapshot.mask.map((row, rowIndex) => (
        Array.isArray(row)
          ? row.map(Boolean)
          : (values[rowIndex] || []).map((value) => value !== null && value !== undefined)
      ))
      : values.map((row) => row.map((value) => value !== null && value !== undefined));
    state.dirty.clear();
    state.model = {
      ...snapshot,
      origin_labels: originLabels,
      dev_labels: developmentLabels,
      values,
      mask,
      data_format: String(snapshot.data_format || "Triangle"),
      source_kind: String(snapshot.source_kind || "dfm-snapshot"),
    };
    state.headerLabels = originLabels.slice();
    state.devHeaderLabels = developmentLabels.slice();
    runtime.currentDatasetSidecarSourceKind = state.model.source_kind;
    runtime.currentDatasetSidecarDataFormat = state.model.data_format;
    runtime.currentDatasetPrecedents = [];
    runtime.isSidecarReadOnlyDataset = true;
    setDatasetRenderNumberFormatSettings({
      number_format: snapshot.number_format,
      decimal_places: snapshot.decimal_places,
    });
    renderTable();
    renderChart();
    applyGridSelectionFromState();
    const meta = document.getElementById("dsMeta");
    if (meta) meta.textContent = `snapshot | origins=${originLabels.length} | dev=${developmentLabels.length}`;
    return { ok: true, data: state.model };
  }

  function isRunInFlight() {
    return runtime.datasetRunController.isRunInFlight();
  }

  function updateCurrentTabTitle() {
    if (isDfmDataTabHost()) return null;
    const triangleName = document.getElementById("triInput")?.value?.trim();
    if (!triangleName) return null;

    window.parent.postMessage(
      {
        type: "arcrho:update-active-tab-title",
        title: `${triangleName}`,
      },
      "*"
    );

    return triangleName;
  }

  function setStatus(text, tone = "") {
    try {
      window.parent.postMessage({ type: "arcrho:status", text, tone }, "*");
    } catch {
      // ignore
    }
  }

  function collectCalculationSteps(report) {
    if (!report || typeof report !== "object") return [];
    const steps = [];
    if (Array.isArray(report.steps)) steps.push(...report.steps);
    if (Array.isArray(report.chains)) {
      report.chains.forEach((chain) => {
        if (Array.isArray(chain?.steps)) steps.push(...chain.steps);
      });
    }
    if (!steps.length && Array.isArray(report.updated)) steps.push(...report.updated.map((item) => ({ ...item, status: "updated" })));
    if (!steps.length && Array.isArray(report.skipped)) steps.push(...report.skipped.map((item) => ({ ...item, status: "skipped" })));
    steps.push(...collectDatasetPropagationFailures(report).map(datasetPropagationFailureStep));
    const seen = new Set();
    return steps.filter((step) => {
      const key = [
        String(step?.reserving_class || report.reserving_class || ""),
        String(step?.dataset_type_name || ""),
        isCalculationStepUpdated(step) ? "updated" : "not-updated",
        String(step?.reason || ""),
      ].join("\u0001");
      if (seen.has(key)) return false;
      seen.add(key);
      return String(step?.dataset_type_name || "").trim() || String(step?.reason || "").trim();
    });
  }

  function publishCalculatedDatasetUpdates(report, source = "Dataset save") {
    if (!collectCalculationSteps(report).some(isCalculationStepUpdated) && !hasResultSelectionUpdates(report)) return;
    try {
      window.parent.postMessage({
        type: CALCULATED_DATASETS_UPDATED_MESSAGE,
        report,
        source,
      }, "*");
    } catch {
      // ignore
    }
  }

  // The post-save "Saved" notice (shared with the method pages) is the only
  // dialog a dataset save opens; this hands the report to the Project
  // Instance so its dataset table and open windows refresh.
  function handleCalculationUpdates(report, source = "Dataset save") {
    publishCalculatedDatasetUpdates(report, source);
  }

  runtime.datasetHeadersService = createDatasetHeadersService({
    state,
    setStatus,
  });

  runtime.datasetRunController = createDatasetRunController({
    config,
    state,
    $,
    logLine,
    getDataset,
    patchDataset,
    renderTable,
    renderChart,
    notifyDatasetUpdated,
    isForceRebuildEnabled,
    validateTriInputsBeforeRun,
    getTriInputs,
    buildTriRequestPayload,
    buildVecRequestPayload,
    getDatasetRunDataFormat,
    clearHeadersCacheForProject: (project, options = {}) =>
      runtime.datasetHeadersService.clearHeadersCacheForProject(project, options),
    ensureHeadersForProject: (project, options = {}) =>
      runtime.datasetHeadersService.ensureHeadersForProject(project, options),
    ensureDevHeadersForProject: (project, options = {}) =>
      runtime.datasetHeadersService.ensureDevHeadersForProject(project, options),
    saveLastDsId,
    recordDatasetBrowsingHistory,
    syncSidecarForCurrentDataset,
    invalidateDatasetContextLoads,
    updateCurrentTabTitle,
    setStatus,
    onCalculatedUpdates: (report, source) => handleCalculationUpdates(report, source),
    applyGridSelectionFromState,
    stepId,
    suppressLoadingPopup: isDfmDataTabHost(),
    isDatasetReadOnly,
    datasetReadOnlyMessage: getDatasetReadOnlyMessage,
    datasetCoarseDevelopmentNote,
  });


  Object.assign(runtime, {
    wireDataTabHostLifecycle,
    buildFontStack,
    applyAppFont,
    loadAppFontFromStorage,
    isForceRebuildEnabled,
    notifyDatasetUpdated,
    requestProjectInstanceDatasetTableRefresh,
    numberOrNull,
    latestDiagonalValues,
    vectorValues,
    cloneDatasetMatrixValues,
    cloneDatasetMask,
    datasetDependencySourceValues,
    buildDatasetDependencySourceMessage,
    postDatasetDependencySourceMessage,
    publishDatasetDependencyPreview,
    clearDatasetDependencyPreview,
    postCalculatedDependencyPreviewTarget,
    clearCalculatedDependencyPreviewTargets,
    scheduleCalculatedDependencyPreview,
    publishCalculatedDependencyPreview,
    dependencyMessageSourceKey,
    dependencyMessageNames,
    dependencyMessageMatchesCurrentContext,
    previewMatrixFromDependencyMessage,
    labelsFromDependencyMessage,
    buildDependencyPreviewMask,
    applyDependencySourcePreview,
    clearDependencySourcePreview,
    normalizeDatasetMatchText,
    publishDatasetReferencePickBegin,
    publishDatasetReferencePickEnd,
    publishDatasetReferencePick,
    handleDatasetReferencePickBegin,
    handleDatasetReferencePickEnd,
    collectCurrentDatasetNamesForMatch,
    isCalculationStepUpdated,
    calculationContextMatches,
    calculationStepMatchesCurrentDataset,
    calculationReportTargetsCurrentDataset,
    handleCalculatedDatasetsUpdatedMessage,
    requestCloseActiveTab,
    scheduleAutoRun,
    bindAutoRunOnEnter,
    runArcRhoTri,
    loadProjectInstanceCachedDataset,
    refreshDfmDatasetForCurrentInputs,
    applyDfmDatasetSnapshot,
    isRunInFlight,
    updateCurrentTabTitle,
    setStatus,
    collectCalculationSteps,
    publishCalculatedDatasetUpdates,
    handleCalculationUpdates,
  });
}
