/*
===============================================================================
DFM Tabs - Orchestrator
Initializes all DFM tabs, wires event handlers, and coordinates modules.
===============================================================================
*/
import {
  applyTabbedPageSaveBar,
  createTabbedPage,
  requestTabbedPageWindowClose,
  updateTabbedPageSaveControls,
} from "/ui/shared/tabbed_page/tabbed_page.js?v=20260816a";
import { syncDetailsLabelWidth } from "/ui/shared/tabs/details/details_form_layout.js?v=20260820b";
import { applyHostFixedDetailsFields } from "/ui/shared/tabs/details/details_host_fields.js?v=20260820b";
import { createPageCloseConfirm } from "/ui/shared/components/close_confirm/close_confirm.js";
import { showSavedDependentsNotice } from "/ui/shared/components/progress_popup/save_progress.js?v=20260831a";
import { setStorageInstance, loadNaBorders } from "/ui/method_pages/dfm/dfm_storage.js";
import {
  state as dfmState,
  getDfmInst,
  setShowNaBorders,
  setCachedRootPath,
  setCurrentDfmTab,
  getCurrentDfmTab,
  getDfmIsDirty,
  markDfmDirty,
  notifyDfmEditState,
  consumePendingDfmPropagationJobId,
} from "/ui/method_pages/dfm/dfm_state.js";
import { ALLOWED_DFM_TABS, DFM_TAB_DEFS } from "/ui/method_pages/dfm/dfm_tab_config.js?v=20260903a";
import { initDfmAuditLog, refreshDfmAuditLog } from "/ui/method_pages/dfm/dfm_audit_log.js?v=20260726a";
import {
  renderRatioTable,
  wireRatioStrikeToggle,
  wireRatioChartModal,
  wireRatioContextMenu,
  wireDfmSpinnerControls,
  excludeExtremeInActiveCol,
  includeAllInActiveCol,
  isRatioChartOpen,
  scheduleRatioChartRender,
  restoreRatioHistoryUi,
} from "/ui/method_pages/dfm/dfm_ratios_tab.js?v=20260907a";
import {
  renderResultsTable,
  wireResultsRatioBasisControls,
  buildPercentDevelopedVector,
  buildResultsVector,
} from "/ui/method_pages/dfm/dfm_results_tab.js?v=20260907a";
import { wireNotesInput } from "/ui/method_pages/dfm/dfm_notes_tab.js?v=20260714a";
import { initDfmCurvesTab, renderDfmCurvesTab } from "/ui/method_pages/dfm/dfm_curves_tab.js?v=20260907a";
import { initDfmLinks, refreshDfmLinks } from "/ui/method_pages/dfm/dfm_links_tab.js?v=20260901a";
import {
  syncMethodNameFromInputs,
  syncOutputTypeFromProject,
  wireMethodName,
  wireDfmInstanceCreationNotice,
  wireDetailsThresholdReset,
} from "/ui/method_pages/dfm/dfm_details.js?v=20260907a";
import {
  scheduleRatioSelectionLoad,
  saveRatioSelectionPattern,
  restoreCleanDfmMethodState,
  recordCurrentDfmCleanState,
  saveDfmTemplate,
  applyDfmOwnedPatchPayload,
  buildDfmAssistantContextPayload,
  resolveCurrentDfmMethodSavePath,
  startDfmMethodFileWatcher,
  stopDfmMethodFileWatcher,
  scheduleDfmMethodPreview,
  cancelDfmMethodAsyncTasks,
} from "/ui/method_pages/dfm/dfm_persistence.js?v=20260907a";
import { wireRatioSyncChannel, requestRatioStateSync } from "/ui/method_pages/dfm/dfm_sync.js?v=20260907a";
import { wireDfmRpcBridgeTabBar } from "/ui/method_pages/dfm/dfm_rpc_bridge_tabbar.js?v=20260907a";
import { reviewArcBotDfmEditApproval } from "/ui/method_pages/dfm/dfm_rpc_bridge_client.js?v=20260907a";
import { wireDfmTabPopoutWindows } from "/ui/method_pages/dfm/dfm_tab_popout_window.js?v=20260903a";
import {
  clearRatioHistoryTempSession,
  getRatioHistoryState,
  initRatioHistory,
  runRatioRedo,
  runRatioUndo,
} from "/ui/method_pages/dfm/dfm_ratio_history.js";
import { readDatasetInputQueryValues } from "/ui/shared/tabs/data/data_tab_query_inputs.js";

const DEFAULT_TOKEN = "__DEFAULT__";
let dfmSaveInFlight = false;
const dfmCloseConfirm = createPageCloseConfirm({ subject: "DFM" });
let dependencyPreviewTimer = 0;
const persistedDfmBootstrap = Boolean(
  new URLSearchParams(globalThis.location?.search || "").get("method_name"),
);

function wireDfmScrollbarActivity(scrollHost) {
  if (!scrollHost || scrollHost.dataset.scrollbarActivityWired === "1") return;
  scrollHost.dataset.scrollbarActivityWired = "1";

  let idleTimer = 0;
  const syncScrollbarHover = (event) => {
    const rect = scrollHost.getBoundingClientRect();
    const verticalScrollbarWidth = Math.max(0, scrollHost.offsetWidth - scrollHost.clientWidth);
    const horizontalScrollbarHeight = Math.max(0, scrollHost.offsetHeight - scrollHost.clientHeight);
    const nearVerticalScrollbar = scrollHost.scrollHeight > scrollHost.clientHeight
      && verticalScrollbarWidth > 0
      && event.clientX >= rect.right - Math.max(verticalScrollbarWidth, 16);
    const nearHorizontalScrollbar = scrollHost.scrollWidth > scrollHost.clientWidth
      && horizontalScrollbarHeight > 0
      && event.clientY >= rect.bottom - Math.max(horizontalScrollbarHeight, 16);

    scrollHost.classList.toggle("isScrollbarHover", nearVerticalScrollbar || nearHorizontalScrollbar);
  };

  scrollHost.addEventListener("scroll", () => {
    scrollHost.classList.add("isScrolling");
    if (idleTimer) clearTimeout(idleTimer);
    idleTimer = setTimeout(() => {
      scrollHost.classList.remove("isScrolling");
    }, 550);
  }, { passive: true });
  scrollHost.addEventListener("pointermove", syncScrollbarHover, { passive: true });
  scrollHost.addEventListener("pointerleave", () => {
    scrollHost.classList.remove("isScrollbarHover");
  }, { passive: true });
}

function getDfmInputSnapshotSafe() {
  try {
    if (typeof window.ADA_GET_DFM_INPUTS === "function") {
      return window.ADA_GET_DFM_INPUTS();
    }
  } catch {
    // ignore
  }
  const project = document.getElementById("projectSelect")?.value?.trim() || "";
  const reservingClass = document.getElementById("pathInput")?.value?.trim() || "";
  return {
    resolved: { project, reservingClass },
    display: { project, reservingClass },
    defaults: { projectDefault: false, reservingClassDefault: false },
  };
}

function handleDatasetUpdated() {
  refreshDfmTabContent("dataset-updated");
}

function normalizeDfmIdentity(value) {
  return String(value || "").trim().toLowerCase();
}

function handleDfmPropagationReport(report) {
  const updates = report?.dfm_updates;
  if (!updates || typeof updates !== "object") return false;
  const currentProject = normalizeDfmIdentity(getDfmInputSnapshotSafe().resolved?.project);
  const currentClass = normalizeDfmIdentity(getDfmInputSnapshotSafe().resolved?.reservingClass);
  if (
    updates.project_name
    && normalizeDfmIdentity(updates.project_name) !== currentProject
  ) return false;
  if (
    updates.reserving_class
    && normalizeDfmIdentity(updates.reserving_class) !== currentClass
  ) return false;

  const query = new URLSearchParams(globalThis.location?.search || "");
  const identities = new Set([
    normalizeDfmIdentity(query.get("output_dataset")),
    normalizeDfmIdentity(query.get("method_name")),
    normalizeDfmIdentity(document.getElementById("dfmMethodName")?.value),
  ].filter(Boolean));
  const all = [
    ...(Array.isArray(updates.updated) ? updates.updated : []),
    ...(Array.isArray(updates.status_refreshed) ? updates.status_refreshed : []),
    ...(Array.isArray(updates.errors) ? updates.errors : []),
  ];
  const matched = all.filter((item) => identities.has(normalizeDfmIdentity(item?.dataset_name)));
  if (!matched.length) return false;
  const failed = matched.find((item) => item?.reason);
  if (failed) {
    window.parent.postMessage({
      type: "arcrho:status",
      text: `DFM refresh requires review: ${String(failed.reason || "upstream refresh failed")}`,
      tone: "warn",
    }, "*");
  }
  if (getDfmIsDirty()) {
    window.parent.postMessage({
      type: "arcrho:status",
      text: "Upstream DFM data changed while this window has edits. Save will rebase the owned edits onto the latest derived state.",
      tone: "warn",
    }, "*");
    return true;
  }
  scheduleRatioSelectionLoad("upstream-refresh");
  return true;
}

function refreshDfmTabContent(reason = "") {
  renderRatioTable();
  renderDfmCurvesTab();
  renderResultsTable();
  if (!persistedDfmBootstrap) {
    syncMethodNameFromInputs();
    syncOutputTypeFromProject();
  }
  if (!getDfmIsDirty()) {
    scheduleRatioSelectionLoad(reason || "dfm-refresh");
  }
  if (getCurrentDfmTab() === "audit") {
    void refreshDfmAuditLog();
  }
  if (isRatioChartOpen()) scheduleRatioChartRender();
}

async function buildAssistantContext() {
  let methodPath = "";
  let pathError = "";
  let activeJson = null;
  let activeJsonError = "";
  try {
    methodPath = await resolveCurrentDfmMethodSavePath();
  } catch (err) {
    pathError = String(err?.message || err || "Could not resolve DFM method path.");
  }
  try {
    activeJson = await buildDfmAssistantContextPayload({ persistSummaryOrder: false });
  } catch (err) {
    activeJsonError = String(err?.message || err || "Could not build active DFM method payload.");
  }
  const inputSnap = getDfmInputSnapshotSafe();
  return {
    available: true,
    pageType: "dfm",
    activeDfmTab: getCurrentDfmTab(),
    methodPath,
    pathError,
    activeJson,
    activeJsonSource: activeJson ? "dfm-ui-state" : "",
    activeJsonError,
    dirty: getDfmIsDirty(),
    fields: {
      project: inputSnap.resolved?.project || document.getElementById("projectSelect")?.value?.trim() || "",
      reservingClass: inputSnap.resolved?.reservingClass || document.getElementById("pathInput")?.value?.trim() || "",
      methodName: document.getElementById("dfmMethodName")?.value?.trim() || "",
      outputVector: document.getElementById("dfmOutputVector")?.value?.trim() || "",
      inputTriangle: document.getElementById("triInput")?.value?.trim() || "",
      originLength: document.getElementById("originLenSelect")?.value?.trim() || "",
      developmentLength: document.getElementById("devLenSelect")?.value?.trim() || "",
    },
  };
}

function postDfmStatus(text, tone = "") {
  window.parent.postMessage({ type: "arcrho:status", text: String(text || ""), tone }, "*");
}

function updateDfmSaveUi() {
  const saveBtn = document.getElementById("dfmSaveBtn");
  const cancelBtn = document.getElementById("dfmCancelBtn");
  const dirty = getDfmIsDirty();
  updateTabbedPageSaveControls({
    saveButton: saveBtn,
    cancelButton: cancelBtn,
    dirty,
    saving: dfmSaveInFlight,
  });
}

function requestConfirmedDfmClose() {
  clearDfmDependencyPreview("close-discard");
  requestTabbedPageWindowClose({
    messageType: "arcrho:dfm-close-confirmed",
    inst: getDfmInst(),
  });
}

function postCurrentDfmDirtyState() {
  const dirty = getDfmIsDirty();
  try {
    window.parent?.postMessage({
      type: "arcrho:dfm-dirty",
      inst: getDfmInst(),
      dirty,
    }, "*");
  } catch {}
}

function buildDfmDependencySourceMessage(type, reason = "") {
  const inputSnap = getDfmInputSnapshotSafe();
  const outputVector = document.getElementById("dfmOutputVector")?.value?.trim() || "";
  const methodName = document.getElementById("dfmMethodName")?.value?.trim() || outputVector;
  const model = dfmState?.model || null;
  const payload = {
    type,
    inst: getDfmInst(),
    project: inputSnap.resolved?.project || document.getElementById("projectSelect")?.value?.trim() || "",
    reservingClass: inputSnap.resolved?.reservingClass || document.getElementById("pathInput")?.value?.trim() || "",
    datasetName: outputVector || methodName,
    datasetTypeName: outputVector || methodName,
    names: [outputVector, methodName].filter(Boolean),
    methodType: "DFM",
    sourceKind: "dfm",
    dataFormat: "Vector",
    reason,
  };
  if (type === "arcrho:dependency-source-preview") {
    payload.values = buildResultsVector();
    // A dependent method reads its percentage developed from this pattern, so a
    // dirty DFM has to preview the pattern alongside the ultimates.
    payload.percentageDeveloped = buildPercentDevelopedVector();
    payload.originLabels = Array.isArray(model?.origin_labels) ? model.origin_labels.map(String) : [];
  }
  if (type === "arcrho:dependency-source-cleared") {
    // Present only when this clean transition came from a save that enqueued
    // an Engine propagation job; Project Instance defers the downstream
    // preview clear until that job reaches a terminal status.
    payload.propagationJobId = consumePendingDfmPropagationJobId();
  }
  return payload;
}

function postDfmDependencySourceMessage(type, reason = "") {
  const message = buildDfmDependencySourceMessage(type, reason);
  if (!message.datasetName && !message.names.length) return;
  try {
    window.parent?.postMessage(message, "*");
  } catch {}
}

function scheduleDfmDependencyPreview() {
  window.clearTimeout(dependencyPreviewTimer);
  dependencyPreviewTimer = window.setTimeout(() => {
    dependencyPreviewTimer = 0;
    if (getDfmIsDirty()) postDfmDependencySourceMessage("arcrho:dependency-source-preview", "dirty");
  }, 120);
}

function clearDfmDependencyPreview(reason = "") {
  window.clearTimeout(dependencyPreviewTimer);
  dependencyPreviewTimer = 0;
  postDfmDependencySourceMessage("arcrho:dependency-source-cleared", reason || "clean");
}

function requestDfmCloseFromShell() {
  if (dfmSaveInFlight) {
    postDfmStatus("Finish the current DFM save before closing the tab.", "error");
    return true;
  }
  if (!getDfmIsDirty()) return false;
  if (dfmCloseConfirm.isOpen) return true;
  void (async () => {
    const discard = await dfmCloseConfirm.confirm({ reason: "close" });
    if (discard) requestConfirmedDfmClose();
  })();
  return true;
}

async function saveCurrentDfmMethodFromBar() {
  if (dfmSaveInFlight) return;
  dfmSaveInFlight = true;
  updateDfmSaveUi();
  let savedCleanly = false;
  let refreshedDatasets = [];
  let linkWarnings = [];
  try {
    const result = await saveRatioSelectionPattern(false);
    if (!result?.ok && result?.error) {
      postDfmStatus(`DFM save failed: ${result.error}`, "error");
    }
    // A save keeps the window open; only after a clean dependent walk is there
    // a refreshed-dependents list to report.
    savedCleanly = Boolean(result?.ok && result?.propagationClean);
    refreshedDatasets = result?.refreshedDatasets || [];
    linkWarnings = result?.linkWarnings || [];
  } finally {
    dfmSaveInFlight = false;
    updateDfmSaveUi();
  }
  if (savedCleanly) await showSavedDependentsNotice(refreshedDatasets, { linkWarnings });
}

async function cancelCurrentDfmChangesFromBar() {
  if (dfmSaveInFlight) return;
  if (!getDfmIsDirty()) {
    requestConfirmedDfmClose();
    return;
  }
  const discard = await dfmCloseConfirm.confirm({ reason: "close" });
  if (!discard) return;
  const result = await restoreCleanDfmMethodState();
  if (result?.ok) {
    postDfmStatus("DFM changes discarded.");
    requestConfirmedDfmClose();
  } else {
    postDfmStatus(`DFM cancel failed: ${result?.error || "Could not restore saved method."}`, "error");
  }
  updateDfmSaveUi();
}

function wireDfmSaveControls() {
  document.getElementById("dfmSaveBtn")?.addEventListener("click", () => {
    void saveCurrentDfmMethodFromBar();
  });
  document.getElementById("dfmCancelBtn")?.addEventListener("click", () => {
    void cancelCurrentDfmChangesFromBar();
  });
  window.addEventListener("arcrho:dfm-dirty-state", updateDfmSaveUi);
  window.addEventListener("arcrho:dfm-dirty-state", (event) => {
    if (event?.detail?.dirty) scheduleDfmDependencyPreview();
    else clearDfmDependencyPreview("clean");
  });
  window.__arcrho_request_close = requestDfmCloseFromShell;
  window.__arcrho_consume_close_shortcut = requestDfmCloseFromShell;
  updateDfmSaveUi();
}

function openPathViaShellBridge(targetPath, preferredApp = "") {
  return new Promise((resolve) => {
    if (!targetPath || !window.parent || window.parent === window) {
      resolve({ ok: false, error: "Open path requires desktop app." });
      return;
    }
    const requestId = `dfm-open-json-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    let done = false;
    let timeoutId = null;
    const finish = (result) => {
      if (done) return;
      done = true;
      if (timeoutId != null) window.clearTimeout(timeoutId);
      window.removeEventListener("message", onMessage);
      resolve(result || { ok: false, error: "Open path failed." });
    };
    const onMessage = (evt) => {
      const msg = evt?.data;
      if (!msg || msg.type !== "arcrho:open-path-result") return;
      if (String(msg.requestId || "") !== requestId) return;
      finish({ ok: !!msg.ok, error: String(msg.error || "") });
    };
    window.addEventListener("message", onMessage);
    timeoutId = window.setTimeout(() => {
      finish({ ok: false, error: "Open path timed out." });
    }, 5000);
    try {
      window.parent.postMessage({ type: "arcrho:open-path", requestId, path: targetPath, preferredApp }, "*");
    } catch {
      finish({ ok: false, error: "Open path requires desktop app." });
    }
  });
}

function forwardChildOpenPathRequest(message, sourceWindow) {
  const source = sourceWindow || null;
  if (!source || source === window || source === window.parent) return false;

  const requestId = String(message?.requestId || "").trim();
  const path = String(message?.path || "").trim();
  const preferredApp = String(message?.preferredApp || "").trim();
  const readOnly = !!message?.readOnly;
  if (!requestId) return true;

  const replyToSource = (payload) => {
    try {
      source.postMessage({ type: "arcrho:open-path-result", requestId, ...payload }, "*");
    } catch {}
  };
  if (!path) {
    replyToSource({ ok: false, error: "Empty path." });
    return true;
  }
  if (!window.parent || window.parent === window) {
    replyToSource({ ok: false, error: "Open path requires desktop app." });
    return true;
  }

  let done = false;
  let timeoutId = null;
  const finish = (payload) => {
    if (done) return;
    done = true;
    if (timeoutId != null) window.clearTimeout(timeoutId);
    window.removeEventListener("message", onMessage);
    replyToSource(payload || { ok: false, error: "Open path failed." });
  };
  const onMessage = (evt) => {
    const msg = evt?.data;
    if (!msg || msg.type !== "arcrho:open-path-result") return;
    if (String(msg.requestId || "") !== requestId) return;
    finish({ ok: !!msg.ok, error: String(msg.error || "") });
  };
  window.addEventListener("message", onMessage);
  timeoutId = window.setTimeout(() => {
    finish({ ok: false, error: "Open path timed out." });
  }, 6000);
  try {
    window.parent.postMessage({ type: "arcrho:open-path", requestId, path, preferredApp, readOnly }, "*");
  } catch {
    finish({ ok: false, error: "Open path requires desktop app." });
  }
  return true;
}

async function openCurrentDfmMethodJson() {
  let methodPath = "";
  try {
    methodPath = await resolveCurrentDfmMethodSavePath();
  } catch (err) {
    postDfmStatus(`Open DFM JSON failed: ${String(err?.message || err)}`, "error");
    return;
  }
  if (!methodPath) {
    postDfmStatus("Open DFM JSON failed: no DFM JSON path is available.", "error");
    return;
  }
  try {
    const hostApi = window.ADAHost || null;
    const result = hostApi && typeof hostApi.openPath === "function"
      ? await hostApi.openPath({ path: methodPath, preferredApp: "arcode" })
      : await openPathViaShellBridge(methodPath, "arcode");
    if (result?.ok) {
      postDfmStatus(`Opened DFM JSON: ${methodPath}`);
    } else {
      postDfmStatus(`Open DFM JSON failed: ${result?.error || methodPath}`, "error");
    }
  } catch (err) {
    postDfmStatus(`Open DFM JSON failed: ${String(err?.message || err)}`, "error");
  }
}

function initDfmTabs() {
  const detailsPage = document.getElementById("dfmDetailsPage");
  const dataPage = document.getElementById("dfmDataPage");
  const ratiosPage = document.getElementById("dfmRatiosPage");
  const curvesPage = document.getElementById("dfmCurvesPage");
  const resultsPage = document.getElementById("dfmResultsPage");
  const notesPage = document.getElementById("dfmNotesPage");
  const linksPage = document.getElementById("dfmLinksPage");
  const auditPage = document.getElementById("dfmAuditPage");
  if (!detailsPage || !dataPage || !ratiosPage || !curvesPage || !resultsPage || !notesPage || !linksPage || !auditPage) return;

  applyHostFixedDetailsFields({ root: detailsPage });
  syncDetailsLabelWidth({
    root: detailsPage,
    labelSelector: ".arDetailsLabel",
  });
  wireDfmScrollbarActivity(detailsPage);
  wireDfmScrollbarActivity(document.getElementById("ratioWrapHost"));
  wireDfmScrollbarActivity(document.getElementById("dfmCurvesWrapHost"));
  wireDfmScrollbarActivity(document.getElementById("resultsWrap"));
  setShowNaBorders(loadNaBorders());

  wireDfmSpinnerControls();
  wireMethodName();
  wireDfmInstanceCreationNotice();
  wireNotesInput();
  initDfmLinks();
  wireDfmSaveControls();
  wireDetailsThresholdReset();
  wireRatioStrikeToggle();
  wireRatioChartModal();
  wireRatioContextMenu();
  wireResultsRatioBasisControls();
  initDfmCurvesTab();
  initDfmAuditLog();

  const params = new URLSearchParams(window.location.search);
  const urlTab = params.get("tab");
  const initialTab = ALLOWED_DFM_TABS.has(urlTab) ? urlTab : "details";

  const tabSystem = createTabbedPage(document.body, {
    tabs: DFM_TAB_DEFS,
    cssPrefix: "dfm",
    initialTab,
    injectTabBar: false,
    shortcutBlockedSelector: ".dfmRpcOverlay, [aria-modal='true']",
    previousTabMessageTypes: ["arcrho:dfm-tab-prev"],
    nextTabMessageTypes: ["arcrho:dfm-tab-next"],
    onTabChange: (tabId) => {
      setCurrentDfmTab(tabId);
      if (tabId === "ratios") renderRatioTable();
      if (tabId === "curves") renderDfmCurvesTab();
      if (tabId === "results") renderResultsTable();
      if (tabId === "links") refreshDfmLinks();
      if (tabId === "audit") refreshDfmAuditLog();
      notifyDfmEditState();
      if (tabId === "details" && !persistedDfmBootstrap) {
        syncMethodNameFromInputs();
        syncOutputTypeFromProject();
      }
      const inst = getDfmInst();
      window.parent.postMessage({ type: "arcrho:dfm-tab-changed", inst, tab: tabId }, "*");
    }
  });
  applyTabbedPageSaveBar(document.getElementById("dfmSaveBar"));

  window.dfmTabSystem = tabSystem;
  wireDfmTabPopoutWindows({
    onPopoutTab: (tabId) => {
      if (tabId === "ratios") renderRatioTable();
      if (tabId === "curves") renderDfmCurvesTab();
      if (tabId === "results") renderResultsTable();
      if (tabId === "links") refreshDfmLinks();
      if (tabId === "audit") refreshDfmAuditLog();
      notifyDfmEditState();
    },
  });
}

export function initDfmRatios() {
  setStorageInstance(getDfmInst());
  initDfmTabs();
  notifyDfmEditState();
  if (!persistedDfmBootstrap) {
    syncMethodNameFromInputs();
    syncOutputTypeFromProject();
  }
  wireDfmRpcBridgeTabBar();
  if (!persistedDfmBootstrap) {
    setTimeout(() => {
      syncOutputTypeFromProject();
    }, 500);
  }

  window.addEventListener("arcrho:workflow-defaults-updated", () => {
    if (persistedDfmBootstrap) return;
    syncMethodNameFromInputs();
    syncOutputTypeFromProject();
  });
  wireRatioSyncChannel();
  requestRatioStateSync();
  initRatioHistory({
    afterRestore: () => {
      restoreRatioHistoryUi();
      if (isRatioChartOpen()) scheduleRatioChartRender();
    },
  });
  startDfmMethodFileWatcher();
  window.addEventListener("beforeunload", () => {
    stopDfmMethodFileWatcher();
    cancelDfmMethodAsyncTasks();
    clearRatioHistoryTempSession();
  }, { once: true });

  window.addEventListener("arcrho:dataset-updated", handleDatasetUpdated);
  window.addEventListener("arcrho:dfm-owned-state-mutated", scheduleDfmMethodPreview);
  window.addEventListener("arcrho:dfm-dirty-state", (event) => {
    if (event?.detail?.dirty) scheduleDfmMethodPreview();
  });

  /* ---- Apply project/class from URL params when embedded in workflow ---- */
  const _qs = new URLSearchParams(window.location.search);
  const {
    project: _urlProject,
    path: _urlClass,
    methodName: _urlMethodName,
    tri: _urlInputTriangle,
  } = readDatasetInputQueryValues(_qs);
  if (_urlProject || _urlClass || _urlMethodName || _urlInputTriangle) {
    const projEl = document.getElementById("projectSelect");
    const classEl = document.getElementById("pathInput");
    const methodEl = document.getElementById("dfmMethodName");
    const triEl = document.getElementById("triInput");
    if (_urlProject && projEl) projEl.value = _urlProject;
    if (_urlClass && classEl) classEl.value = _urlClass;
    if (_urlMethodName && methodEl) methodEl.value = _urlMethodName;
    if (_urlInputTriangle && triEl) triEl.value = _urlInputTriangle;
    if (!persistedDfmBootstrap) {
      syncMethodNameFromInputs();
      syncOutputTypeFromProject({ forceReload: true });
    }
  }
  refreshDfmTabContent("dfm-open");
  recordCurrentDfmCleanState();

  window.addEventListener("message", (e) => {
    if (e?.data?.type === "arcrho:open-path" && forwardChildOpenPathRequest(e.data, e.source)) {
      return;
    }
    if (e?.data?.type === "arcrho:calculated-datasets-updated") {
      handleDfmPropagationReport(e.data?.report || null);
      return;
    }
    /* Respond to workflow requesting DFM step settings for snapshot */
    if (e?.data?.type === "arcrho:get-dfm-settings") {
      const inputSnap = getDfmInputSnapshotSafe();
      const settings = {
        project: inputSnap.defaults?.projectDefault
          ? DEFAULT_TOKEN
          : (inputSnap.resolved?.project || document.getElementById("projectSelect")?.value?.trim() || ""),
        reservingClass: inputSnap.defaults?.reservingClassDefault
          ? DEFAULT_TOKEN
          : (inputSnap.resolved?.reservingClass || document.getElementById("pathInput")?.value?.trim() || ""),
        objectName: document.getElementById("dfmMethodName")?.value?.trim() || "",
        outputType: document.getElementById("dfmOutputVector")?.value?.trim() || "",
        originLen: document.getElementById("originLenSelect")?.value?.trim() || "",
        devLen: document.getElementById("devLenSelect")?.value?.trim() || "",
      };
      window.parent.postMessage({ type: "arcrho:dfm-settings", settings, requestId: e.data.requestId }, "*");
      return;
    }
    /* Handle global control changes from workflow */
    if (e?.data?.type === "arcrho:workflow-global-changed") {
      const inputSnap = getDfmInputSnapshotSafe();
      if (!inputSnap.defaults?.projectDefault && !inputSnap.defaults?.reservingClassDefault) return;
      syncMethodNameFromInputs();
      syncOutputTypeFromProject({ forceReload: true });
      scheduleRatioSelectionLoad("global-changed");
      return;
    }
    if (e?.data?.type === "arcrho:server-connection-updated") {
      setCachedRootPath(e.data.config?.workspace_root || "");
      window.parent.postMessage({ type: "arcrho:status", text: "Server connection updated." }, "*");
      return;
    }
    if (e?.data?.type === "arcrho:assistant-context-request") {
      const requestId = e.data.requestId || "";
      buildAssistantContext()
        .then((context) => {
          window.parent.postMessage({ type: "arcrho:assistant-context-result", requestId, context }, "*");
        })
        .catch((err) => {
          window.parent.postMessage({
            type: "arcrho:assistant-context-result",
            requestId,
            context: {
              available: false,
              pageType: "dfm",
              error: String(err?.message || err || "DFM assistant context failed."),
            },
          }, "*");
        });
      return;
    }
    if (e?.data?.type === "arcrho:assistant-json-updated") {
      scheduleRatioSelectionLoad("assistant-edit");
      return;
    }
    if (e?.data?.type === "arcrho:assistant-dfm-edit-approval") {
      const requestId = e.data.requestId || "";
      const reply = (payload) => {
        try {
          window.parent.postMessage({
            type: "arcrho:assistant-dfm-edit-approval-result",
            requestId,
            ...payload,
          }, "*");
        } catch {
          // ignore stale shell messaging
        }
      };
      reviewArcBotDfmEditApproval({
        targetPath: e.data.targetPath || "",
        originalJson: e.data.originalJson || null,
        proposedJson: e.data.proposedJson || null,
        reply: e.data.reply || "",
      })
        .then(reply)
        .catch((err) => reply({ ok: false, error: String(err?.message || err || "Could not review ArcBot DFM edit.") }));
      return;
    }
    if (e?.data?.type === "arcrho:dfm-apply-method-payload") {
      const requestId = e.data.requestId || "";
      const reply = (payload) => {
        try {
          window.parent.postMessage({
            type: "arcrho:dfm-apply-method-payload-result",
            requestId,
            ...payload,
          }, "*");
        } catch {
          // ignore stale shell messaging
        }
      };
      applyDfmOwnedPatchPayload(e.data.payload, { reason: "macro" })
        .then((applied) => {
          if (applied?.ok) {
            markDfmDirty();
            postDfmStatus("Macro applied to active DFM.");
            reply({ ok: true });
          } else {
            reply({ ok: false, error: "Could not apply macro result to DFM tab." });
          }
        })
        .catch((err) => reply({ ok: false, error: String(err?.message || err || "Could not apply macro result.") }));
      return;
    }
    if (e?.data?.type === "arcrho:dfm-request-state") {
      notifyDfmEditState();
      postCurrentDfmDirtyState();
      const history = getRatioHistoryState();
      window.parent.postMessage({
        type: "arcrho:dfm-history-state",
        inst: getDfmInst(),
        canUndo: history.canUndo,
        canRedo: history.canRedo,
      }, "*");
      return;
    }
    if (e?.data?.type === "arcrho:dfm-tab-activated") {
      notifyDfmEditState();
      const history = getRatioHistoryState();
      window.parent.postMessage({
        type: "arcrho:dfm-history-state",
        inst: getDfmInst(),
        canUndo: history.canUndo,
        canRedo: history.canRedo,
      }, "*");
      return;
    }
    if (e?.data?.type === "arcrho:dfm-exclude-high") {
      const ratiosPage = document.getElementById("dfmRatiosPage");
      if (!ratiosPage || ratiosPage.style.display === "none") return;
      excludeExtremeInActiveCol("high");
      return;
    }
    if (e?.data?.type === "arcrho:dfm-exclude-low") {
      const ratiosPage = document.getElementById("dfmRatiosPage");
      if (!ratiosPage || ratiosPage.style.display === "none") return;
      excludeExtremeInActiveCol("low");
      return;
    }
    if (e?.data?.type === "arcrho:dfm-include-all") {
      const ratiosPage = document.getElementById("dfmRatiosPage");
      if (!ratiosPage || ratiosPage.style.display === "none") return;
      includeAllInActiveCol();
      return;
    }
    if (e?.data?.type === "arcrho:dfm-undo") {
      const ratiosPage = document.getElementById("dfmRatiosPage");
      if (!ratiosPage || ratiosPage.style.display === "none") return;
      runRatioUndo();
      return;
    }
    if (e?.data?.type === "arcrho:dfm-redo") {
      const ratiosPage = document.getElementById("dfmRatiosPage");
      if (!ratiosPage || ratiosPage.style.display === "none") return;
      runRatioRedo();
      return;
    }
    if (e?.data?.type === "arcrho:dfm-tab-closing") {
      clearRatioHistoryTempSession();
      return;
    }
    if (e?.data?.type === "arcrho:dfm-open-method-json") {
      openCurrentDfmMethodJson();
      return;
    }
    if (e?.data?.type === "arcrho:dfm-save") {
      // Route through the bar handler so the explicit Save command shares the
      // in-flight guard and the close-on-clean-propagation behavior.
      void saveCurrentDfmMethodFromBar();
      return;
    }
    if (e?.data?.type === "arcrho:dfm-save-as") {
      saveRatioSelectionPattern(true);
      return;
    }
    if (e?.data?.type === "arcrho:dfm-save-template") {
      saveDfmTemplate();
      return;
    }
  });

  window.addEventListener("keydown", (e) => {
    if (!e.ctrlKey || e.altKey || e.metaKey) return;
    const key = (e.key || "").toLowerCase();
    if (key !== "h" && key !== "l" && key !== "i" && key !== "z" && key !== "y") return;
    const tag = e.target?.tagName?.toLowerCase();
    if (tag === "input" || tag === "textarea" || tag === "select" || e.target?.isContentEditable) return;
    const ratiosPage = document.getElementById("dfmRatiosPage");
    if (!ratiosPage || ratiosPage.style.display === "none") return;
    e.preventDefault();
    if (key === "h") excludeExtremeInActiveCol("high");
    if (key === "l") excludeExtremeInActiveCol("low");
    if (key === "i") includeAllInActiveCol();
    if (key === "z") runRatioUndo();
    if (key === "y") runRatioRedo();
  }, { capture: true });
}
