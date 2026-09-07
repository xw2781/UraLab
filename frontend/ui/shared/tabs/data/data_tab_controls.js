import {
  applyDecimalPlacesToDatasetNumberFormat,
  clampDatasetDecimalPlaces,
  getDatasetNumberFormatDecimalPlaces,
  normalizeDatasetNumberFormat,
} from "/ui/shared/dataset/dataset_number_format.js";
import { wireNumberFormatField } from "/ui/shared/components/pickers/number_format_field.js?v=20260817a";

function wireChartPanelResize(redrawChartSafely) {
  const panel = document.getElementById("chartPanel");
  if (!panel || typeof ResizeObserver !== "function") return;
  if (panel.__datasetChartResizeObserver) return;

  let resizeFrame = 0;
  const scheduleRedraw = () => {
    if (resizeFrame) return;
    resizeFrame = requestAnimationFrame(() => {
      resizeFrame = 0;
      redrawChartSafely();
    });
  };

  const observer = new ResizeObserver(scheduleRedraw);
  observer.observe(panel);
  panel.__datasetChartResizeObserver = observer;
}

export function wireDatasetInputController(deps) {
  const {
    state,
    $,
    loadDataset,
    isRunInFlight,
    setStatus,
    runArcRhoTri,
    savePatch,
    toggleBlanks,
    wireLenDropdowns,
    syncDetailDatasetTypeFromTopInput,
    clearInputInvalid,
    showProjectDropdown,
    openProjectNameTreeForDataset,
    showDatasetDropdown,
    openDatasetNameTreeForDataset,
    saveTriInputsToStorage,
    scheduleAutoRun,
    renderTable,
    notifyDatasetUpdated,
    renderChart,
    isDefaultTokenValue,
    setInputDefaultBound,
    getResolvedProjectValue,
    validateAndNormalizeReservingClassInput,
    filterDatasetOptions,
    getActiveDatasetIndex,
    setActiveDatasetIndex,
    chooseActiveDataset,
    validateAndNormalizeDatasetInput,
    validateDatasetTypeDependencies,
    handleDatasetSelection,
    setLastDatasetSelection,
    filterProjectOptions,
    getProjectFilterQuery,
    getActiveProjectIndex,
    setActiveProjectIndex,
    chooseActiveProject,
    handleProjectSelection,
    setLastProjectSelection,
    LEN_DROPDOWN_CONFIG,
    closeAllLenDropdowns,
    enforceDevLenRule,
    ensureHeadersForProject,
    ensureDevHeadersForProject,
    bindAutoRunOnEnter,
    redrawChartSafely,
    wireDatasetHostBridge,
    getTriInputsForStorage,
    syncSidecarForCurrentDataset,
    instanceId,
    wireGridInteractions,
    isProjectInstanceDraft = false,
    refreshProjectInstanceDraftModel = null,
    validateManualDatasetLengthChange = null,
    isManualDatasetModeLocked = null,
    restoreManualDatasetModeControls = null,
    chooseStoredDevelopmentLength = null,
    refreshDatasetSettingsDirty = null,
  } = deps;

  async function refreshDraftModelAfterInputChange() {
    if (!isProjectInstanceDraft || typeof refreshProjectInstanceDraftModel !== "function") return false;
    await refreshProjectInstanceDraftModel();
    return true;
  }

  document.getElementById("reloadBtn")?.addEventListener("click", loadDataset);
  document.getElementById("clearCacheReloadBtn")?.addEventListener("click", () => {
    if (isRunInFlight()) return;
    setStatus("Clearing cache and reloading dataset...");
    void runArcRhoTri({ clearCache: true, showValidationMessage: true });
  });
  const saveBtn = $("saveBtn");
  if (window.location.search.includes("readonly=1")) {
    saveBtn.disabled = true;
    saveBtn.title = "Generated datasets are read-only.";
  }
  saveBtn.addEventListener("click", savePatch);
  $("toggleBlankBtn").addEventListener("click", toggleBlanks);

  const pathInput = document.getElementById("pathInput");
  const triInput = document.getElementById("triInput");
  const datasetTreeBtn = document.getElementById("datasetTreeBtn");
  const projectSelect = document.getElementById("projectSelect");
  const projectTreeBtn = document.getElementById("projectTreeBtn");
  const originSel = document.getElementById("originLenSelect");
  const devSel = document.getElementById("devLenSelect");
  wireLenDropdowns();

  // Name is auto-copied only when Dataset Type switches. On this first call the
  // remembered type is still blank, so every load would read as a switch. Seed the
  // remembered type instead, and only copy into an empty Name (a new dataset draft)
  // so a loaded instance name that differs from its type survives.
  if (triInput) {
    const detailName = document.getElementById("dsDetailName");
    const hasInstanceName = !!String(detailName?.value || "").trim();
    syncDetailDatasetTypeFromTopInput(triInput.value, { syncName: !hasInstanceName });
  }

  if (projectTreeBtn && projectSelect) {
    projectTreeBtn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      showProjectDropdown(false);
      void openProjectNameTreeForDataset(projectSelect);
    });
  }

  if (datasetTreeBtn && triInput) {
    datasetTreeBtn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      showDatasetDropdown(false);
      void openDatasetNameTreeForDataset(triInput);
    });
  }

  const cumulativeChk = document.getElementById("cumulativeChk");
  if (cumulativeChk) {
    cumulativeChk.addEventListener("change", () => {
      if (typeof isManualDatasetModeLocked === "function" && isManualDatasetModeLocked()) {
        if (typeof restoreManualDatasetModeControls === "function") restoreManualDatasetModeControls();
        setStatus("Manual input Triangle/Vector datasets keep their cumulative mode fixed.");
        return;
      }
      saveTriInputsToStorage();
      scheduleAutoRun(0);
    });
  }

  const transposedChk = document.getElementById("transposedChk");
  if (transposedChk) {
    transposedChk.addEventListener("change", () => {
      state.activeCell = null;
      state.selRanges = [];
      saveTriInputsToStorage();
      renderTable();
      notifyDatasetUpdated();
      renderChart();
    });
  }

  const timeModeInputs = Array.from(document.querySelectorAll('input[name="timeMode"]'));
  for (const input of timeModeInputs) {
    input.addEventListener("change", async () => {
      if (!input.checked) return;
      if (typeof isManualDatasetModeLocked === "function" && isManualDatasetModeLocked()) {
        if (typeof restoreManualDatasetModeControls === "function") restoreManualDatasetModeControls();
        setStatus("Manual input Triangle/Vector datasets keep their development/calendar mode fixed.");
        return;
      }
      saveTriInputsToStorage();
      const project = getResolvedProjectValue();
      if (project) {
        await ensureDevHeadersForProject(project, { forceRefresh: true });
      }
      renderTable();
      notifyDatasetUpdated();
      setStatus("Loading dataset...");
      scheduleAutoRun(0);
    });
  }

  const dec = document.getElementById("decimalPlaces");
  const numberFormatSelect = document.getElementById("numberFormatSelect");
  const numberFormatWrap = document.getElementById("numberFormatWrap");
  const numberFormatDropdownBtn = document.getElementById("numberFormatDropdownBtn");
  const numberFormatDropdown = document.getElementById("numberFormatDropdown");
  const decimalPlacesUpBtn = document.getElementById("decimalPlacesUpBtn");
  const decimalPlacesDownBtn = document.getElementById("decimalPlacesDownBtn");

  function refreshNumberDisplaySettings() {
    saveTriInputsToStorage();
    renderTable();
    notifyDatasetUpdated();
    renderChart();
  }

  function syncNumberFormatFromDecimalPlaces() {
    if (!numberFormatSelect || !dec) return;
    const places = clampDatasetDecimalPlaces(dec.value);
    dec.value = String(places);
    numberFormatSelect.value = applyDecimalPlacesToDatasetNumberFormat(numberFormatSelect.value, places);
  }

  function syncDecimalPlacesFromNumberFormat() {
    if (!numberFormatSelect || !dec) return;
    numberFormatSelect.value = normalizeDatasetNumberFormat(numberFormatSelect.value);
    dec.value = String(clampDatasetDecimalPlaces(getDatasetNumberFormatDecimalPlaces(numberFormatSelect.value)));
  }

  function stepDecimalPlaces(delta) {
    if (!dec) return;
    const next = clampDatasetDecimalPlaces((Number.parseInt(dec.value, 10) || 0) + delta);
    dec.value = String(next);
    syncNumberFormatFromDecimalPlaces();
    refreshNumberDisplaySettings();
    dec.focus();
  }

  const numberFormatField = wireNumberFormatField({
    input: numberFormatSelect,
    field: numberFormatWrap,
    toggle: numberFormatDropdownBtn,
    menu: numberFormatDropdown,
    onApply: (preset) => {
      if (!numberFormatSelect) return;
      numberFormatSelect.value = preset;
      syncDecimalPlacesFromNumberFormat();
      refreshNumberDisplaySettings();
    },
  });
  const closeNumberFormatDropdown = () => numberFormatField?.close();

  if (dec && numberFormatSelect) {
    dec.addEventListener("change", () => {
      syncNumberFormatFromDecimalPlaces();
      refreshNumberDisplaySettings();
    });
    dec.addEventListener("input", () => {
      syncNumberFormatFromDecimalPlaces();
      refreshNumberDisplaySettings();
    });
  }
  decimalPlacesUpBtn?.addEventListener("click", (e) => {
    e.preventDefault();
    stepDecimalPlaces(1);
  });
  decimalPlacesDownBtn?.addEventListener("click", (e) => {
    e.preventDefault();
    stepDecimalPlaces(-1);
  });

  if (numberFormatSelect) {
    numberFormatSelect.addEventListener("change", () => {
      syncDecimalPlacesFromNumberFormat();
      refreshNumberDisplaySettings();
      closeNumberFormatDropdown();
    });
    numberFormatSelect.addEventListener("input", () => {
      refreshNumberDisplaySettings();
    });
  }
  syncNumberFormatFromDecimalPlaces();

  // Chart mode toggle
  const chartToggle = document.getElementById("chartModeToggle");
  if (chartToggle) {
    chartToggle.addEventListener("click", (e) => {
      const btn = e.target.closest(".chartToggleBtn");
      if (!btn) return;
      const mode = btn.dataset.mode;
      if (mode && mode !== state.chartMode) {
        // Reset legend state when switching modes
        const legendEl = document.getElementById("devChartLegend");
        if (legendEl?.__chartLegendState) {
          legendEl.__chartLegendState.hoverIndex = null;
          legendEl.__chartLegendState.selectedIndex = null;
          legendEl.__chartLegendState.hiddenSet = new Set();
        }
        state.chartMode = mode;
        renderChart();
      }
    });
  }

  // change -> auto run
  if (pathInput) {
    pathInput.addEventListener("change", async () => {
      if (isDefaultTokenValue(pathInput.value)) {
        setInputDefaultBound(pathInput, true);
      } else {
        setInputDefaultBound(pathInput, false);
      }
      const project = getResolvedProjectValue();
      const pathResult = await validateAndNormalizeReservingClassInput(project, { strict: true, showMessage: true });
      if (!pathResult.ok) return;
      saveTriInputsToStorage();
      await syncSidecarForCurrentDataset?.({ applyLengths: true });
      setStatus("Loading dataset...");
      scheduleAutoRun();
    });
    pathInput.addEventListener("input", () => {
      if (!isDefaultTokenValue(pathInput.value)) {
        setInputDefaultBound(pathInput, false);
      }
      clearInputInvalid(pathInput);
    });
  }
  if (triInput) {
    // Typing filters the list and ArrowDown opens it; putting the caret in the
    // field does not, so the browse button stays the only pointer path in.
    triInput.addEventListener("keydown", (e) => {
      if (e.key === "ArrowDown" || e.key === "ArrowUp") {
        const list = document.getElementById("datasetDropdown");
        if (!list || !list.classList.contains("open")) {
          filterDatasetOptions(triInput.value);
        }
        const dir = e.key === "ArrowDown" ? 1 : -1;
        const activeDatasetIndex = getActiveDatasetIndex();
        if (activeDatasetIndex === -1) {
          setActiveDatasetIndex(dir > 0 ? 0 : -1);
        } else {
          setActiveDatasetIndex(activeDatasetIndex + dir);
        }
        e.preventDefault();
        return;
      }

      if (e.key === "Enter") {
        if (chooseActiveDataset()) {
          e.preventDefault();
          return;
        }
        const datasetResult = validateAndNormalizeDatasetInput({ strict: true, showMessage: true });
        if (!datasetResult.ok) {
          e.preventDefault();
          return;
        }
        void (async () => {
          const dependencyResult = await validateDatasetTypeDependencies(datasetResult.value, { showMessage: true });
          if (!dependencyResult.ok) return;
          saveTriInputsToStorage();
          await syncSidecarForCurrentDataset?.({ applyLengths: true });
          setStatus("Loading dataset...");
          scheduleAutoRun(0);
        })();
        return;
      }

      if (e.key === "Escape") {
        showDatasetDropdown(false);
      }
    });

    triInput.addEventListener("input", () => {
      clearInputInvalid(triInput);
      filterDatasetOptions(triInput.value);
      if (!triInput.value.trim()) setLastDatasetSelection("");
      void handleDatasetSelection(triInput.value);
    });

    triInput.addEventListener("change", async () => {
      const datasetResult = validateAndNormalizeDatasetInput({ strict: true, showMessage: true });
      if (!datasetResult.ok) {
        showDatasetDropdown(false);
        return;
      }
      syncDetailDatasetTypeFromTopInput(datasetResult.value, { syncName: true });
      const dependencyResult = await validateDatasetTypeDependencies(datasetResult.value, { showMessage: true });
      if (!dependencyResult.ok) {
        showDatasetDropdown(false);
        return;
      }
      setLastDatasetSelection(datasetResult.value);
      saveTriInputsToStorage();
      await syncSidecarForCurrentDataset?.({ applyLengths: true });
      setStatus("Loading dataset...");
      scheduleAutoRun();
      showDatasetDropdown(false);
    });
  }

  if (projectSelect) {
    projectSelect.addEventListener("keydown", (e) => {
      if (e.key === "ArrowDown" || e.key === "ArrowUp") {
        const list = document.getElementById("projectDropdown");
        if (!list || !list.classList.contains("open")) {
          filterProjectOptions(getProjectFilterQuery(projectSelect));
        }
        const dir = e.key === "ArrowDown" ? 1 : -1;
        const activeProjectIndex = getActiveProjectIndex();
        if (activeProjectIndex === -1) {
          setActiveProjectIndex(dir > 0 ? 0 : -1);
        } else {
          setActiveProjectIndex(activeProjectIndex + dir);
        }
        e.preventDefault();
        return;
      }

      if (e.key === "Enter") {
        if (chooseActiveProject()) {
          e.preventDefault();
          return;
        }
        void (async () => {
          const ok = await handleProjectSelection(projectSelect.value, { strict: true, showMessage: true });
          if (ok) setStatus("Loading dataset...");
        })();
        e.preventDefault();
        return;
      }

      if (e.key === "Escape") {
        showProjectDropdown(false);
      }
    });

    projectSelect.addEventListener("input", () => {
      if (!isDefaultTokenValue(projectSelect.value)) {
        setInputDefaultBound(projectSelect, false);
      }
      clearInputInvalid(projectSelect);
      filterProjectOptions(getProjectFilterQuery(projectSelect));
      if (!projectSelect.value.trim()) setLastProjectSelection("");
      void handleProjectSelection(projectSelect.value);
    });

    projectSelect.addEventListener("change", async () => {
      if (!projectSelect.value.trim()) return;
      const ok = await handleProjectSelection(projectSelect.value, { strict: true, showMessage: true });
      if (!ok) return;
      setStatus("Loading dataset...");
    });
  }

  document.addEventListener("mousedown", (e) => {
    const projectWrap = document.querySelector(".projectSelectWrap");
    if (projectWrap && !projectWrap.contains(e.target)) {
      showProjectDropdown(false);
    }
    const datasetWrap = document.querySelector(".datasetSelectWrap");
    if (datasetWrap && !datasetWrap.contains(e.target)) {
      showDatasetDropdown(false);
    }
    const inLenWrap = Object.values(LEN_DROPDOWN_CONFIG).some((cfg) => {
      const wrap = document.getElementById(cfg.wrapId);
      return !!wrap && wrap.contains(e.target);
    });
    if (!inLenWrap) closeAllLenDropdowns();
  });

  // Origin change -> enforce rule -> refresh headers -> auto run
  if (originSel) {
    originSel.addEventListener("change", async () => {
      enforceDevLenRule({ source: "origin" });
      if (typeof validateManualDatasetLengthChange === "function" && !validateManualDatasetLengthChange()) {
        originSel.blur();
        return;
      }
      saveTriInputsToStorage();

      if (await refreshDraftModelAfterInputChange()) {
        originSel.blur();
        return;
      }

      const project = getResolvedProjectValue();
      await ensureHeadersForProject(project);
      await ensureDevHeadersForProject(project);
      renderTable();
      notifyDatasetUpdated();
      setStatus("Loading dataset...");
      scheduleAutoRun(0);
      originSel.blur();
    });
  }

  // Dev change -> enforce rule -> refresh dev headers -> auto run
  if (devSel) {
    devSel.addEventListener("change", async () => {
      enforceDevLenRule({ source: "dev" });
      if (typeof validateManualDatasetLengthChange === "function" && !validateManualDatasetLengthChange()) {
        devSel.blur();
        return;
      }
      saveTriInputsToStorage();

      if (await refreshDraftModelAfterInputChange()) {
        devSel.blur();
        return;
      }

      const project = getResolvedProjectValue();
      if (project) {
        await ensureHeadersForProject(project);
        await ensureDevHeadersForProject(project);
      }

      renderTable();
      notifyDatasetUpdated();
      setStatus("Loading dataset...");
      scheduleAutoRun(0);
      devSel.blur();
    });
  }

  // Lowering the development `Stored at` leaves the display alone: the grid
  // keeps the shape on screen, and the finer periods exist only in the file the
  // next save writes.
  const devStoredSel = document.getElementById("devStoredLenSelect");
  if (devStoredSel && typeof chooseStoredDevelopmentLength === "function") {
    devStoredSel.addEventListener("change", () => {
      chooseStoredDevelopmentLength(devStoredSel.value);
      if (typeof refreshDatasetSettingsDirty === "function") refreshDatasetSettingsDirty();
      devStoredSel.blur();
    });
  }

  // Enter -> auto run
  bindAutoRunOnEnter(pathInput);
  // Run button still as fallback
  const runBtn = document.getElementById("runArcRhoTriBtn");
  if (runBtn) {
    runBtn.addEventListener("click", () => {
      void runArcRhoTri({ showValidationMessage: true });
    });
  }

  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) {
      // wait for layout to settle
      requestAnimationFrame(() => {
        requestAnimationFrame(redrawChartSafely);
      });
    }
  });

  window.addEventListener("resize", () => {
    requestAnimationFrame(redrawChartSafely);
  });
  wireChartPanelResize(redrawChartSafely);

  wireDatasetHostBridge({
    getTriInputsForStorage,
    instanceId,
    redrawChartSafely,
  });

  wireGridInteractions();
}
