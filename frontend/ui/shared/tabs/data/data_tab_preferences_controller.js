// Owns Dataset Viewer preferences, saved inputs, browsing history, and name availability.

export function registerDataTabPreferencesController(runtime) {
  const { BROWSING_HISTORY_MAX_ENTRIES, DEFAULT_PATH_DISPLAY, DEFAULT_PROJECT_DISPLAY, DEFAULT_TOKEN, instanceId, isProjectInstanceDraft, isReadOnlyDatasetViewer, isTemporaryDatasetView, LOCAL_PROJECT_PREFS_ENDPOINT, LS_DS_KEY, LS_FORM_KEY, scopedKey, WF_GLOBAL_CTRL_PREFIX, workflowId } = runtime;
  const defer = (name) => (...args) => runtime[name](...args);
  const { normalizeReservingClassPath, isDfmDataTabHost, normalizeProjectText, loadProjectUserPreferences, scheduleProjectUserPreferencesSave, updateDatasetSaveUi, normalizeBrowsingHistoryEntry, pushBrowsingHistoryEntry, getDatasetDecimalPlacesValue, getDatasetSyncedNumberFormatValue, findExactProjectMatch, setLastViewedDatasetInputs, refreshDatasetSettingsDirty, getLastViewedDatasetInputs, readDatasetInputsFromQueryParams, refreshLenDropdowns, datasetOriginDisplayIsCoarserThanStored, datasetCoarserViewMessage, setDatasetDecimalPlacesValue, setDatasetNumberFormatValue, ensureHeadersForProject, ensureDevHeadersForProject, refreshDatasetTypesForProject, refreshReservingClassPathsForProject, renderProjectOptions, scheduleAutoRun } = new Proxy({}, { get: (_target, name) => defer(name) });
  const datasetProjectPrefs = new Map();
  let localDatasetViewerPrefsLoadPromise = null;
  let localDatasetViewerProjectSaved = "";
  let cachedDatasetInstanceRows = [];
  let cachedDatasetInstanceKey = "";
  let cachedDatasetInstanceLoadPromise = null;
  function normalizeDatasetViewerPrefs(raw, projectFallback = "", sharedReservingClassPath = "") {
    const source = raw && typeof raw === "object" ? raw : {};
    const project = String(source.project || source.project_name || projectFallback || "").trim();
    const path = normalizeReservingClassPath(
      sharedReservingClassPath
      || source.path
      || source.reservingClass
      || source.reserving_class
      || "",
    );
    const tri = String(source.tri || source.datasetName || source.dataset_name || "").trim();
    if (!project) return null;
    return { project, path, tri };
  }

  function normalizeLocalDatasetViewerPrefs(raw) {
    const prefs = raw && typeof raw === "object" ? raw : {};
    const project = String(
      prefs.projectName
      || prefs.project_name
      || prefs.project
      || "",
    ).trim();
    return { project };
  }

  async function loadLastDatasetViewerProjectFromAppData() {
    if (isDfmDataTabHost()) return "";
    if (localDatasetViewerPrefsLoadPromise) return localDatasetViewerPrefsLoadPromise;
    localDatasetViewerPrefsLoadPromise = (async () => {
      try {
        const res = await fetch(LOCAL_PROJECT_PREFS_ENDPOINT, { cache: "no-store" });
        if (!res.ok) return "";
        const payload = await res.json().catch(() => ({}));
        const normalized = normalizeLocalDatasetViewerPrefs(payload?.preferences || payload);
        localDatasetViewerProjectSaved = normalized.project;
        return normalized.project;
      } catch {
        return "";
      } finally {
        localDatasetViewerPrefsLoadPromise = null;
      }
    })();
    return localDatasetViewerPrefsLoadPromise;
  }

  function saveLastDatasetViewerProjectToAppData(projectName) {
    if (isDfmDataTabHost()) return;
    const project = String(projectName || "").trim();
    if (!project || normalizeProjectText(project) === normalizeProjectText(localDatasetViewerProjectSaved)) return;
    localDatasetViewerProjectSaved = project;
    void (async () => {
      try {
        const res = await fetch(LOCAL_PROJECT_PREFS_ENDPOINT, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            projectName: project,
            updated_at: new Date().toISOString(),
          }),
        });
        if (!res.ok) localDatasetViewerProjectSaved = "";
      } catch {
        localDatasetViewerProjectSaved = "";
      }
    })();
  }

  async function loadDatasetProjectPrefs(projectName, options = {}) {
    const project = String(projectName || "").trim();
    if (!project) return null;
    const key = normalizeProjectText(project);
    if (!options?.forceReload && datasetProjectPrefs.has(key)) return datasetProjectPrefs.get(key);
    try {
      const prefs = await loadProjectUserPreferences(project, options);
      const normalized = normalizeDatasetViewerPrefs(
        prefs?.datasetViewer,
        project,
        prefs?.lastReservingClassPath || prefs?.last_reserving_class_path || "",
      );
      datasetProjectPrefs.set(key, normalized);
      return normalized;
    } catch {
      datasetProjectPrefs.set(key, null);
      return null;
    }
  }

  function saveDatasetProjectPrefs(raw) {
    const normalized = normalizeDatasetViewerPrefs(raw);
    if (!normalized) return;
    const key = normalizeProjectText(normalized.project);
    datasetProjectPrefs.set(key, normalized);
    scheduleProjectUserPreferencesSave(normalized.project, {
      lastReservingClassPath: normalized.path,
      datasetViewer: {
        datasetName: normalized.tri,
        updated_at: new Date().toISOString(),
      },
    });
  }

  function getDefaultDisplayLabelForInput(input) {
    if (input?.id === "projectSelect") return DEFAULT_PROJECT_DISPLAY;
    if (input?.id === "pathInput") return DEFAULT_PATH_DISPLAY;
    return "Default";
  }

  function buildDefaultDisplayValue(input, raw) {
    const resolved = String(raw || "").trim();
    const label = getDefaultDisplayLabelForInput(input);
    return resolved ? `${label} (${resolved})` : label;
  }

  function getDefaultValueForInput(input) {
    const defaults = loadWorkflowDefaults();
    if (!defaults || !input) return "";
    if (input.id === "projectSelect") return defaults.project || "";
    if (input.id === "pathInput") return defaults.reservingClass || "";
    return "";
  }

  function isDefaultTokenValue(value) {
    const v = String(value || "").trim();
    if (!v) return false;
    const lower = v.toLowerCase();
    if (lower === DEFAULT_TOKEN.toLowerCase() || lower === "default") return true;
    const defaultLabels = [DEFAULT_PROJECT_DISPLAY, DEFAULT_PATH_DISPLAY];
    return defaultLabels.some((label) => {
      const labelLower = label.toLowerCase();
      return lower === labelLower || (lower.startsWith(`${labelLower} (`) && lower.endsWith(")"));
    });
  }

  function isInputDefaultBound(input) {
    if (!input) return false;
    if (input.dataset?.globalDefault === "1") return true;
    return isDefaultTokenValue(input.value);
  }

  function setInputDefaultBound(input, bound) {
    if (!input) return;
    if (bound) {
      input.dataset.globalDefault = "1";
      input.value = buildDefaultDisplayValue(input, getDefaultValueForInput(input));
    } else {
      delete input.dataset.globalDefault;
    }
  }

  function getWorkflowVarValue(vars, key, fallbackName) {
    if (!Array.isArray(vars)) return "";
    const byKey = vars.find((v) => v && typeof v === "object" && String(v.key || "") === key);
    if (byKey && typeof byKey.value === "string") return byKey.value.trim();
    const target = String(fallbackName || "").trim().toLowerCase();
    if (!target) return "";
    const byName = vars.find((v) => {
      if (!v || typeof v !== "object") return false;
      const name = String(v.name || "").trim().toLowerCase();
      return name === target;
    });
    if (byName && typeof byName.value === "string") return byName.value.trim();
    return "";
  }

  function getDatasetInstanceNameValue() {
    const detailName = String(document.getElementById("dsDetailName")?.value || "").trim();
    return detailName || String(document.getElementById("triInput")?.value || "").trim();
  }

  function normalizeDatasetInstanceKey(value) {
    return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
  }

  function getCachedInstanceNamesFromItem(item = {}) {
    const names = [];
    const add = (value) => {
      const text = String(value || "").trim();
      if (text) names.push(text);
    };
    add(item.name);
    return names;
  }

  function cachedInstanceMatchesName(item, instanceName) {
    const instanceKey = normalizeDatasetInstanceKey(instanceName);
    if (!instanceKey) return false;
    const itemNames = getCachedInstanceNamesFromItem(item).map(normalizeDatasetInstanceKey);
    return itemNames.includes(instanceKey);
  }

  async function loadCachedDatasetInstancesForCurrentContext() {
    const project = getResolvedProjectValue();
    const path = getResolvedReservingClassValue();
    if (!project || !path) {
      cachedDatasetInstanceRows = [];
      cachedDatasetInstanceKey = "";
      return [];
    }
    const key = `${normalizeProjectText(project)}\u001f${normalizeReservingClassPath(path).toLowerCase()}`;
    if (cachedDatasetInstanceKey === key) return cachedDatasetInstanceRows;
    if (cachedDatasetInstanceLoadPromise) return cachedDatasetInstanceLoadPromise;
    cachedDatasetInstanceLoadPromise = (async () => {
      try {
        const url = new URL("/datasets/cached", window.location.origin);
        url.searchParams.set("project_name", project);
        url.searchParams.set("reserving_class", path);
        const resp = await fetch(url.toString(), { cache: "no-store" });
        const payload = await resp.json().catch(() => ({}));
        if (!resp.ok || payload?.ok === false) throw new Error(payload?.detail || "Cached dataset lookup failed.");
        cachedDatasetInstanceRows = Array.isArray(payload?.files) ? payload.files : [];
        cachedDatasetInstanceKey = key;
        return cachedDatasetInstanceRows;
      } catch {
        cachedDatasetInstanceRows = [];
        cachedDatasetInstanceKey = key;
        return cachedDatasetInstanceRows;
      } finally {
        cachedDatasetInstanceLoadPromise = null;
      }
    })();
    return cachedDatasetInstanceLoadPromise;
  }

  function setDatasetInstanceNameConflict(hasConflict, message = "") {
    runtime.datasetInstanceNameConflict = !!hasConflict;
    runtime.datasetInstanceNameConflictMessage = runtime.datasetInstanceNameConflict ? String(message || "Name already exists.") : "";
    const warning = document.getElementById("dsDetailNameWarning");
    if (warning) {
      warning.textContent = runtime.datasetInstanceNameConflictMessage;
      warning.hidden = !runtime.datasetInstanceNameConflict;
    }
    const input = document.getElementById("dsDetailName");
    if (input) {
      input.setCustomValidity(runtime.datasetInstanceNameConflict ? runtime.datasetInstanceNameConflictMessage : "");
      input.classList.toggle("invalid", runtime.datasetInstanceNameConflict);
    }
    updateDatasetSaveUi();
  }

  function invalidateCachedDatasetInstances() {
    cachedDatasetInstanceRows = [];
    cachedDatasetInstanceKey = "";
    cachedDatasetInstanceLoadPromise = null;
  }

  async function refreshDatasetInstanceNameConflict() {
    if (!isProjectInstanceDraft) {
      setDatasetInstanceNameConflict(false);
      return false;
    }
    const instanceName = getDatasetInstanceNameValue();
    if (!instanceName) {
      setDatasetInstanceNameConflict(false);
      return false;
    }
    if (runtime.savedProjectInstanceDraftName && normalizeDatasetInstanceKey(instanceName) === normalizeDatasetInstanceKey(runtime.savedProjectInstanceDraftName)) {
      setDatasetInstanceNameConflict(false);
      return false;
    }
    const rows = await loadCachedDatasetInstancesForCurrentContext();
    const conflict = rows.some((item) => cachedInstanceMatchesName(item, instanceName));
    setDatasetInstanceNameConflict(
      conflict,
      conflict ? "Name already exists in this reserving class path." : "",
    );
    return conflict;
  }

  function recordDatasetBrowsingHistory(entry) {
    if (isDfmDataTabHost()) return;
    const normalized = normalizeBrowsingHistoryEntry(entry);
    if (!normalized) return;
    const out = pushBrowsingHistoryEntry(normalized, { maxEntries: BROWSING_HISTORY_MAX_ENTRIES });
    try {
      window.parent.postMessage(
        {
          type: "arcrho:browsing-history-updated",
          entry: out?.entry || normalized,
        },
        "*",
      );
    } catch {
      // ignore
    }
  }

  function saveLastDsId(dsId) {
    if (!dsId) return;
    try {
      localStorage.setItem(scopedKey(LS_DS_KEY), String(dsId));
    } catch {
      // ignore
    }
  }

  function loadLastDsId() {
    try {
      return localStorage.getItem(scopedKey(LS_DS_KEY)) || "";
    } catch {
      return "";
    }
  }

  // Persist ArcRhoTri input controls so refresh doesn't reset them.
  function saveTriInputsToStorage() {
    try {
      const projectInput = document.getElementById("projectSelect");
      const pathInput = document.getElementById("pathInput");
      const triInput = document.getElementById("triInput");
      const payload = {
        project: getStoredInputValue(projectInput),
        path: getStoredInputValue(pathInput),
        tri: triInput?.value || "",
        instanceName: getDatasetInstanceNameValue(),
        originLen: document.getElementById("originLenSelect")?.value || "",
        devLen: document.getElementById("devLenSelect")?.value || "",
        cumulative: !!document.getElementById("cumulativeChk")?.checked,
        transposed: !!document.getElementById("transposedChk")?.checked,
        calendar: document.querySelector('input[name="timeMode"][value="calendar"]')?.checked === true,
        decimalPlaces: getDatasetDecimalPlacesValue(),
        numberFormat: getDatasetSyncedNumberFormatValue(),
      };
      const resolvedInputs = normalizeBrowsingHistoryEntry({
        project: getResolvedProjectValue(),
        path: getResolvedReservingClassValue(),
        tri: String(triInput?.value || "").trim(),
        instanceName: getDatasetInstanceNameValue(),
      });
      localStorage.setItem(scopedKey(LS_FORM_KEY), JSON.stringify(payload));
      if (!isDfmDataTabHost()) {
        saveDatasetProjectPrefs(resolvedInputs);
        const matchedProject = findExactProjectMatch(getResolvedProjectValue());
        if (matchedProject) saveLastDatasetViewerProjectToAppData(matchedProject);
      }
      if (!isDfmDataTabHost() && resolvedInputs) {
        setLastViewedDatasetInputs(resolvedInputs);
      }
      try {
        window.parent.postMessage({
          type: "arcrho:dataset-settings-changed",
          stepId: instanceId,
          settings: payload,
          resolved: resolvedInputs || null,
        }, "*");
      } catch {
        // ignore
      }
      refreshDatasetSettingsDirty();
    } catch {
      // ignore
    }
  }

  const GENERATED_DATASET_READ_ONLY_MESSAGE = "Generated datasets are read-only.";

  function isDatasetReadOnly() {
    return isTemporaryDatasetView
      || isReadOnlyDatasetViewer
      || runtime.isSidecarReadOnlyDataset
      || runtime.datasetSaveInFlight
      // Only the origin axis locks the grid. A coarse origin row has no single
      // stored cell behind it, while a coarse development column does, so
      // typing, paste and links stay live there and the save scatters them.
      || datasetOriginDisplayIsCoarserThanStored();
  }

  // Every refusal the grid, the Links tab and the patch save report comes from
  // here, so the reason a reader sees always matches the rule that stopped them.
  function getDatasetReadOnlyMessage() {
    const coarseOnly = !isTemporaryDatasetView
      && !isReadOnlyDatasetViewer
      && !runtime.isSidecarReadOnlyDataset
      && datasetOriginDisplayIsCoarserThanStored();
    return coarseOnly ? datasetCoarserViewMessage() : GENERATED_DATASET_READ_ONLY_MESSAGE;
  }

  async function restoreTriInputsFromStorage() {
    let s = null;
    try {
      const raw = localStorage.getItem(scopedKey(LS_FORM_KEY)) || "";
      if (raw) s = JSON.parse(raw);
    } catch {
      s = null;
    }
    if (!isDfmDataTabHost() && !workflowId) {
      const localProject = await loadLastDatasetViewerProjectFromAppData();
      const matchedProject = findExactProjectMatch(localProject);
      if (matchedProject) {
        const base = s && typeof s === "object" ? s : {};
        const sameBaseProject = normalizeProjectText(base.project) === normalizeProjectText(matchedProject);
        const prefs = await loadDatasetProjectPrefs(matchedProject);
        s = {
          ...base,
          project: matchedProject,
          path: prefs?.path || (sameBaseProject ? (base.path || "") : ""),
          tri: prefs?.tri || (sameBaseProject ? (base.tri || "") : ""),
        };
      }
    }
    if (s && typeof s === "object") {
      const project = isDefaultTokenValue(s.project)
        ? String(loadWorkflowDefaults()?.project || "").trim()
        : String(s.project || "").trim();
      const prefs = await loadDatasetProjectPrefs(project);
      if (prefs) {
        s = {
          ...s,
          path: prefs.path || s.path || "",
          tri: prefs.tri || s.tri || "",
        };
      }
    }
    if ((!s || typeof s !== "object") && !isDfmDataTabHost()) {
      s = getLastViewedDatasetInputs();
      const prefs = await loadDatasetProjectPrefs(s?.project || "");
      if (prefs) s = prefs;
    }
    if (!s || typeof s !== "object") return;

    const projectInput = document.getElementById("projectSelect");
    const pathInput = document.getElementById("pathInput");
    const triInput = document.getElementById("triInput");
    const detailNameInput = document.getElementById("dsDetailName");
    const originSel = document.getElementById("originLenSelect");
    const devSel = document.getElementById("devLenSelect");

    // Only restore if the saved value is valid in the current UI.
    if (projectInput && typeof s.project === "string") {
      if (isDefaultTokenValue(s.project)) {
        setInputDefaultBound(projectInput, true);
      } else if (s.project.trim()) {
        setInputDefaultBound(projectInput, false);
        const match = findExactProjectMatch(s.project);
        projectInput.value = match || s.project;
      }
    }
    if (pathInput && typeof s.path === "string") {
      if (isDefaultTokenValue(s.path)) {
        setInputDefaultBound(pathInput, true);
      } else if (s.path.trim()) {
        setInputDefaultBound(pathInput, false);
        pathInput.value = normalizeReservingClassPath(s.path);
      }
    }
    if (triInput && typeof s.tri === "string" && s.tri.trim()) triInput.value = s.tri;
    if (detailNameInput && typeof s.instanceName === "string" && s.instanceName.trim()) {
      detailNameInput.value = s.instanceName.trim();
    }

    if (originSel && s.originLen && [...originSel.options].some(o => o.value === String(s.originLen))) {
      originSel.value = String(s.originLen);
    }
    if (devSel && s.devLen && [...devSel.options].some(o => o.value === String(s.devLen))) {
      devSel.value = String(s.devLen);
    }
    refreshLenDropdowns();

    const cumChk = document.getElementById("cumulativeChk");
    if (cumChk && typeof s.cumulative === "boolean") cumChk.checked = s.cumulative;

    const transposedChk = document.getElementById("transposedChk");
    if (transposedChk && typeof s.transposed === "boolean") transposedChk.checked = s.transposed;

    if (typeof s.calendar === "boolean") {
      const mode = s.calendar ? "calendar" : "development";
      const modeInput = document.querySelector(`input[name="timeMode"][value="${mode}"]`);
      if (modeInput) modeInput.checked = true;
    }
    if (s.decimalPlaces !== undefined || s.decimal_places !== undefined) {
      setDatasetDecimalPlacesValue(s.decimalPlaces ?? s.decimal_places);
    }
    if (typeof s.numberFormat === "string") {
      setDatasetNumberFormatValue(s.numberFormat);
    }
  }

  function applyTriInputsFromQueryParams() {
    const queryInputs = readDatasetInputsFromQueryParams();
    if (!queryInputs) return false;

    const projectInput = document.getElementById("projectSelect");
    const pathInput = document.getElementById("pathInput");
    const triInput = document.getElementById("triInput");
    const detailNameInput = document.getElementById("dsDetailName");
    const originSel = document.getElementById("originLenSelect");
    const devSel = document.getElementById("devLenSelect");
    if (projectInput && queryInputs.project) {
      setInputDefaultBound(projectInput, false);
      projectInput.value = queryInputs.project;
    }
    if (pathInput && queryInputs.path) {
      setInputDefaultBound(pathInput, false);
      pathInput.value = queryInputs.path;
    }
    if (triInput && queryInputs.tri) {
      triInput.value = queryInputs.tri;
    }
    if (detailNameInput && queryInputs.instanceName) {
      detailNameInput.value = queryInputs.instanceName;
    } else if (detailNameInput && queryInputs.tri && !String(detailNameInput.value || "").trim()) {
      detailNameInput.value = queryInputs.tri;
    }
    if (originSel && queryInputs.originLen && [...originSel.options].some(o => o.value === String(queryInputs.originLen))) {
      originSel.value = String(queryInputs.originLen);
    }
    if (devSel && queryInputs.devLen && [...devSel.options].some(o => o.value === String(queryInputs.devLen))) {
      devSel.value = String(queryInputs.devLen);
    }
    if (queryInputs.decimalPlaces !== undefined || queryInputs.decimal_places !== undefined) {
      setDatasetDecimalPlacesValue(queryInputs.decimalPlaces ?? queryInputs.decimal_places);
    }
    if (typeof queryInputs.numberFormat === "string") {
      setDatasetNumberFormatValue(queryInputs.numberFormat);
    }
    refreshLenDropdowns();
    if (!isDfmDataTabHost()) {
      setLastViewedDatasetInputs(queryInputs);
    }
    return true;
  }

  function hasScopedTriInputs() {
    try {
      return !!localStorage.getItem(scopedKey(LS_FORM_KEY));
    } catch {
      return false;
    }
  }

  function loadWorkflowDefaults() {
    if (!workflowId) return null;
    try {
      const raw = localStorage.getItem(`${WF_GLOBAL_CTRL_PREFIX}${workflowId}`) || "";
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object") return null;
      const vars = Array.isArray(parsed.vars) ? parsed.vars : null;
      const project = vars
        ? (getWorkflowVarValue(vars, "project", "Default Project") || getWorkflowVarValue(vars, "project", "Project"))
        : (typeof parsed.project === "string" ? parsed.project : "");
      const reservingClass = vars
        ? (getWorkflowVarValue(vars, "reservingClass", "Default Path") || getWorkflowVarValue(vars, "reservingClass", "Reserving Class"))
        : (typeof parsed.reservingClass === "string" ? parsed.reservingClass : "");
      return { project, reservingClass, vars: vars || [] };
    } catch {
      return null;
    }
  }

  function applyWorkflowDefaultsIfNew() {
    if (!workflowId) return;
    if (hasScopedTriInputs()) return;

    const defaults = loadWorkflowDefaults();
    if (!defaults) return;

    const projectInput = document.getElementById("projectSelect");
    const pathInput = document.getElementById("pathInput");

    if (projectInput && defaults.project) {
      setInputDefaultBound(projectInput, true);
    }
    if (pathInput && defaults.reservingClass) {
      setInputDefaultBound(pathInput, true);
    }
    if (defaults.project) {
      void applyResolvedProjectDefaults(defaults.project);
    }
    saveTriInputsToStorage();
  }

  function getResolvedProjectValue() {
    const input = document.getElementById("projectSelect");
    const raw = (input?.value || "").trim();
    if (isInputDefaultBound(input)) {
      const defaults = loadWorkflowDefaults();
      return (defaults?.project || "").trim();
    }
    return raw;
  }

  function getResolvedReservingClassValue() {
    const input = document.getElementById("pathInput");
    const raw = normalizeReservingClassPath(input?.value || "");
    if (isInputDefaultBound(input)) {
      const defaults = loadWorkflowDefaults();
      return normalizeReservingClassPath(defaults?.reservingClass || "");
    }
    return raw;
  }

  function getStoredInputValue(input) {
    if (!input) return "";
    if (isInputDefaultBound(input)) return DEFAULT_TOKEN;
    return input.value || "";
  }

  async function applyResolvedProjectDefaults(project) {
    if (!project) return;
    if (project === runtime.lastProjectSelection) return;
    runtime.lastProjectSelection = project;
    await ensureHeadersForProject(project);
    await ensureDevHeadersForProject(project);
    await refreshDatasetTypesForProject(project);
    await refreshReservingClassPathsForProject(project);
  }

  function extractDefaultsFromControl(control) {
    if (!control || typeof control !== "object") return null;
    const vars = Array.isArray(control.vars) ? control.vars : null;
    const project = vars
      ? (getWorkflowVarValue(vars, "project", "Default Project") || getWorkflowVarValue(vars, "project", "Project"))
      : (typeof control.project === "string" ? control.project : "");
    const reservingClass = vars
      ? (getWorkflowVarValue(vars, "reservingClass", "Default Path") || getWorkflowVarValue(vars, "reservingClass", "Reserving Class"))
      : (typeof control.reservingClass === "string" ? control.reservingClass : "");
    return { project, reservingClass };
  }

  function handleWorkflowGlobalChange(control = null) {
    if (!workflowId) return;
    const projectInput = document.getElementById("projectSelect");
    const pathInput = document.getElementById("pathInput");
    const projectDefault = isInputDefaultBound(projectInput);
    const pathDefault = isInputDefaultBound(pathInput);
    if (!projectDefault && !pathDefault) return;

    const defaults = control ? extractDefaultsFromControl(control) : loadWorkflowDefaults();
    if (!defaults) return;

    if (projectDefault && projectInput) {
      setInputDefaultBound(projectInput, true);
    }
    if (pathDefault && pathInput) {
      setInputDefaultBound(pathInput, true);
    }

    if (projectDefault && defaults.project) {
      void applyResolvedProjectDefaults(defaults.project);
    }

    if (projectDefault || pathDefault) {
      const currentProjectValue = projectDefault ? DEFAULT_TOKEN : (projectInput?.value || "");
      renderProjectOptions(runtime.allProjects, currentProjectValue);
      saveTriInputsToStorage();
      scheduleAutoRun(0);
      try {
        window.dispatchEvent(new CustomEvent("arcrho:workflow-defaults-updated", { detail: defaults }));
      } catch {
        // ignore
      }
    }
  }
  function wireDatasetInstanceNameInput() {
    const input = document.getElementById("dsDetailName");
    if (!input || input.dataset.instanceNameWired === "1") return;
    input.dataset.instanceNameWired = "1";
    input.addEventListener("input", () => {
      saveTriInputsToStorage();
      refreshDatasetSettingsDirty();
      void refreshDatasetInstanceNameConflict();
    });
    input.addEventListener("change", () => {
      void refreshDatasetInstanceNameConflict();
    });
  }


  Object.assign(runtime, {
    normalizeDatasetViewerPrefs,
    normalizeLocalDatasetViewerPrefs,
    loadLastDatasetViewerProjectFromAppData,
    saveLastDatasetViewerProjectToAppData,
    loadDatasetProjectPrefs,
    saveDatasetProjectPrefs,
    getDefaultDisplayLabelForInput,
    buildDefaultDisplayValue,
    getDefaultValueForInput,
    isDefaultTokenValue,
    isInputDefaultBound,
    setInputDefaultBound,
    getWorkflowVarValue,
    getDatasetInstanceNameValue,
    normalizeDatasetInstanceKey,
    getCachedInstanceNamesFromItem,
    cachedInstanceMatchesName,
    loadCachedDatasetInstancesForCurrentContext,
    setDatasetInstanceNameConflict,
    invalidateCachedDatasetInstances,
    refreshDatasetInstanceNameConflict,
    recordDatasetBrowsingHistory,
    saveLastDsId,
    loadLastDsId,
    saveTriInputsToStorage,
    isDatasetReadOnly,
    getDatasetReadOnlyMessage,
    restoreTriInputsFromStorage,
    applyTriInputsFromQueryParams,
    hasScopedTriInputs,
    loadWorkflowDefaults,
    applyWorkflowDefaultsIfNew,
    getResolvedProjectValue,
    getResolvedReservingClassValue,
    getStoredInputValue,
    applyResolvedProjectDefaults,
    extractDefaultsFromControl,
    handleWorkflowGlobalChange,
    wireDatasetInstanceNameInput,
  });
}
