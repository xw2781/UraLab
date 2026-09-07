// Owns project, reserving-class, and dataset valid-value selection.

export function registerDataTabInputsController(runtime) {
  const { state, workflowId, DEFAULT_TOKEN } = runtime;
  const defer = (name) => (...args) => runtime[name](...args);
  const { normalizeDatasetNumberFormat, clampDatasetDecimalPlaces, applyDecimalPlacesToDatasetNumberFormat, loadWorkflowDefaults, buildDefaultDisplayValue, isDefaultTokenValue, setInputDefaultBound, isInputDefaultBound, renderDetailFormula, refreshDatasetInstanceNameConflict, setStatus, buildReservingClassPathPartLookup, normalizeReservingClassPath, normalizeReservingClassPathByPartLookup, normalizeReservingClassPathKey, loadDatasetValidValueList, loadReservingClassValidValueList, getResolvedProjectValue, getResolvedReservingClassValue, getTriInputs, isDfmDataTabHost, syncSidecarForCurrentDataset, enforceDevLenRule, ensureHeadersForProject, ensureDevHeadersForProject, loadDatasetProjectPrefs, saveLastDatasetViewerProjectToAppData, saveTriInputsToStorage, scheduleAutoRun, isRunInFlight, showDatasetLoadingPopup, hideDatasetLoadingPopup, validateReservingClassPathByTypeNames, loadProjectsDropdown } = new Proxy({}, { get: (_target, name) => defer(name) });
  const applyResolvedProjectDefaults = defer("applyResolvedProjectDefaults");
  const LEN_DROPDOWN_CONFIG = {
    originLenSelect: {
      wrapId: "originLenWrap",
      buttonId: "originLenDisplay",
      dropdownId: "originLenDropdown",
    },
    devLenSelect: {
      wrapId: "devLenWrap",
      buttonId: "devLenDisplay",
      dropdownId: "devLenDropdown",
    },
    // The `Stored at` value beside each length is the same control, so it opens
    // the same list and takes the same lock and tooltip treatment.
    originStoredLenSelect: {
      wrapId: "originStoredLenWrap",
      buttonId: "originStoredLenDisplay",
      dropdownId: "originStoredLenDropdown",
    },
    devStoredLenSelect: {
      wrapId: "devStoredLenWrap",
      buttonId: "devStoredLenDisplay",
      dropdownId: "devStoredLenDropdown",
    },
  };

  let activeProjectIndex = -1;
  let activeDatasetIndex = -1;
  let lastDatasetSelection = "";
  let allReservingClassPaths = [];
  let reservingClassPathByKey = new Map();
  let reservingClassPathPartByKey = new Map();
  let lastReservingClassSelection = "";
  let inputLifecycleWired = false;
  function setLastProjectSelection(value) {
    runtime.lastProjectSelection = String(value || "");
  }

  function notifyProjectSelectionCommitted(projectName, source = "") {
    const projectInput = document.getElementById("projectSelect");
    const project = String(projectName || "").trim();
    if (!projectInput || !project) return;
    projectInput.dispatchEvent(new CustomEvent("arcrho:project-selected", {
      bubbles: true,
      detail: { projectName: project, source },
    }));
  }

  function setLastDatasetSelection(value) {
    lastDatasetSelection = String(value || "");
  }

  function setLastReservingClassSelection(value) {
    lastReservingClassSelection = String(value || "");
  }

  function normalizeProjectText(s) {
    return String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
  }

  function getDatasetNumberFormatValue() {
    return normalizeDatasetNumberFormat(document.getElementById("numberFormatSelect")?.value);
  }

  function setDatasetNumberFormatValue(value) {
    const input = document.getElementById("numberFormatSelect");
    if (!input) return;
    input.value = normalizeDatasetNumberFormat(value);
  }

  function getDatasetDecimalPlacesValue() {
    return clampDatasetDecimalPlaces(document.getElementById("decimalPlaces")?.value);
  }

  function getDatasetSyncedNumberFormatValue() {
    return applyDecimalPlacesToDatasetNumberFormat(getDatasetNumberFormatValue(), getDatasetDecimalPlacesValue());
  }

  function setDatasetDecimalPlacesValue(value) {
    const input = document.getElementById("decimalPlaces");
    if (!input) return;
    input.value = String(clampDatasetDecimalPlaces(value));
  }

  function normalizeSearchTokens(q) {
    return normalizeProjectText(q).split(" ").filter(Boolean);
  }

  function matchesProject(name, tokens) {
    if (!tokens.length) return true;
    const hay = normalizeProjectText(name);
    return tokens.every(t => hay.includes(t));
  }

  function getActiveProjectValue() {
    const list = document.getElementById("projectDropdown");
    if (!list) return "";
    const opt = list.children[activeProjectIndex];
    return opt?.dataset?.value || "";
  }

  function renderProjectOptions(projects, activeValue = "") {
    const list = document.getElementById("projectDropdown");
    if (!list) return;
    list.innerHTML = "";
    const defaults = loadWorkflowDefaults();
    const defaultProject = (defaults?.project || "").trim();
    const options = [];
    if (workflowId && defaultProject) {
      options.push({
        label: buildDefaultDisplayValue(document.getElementById("projectSelect"), defaultProject),
        value: DEFAULT_TOKEN,
      });
    }
    for (const p of projects) {
      options.push({ label: p, value: p });
    }

    options.forEach((optData, i) => {
      const opt = document.createElement("div");
      opt.className = "projectOption";
      opt.textContent = optData.label;
      opt.dataset.value = optData.value;
      opt.dataset.index = String(i);
      opt.addEventListener("mouseenter", () => {
        setActiveProjectIndex(i);
      });
      opt.addEventListener("mousedown", (e) => {
        e.preventDefault();
        const projectInput = document.getElementById("projectSelect");
        if (projectInput) {
          if (isDefaultTokenValue(optData.value)) {
            setInputDefaultBound(projectInput, true);
          } else {
            setInputDefaultBound(projectInput, false);
            projectInput.value = optData.value;
          }
        }
        showProjectDropdown(false);
        void handleProjectSelection(optData.value);
      });
      list.appendChild(opt);
    });

    activeProjectIndex = -1;
    if (options.length) {
      let idx = 0;
      if (activeValue) {
        const found = options.findIndex((o) => o.value === activeValue);
        if (found >= 0) idx = found;
      }
      setActiveProjectIndex(idx);
    }
  }

  function showProjectDropdown(open) {
    const list = document.getElementById("projectDropdown");
    if (!list) return;
    const hasItems = !!list.children.length;
    if (open && hasItems) list.classList.add("open");
    else list.classList.remove("open");
  }

  function filterProjectOptions(query) {
    const tokens = normalizeSearchTokens(query);
    const filtered = tokens.length
      ? runtime.allProjects.filter(p => matchesProject(p, tokens))
      : runtime.allProjects.slice();
    const activeValue = getActiveProjectValue();
    renderProjectOptions(filtered, activeValue);
    showProjectDropdown(true);
  }

  function getProjectFilterQuery(input) {
    if (isInputDefaultBound(input)) return "";
    return input?.value || "";
  }

  function getProjectOptionsList() {
    const list = document.getElementById("projectDropdown");
    if (!list) return [];
    return Array.from(list.children);
  }

  function setActiveProjectIndex(idx) {
    const opts = getProjectOptionsList();
    if (!opts.length) {
      activeProjectIndex = -1;
      return;
    }
    let next = idx;
    if (next < 0) next = opts.length - 1;
    if (next >= opts.length) next = 0;
    activeProjectIndex = next;
    opts.forEach((el, i) => el.classList.toggle("active", i === activeProjectIndex));
    opts[activeProjectIndex].scrollIntoView({ block: "nearest" });
  }

  function getActiveProjectIndex() {
    return activeProjectIndex;
  }

  function chooseActiveProject() {
    const opts = getProjectOptionsList();
    if (activeProjectIndex < 0 || activeProjectIndex >= opts.length) return false;
    const value = opts[activeProjectIndex].dataset.value || opts[activeProjectIndex].textContent;
    if (!value) return false;
    const projectInput = document.getElementById("projectSelect");
    if (projectInput) {
      if (isDefaultTokenValue(value)) {
        setInputDefaultBound(projectInput, true);
      } else {
        setInputDefaultBound(projectInput, false);
        projectInput.value = value;
      }
    }
    showProjectDropdown(false);
    void handleProjectSelection(value);
    return true;
  }

  function findExactProjectMatch(value) {
    const v = normalizeProjectText(value);
    if (!v) return "";
    return runtime.allProjects.find(p => normalizeProjectText(p) === v) || "";
  }

  function getActiveDatasetValue() {
    const list = document.getElementById("datasetDropdown");
    if (!list) return "";
    const opt = list.children[activeDatasetIndex];
    return opt?.dataset?.value || "";
  }

  function renderDatasetOptions(items, activeValue = "") {
    const list = document.getElementById("datasetDropdown");
    if (!list) return;
    list.innerHTML = "";
    items.forEach((name, i) => {
      const opt = document.createElement("div");
      opt.className = "datasetOption";
      opt.textContent = name;
      opt.dataset.value = name;
      opt.dataset.index = String(i);
      opt.addEventListener("mouseenter", () => {
        setActiveDatasetIndex(i);
      });
      opt.addEventListener("mousedown", (e) => {
        e.preventDefault();
        const triInput = document.getElementById("triInput");
        if (triInput) triInput.value = name;
        showDatasetDropdown(false);
        void handleDatasetSelection(name);
      });
      list.appendChild(opt);
    });

    activeDatasetIndex = -1;
    if (items.length) {
      const idx = activeValue ? Math.max(0, items.indexOf(activeValue)) : 0;
      setActiveDatasetIndex(idx);
    }
  }

  function showDatasetDropdown(open) {
    const list = document.getElementById("datasetDropdown");
    if (!list) return;
    const hasItems = !!list.children.length;
    if (open && hasItems) list.classList.add("open");
    else list.classList.remove("open");
  }

  function filterDatasetOptions(query) {
    if (!runtime.allDatasetTypes.length) {
      showDatasetDropdown(false);
      return;
    }
    const tokens = normalizeSearchTokens(query);
    const filtered = tokens.length
      ? runtime.allDatasetTypes.filter(name => matchesProject(name, tokens))
      : runtime.allDatasetTypes;
    const activeValue = getActiveDatasetValue();
    renderDatasetOptions(filtered, activeValue);
    showDatasetDropdown(true);
  }

  function getDatasetOptionsList() {
    const list = document.getElementById("datasetDropdown");
    if (!list) return [];
    return Array.from(list.children);
  }

  function setActiveDatasetIndex(idx) {
    const opts = getDatasetOptionsList();
    if (!opts.length) {
      activeDatasetIndex = -1;
      return;
    }
    let next = idx;
    if (next < 0) next = opts.length - 1;
    if (next >= opts.length) next = 0;
    activeDatasetIndex = next;
    opts.forEach((el, i) => el.classList.toggle("active", i === activeDatasetIndex));
    opts[activeDatasetIndex].scrollIntoView({ block: "nearest" });
  }

  function getActiveDatasetIndex() {
    return activeDatasetIndex;
  }

  function chooseActiveDataset() {
    const opts = getDatasetOptionsList();
    if (activeDatasetIndex < 0 || activeDatasetIndex >= opts.length) return false;
    const value = opts[activeDatasetIndex].dataset.value || opts[activeDatasetIndex].textContent;
    if (!value) return false;
    const triInput = document.getElementById("triInput");
    if (triInput) triInput.value = value;
    showDatasetDropdown(false);
    void handleDatasetSelection(value);
    return true;
  }

  function findExactDatasetMatch(value) {
    const v = normalizeProjectText(value);
    if (!v) return "";
    return runtime.allDatasetTypes.find(name => normalizeProjectText(name) === v) || "";
  }

  function ensureDatasetTypeOption(value) {
    const name = String(value || "").trim();
    if (!name) return "";
    const key = normalizeProjectText(name);
    const existing = runtime.allDatasetTypes.find((item) => normalizeProjectText(item) === key);
    if (existing) return existing;

    runtime.allDatasetTypes = [...runtime.allDatasetTypes, name].sort((a, b) =>
      String(a || "").localeCompare(String(b || ""), undefined, { sensitivity: "base", numeric: true }),
    );
    renderDatasetOptions(runtime.allDatasetTypes, name);
    return name;
  }

  function getDatasetTypeFormulaByName(datasetTypeName) {
    const key = normalizeProjectText(datasetTypeName);
    if (!key) return "";
    const formulaMap = state.datasetTypeFormulaByKey instanceof Map ? state.datasetTypeFormulaByKey : null;
    if (!formulaMap) return "";
    return String(formulaMap.get(key) || "").trim();
  }

  function getDatasetTypeDataFormatByName(datasetTypeName) {
    const key = normalizeProjectText(datasetTypeName);
    if (!key) return "";
    const dataFormatMap = state.datasetTypeDataFormatByKey instanceof Map ? state.datasetTypeDataFormatByKey : null;
    if (!dataFormatMap) return "";
    return String(dataFormatMap.get(key) || "").trim();
  }

  function resizeDetailFormulaInput() {
    const formulaBox = document.getElementById("dsDetailFormulaBox");
    if (!formulaBox) return;
    formulaBox.style.maxHeight = "140px";
  }

  function wireDataTabInputLifecycle() {
    if (inputLifecycleWired) return;
    inputLifecycleWired = true;
    window.addEventListener("resize", resizeDetailFormulaInput);
  }

  function syncDetailFormulaFromDatasetType(datasetTypeName) {
    const formula = getDatasetTypeFormulaByName(datasetTypeName);
    renderDetailFormula(formula, runtime.currentDatasetPrecedents);
    resizeDetailFormulaInput();
  }

  function syncDetailDatasetTypeFromTopInput(rawValue, options = {}) {
    const syncName = !!options?.syncName;
    const dsDetailName = document.getElementById("dsDetailName");
    const prevType = String(dsDetailName?.dataset?.datasetType || "").trim();
    const raw = String(rawValue || "").trim();
    const canonical = raw ? (ensureDatasetTypeOption(raw) || raw) : "";
    const nextType = String(canonical || "").trim();

    if (dsDetailName) {
      if (syncName) {
        const currentName = String(dsDetailName.value || "").trim();
        if (!currentName || normalizeProjectText(prevType) !== normalizeProjectText(nextType)) {
          dsDetailName.value = nextType;
        }
      }
      dsDetailName.dataset.datasetType = nextType;
    }

    syncDetailFormulaFromDatasetType(nextType);
    void refreshDatasetInstanceNameConflict();
  }

  function loadDatasetTypeDependencyModel(projectName, options = {}) {
    return runtime.datasetDependencyGuard.loadDatasetTypeDependencyModel(projectName, options);
  }

  function validateDatasetTypeDependencies(datasetType, options = {}) {
    return runtime.datasetDependencyGuard.validateDatasetTypeDependencies(datasetType, options);
  }

  function setInputInvalid(input, message) {
    if (!input) return;
    input.setCustomValidity(String(message || "Invalid value."));
  }

  function clearInputInvalid(input) {
    if (!input) return;
    input.setCustomValidity("");
  }

  function reportInputInvalid(input, message, statusText = "") {
    if (!input) return;
    setInputInvalid(input, message);
    try { input.reportValidity(); } catch {}
    if (statusText) setStatus(statusText);
  }

  function rebuildReservingClassPathLookup(paths) {
    reservingClassPathByKey = new Map();
    reservingClassPathPartByKey = buildReservingClassPathPartLookup(paths);
    for (const raw of Array.isArray(paths) ? paths : []) {
      const normalized = normalizeReservingClassPath(raw);
      if (!normalized) continue;
      const key = normalizeReservingClassPathKey(normalized);
      if (!key || reservingClassPathByKey.has(key)) continue;
      reservingClassPathByKey.set(key, normalized);
    }
  }

  function findExactReservingClassMatch(value) {
    const normalized = normalizeReservingClassPath(value);
    const key = normalizeReservingClassPathKey(normalized);
    if (!key) return "";
    const exact = reservingClassPathByKey.get(key);
    if (exact) return exact;
    return normalizeReservingClassPathByPartLookup(normalized, reservingClassPathPartByKey);
  }

  function ensureReservingClassOption(value) {
    const normalized = normalizeReservingClassPath(value);
    if (!normalized) return "";
    const existing = findExactReservingClassMatch(normalized);
    if (existing) return existing;
    allReservingClassPaths = [...allReservingClassPaths, normalized].sort((a, b) =>
      String(a || "").localeCompare(String(b || ""), undefined, { sensitivity: "base", numeric: true }),
    );
    rebuildReservingClassPathLookup(allReservingClassPaths);
    return normalized;
  }

  async function refreshDatasetTypesForProject(project, useCache = true) {
    runtime.datasetDependencyGuard.clearProjectCache(project);

    if (!project) {
      runtime.allDatasetTypes = [];
      state.datasetTypeSourceByKey = new Map();
      state.datasetTypeFormulaByKey = new Map();
      state.datasetTypeDataFormatByKey = new Map();
      renderDatasetOptions([]);
      syncDetailDatasetTypeFromTopInput(document.getElementById("triInput")?.value || "", { syncName: false });
      showDatasetDropdown(false);
      return;
    }

    let items = [];
    try {
      items = await loadDatasetValidValueList(project, { forceReload: !useCache });
    } catch (err) {
      console.error(`Failed to load dataset types for project "${project}":`, err);
      items = [];
    }
    runtime.allDatasetTypes = Array.isArray(items) ? items : [];
    try {
      await loadDatasetTypeDependencyModel(project, { forceReload: !useCache });
    } catch {
      state.datasetTypeSourceByKey = new Map();
      state.datasetTypeFormulaByKey = new Map();
      state.datasetTypeDataFormatByKey = new Map();
    }
    renderDatasetOptions(runtime.allDatasetTypes);
    syncDetailDatasetTypeFromTopInput(document.getElementById("triInput")?.value || "", { syncName: false });
    showDatasetDropdown(false);
  }

  async function refreshReservingClassPathsForProject(project, useCache = true) {
    if (!project) {
      allReservingClassPaths = [];
      rebuildReservingClassPathLookup([]);
      return;
    }

    let items = [];
    try {
      items = await loadReservingClassValidValueList(project, { forceReload: !useCache });
    } catch (err) {
      console.error(`Failed to load reserving class values for project "${project}":`, err);
      items = [];
    }
    allReservingClassPaths = Array.isArray(items) ? items : [];
    rebuildReservingClassPathLookup(allReservingClassPaths);
  }

  async function handleDatasetSelection(value, options = {}) {
    const strict = !!options?.strict;
    const showMessage = !!options?.showMessage;
    const name = findExactDatasetMatch(value);
    const triInput = document.getElementById("triInput");
    if (!name) {
      if (strict && triInput) {
        if (lastDatasetSelection) triInput.value = lastDatasetSelection;
        else triInput.value = "";
        clearInputInvalid(triInput);
        if (showMessage) {
          reportInputInvalid(
            triInput,
            "Dataset Type is not in the valid list for this project.",
            "Invalid Dataset Type. Please select a value from the valid list.",
          );
        }
      }
      return false;
    }
    const switched = name !== lastDatasetSelection;

    if (triInput) triInput.value = name;
    syncDetailDatasetTypeFromTopInput(name, { syncName: switched });
    const dependencyResult = await validateDatasetTypeDependencies(name, {
      showMessage: switched || showMessage || strict,
    });
    if (!dependencyResult.ok) {
      showDatasetDropdown(false);
      return false;
    }
    lastDatasetSelection = name;
    clearInputInvalid(triInput);
    showDatasetDropdown(false);
    if (switched) {
      saveTriInputsToStorage();
      await syncSidecarForCurrentDataset({ applyLengths: true });
      enforceDevLenRule({ source: "origin" });
      scheduleAutoRun();
    }
    return true;
  }

  function validateAndNormalizeProjectInput(options = {}) {
    const strict = !!options?.strict;
    const showMessage = !!options?.showMessage;
    const input = document.getElementById("projectSelect");
    if (!input) return { ok: false, value: "" };

    if (isInputDefaultBound(input)) {
      const resolvedDefault = getResolvedProjectValue();
      const matchedDefault = findExactProjectMatch(resolvedDefault);
      if (!matchedDefault) {
        if (strict && showMessage) {
          reportInputInvalid(
            input,
            "Default Project is not in the valid list.",
            "Invalid Project Name. Please select a valid project.",
          );
        }
        return { ok: false, value: "" };
      }
      clearInputInvalid(input);
      return { ok: true, value: matchedDefault };
    }

    const raw = String(input.value || "").trim();
    const matched = findExactProjectMatch(raw);
    if (!matched) {
      if (strict) {
        if (runtime.lastProjectSelection) input.value = runtime.lastProjectSelection;
        else input.value = "";
        clearInputInvalid(input);
        if (showMessage) {
          reportInputInvalid(
            input,
            "Project Name is not in the valid list.",
            "Invalid Project Name. Please select a valid project.",
          );
        }
      }
      return { ok: false, value: "" };
    }
    input.value = matched;
    clearInputInvalid(input);
    return { ok: true, value: matched };
  }

  async function ensureProjectValidationOptions() {
    const project = String(getResolvedProjectValue() || "").trim();
    if (!project || runtime.allProjects.length) return true;
    const result = await loadProjectsDropdown();
    return result?.ok !== false;
  }

  function validateAndNormalizeDatasetInput(options = {}) {
    const strict = !!options?.strict;
    const showMessage = !!options?.showMessage;
    const input = document.getElementById("triInput");
    if (!input) return { ok: false, value: "" };
    const matched = findExactDatasetMatch(input.value);
    if (!matched) {
      if (strict) {
        if (lastDatasetSelection) input.value = lastDatasetSelection;
        else input.value = "";
        clearInputInvalid(input);
        if (showMessage) {
          reportInputInvalid(
            input,
            "Dataset Type is not in the valid list for this project.",
            "Invalid Dataset Type. Please select a value from the valid list.",
          );
        }
      }
      return { ok: false, value: "" };
    }
    input.value = matched;
    clearInputInvalid(input);
    return { ok: true, value: matched };
  }

  async function validateAndNormalizeReservingClassInput(projectName, options = {}) {
    const strict = !!options?.strict;
    const showMessage = !!options?.showMessage;
    const input = document.getElementById("pathInput");
    if (!input) return { ok: false, value: "" };
    const project = String(projectName || "").trim();

    if (isInputDefaultBound(input)) {
      const resolvedDefault = getResolvedReservingClassValue();
      const normalizedDefault = normalizeReservingClassPath(resolvedDefault);
      if (!normalizedDefault) {
        if (strict && showMessage) {
          reportInputInvalid(
            input,
            "Default Path is empty.",
            "Invalid Reserving Class. Please select a value from the valid list.",
          );
        }
        return { ok: false, value: "" };
      }
      const validatedDefault = await validateReservingClassPathByTypeNames(project, normalizedDefault);
      if (!validatedDefault?.ok || !validatedDefault?.path) {
        if (strict && showMessage) {
          reportInputInvalid(
            input,
            "Default Path is not in the valid list for this project.",
            "Invalid Reserving Class. Please select a value from the valid list.",
          );
        }
        return { ok: false, value: "" };
      }
      const canonicalDefault = normalizeReservingClassPath(validatedDefault.path);
      clearInputInvalid(input);
      lastReservingClassSelection = canonicalDefault;
      return { ok: true, value: canonicalDefault };
    }

    const normalizedInput = normalizeReservingClassPath(input.value);
    if (!normalizedInput) {
      if (strict) {
        if (lastReservingClassSelection) input.value = lastReservingClassSelection;
        else input.value = "";
        clearInputInvalid(input);
        if (showMessage) {
          reportInputInvalid(
            input,
            "Reserving Class is not in the valid list for this project.",
            "Invalid Reserving Class. Please select a value from the valid list.",
          );
        }
      }
      return { ok: false, value: "" };
    }

    const validatedInput = await validateReservingClassPathByTypeNames(project, normalizedInput);
    if (!validatedInput?.ok || !validatedInput?.path) {
      if (strict) {
        if (lastReservingClassSelection) input.value = lastReservingClassSelection;
        else input.value = "";
        clearInputInvalid(input);
        if (showMessage) {
          reportInputInvalid(
            input,
            "Reserving Class is not in the valid list for this project.",
            "Invalid Reserving Class. Please select a value from the valid list.",
          );
        }
      }
      return { ok: false, value: "" };
    }

    input.value = normalizeReservingClassPath(validatedInput.path);
    clearInputInvalid(input);
    lastReservingClassSelection = input.value;
    return { ok: true, value: input.value };
  }

  async function validateTriInputsBeforeRun(options = {}) {
    const showMessage = !!options?.showMessage;
    const hasNameConflict = await refreshDatasetInstanceNameConflict();
    if (hasNameConflict) {
      setStatus(runtime.datasetInstanceNameConflictMessage || "Dataset instance name already exists.");
      return { ok: false };
    }
    if (!await ensureProjectValidationOptions()) return { ok: false };
    const projectResult = validateAndNormalizeProjectInput({ strict: true, showMessage });
    if (!projectResult.ok || !projectResult.value) return { ok: false };

    const project = projectResult.value;
    await Promise.all([
      refreshDatasetTypesForProject(project),
      refreshReservingClassPathsForProject(project),
    ]);

    const reservingResult = await validateAndNormalizeReservingClassInput(project, { strict: true, showMessage });
    if (!reservingResult.ok || !reservingResult.value) return { ok: false };

    const datasetResult = validateAndNormalizeDatasetInput({ strict: true, showMessage });
    if (!datasetResult.ok || !datasetResult.value) return { ok: false };
    const triInputs = getTriInputs();
    const dependencyResult = await validateDatasetTypeDependencies(datasetResult.value, {
      showMessage,
      precheckInputs: {
        project,
        path: reservingResult.value,
        tri: datasetResult.value,
        instanceName: triInputs.instanceName,
        cumulative: triInputs.cumulative,
        calendar: triInputs.calendar,
        originLen: triInputs.originLen,
        devLen: triInputs.devLen,
      },
    });
    if (!dependencyResult.ok) return { ok: false };

    saveTriInputsToStorage();
    return {
      ok: true,
      project,
      path: reservingResult.value,
      tri: datasetResult.value,
      instanceName: triInputs.instanceName,
      dependencyBypassedByExistingCsv: !!dependencyResult?.bypassedByExistingCsv,
    };
  }

  async function handleProjectSelection(value, options = {}) {
    const strict = !!options?.strict;
    const showMessage = !!options?.showMessage;
    const projectInput = document.getElementById("projectSelect");
    if (isDefaultTokenValue(value)) {
      if (projectInput) setInputDefaultBound(projectInput, true);
      clearInputInvalid(projectInput);
      const defaults = loadWorkflowDefaults();
      if (defaults?.project) {
        await applyResolvedProjectDefaults(defaults.project);
      }
      saveTriInputsToStorage();
      await syncSidecarForCurrentDataset({ applyLengths: true });
      scheduleAutoRun(0);
      return true;
    }

    if (projectInput) setInputDefaultBound(projectInput, false);

    const project = findExactProjectMatch(value);
    if (!project) {
      if (strict && projectInput) {
        if (runtime.lastProjectSelection) projectInput.value = runtime.lastProjectSelection;
        else projectInput.value = "";
        clearInputInvalid(projectInput);
        if (showMessage) {
          reportInputInvalid(
            projectInput,
            "Project Name is not in the valid list.",
            "Invalid Project Name. Please select a valid project.",
          );
        }
      }
      return false;
    }
    clearInputInvalid(projectInput);
    if (project === runtime.lastProjectSelection) {
      notifyProjectSelectionCommitted(project, "project-selection");
      return true;
    }

    runtime.lastProjectSelection = project;
    if (!isDfmDataTabHost()) {
      saveLastDatasetViewerProjectToAppData(project);
    }

    if (projectInput) projectInput.value = project;
    notifyProjectSelectionCommitted(project, "project-selection");
    showProjectDropdown(false);

    saveTriInputsToStorage();
    const showProjectSwitchPopup = !isRunInFlight();
    if (showProjectSwitchPopup) {
      showDatasetLoadingPopup("Validating Reserving Class");
    }
    try {
      await ensureHeadersForProject(project);
      await ensureDevHeadersForProject(project);
      await refreshDatasetTypesForProject(project);
      await refreshReservingClassPathsForProject(project);

      if (options?.applyProjectUserPreferences !== false && !isDfmDataTabHost()) {
        const prefs = await loadDatasetProjectPrefs(project);
        const pathInputForPrefs = document.getElementById("pathInput");
        const triInputForPrefs = document.getElementById("triInput");
        if (prefs?.path && pathInputForPrefs && !isInputDefaultBound(pathInputForPrefs)) {
          pathInputForPrefs.value = prefs.path;
        }
        if (prefs?.tri && triInputForPrefs) {
          triInputForPrefs.value = prefs.tri;
          setLastDatasetSelection(prefs.tri);
        }
      }

      const pathInput = document.getElementById("pathInput");
      if (pathInput) {
        const pathIsDefault = isInputDefaultBound(pathInput);
        const currentPath = pathIsDefault
          ? getResolvedReservingClassValue()
          : pathInput.value;
        const normalizedPath = normalizeReservingClassPath(currentPath);
        let validatedPath = "";
        if (normalizedPath) {
          const validated = await validateReservingClassPathByTypeNames(project, normalizedPath);
          if (validated?.ok && validated?.path) {
            validatedPath = ensureReservingClassOption(validated.path) || normalizeReservingClassPath(validated.path);
          }
        }

        if (validatedPath) {
          lastReservingClassSelection = validatedPath;
          if (!pathIsDefault) {
            pathInput.value = validatedPath;
          }
        } else {
          if (pathIsDefault) {
            setInputDefaultBound(pathInput, false);
          }
          pathInput.value = "";
          lastReservingClassSelection = "";
        }
        clearInputInvalid(pathInput);
      }

      const triInput = document.getElementById("triInput");
      if (triInput) {
        const matchedTri = findExactDatasetMatch(triInput.value);
        if (matchedTri) {
          triInput.value = matchedTri;
          lastDatasetSelection = matchedTri;
        } else {
          triInput.value = "";
          lastDatasetSelection = "";
        }
        clearInputInvalid(triInput);
      }

      await syncSidecarForCurrentDataset({ applyLengths: true });
      scheduleAutoRun();
      return true;
    } finally {
      if (showProjectSwitchPopup && !isRunInFlight()) {
        hideDatasetLoadingPopup();
      }
    }
  }


  runtime.LEN_DROPDOWN_CONFIG = LEN_DROPDOWN_CONFIG;

  Object.assign(runtime, {
    wireDataTabInputLifecycle,
    setLastProjectSelection,
    notifyProjectSelectionCommitted,
    setLastDatasetSelection,
    setLastReservingClassSelection,
    normalizeProjectText,
    getDatasetNumberFormatValue,
    setDatasetNumberFormatValue,
    getDatasetDecimalPlacesValue,
    getDatasetSyncedNumberFormatValue,
    setDatasetDecimalPlacesValue,
    normalizeSearchTokens,
    matchesProject,
    getActiveProjectValue,
    renderProjectOptions,
    showProjectDropdown,
    filterProjectOptions,
    getProjectFilterQuery,
    getProjectOptionsList,
    setActiveProjectIndex,
    getActiveProjectIndex,
    chooseActiveProject,
    findExactProjectMatch,
    getActiveDatasetValue,
    renderDatasetOptions,
    showDatasetDropdown,
    filterDatasetOptions,
    getDatasetOptionsList,
    setActiveDatasetIndex,
    getActiveDatasetIndex,
    chooseActiveDataset,
    findExactDatasetMatch,
    ensureDatasetTypeOption,
    getDatasetTypeFormulaByName,
    getDatasetTypeDataFormatByName,
    resizeDetailFormulaInput,
    syncDetailFormulaFromDatasetType,
    syncDetailDatasetTypeFromTopInput,
    loadDatasetTypeDependencyModel,
    validateDatasetTypeDependencies,
    setInputInvalid,
    clearInputInvalid,
    reportInputInvalid,
    rebuildReservingClassPathLookup,
    findExactReservingClassMatch,
    ensureReservingClassOption,
    refreshDatasetTypesForProject,
    refreshReservingClassPathsForProject,
    handleDatasetSelection,
    validateAndNormalizeProjectInput,
    ensureProjectValidationOptions,
    validateAndNormalizeDatasetInput,
    validateAndNormalizeReservingClassInput,
    validateTriInputsBeforeRun,
    handleProjectSelection,
  });
}
