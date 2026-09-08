/*
===============================================================================
DFM Details Tab - method name, project selection, path bar, threshold reset
===============================================================================
*/
import {
  getDfmInst,
  getDefaultMethodName,
  getDfmIsDirty,
  getResolvedProjectName,
  getResolvedReservingClass,
  markDfmDirty,
  sanitizeDfmMethodFilePart,
} from "/ui/method_pages/dfm/dfm_state.js";
import { resetRatioChartThresholds } from "/ui/method_pages/dfm/dfm_ratios_tab.js?v=20260907b";
import {
  scheduleRatioSelectionLoad,
} from "/ui/method_pages/dfm/dfm_persistence.js?v=20260907a";
import { openDatasetNamePicker } from "/ui/shared/components/pickers/dataset_name_picker.js";
import {
  loadProjectUserPreferences,
  scheduleProjectUserPreferencesSave,
} from "/ui/shared/services/project_user_preferences.js";
import { fetchProjectDatasetTypeItems } from "/ui/shared/dataset/dataset_types_source.js";

const outputTypeNamesByProject = new Map();
const dfmMethodNamesByProjectPath = new Map();
let outputTypeRequestSeq = 0;
let dfmMethodNameRequestSeq = 0;
let pendingOutputTypeFromUrl = null;
let localProjectPreferenceLoadPromise = null;

function toText(value) {
  return String(value ?? "").trim();
}

function normalizeKey(value) {
  return toText(value).toLowerCase();
}

function getInputSnapshotSafe() {
  try {
    if (typeof window.ADA_GET_DFM_INPUTS === "function") {
      return window.ADA_GET_DFM_INPUTS();
    }
  } catch {
    // ignore
  }
  return null;
}

function isProjectDefaultBound() {
  return !!getInputSnapshotSafe()?.defaults?.projectDefault;
}

function isReservingClassDefaultBound() {
  return !!getInputSnapshotSafe()?.defaults?.reservingClassDefault;
}

async function loadOutputTypeNames(projectName, _pathValue = "", options = {}) {
  const project = toText(projectName);
  if (!project) return [];
  const cacheKey = normalizeKey(project);
  if (!options?.forceReload && outputTypeNamesByProject.has(cacheKey)) {
    return outputTypeNamesByProject.get(cacheKey);
  }
  const payload = await fetchProjectDatasetTypeItems(project, { dedupeByName: true });
  const seen = new Set();
  const names = [];
  for (const item of Array.isArray(payload?.items) ? payload.items : []) {
    if (normalizeKey(item?.dataFormat) !== "vector") continue;
    const name = toText(item?.name);
    const key = normalizeKey(name);
    if (!name || !key || seen.has(key)) continue;
    seen.add(key);
    names.push(name);
  }
  names.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base", numeric: true }));
  outputTypeNamesByProject.set(cacheKey, names);
  return names;
}

function closeOutputTypeDropdown() {
  const dropdown = document.getElementById("dfmOutputVectorDropdown");
  if (!dropdown) return;
  dropdown.classList.remove("open");
  dropdown.innerHTML = "";
}

function closeTriangleTypeDropdown() {
  const dropdown = document.getElementById("dfmTriTypeDropdown");
  if (!dropdown) return;
  dropdown.classList.remove("open");
  dropdown.innerHTML = "";
}

function closeDfmMethodNameDropdown() {
  const dropdown = document.getElementById("dfmMethodNameDropdown");
  if (!dropdown) return;
  dropdown.classList.remove("open");
  dropdown.innerHTML = "";
}

function postDfmStatus(text, options = {}) {
  try {
    window.parent.postMessage(
      {
        type: "arcrho:status",
        text: String(text || ""),
        ...(options?.tone ? { tone: options.tone } : {}),
      },
      "*",
    );
  } catch {
    // ignore
  }
}

function syncMethodNameToOutputType(value, options = {}) {
  const next = toText(value);
  const methodInput = document.getElementById("dfmMethodName");
  if (!methodInput) return false;
  if (toText(methodInput.value)) {
    updateAppTabTitle(toText(methodInput.value) || getDefaultMethodName(), !options?.silent);
    return false;
  }
  const changed = toText(methodInput.value) !== next;
  if (changed) methodInput.value = next;
  updateAppTabTitle(next || getDefaultMethodName(), !options?.silent);
  if (changed && !options?.silent) {
    // Name is updated programmatically here, so the normal Name input change/blur
    // pipeline may not fire. Trigger local method lookup explicitly.
    queueMicrotask(() => scheduleRatioSelectionLoad("details-change"));
  }
  return changed;
}

function applyOutputTypeSelection(value, options = {}) {
  const input = document.getElementById("dfmOutputVector");
  if (!input) return;
  const next = toText(value);
  const prev = toText(input.value);
  const outputChanged = next !== prev;
  if (outputChanged) input.value = next;
  const methodChanged = syncMethodNameToOutputType(next, options);
  if (!outputChanged && !methodChanged) return;
  if (options?.silent) return;
  markDfmDirty();
  scheduleRatioSelectionLoad("details-change");
}

function applyTriangleSelection(value) {
  const input = document.getElementById("triInput");
  if (!input) return;
  const next = toText(value);
  if (!next) return;
  if (toText(input.value) === next) return;
  input.value = next;
  input.dispatchEvent(new Event("input", { bubbles: true }));
  input.dispatchEvent(new Event("change", { bubbles: true }));
}

function normalizeLocalProjectPreference(raw) {
  const source = raw && typeof raw === "object" ? raw : {};
  return String(
    source.projectName
    || source.project_name
    || source.project
    || "",
  ).trim();
}

async function loadLastLocalProjectName() {
  if (localProjectPreferenceLoadPromise) return localProjectPreferenceLoadPromise;
  localProjectPreferenceLoadPromise = (async () => {
    try {
      const response = await fetch("/local-project/preferences", { cache: "no-store" });
      if (!response.ok) return "";
      const payload = await response.json().catch(() => ({}));
      const project = normalizeLocalProjectPreference(payload?.preferences || payload);
      return project;
    } catch {
      return "";
    } finally {
      localProjectPreferenceLoadPromise = null;
    }
  })();
  return localProjectPreferenceLoadPromise;
}

function getLastReservingClassPathFromProjectPrefs(prefs) {
  const source = prefs && typeof prefs === "object" ? prefs : {};
  const direct = toText(source.lastReservingClassPath || source.last_reserving_class_path);
  if (direct) return direct;
  for (const sectionName of ["dfmObject", "datasetViewer"]) {
    const section = source[sectionName];
    if (!section || typeof section !== "object") continue;
    const path = toText(section.reservingClass || section.reserving_class || section.path);
    if (path) return path;
  }
  return "";
}

function getDfmObjectPrefsFromProjectPrefs(prefs) {
  const source = prefs && typeof prefs === "object" ? prefs : {};
  const dfmObject = source.dfmObject && typeof source.dfmObject === "object" ? source.dfmObject : {};
  return {
    methodName: toText(dfmObject.methodName || dfmObject.method_name),
    outputVector: toText(dfmObject.outputVector || dfmObject.output_vector),
    inputTriangle: toText(dfmObject.inputTriangle || dfmObject.input_triangle || dfmObject.datasetName || dfmObject.dataset_name),
    originLength: toText(dfmObject.originLength || dfmObject.origin_length),
    developmentLength: toText(dfmObject.developmentLength || dfmObject.development_length),
    decimalPlaces: toText(dfmObject.decimalPlaces || dfmObject.decimal_places),
  };
}

function applyTextInputValue(id, value, options = {}) {
  const input = document.getElementById(id);
  const next = toText(value);
  if (!input || !next) return false;
  if (!options?.replace && toText(input.value)) return false;
  if (toText(input.value) === next) return false;
  if (options?.programmatic) input.dataset.programmatic = "1";
  input.value = next;
  if (options?.dispatchInput) input.dispatchEvent(new Event("input", { bubbles: true }));
  if (options?.dispatchSelection) {
    input.dispatchEvent(new CustomEvent("arcrho:output-type-selected", { detail: { value: next } }));
  }
  return true;
}

function applySelectValue(id, value, options = {}) {
  const select = document.getElementById(id);
  const next = toText(value);
  if (!select || !next) return false;
  if (!options?.replace && toText(select.value)) return false;
  if (![...select.options].some((opt) => String(opt.value) === next)) {
    const opt = document.createElement("option");
    opt.value = next;
    opt.textContent = next;
    select.appendChild(opt);
  }
  if (toText(select.value) === next) return false;
  select.value = next;
  return true;
}

function applyDfmObjectPrefsToDetails(prefs, options = {}) {
  const dfmPrefs = getDfmObjectPrefsFromProjectPrefs(prefs);
  let changed = false;
  changed = applyTextInputValue("dfmMethodName", dfmPrefs.methodName, {
    replace: !!options?.replace,
    programmatic: true,
    dispatchInput: true,
  }) || changed;
  changed = applyTextInputValue("dfmOutputVector", dfmPrefs.outputVector, {
    replace: !!options?.replace,
    dispatchSelection: true,
  }) || changed;
  changed = applyTextInputValue("triInput", dfmPrefs.inputTriangle, {
    replace: !!options?.replace,
  }) || changed;
  changed = applySelectValue("originLenSelect", dfmPrefs.originLength, options) || changed;
  changed = applySelectValue("devLenSelect", dfmPrefs.developmentLength, options) || changed;
  changed = applyTextInputValue("decimalPlaces", dfmPrefs.decimalPlaces, options) || changed;

  const methodName = toText(document.getElementById("dfmMethodName")?.value);
  updateAppTabTitle(methodName || getDefaultMethodName());
  return changed;
}

function applyLastReservingClassPathFromProjectPrefs(prefs, options = {}) {
  const input = document.getElementById("pathInput");
  if (!input) return false;
  if (!options?.replace && toText(input.value)) return false;
  const path = getLastReservingClassPathFromProjectPrefs(prefs);
  if (!path) return false;
  if (toText(input.value) === path) return false;
  input.value = path;
  return true;
}

async function applyDfmProjectUserPreferences(projectName, options = {}) {
  const project = toText(projectName);
  if (!project) return false;
  try {
    const prefs = await loadProjectUserPreferences(project);
    const pathChanged = applyLastReservingClassPathFromProjectPrefs(prefs, options);
    const detailsChanged = applyDfmObjectPrefsToDetails(prefs, options);
    if (pathChanged || detailsChanged) {
      scheduleRatioSelectionLoad("details-change");
    }
    return pathChanged || detailsChanged;
  } catch {
    return false;
  }
}

function saveLastReservingClassPathForCurrentProject() {
  if (isProjectDefaultBound() || isReservingClassDefaultBound()) return;
  const project = toText(getResolvedProjectName());
  const path = toText(getResolvedReservingClass());
  if (!project || !path) return;
  scheduleProjectUserPreferencesSave(project, {
    lastReservingClassPath: path,
  });
}

function commitSelectedDfmProject(projectName, options = {}) {
  if (isProjectDefaultBound()) return;
  const project = toText(projectName || getResolvedProjectName());
  if (!project) return;
  if (options?.applyPreferences !== false) {
    void applyDfmProjectUserPreferences(project, { replace: true });
  }
}

async function applyLastLocalProjectNameIfBlank() {
  const input = document.getElementById("projectSelect");
  if (!input || toText(input.value)) return;
  const project = await loadLastLocalProjectName();
  if (!project || toText(input.value)) return;
  input.value = project;
  input.dispatchEvent(new Event("input", { bubbles: true }));
  void applyDfmProjectUserPreferences(project);
}

function normalizeDfmMethodIndexNames(payload) {
  const seen = new Set();
  const out = [];
  for (const item of Array.isArray(payload?.files) ? payload.files : []) {
    if (normalizeKey(item?.method_type) !== "dfm") continue;
    const name = toText(item?.method_name || item?.name);
    if (!name) continue;
    const key = normalizeKey(name);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push(name);
  }
  out.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base", numeric: true }));
  return out;
}

async function loadDfmMethodNames(projectName, pathValue, options = {}) {
  const project = toText(projectName);
  const pathPart = sanitizeDfmMethodFilePart(pathValue, "");
  if (!project || !pathPart) return [];
  const cacheKey = `${normalizeKey(project)}\n${normalizeKey(pathPart)}`;
  if (!options?.forceReload && dfmMethodNamesByProjectPath.has(cacheKey)) {
    return dfmMethodNamesByProjectPath.get(cacheKey);
  }
  const query = new URLSearchParams({
    project_name: project,
    reserving_class: pathValue,
    refresh: options?.forceReload ? "true" : "false",
  });
  const response = await fetch(`/dfm/method-index?${query.toString()}`);
  if (!response.ok) {
    let detail = "";
    try {
      detail = toText(await response.text());
    } catch {}
    throw new Error(detail || `HTTP ${response.status}`);
  }
  const payload = await response.json().catch(() => ({}));
  const names = normalizeDfmMethodIndexNames(payload);
  dfmMethodNamesByProjectPath.set(cacheKey, names);
  return names;
}

function renderDfmMethodNameDropdown(names) {
  const dropdown = document.getElementById("dfmMethodNameDropdown");
  const input = document.getElementById("dfmMethodName");
  if (!dropdown || !input) return;
  dropdown.innerHTML = "";
  if (!Array.isArray(names) || names.length === 0) {
    const option = document.createElement("div");
    option.className = "datasetOption";
    option.textContent = "No DFM methods found for this path.";
    option.style.cursor = "default";
    option.style.color = "#666";
    dropdown.appendChild(option);
    dropdown.classList.add("open");
    return;
  }
  const selectedKey = normalizeKey(input.value);
  for (const name of names) {
    const option = document.createElement("div");
    option.className = "datasetOption";
    option.textContent = name;
    if (normalizeKey(name) === selectedKey) option.classList.add("active");
    option.addEventListener("mousedown", (e) => {
      e.preventDefault();
    });
    option.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      input.dataset.programmatic = "1";
      input.value = name;
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new Event("change", { bubbles: true }));
      scheduleRatioSelectionLoad("details-change");
      closeDfmMethodNameDropdown();
    });
    dropdown.appendChild(option);
  }
  dropdown.classList.add("open");
}

async function syncOutputTypeForCurrentProject(options = {}) {
  const projectName = toText(getResolvedProjectName());
  const input = document.getElementById("dfmOutputVector");
  if (!input) return;
  const wasFocusedAtStart = document.activeElement === input;
  const valueAtStart = toText(input.value);

  if (!projectName) {
    applyOutputTypeSelection("", { silent: true });
    closeOutputTypeDropdown();
    return;
  }

  const requestSeq = ++outputTypeRequestSeq;
  try {
    const names = await loadOutputTypeNames(projectName, "", { forceReload: !!options?.forceReload });
    if (requestSeq !== outputTypeRequestSeq) return;
    const allowedKeys = new Set(names.map((name) => normalizeKey(name)));

    const pending = toText(pendingOutputTypeFromUrl);
    if (pending) {
      const matched = names.find((name) => normalizeKey(name) === normalizeKey(pending)) || "";
      applyOutputTypeSelection(matched, { silent: true });
      pendingOutputTypeFromUrl = "";
      return;
    }

    const current = toText(input.value);
    const isActivelyEditing = document.activeElement === input;
    // Avoid clobbering user typing if the async dataset-type response returns
    // while Output Vector is focused/being edited after page refresh.
    if (isActivelyEditing || (wasFocusedAtStart && current !== valueAtStart)) {
      return;
    }
    if (current && !allowedKeys.has(normalizeKey(current))) {
      applyOutputTypeSelection("", { silent: true });
    }
  } catch (err) {
    if (requestSeq !== outputTypeRequestSeq) return;
    console.error("Failed to load output vectors:", err);
  }
}

export function syncOutputTypeFromProject(options = {}) {
  void syncOutputTypeForCurrentProject(options);
}

export function updateAppTabTitle(title, userAction) {
  if (!title) return;
  const inst = getDfmInst();
  window.parent.postMessage({ type: "arcrho:update-active-tab-title", title, inst, userAction: !!userAction }, "*");
}

export function syncMethodNameFromInputs() {
  const input = document.getElementById("dfmMethodName");
  if (!input) return;
  const outputVector = toText(document.getElementById("dfmOutputVector")?.value);
  const current = toText(input.value);
  const next = current || outputVector || "";
  if (input.value !== next) input.value = next;
  updateAppTabTitle(next || getDefaultMethodName());
}

export function wireMethodName() {
  const input = document.getElementById("dfmMethodName");
  if (!input || input.dataset.wired === "1") return;
  input.dataset.wired = "1";
  let lastSeenValue = input.value.trim();
  let lastLookupCommittedValue = input.value.trim();

  const commitValue = (options = {}) => {
    const raw = input.value.trim();
    const programmatic = input.dataset.programmatic === "1";
    if (programmatic) delete input.dataset.programmatic;
    const valueChanged = raw !== lastSeenValue;
    const lookupValueChanged = raw !== lastLookupCommittedValue;
    updateAppTabTitle(raw || getDefaultMethodName(), true);
    if (valueChanged) {
      if (!programmatic) markDfmDirty();
      lastSeenValue = raw;
    }
    if (programmatic) {
      lastLookupCommittedValue = raw;
      return;
    }
    if (options?.triggerLoad && lookupValueChanged && !getDfmIsDirty()) {
      lastLookupCommittedValue = raw;
      scheduleRatioSelectionLoad("details-change");
    }
  };

  input.addEventListener("input", () => commitValue({ triggerLoad: false }));
  input.addEventListener("change", () => commitValue({ triggerLoad: true }));
  input.addEventListener("blur", () => commitValue({ triggerLoad: true }));

  const triInput = document.getElementById("triInput");
  const pathInput = document.getElementById("pathInput");
  const projectInput = document.getElementById("projectSelect");
  const originLen = document.getElementById("originLenSelect");
  const devLen = document.getElementById("devLenSelect");
  triInput?.addEventListener("change", syncMethodNameFromInputs);
  triInput?.addEventListener("input", syncMethodNameFromInputs);
  pathInput?.addEventListener("change", syncMethodNameFromInputs);
  originLen?.addEventListener("change", syncMethodNameFromInputs);
  devLen?.addEventListener("change", syncMethodNameFromInputs);

  projectInput?.addEventListener("change", () => {
    commitSelectedDfmProject(projectInput.value);
  });
  projectInput?.addEventListener("arcrho:project-selected", (event) => {
    commitSelectedDfmProject(event?.detail?.projectName || projectInput.value);
  });
  pathInput?.addEventListener("change", saveLastReservingClassPathForCurrentProject);

  const markDirtyOnChange = () => markDfmDirty();
  triInput?.addEventListener("change", markDirtyOnChange);
  pathInput?.addEventListener("change", markDirtyOnChange);
  projectInput?.addEventListener("change", markDirtyOnChange);
  originLen?.addEventListener("change", markDirtyOnChange);
  devLen?.addEventListener("change", markDirtyOnChange);

  const triggerLoad = () => scheduleRatioSelectionLoad("details-change");
  pathInput?.addEventListener("change", triggerLoad);
  projectInput?.addEventListener("change", triggerLoad);

  wireDfmMethodNamePicker();
  wireOutputTypePicker();
  wireTriangleTypePicker();
  void applyLastLocalProjectNameIfBlank();
}

function wireDfmMethodNamePicker() {
  const input = document.getElementById("dfmMethodName");
  const button = document.getElementById("dfmMethodNameBtn");
  const dropdown = document.getElementById("dfmMethodNameDropdown");
  if (!input || !button || !dropdown || button.dataset.wired === "1") return;
  button.dataset.wired = "1";

  const openPicker = async (options = {}) => {
    const projectName = toText(getResolvedProjectName());
    const pathValue = toText(getResolvedReservingClass());
    if (!projectName) {
      closeDfmMethodNameDropdown();
      postDfmStatus("Select a project first.", { tone: "warn" });
      return;
    }
    if (!pathValue) {
      closeDfmMethodNameDropdown();
      postDfmStatus("Select a reserving class first.", { tone: "warn" });
      return;
    }
    button.disabled = true;
    const requestSeq = ++dfmMethodNameRequestSeq;
    try {
      const names = await loadDfmMethodNames(projectName, pathValue, { forceReload: !!options?.forceReload });
      if (requestSeq !== dfmMethodNameRequestSeq) return;
      renderDfmMethodNameDropdown(names);
    } catch (err) {
      console.error("Failed to load DFM method names:", err);
      closeDfmMethodNameDropdown();
      postDfmStatus(`Error loading DFM method names: ${String(err?.message || err)}`, { tone: "error" });
    } finally {
      if (requestSeq === dfmMethodNameRequestSeq) button.disabled = false;
    }
  };

  button.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    void openPicker({ forceReload: true });
  });

  input.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      closeDfmMethodNameDropdown();
      return;
    }
    if (e.key === "ArrowDown" && !dropdown.classList.contains("open")) {
      e.preventDefault();
      void openPicker();
    }
  });

  document.addEventListener("click", (e) => {
    if (dropdown.contains(e.target) || button.contains(e.target) || input.contains(e.target)) return;
    closeDfmMethodNameDropdown();
  });
}

function setDfmInstanceMissingNoticeVisible(visible) {
  const notice = document.getElementById("dfmInstanceMissingNotice");
  if (!notice) return;
  notice.classList.toggle("show", !!visible);
}

export function wireDfmInstanceCreationNotice() {
  const notice = document.getElementById("dfmInstanceMissingNotice");
  if (!notice || notice.dataset.wired === "1") return;
  notice.dataset.wired = "1";
  setDfmInstanceMissingNoticeVisible(false);
  // Inline "missing instance" notice beside Name is intentionally disabled.
  // New-object guidance is shown in the shell status bar (yellow warning) instead.
}

function wireOutputTypePicker() {
  const input = document.getElementById("dfmOutputVector");
  const button = document.getElementById("dfmOutputVectorBtn");
  const dropdown = document.getElementById("dfmOutputVectorDropdown");
  if (!input || !button || !dropdown || button.dataset.wired === "1") return;
  button.dataset.wired = "1";
  input.readOnly = false;

  if (pendingOutputTypeFromUrl == null) {
    const query = new URLSearchParams(window.location.search);
    pendingOutputTypeFromUrl = toText(query.get("output_type"));
  }

  let pickerProjectKey = "";
  let pickerNames = [];
  let pickerLoaded = false;
  let committedOutputType = toText(input.value);

  const resetPickerCache = () => {
    pickerProjectKey = "";
    pickerNames = [];
    pickerLoaded = false;
  };

  const ensurePickerNames = async (options = {}) => {
    const projectName = toText(getResolvedProjectName());
    if (!projectName) {
      resetPickerCache();
      return { projectName: "", names: [] };
    }
    const projectKey = normalizeKey(projectName);
    const forceReload = !!options?.forceReload;
    const projectChanged = projectKey !== pickerProjectKey;
    if (forceReload || projectChanged || !pickerLoaded) {
      const requestSeq = ++outputTypeRequestSeq;
      const names = await loadOutputTypeNames(projectName, "", {
        forceReload: forceReload || projectChanged || !pickerLoaded,
      });
      if (requestSeq !== outputTypeRequestSeq) return null;
      pickerProjectKey = projectKey;
      pickerNames = Array.isArray(names) ? names : [];
      pickerLoaded = true;
    }
    return { projectName, names: pickerNames };
  };

  const openPicker = async (options = {}) => {
    const projectName = toText(getResolvedProjectName());
    if (!projectName) {
      closeOutputTypeDropdown();
      if (options?.alertOnProjectMissing) alert("Select a project first.");
      return;
    }
    button.disabled = true;
    try {
      closeOutputTypeDropdown();
      await openDatasetNamePicker({
        projectName,
        initialName: input.value,
        anchorElement: input,
        title: "Select Output Type",
        allowedDataFormats: ["Vector"],
        forceReload: !!options?.forceReload,
        emptyMessage: "No output types found (Vector).",
        setStatus: (message) => {
          const text = toText(message);
          if (text) postDfmStatus(text, { tone: "warn" });
        },
        onError: (err) => {
          console.error("Failed to open output vector picker:", err);
          postDfmStatus(`Error loading output vectors: ${String(err?.message || err)}`, { tone: "error" });
        },
        onSelect: (name) => {
          const selected = toText(name);
          if (!selected) return;
          applyOutputTypeSelection(selected);
          committedOutputType = selected;
          input.dispatchEvent(new CustomEvent("arcrho:output-type-selected", { detail: { value: selected } }));
        },
      });
    } catch (err) {
      console.error("Failed to load output vector options:", err);
      alert(`Error loading output vectors: ${err?.message || err}`);
    } finally {
      button.disabled = false;
    }
  };

  const commitTypedOutputTypeIfNeeded = async () => {
    const out = await ensurePickerNames({ forceReload: false });
    if (!out) return;
    const typed = toText(input.value);
    if (!typed) {
      if (committedOutputType) applyOutputTypeSelection("");
      committedOutputType = "";
      return;
    }
    const exact = out.names.find((name) => normalizeKey(name) === normalizeKey(typed));
    if (exact) {
      if (normalizeKey(exact) !== normalizeKey(committedOutputType)) {
        applyOutputTypeSelection(exact);
      } else if (toText(input.value) !== exact) {
        input.value = exact;
      }
      committedOutputType = exact;
      return;
    }
    input.value = committedOutputType;
  };

  button.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    void openPicker({ forceReload: true, alertOnProjectMissing: true });
  });

  // The picker button is the only pointer path that opens the list. Clicking or
  // tabbing into the field puts a caret in it, which is what a text box should
  // do; ArrowDown below still opens the list from the keyboard.
  input.addEventListener("focus", () => {
    committedOutputType = toText(input.value);
  });

  input.addEventListener("input", () => {
    closeOutputTypeDropdown();
  });

  input.addEventListener("change", () => {
    void commitTypedOutputTypeIfNeeded();
  });

  input.addEventListener("blur", () => {
    setTimeout(() => {
      if (!dropdown.contains(document.activeElement) && !button.contains(document.activeElement)) {
        void commitTypedOutputTypeIfNeeded();
        closeOutputTypeDropdown();
      }
    }, 0);
  });

  input.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      closeOutputTypeDropdown();
      return;
    }
    if (e.key === "ArrowDown") {
      e.preventDefault();
      void openPicker({ forceReload: false, alertOnProjectMissing: false });
    }
  });

  document.addEventListener("mousedown", (e) => {
    if (!dropdown.classList.contains("open")) return;
    const target = e.target;
    if (dropdown.contains(target) || button.contains(target) || input.contains(target)) return;
    closeOutputTypeDropdown();
  }, true);

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeOutputTypeDropdown();
  }, true);

  const projectInput = document.getElementById("projectSelect");
  const pathInput = document.getElementById("pathInput");
  projectInput?.addEventListener("change", () => {
    resetPickerCache();
    committedOutputType = toText(input.value);
    closeOutputTypeDropdown();
    void syncOutputTypeForCurrentProject({ forceReload: true });
  });
  projectInput?.addEventListener("input", () => {
    resetPickerCache();
    closeOutputTypeDropdown();
  });
  pathInput?.addEventListener("change", () => {
    resetPickerCache();
    committedOutputType = toText(input.value);
    closeOutputTypeDropdown();
    void syncOutputTypeForCurrentProject({ forceReload: true });
  });
  pathInput?.addEventListener("input", () => {
    resetPickerCache();
    closeOutputTypeDropdown();
  });

  input.addEventListener("arcrho:output-type-selected", () => {
    committedOutputType = toText(input.value);
  });

  void syncOutputTypeForCurrentProject();
}

function wireTriangleTypePicker() {
  const triInput = document.getElementById("triInput");
  const button = document.getElementById("dfmTriTypeBtn");
  const dropdown = document.getElementById("dfmTriTypeDropdown");
  if (!triInput || !button || !dropdown || button.dataset.wired === "1") return;
  button.dataset.wired = "1";

  const resetPickerCache = () => {
    closeTriangleTypeDropdown();
    const nativeDatasetDropdown = document.getElementById("datasetDropdown");
    nativeDatasetDropdown?.classList.remove("open");
  };

  const openPicker = async (options = {}) => {
    const projectName = toText(getResolvedProjectName());
    closeTriangleTypeDropdown();
    const nativeDatasetDropdown = document.getElementById("datasetDropdown");
    nativeDatasetDropdown?.classList.remove("open");
    if (!projectName) {
      if (options?.alertOnContextMissing) alert("Select a project first.");
      return;
    }
    button.disabled = true;
    try {
      await openDatasetNamePicker({
        projectName,
        initialName: triInput.value,
        anchorElement: triInput,
        title: "Select Input Triangle",
        allowedDataFormats: ["Triangle"],
        forceReload: !!options?.forceReload,
        emptyMessage: "No input triangles found (Triangle).",
        setStatus: (message) => {
          const text = toText(message);
          if (text) postDfmStatus(text, { tone: "warn" });
        },
        onError: (err) => {
          console.error("Failed to open input triangle picker:", err);
          postDfmStatus(`Error loading triangle names: ${String(err?.message || err)}`, { tone: "error" });
        },
        onSelect: (name) => {
          applyTriangleSelection(name);
        },
      });
    } catch (err) {
      console.error("Failed to load input-triangle options:", err);
      postDfmStatus(`Error loading triangle names: ${String(err?.message || err)}`, { tone: "error" });
    } finally {
      button.disabled = false;
    }
  };

  button.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    void openPicker({ forceReload: true, alertOnContextMissing: true });
  });

  triInput.addEventListener("input", () => {
    closeTriangleTypeDropdown();
    const nativeDatasetDropdown = document.getElementById("datasetDropdown");
    nativeDatasetDropdown?.classList.remove("open");
  });

  triInput.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      closeTriangleTypeDropdown();
      return;
    }
    if (e.key === "ArrowDown") {
      e.preventDefault();
      void openPicker({ forceReload: false, alertOnContextMissing: false });
    }
  });

  document.addEventListener("mousedown", (e) => {
    if (!dropdown.classList.contains("open")) return;
    const target = e.target;
    if (dropdown.contains(target) || button.contains(target)) return;
    closeTriangleTypeDropdown();
  }, true);

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeTriangleTypeDropdown();
  }, true);

  const projectInput = document.getElementById("projectSelect");
  const pathInput = document.getElementById("pathInput");
  projectInput?.addEventListener("change", () => {
    resetPickerCache();
    closeTriangleTypeDropdown();
  });
  projectInput?.addEventListener("input", () => {
    resetPickerCache();
    closeTriangleTypeDropdown();
  });
  pathInput?.addEventListener("change", () => {
    resetPickerCache();
    closeTriangleTypeDropdown();
  });
  pathInput?.addEventListener("input", () => {
    resetPickerCache();
    closeTriangleTypeDropdown();
  });
}

export function wireDetailsThresholdReset() {
  const detailsPage = document.getElementById("dfmDetailsPage");
  if (!detailsPage || detailsPage.dataset.thresholdWired === "1") return;
  detailsPage.dataset.thresholdWired = "1";
  const handleChange = () => {
    resetRatioChartThresholds();
  };
  detailsPage.addEventListener("input", handleChange, { capture: true });
  detailsPage.addEventListener("change", handleChange, { capture: true });
}
