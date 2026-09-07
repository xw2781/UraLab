// Owns draft models, period controls, run request construction, and header rules.

import {
  beginDatasetGridLoading,
  endDatasetGridLoading,
  renderDatasetGridPlaceholder,
  setDatasetGridError,
} from "/ui/shared/tabs/data/dataset_grid_placeholder.js?v=20260809a";
import { attachArcrhoTooltip } from "/ui/shared/components/tooltip/tooltip.js?v=20260812a";

export function registerDataTabRequestController(runtime) {
  const { state, config, isTemporaryDatasetView, qs, temporaryDatasetSessionId } = runtime;
  const defer = (name) => (...args) => runtime[name](...args);
  const { readDatasetInputQueryValues, normalizeReservingClassPath, normalizeBrowsingHistoryEntry, validateDatasetOriginLabels, updateDatasetSaveUi, getDatasetInstanceNameValue, getResolvedProjectValue, getResolvedReservingClassValue, notifyDatasetUpdated, renderTable, renderChart, setStatus, showDatasetDropdown, showProjectDropdown, getDatasetDecimalPlacesValue, getDatasetSyncedNumberFormatValue, getDatasetTypeDataFormatByName, getStoredInputValue, isDfmDataTabHost, isInputDefaultBound, loadLastDsId, loadProjectValidValueList, renderProjectOptions, saveLastDsId, createDatasetDependencyGuard } = new Proxy({}, { get: (_target, name) => defer(name) });
  const normalizeProjectText = defer("normalizeProjectText");
  const setInputInvalid = defer("setInputInvalid");
  const clearInputInvalid = defer("clearInputInvalid");
  let projectInstanceDraftRefreshSeq = 0;
  const lenDropdownActiveIndexBySelect = new Map();
  function readDatasetInputsFromQueryParams() {
    const {
      project,
      path: rawPath,
      tri,
      instanceName,
      originLen,
      devLen,
      dataFormat,
      numberFormat,
      decimalPlaces,
    } = readDatasetInputQueryValues(qs);
    const path = normalizeReservingClassPath(rawPath);
    const normalized = normalizeBrowsingHistoryEntry({ project, path, tri });
    if (normalized && instanceName) normalized.instanceName = instanceName;
    if (normalized && originLen) normalized.originLen = originLen;
    if (normalized && devLen) normalized.devLen = devLen;
    if (normalized && dataFormat) normalized.dataFormat = dataFormat;
    if (normalized && numberFormat) normalized.numberFormat = numberFormat;
    if (normalized && decimalPlaces) normalized.decimalPlaces = decimalPlaces;
    return normalized;
  }

  function normalizeDraftDataFormat(value) {
    return String(value || "").trim().replace(/\s+/g, " ").toLowerCase() === "vector"
      ? "Vector"
      : "Triangle";
  }

  function getProjectInstanceDraftDataFormat() {
    const queryInputs = readDatasetInputsFromQueryParams();
    return normalizeDraftDataFormat(queryInputs?.dataFormat);
  }

  function numericDevelopmentLabels(count) {
    const safeCount = Number.isFinite(count) && count > 0 ? Math.trunc(count) : 12;
    return Array.from({ length: safeCount }, (_, index) => String(index + 1));
  }

  function resolveDevelopmentLabels(labels, fallbackCount) {
    if (Array.isArray(labels) && labels.length) return labels.map(String);
    return numericDevelopmentLabels(fallbackCount);
  }

  function buildProjectInstanceDraftMask(originCount, devCount, dataFormat) {
    const isVector = normalizeDraftDataFormat(dataFormat) === "Vector";
    return Array.from({ length: originCount }, (_, r) => (
      Array.from({ length: devCount }, (_, c) => isVector || r + c < devCount)
    ));
  }

  function buildProjectInstanceDraftModel() {
    const { originLen, devLen } = getTriInputs();
    const dataFormat = getProjectInstanceDraftDataFormat();
    const isVector = dataFormat === "Vector";
    const originResult = validateDatasetOriginLabels(state.headerLabels, {
      originLen,
      requireMatchingPeriod: true,
    });
    if (!originResult.ok) {
      throw new Error(
        `Cannot create dataset draft: ${originResult.error}. `
        + "Set a valid Origin Start Date in Project Settings, then try again.",
      );
    }
    const originLabels = originResult.labels;
    const projectDevLabels = resolveDevelopmentLabels(state.devHeaderLabels, devLen);
    const devLabels = isVector ? [projectDevLabels[0] || "1"] : projectDevLabels;
    const originCount = Math.max(1, originLabels.length);
    const devCount = Math.max(1, devLabels.length);
    const mask = buildProjectInstanceDraftMask(originCount, devCount, dataFormat);
    const values = mask.map((row) => row.map((hasValue) => (hasValue ? 0 : null)));
    return {
      id: `draft:${getDatasetInstanceNameValue() || getTriInputs().tri || "dataset"}`,
      origin_labels: originLabels,
      dev_labels: devLabels,
      values,
      mask,
      data_format: dataFormat,
      mtime: null,
    };
  }

  function initializeProjectInstanceDraftModel() {
    state.dirty.clear();
    state.fileMtime = null;
    state.model = buildProjectInstanceDraftModel();
    const meta = document.getElementById("dsMeta");
    if (meta) {
      meta.textContent = `draft | origins=${state.model.origin_labels.length} | dev=${state.model.dev_labels.length}`;
    }
  }

  async function refreshProjectInstanceDraftModel() {
    const gridPlaceholderToken = beginDatasetGridLoading({ message: "Preparing dataset draft" });
    try {
      return await refreshProjectInstanceDraftModelOnce();
    } finally {
      endDatasetGridLoading(gridPlaceholderToken);
    }
  }

  async function refreshProjectInstanceDraftModelOnce() {
    const refreshSeq = ++projectInstanceDraftRefreshSeq;
    const project = getResolvedProjectValue();
    const { originLen, devLen } = getTriInputs();
    const isCurrent = () => {
      if (refreshSeq !== projectInstanceDraftRefreshSeq) return false;
      const current = getTriInputs();
      return getResolvedProjectValue() === project
        && String(current.originLen ?? "") === String(originLen ?? "")
        && String(current.devLen ?? "") === String(devLen ?? "");
    };
    try {
      if (!project) throw new Error("Cannot create dataset draft: project name is missing.");
      await ensureHeadersForProject(project, {
        forceRefresh: true,
        throwOnError: true,
        isCurrent,
      });
      if (!isCurrent()) return false;
      await ensureDevHeadersForProject(project, { forceRefresh: true, isCurrent });
      if (!isCurrent()) return false;
      initializeProjectInstanceDraftModel();
      renderTable();
      notifyDatasetUpdated();
      renderChart();
      // Save is blocked while the draft has no placeholder grid, so the save bar
      // has to be refreshed once the grid exists or fails to build.
      updateDatasetSaveUi();
      setStatus("Ready to edit new dataset draft.");
      return true;
    } catch (err) {
      if (!isCurrent()) return false;
      const message = String(err?.message || err || "Origin labels are unavailable.");
      state.model = null;
      state.fileMtime = null;
      state.headerLabels = [];
      setDatasetGridError(message);
      renderDatasetGridPlaceholder(document.getElementById("tableWrap"));
      const meta = document.getElementById("dsMeta");
      if (meta) meta.textContent = "";
      renderChart();
      notifyDatasetUpdated({ publishPreview: false });
      updateDatasetSaveUi();
      setStatus(message);
      return false;
    }
  }

  function getLenDropdownIds(selectId) {
    return runtime.LEN_DROPDOWN_CONFIG[selectId] || null;
  }

  function getLenDropdownElements(selectId) {
    const ids = getLenDropdownIds(selectId);
    if (!ids) return null;
    return {
      select: document.getElementById(selectId),
      wrap: document.getElementById(ids.wrapId),
      button: document.getElementById(ids.buttonId),
      dropdown: document.getElementById(ids.dropdownId),
    };
  }

  function getLenDropdownActiveIndex(selectId) {
    const idx = lenDropdownActiveIndexBySelect.get(selectId);
    return Number.isInteger(idx) ? idx : -1;
  }

  function setLenDropdownActiveIndex(selectId, idx, dir = 1) {
    const parts = getLenDropdownElements(selectId);
    const dropdown = parts?.dropdown;
    if (!dropdown) return;
    const opts = Array.from(dropdown.children);
    if (!opts.length) {
      lenDropdownActiveIndexBySelect.set(selectId, -1);
      return;
    }
    let next = Number.isFinite(idx) ? idx : 0;
    if (next < 0) next = opts.length - 1;
    if (next >= opts.length) next = 0;
    // A muted length is shown but cannot be chosen, so the highlight steps over
    // it in the direction the user was already moving.
    const step = dir >= 0 ? 1 : -1;
    for (let seen = 0; seen < opts.length && opts[next]?.dataset?.unavailable === "1"; seen += 1) {
      next += step;
      if (next < 0) next = opts.length - 1;
      if (next >= opts.length) next = 0;
    }
    if (opts[next]?.dataset?.unavailable === "1") {
      lenDropdownActiveIndexBySelect.set(selectId, -1);
      opts.forEach((el) => el.classList.toggle("active", false));
      return;
    }
    lenDropdownActiveIndexBySelect.set(selectId, next);
    opts.forEach((el, i) => el.classList.toggle("active", i === next));
    opts[next]?.scrollIntoView?.({ block: "nearest" });
  }

  // A length the open dataset has no use for is locked rather than removed, so
  // the top bar keeps its shape. The lock lives on the wrap element instead of a
  // module variable because every label repaint runs through the helpers below:
  // a lock held anywhere else would be overwritten by the next options render.
  function isLenSelectLocked(selectId) {
    return !!getLenDropdownElements(selectId)?.wrap?.hasAttribute("data-locked-value");
  }

  function getLenSelectLockedLabel(selectId) {
    return getLenDropdownElements(selectId)?.wrap?.getAttribute("data-locked-value") || "";
  }

  function setLenSelectLock(selectId, options = {}) {
    const parts = getLenDropdownElements(selectId);
    const wrap = parts?.wrap;
    const button = parts?.button;
    if (!wrap || !button) return;
    const locked = !!options.locked;
    if (locked) {
      wrap.setAttribute("data-locked-value", String(options.displayValue ?? ""));
      wrap.setAttribute("data-locked-reason", String(options.reason || ""));
      showLenDropdown(selectId, false);
    } else {
      wrap.removeAttribute("data-locked-value");
      wrap.removeAttribute("data-locked-reason");
    }
    wrap.classList.toggle("is-locked", locked);
    button.setAttribute("aria-disabled", locked ? "true" : "false");
    // Kept out of the tab order rather than disabled, so the reason tooltip is
    // still reachable by hover.
    button.tabIndex = locked ? -1 : 0;
    syncLenDropdownButtonLabel(selectId);
  }

  function syncLenDropdownButtonLabel(selectId) {
    const parts = getLenDropdownElements(selectId);
    const select = parts?.select;
    const button = parts?.button;
    if (!select || !button) return;
    const label = button.querySelector(".lenSelectValue");
    if (!label) return;
    if (isLenSelectLocked(selectId)) {
      label.textContent = getLenSelectLockedLabel(selectId);
      return;
    }
    const selected = select.options[select.selectedIndex];
    label.textContent = (selected?.textContent || select.value || "").trim();
  }

  // A length the stored period rules out stays in the list and reads as
  // unavailable, so the menu shows the whole ladder and says which rungs this
  // dataset can be displayed at.
  function lenOptionIsUnavailable(option) {
    return option?.dataset?.unavailable === "1" || option?.disabled === true;
  }

  function renderLenDropdownOptions(selectId) {
    const parts = getLenDropdownElements(selectId);
    const select = parts?.select;
    const dropdown = parts?.dropdown;
    if (!select || !dropdown) return;

    dropdown.innerHTML = "";
    const options = Array.from(select.options);
    if (!options.length) {
      lenDropdownActiveIndexBySelect.set(selectId, -1);
      syncLenDropdownButtonLabel(selectId);
      showLenDropdown(selectId, false);
      return;
    }

    options.forEach((opt, i) => {
      const item = document.createElement("div");
      const unavailable = lenOptionIsUnavailable(opt);
      item.className = unavailable ? "datasetOption lenOption is-unavailable" : "datasetOption lenOption";
      item.textContent = String(opt.textContent || opt.value || "");
      item.dataset.value = String(opt.value || "");
      item.dataset.index = String(i);
      if (unavailable) {
        item.dataset.unavailable = "1";
        item.setAttribute("aria-disabled", "true");
        attachArcrhoTooltip(item, () => String(opt.title || ""));
      }
      item.addEventListener("mouseenter", () => {
        if (unavailable) return;
        setLenDropdownActiveIndex(selectId, i);
      });
      item.addEventListener("mousedown", (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (unavailable) return;
        setLenSelectValue(selectId, opt.value, { emitChange: true });
        showLenDropdown(selectId, false);
        parts.button?.focus();
      });
      dropdown.appendChild(item);
    });

    const selectedIdx = options.findIndex((opt) => opt.value === select.value);
    setLenDropdownActiveIndex(selectId, selectedIdx >= 0 ? selectedIdx : 0);
    syncLenDropdownButtonLabel(selectId);
  }

  function refreshLenDropdowns() {
    Object.keys(runtime.LEN_DROPDOWN_CONFIG).forEach((selectId) => {
      renderLenDropdownOptions(selectId);
    });
  }

  function showLenDropdown(selectId, open) {
    const parts = getLenDropdownElements(selectId);
    const wrap = parts?.wrap;
    const dropdown = parts?.dropdown;
    const button = parts?.button;
    if (!wrap || !dropdown || !button) return;

    if (open) {
      Object.keys(runtime.LEN_DROPDOWN_CONFIG).forEach((id) => {
        if (id !== selectId) showLenDropdown(id, false);
      });
    }

    const shouldOpen = !!open && !!dropdown.children.length && !isLenSelectLocked(selectId);
    wrap.classList.toggle("open", shouldOpen);
    dropdown.classList.toggle("open", shouldOpen);
    button.setAttribute("aria-expanded", shouldOpen ? "true" : "false");
  }

  function closeAllLenDropdowns() {
    Object.keys(runtime.LEN_DROPDOWN_CONFIG).forEach((selectId) => {
      showLenDropdown(selectId, false);
    });
  }

  function setLenSelectValue(selectId, value, options = {}) {
    const emitChange = !!options?.emitChange;
    const select = document.getElementById(selectId);
    if (!select) return false;
    const nextValue = String(value ?? "");
    const option = [...select.options].find((opt) => opt.value === nextValue);
    if (!option || lenOptionIsUnavailable(option)) return false;
    const changed = select.value !== nextValue;
    select.value = nextValue;
    syncLenDropdownButtonLabel(selectId);
    renderLenDropdownOptions(selectId);
    if (emitChange && changed) {
      select.dispatchEvent(new Event("change"));
    }
    return true;
  }

  function chooseActiveLenDropdownOption(selectId) {
    const select = document.getElementById(selectId);
    if (!select || !select.options.length) return false;
    const idx = getLenDropdownActiveIndex(selectId);
    let nextIdx = idx;
    if (nextIdx < 0 || nextIdx >= select.options.length) {
      nextIdx = Math.max(0, select.selectedIndex);
    }
    const opt = select.options[nextIdx];
    if (!opt || lenOptionIsUnavailable(opt)) return false;
    const changed = select.value !== opt.value;
    select.value = opt.value;
    syncLenDropdownButtonLabel(selectId);
    renderLenDropdownOptions(selectId);
    showLenDropdown(selectId, false);
    if (changed) select.dispatchEvent(new Event("change"));
    return true;
  }

  function moveLenDropdownActiveOption(selectId, dir) {
    const parts = getLenDropdownElements(selectId);
    const dropdown = parts?.dropdown;
    if (!dropdown || !dropdown.children.length) return;
    const idx = getLenDropdownActiveIndex(selectId);
    const baseIdx = idx >= 0 ? idx : 0;
    setLenDropdownActiveIndex(selectId, baseIdx + dir, dir);
  }

  function cycleLenSelect(selectId, dir) {
    const select = document.getElementById(selectId);
    if (!select) return;
    let idx = select.selectedIndex + dir;
    // The wheel walks past the muted lengths instead of stopping on one.
    while (idx >= 0 && idx < select.options.length && lenOptionIsUnavailable(select.options[idx])) {
      idx += dir;
    }
    if (idx < 0 || idx >= select.options.length) return;
    select.selectedIndex = idx;
    syncLenDropdownButtonLabel(selectId);
    renderLenDropdownOptions(selectId);
    select.dispatchEvent(new Event("change"));
  }

  function wireLenDropdown(selectId) {
    const parts = getLenDropdownElements(selectId);
    const select = parts?.select;
    const button = parts?.button;
    const wrap = parts?.wrap;
    if (!select || !button || !wrap) return;
    if (button.dataset.wired === "1") return;
    button.dataset.wired = "1";

    // The only tooltip a length control carries is the reason it is locked;
    // the period the file is held at reads off the `Stored at` value beside it.
    attachArcrhoTooltip(button, () => wrap.getAttribute("data-locked-reason") || "");

    button.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (isLenSelectLocked(selectId)) return;
      showProjectDropdown(false);
      showDatasetDropdown(false);
      const willOpen = !wrap.classList.contains("open");
      if (willOpen) renderLenDropdownOptions(selectId);
      showLenDropdown(selectId, willOpen);
    });

    button.addEventListener("keydown", (e) => {
      if (isLenSelectLocked(selectId)) return;
      const key = e.key;
      if (key === "ArrowDown" || key === "ArrowUp") {
        e.preventDefault();
        e.stopPropagation();
        if (!wrap.classList.contains("open")) {
          renderLenDropdownOptions(selectId);
          showLenDropdown(selectId, true);
        }
        moveLenDropdownActiveOption(selectId, key === "ArrowDown" ? 1 : -1);
        return;
      }
      if (key === "Enter" || key === " ") {
        e.preventDefault();
        e.stopPropagation();
        if (wrap.classList.contains("open")) {
          chooseActiveLenDropdownOption(selectId);
        } else {
          renderLenDropdownOptions(selectId);
          showLenDropdown(selectId, true);
        }
        return;
      }
      if (key === "Escape" && wrap.classList.contains("open")) {
        e.preventDefault();
        e.stopPropagation();
        showLenDropdown(selectId, false);
      }
    });

    button.addEventListener("wheel", (e) => {
      if (isLenSelectLocked(selectId)) return;
      e.preventDefault();
      e.stopPropagation();
      const dir = e.deltaY > 0 ? 1 : -1;
      cycleLenSelect(selectId, dir);
    }, { passive: false });

    wrap.addEventListener("focusout", (e) => {
      const next = e.relatedTarget;
      if (next && wrap.contains(next)) return;
      showLenDropdown(selectId, false);
    });

    select.addEventListener("change", () => {
      syncLenDropdownButtonLabel(selectId);
      renderLenDropdownOptions(selectId);
    });

    syncLenDropdownButtonLabel(selectId);
    renderLenDropdownOptions(selectId);
    showLenDropdown(selectId, false);
  }

  function wireLenDropdowns() {
    Object.keys(runtime.LEN_DROPDOWN_CONFIG).forEach((selectId) => {
      wireLenDropdown(selectId);
    });
  }

  function initializeDatasetId() {
    const dsFromUrl = qs.get("ds");
    if (dsFromUrl) {
      config.DS_ID = dsFromUrl;
      saveLastDsId(dsFromUrl);
    } else {
      const saved = loadLastDsId();
      if (saved) config.DS_ID = saved;
    }
  }

  const LEN_CHOICES = [12, 6, 3, 1];

  // A length control may only offer a whole multiple of the period the data is
  // stored at: anything finer would ask for cells the file does not hold. The
  // ladder itself stays LEN_CHOICES, so one list decides every length the top
  // bar can ever show, and the stored period only narrows it.
  function lenChoicesForStoredLength(storedLength) {
    const stored = Number(storedLength);
    if (!Number.isFinite(stored) || stored <= 0) return LEN_CHOICES.slice();
    const offered = LEN_CHOICES.filter((choice) => choice % stored === 0);
    return offered.length ? offered : [stored];
  }

  function lenUnavailableReason(storedLength) {
    const stored = Number(storedLength);
    if (!Number.isFinite(stored) || stored <= 0) return "";
    return `This dataset is stored at ${stored}, so it can only be shown at a whole multiple of ${stored}.`;
  }

  // A `Stored at` control runs the other way round: the file may be held finer
  // than it is shown but never coarser, and the period it is held at has to
  // divide the displayed one evenly.
  function lenChoicesForDisplayLength(displayLength) {
    const display = Number(displayLength);
    if (!Number.isFinite(display) || display <= 0) return LEN_CHOICES.slice();
    const offered = LEN_CHOICES.filter((choice) => display % choice === 0);
    return offered.length ? offered : [display];
  }

  function storedLenUnavailableReason(displayLength) {
    const display = Number(displayLength);
    if (!Number.isFinite(display) || display <= 0) return "";
    return `This dataset is shown at ${display}, so it can only be stored at a period that divides ${display}.`;
  }

  // The ladder itself never changes: a length the control rules out keeps its
  // place in the list and is shown muted, so the menu carries the whole ladder
  // and nothing has to be said beside the control.
  function paintLenSelectOptions(selectId, offered, reason, fallback) {
    const select = document.getElementById(selectId);
    if (!select) return false;
    const available = new Set(offered.map((value) => String(value)));
    const shape = (value, unavailable) => `${value}:${unavailable ? "0" : "1"}`;
    const current = Array.from(select.options).map((option) => shape(option.value, lenOptionIsUnavailable(option)));
    const wanted = LEN_CHOICES.map((value) => shape(value, !available.has(String(value))));
    if (current.length === wanted.length && current.every((entry, i) => entry === wanted[i])) return false;
    const previous = select.value;
    select.innerHTML = "";
    for (const value of LEN_CHOICES) {
      const option = document.createElement("option");
      option.value = String(value);
      option.textContent = String(value);
      if (!available.has(String(value))) {
        option.disabled = true;
        option.dataset.unavailable = "1";
        option.title = reason;
      }
      select.appendChild(option);
    }
    // A value that has just been muted lands on the caller's fallback rather
    // than on an arbitrary neighbour.
    select.value = available.has(previous) ? previous : String(fallback);
    syncLenDropdownButtonLabel(selectId);
    renderLenDropdownOptions(selectId);
    return true;
  }

  // The finest length still offered is the stored period itself, which is where
  // the values live.
  function setLenSelectStoredLength(selectId, storedLength) {
    const offered = lenChoicesForStoredLength(storedLength);
    return paintLenSelectOptions(selectId, offered, lenUnavailableReason(storedLength), offered[offered.length - 1]);
  }

  // The coarsest period still offered is the displayed one, which is where a
  // store that has not been lowered sits.
  function setLenSelectDisplayLength(selectId, displayLength) {
    const offered = lenChoicesForDisplayLength(displayLength);
    return paintLenSelectOptions(selectId, offered, storedLenUnavailableReason(displayLength), offered[0]);
  }

  function fillLenDropdowns() {
    const o = document.getElementById("originLenSelect");
    const d = document.getElementById("devLenSelect");
    if (!o || !d) return;

    setLenSelectStoredLength("originLenSelect", 0);
    setLenSelectStoredLength("devLenSelect", 0);
    // The two `Stored at` controls carry the same ladder, narrowed to the
    // periods that divide the length beside them.
    setLenSelectDisplayLength("originStoredLenSelect", 12);
    setLenSelectDisplayLength("devStoredLenSelect", 12);

    // defaults
    o.value = "12";
    d.value = "12";
    refreshLenDropdowns();
  }

  async function loadProjectsDropdown() {
    const input = document.getElementById("projectSelect");
    const list = document.getElementById("projectDropdown");
    if (!input || !list) return { ok: false, projects: [] };

    try {
      runtime.allProjects = await loadProjectValidValueList();
    } catch (err) {
      console.error("Failed to load project names:", err);
      setStatus("Failed to load project names.");
      runtime.allProjects = [];
      renderProjectOptions(runtime.allProjects);
      showProjectDropdown(false);
      return { ok: false, projects: [] };
    }
    renderProjectOptions(runtime.allProjects);
    showProjectDropdown(false);

    // default values you requested
    const pathInput = document.getElementById("pathInput");
    const triInput = document.getElementById("triInput");
    if (!isDfmDataTabHost() && pathInput && !pathInput.value && !isInputDefaultBound(pathInput)) {
      pathInput.value = "PRNJ - PA\\PA\\NJ\\Direct Group\\COL";
    }
    if (!isDfmDataTabHost() && triInput && !triInput.value) triInput.value = "Net Loss--Incurred";

    return { ok: true, projects: runtime.allProjects.slice() };
  }

  function showDatasetLoadingPopup(message = "") {
    runtime.datasetRunController.showDatasetLoadingPopup(message);
  }

  function hideDatasetLoadingPopup() {
    runtime.datasetRunController.hideDatasetLoadingPopup();
  }

  function getTriInputs() {
    enforceDevLenRule();
    const project = getResolvedProjectValue();
    const path = getResolvedReservingClassValue();
    const tri = (document.getElementById("triInput")?.value || "").trim();
    const instanceName = getDatasetInstanceNameValue();
    const originLen = parseInt(document.getElementById("originLenSelect")?.value, 10);
    const devLen = parseInt(document.getElementById("devLenSelect")?.value, 10);
    const cumulative = !!document.getElementById("cumulativeChk")?.checked;
    const transposed = !!document.getElementById("transposedChk")?.checked;
    const calendar = document.querySelector('input[name="timeMode"][value="calendar"]')?.checked === true;

    return {
      project,
      path,
      tri,
      instanceName,
      cumulative,
      transposed,
      calendar,
      originLen: Number.isFinite(originLen) ? originLen : 12,
      devLen: Number.isFinite(devLen) ? devLen : 12,
    };
  }

  function resolveTriRequestInputs(rawInputs = {}) {
    const project = String(rawInputs?.project || "").trim();
    const path = normalizeReservingClassPath(rawInputs?.path || "");
    const tri = String(rawInputs?.tri || "").trim();
    const instanceName = String(rawInputs?.instanceName || rawInputs?.instance_name || "").trim();
    const cumulative = !!rawInputs?.cumulative;
    const calendar = !!rawInputs?.calendar;
    const originRaw = Number(rawInputs?.originLen);
    const devRaw = Number(rawInputs?.devLen);
    return {
      project,
      path,
      tri,
      instanceName,
      cumulative,
      calendar,
      originLen: Number.isFinite(originRaw) ? originRaw : 12,
      devLen: Number.isFinite(devRaw) ? devRaw : 12,
    };
  }

  function buildTriRequestPayload(rawInputs = {}) {
    const resolved = resolveTriRequestInputs(rawInputs);
    return {
      Path: resolved.path,
      TriangleName: resolved.tri,
      DatasetTypeName: resolved.tri,
      InstanceName: resolved.instanceName || resolved.tri,
      ProjectName: resolved.project,
      Cumulative: resolved.cumulative,
      Calendar: resolved.calendar,
      OriginLength: resolved.originLen,
      DevelopmentLength: resolved.devLen,
      LocalOnly: isDfmDataTabHost(),
      AllowDerived: true,
      WriteSidecar: false,
      ...(isTemporaryDatasetView ? { TemporarySessionId: temporaryDatasetSessionId } : {}),
    };
  }

  async function precheckArcRhoTriCsv(rawInputs = {}) {
    const resolved = resolveTriRequestInputs(rawInputs);
    if (!resolved.project || !resolved.path || !resolved.tri) {
      return { ok: false, hasExistingCsv: false, skipped: true, data: null };
    }
    try {
      const precheckResp = await fetch("/arcrho/tri/precheck", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(buildTriRequestPayload(resolved)),
      });
      if (!precheckResp.ok) {
        return { ok: false, hasExistingCsv: false, skipped: false, data: null };
      }
      const data = await precheckResp.json().catch(() => ({}));
      return {
        ok: true,
        hasExistingCsv: data?.need_request === false || data?.cache_exists === true,
        skipped: false,
        data,
      };
    } catch {
      return { ok: false, hasExistingCsv: false, skipped: false, data: null };
    }
  }

  function buildVecRequestPayload(rawInputs = {}) {
    const resolved = resolveTriRequestInputs(rawInputs);
    return {
      Path: resolved.path,
      VectorName: resolved.tri,
      DatasetTypeName: resolved.tri,
      InstanceName: resolved.instanceName || resolved.tri,
      ProjectName: resolved.project,
      PeriodLength: resolved.originLen,
      Cumulative: resolved.cumulative,
      Calendar: resolved.calendar,
      LocalOnly: isDfmDataTabHost(),
      AllowDerived: true,
      WriteSidecar: false,
      ...(isTemporaryDatasetView ? { TemporarySessionId: temporaryDatasetSessionId } : {}),
    };
  }

  async function precheckArcRhoVecCsv(rawInputs = {}) {
    const resolved = resolveTriRequestInputs(rawInputs);
    if (!resolved.project || !resolved.path || !resolved.tri) {
      return { ok: false, hasExistingCsv: false, skipped: true, data: null };
    }
    try {
      const precheckResp = await fetch("/arcrho/vec/precheck", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(buildVecRequestPayload(resolved)),
      });
      if (!precheckResp.ok) {
        return { ok: false, hasExistingCsv: false, skipped: false, data: null };
      }
      const data = await precheckResp.json().catch(() => ({}));
      return {
        ok: true,
        hasExistingCsv: data?.need_request === false || data?.cache_exists === true,
        skipped: false,
        data,
      };
    } catch {
      return { ok: false, hasExistingCsv: false, skipped: false, data: null };
    }
  }

  runtime.datasetDependencyGuard = createDatasetDependencyGuard({
    state,
    normalizeProjectText,
    getResolvedProjectValue,
    getTriInputs,
    precheckArcRhoTriCsv,
    precheckArcRhoVecCsv,
    setInputInvalid,
    clearInputInvalid,
    setStatus,
  });

  function getDatasetRunDataFormat(datasetTypeName = "") {
    const fromDatasetTypes = getDatasetTypeDataFormatByName(datasetTypeName || document.getElementById("triInput")?.value || "");
    if (fromDatasetTypes) return fromDatasetTypes;
    return runtime.currentDatasetSidecarDataFormat || state.model?.data_format || getProjectInstanceDraftDataFormat();
  }

  function getTriInputsForStorage() {
    enforceDevLenRule();
    const projectInput = document.getElementById("projectSelect");
    const pathInput = document.getElementById("pathInput");
    const tri = (document.getElementById("triInput")?.value || "").trim();
    const originLen = parseInt(document.getElementById("originLenSelect")?.value, 10);
    const devLen = parseInt(document.getElementById("devLenSelect")?.value, 10);
    const cumulative = !!document.getElementById("cumulativeChk")?.checked;
    const transposed = !!document.getElementById("transposedChk")?.checked;
    const calendar = document.querySelector('input[name="timeMode"][value="calendar"]')?.checked === true;

    return {
      project: getStoredInputValue(projectInput),
      path: getStoredInputValue(pathInput),
      tri,
      cumulative,
      transposed,
      calendar,
      decimalPlaces: getDatasetDecimalPlacesValue(),
      numberFormat: getDatasetSyncedNumberFormatValue(),
      originLen: Number.isFinite(originLen) ? originLen : 12,
      devLen: Number.isFinite(devLen) ? devLen : 12,
    };
  }

  function loadDataset() {
    return runtime.datasetRunController.loadDataset();
  }

  function savePatch() {
    return runtime.datasetRunController.savePatch();
  }

  function toggleBlanks() {
    return runtime.datasetRunController.toggleBlanks();
  }

  function getValidDevelopmentLengthForOrigin(origin, currentDev) {
    if (!Number.isFinite(origin) || origin <= 0) return "";
    const devSelect = document.getElementById("devLenSelect");
    const candidates = Array.from(devSelect?.options || [])
      .map((opt) => Number.parseInt(String(opt.value || opt.textContent || ""), 10))
      .filter((value) => Number.isFinite(value) && value > 0 && value <= origin && origin % value === 0)
      .sort((a, b) => b - a);
    if (!candidates.length) return "";
    if (Number.isFinite(currentDev) && candidates.includes(currentDev)) return String(currentDev);
    return String(candidates[0]);
  }

  function getValidOriginLengthForDevelopment(dev, currentOrigin) {
    if (!Number.isFinite(dev) || dev <= 0) return "";
    const originSelect = document.getElementById("originLenSelect");
    const candidates = Array.from(originSelect?.options || [])
      .map((opt) => Number.parseInt(String(opt.value || opt.textContent || ""), 10))
      .filter((value) => Number.isFinite(value) && value > 0 && value >= dev && value % dev === 0)
      .sort((a, b) => a - b);
    if (!candidates.length) return "";
    if (Number.isFinite(currentOrigin) && candidates.includes(currentOrigin)) return String(currentOrigin);
    return String(candidates[0]);
  }

  function enforceDevLenRule(options = {}) {
    const source = String(options?.source || "auto");
    const o = document.getElementById("originLenSelect");
    const d = document.getElementById("devLenSelect");
    if (!o || !d) return false;

    let origin = parseInt(o.value, 10);
    let dev = parseInt(d.value, 10);

    const ok =
      Number.isFinite(origin) &&
      Number.isFinite(dev) &&
      dev <= origin &&
      origin % dev === 0;

    let changed = false;
    if (!ok) {
      if (source === "dev" || (source !== "origin" && dev > origin)) {
        const nextOrigin = getValidOriginLengthForDevelopment(dev, origin);
        if (nextOrigin) {
          changed = setLenSelectValue("originLenSelect", nextOrigin) || changed;
          origin = parseInt(o.value, 10);
        }
      } else {
        const nextDev = getValidDevelopmentLengthForOrigin(origin, dev);
        if (nextDev) {
          changed = setLenSelectValue("devLenSelect", nextDev) || changed;
          dev = parseInt(d.value, 10);
        }
      }
      const validAfterFirstPass =
        Number.isFinite(origin) &&
        Number.isFinite(dev) &&
        dev <= origin &&
        origin % dev === 0;
      if (!validAfterFirstPass) {
        const nextDev = getValidDevelopmentLengthForOrigin(origin, dev);
        if (nextDev) {
          changed = setLenSelectValue("devLenSelect", nextDev) || changed;
        }
      }
    }
    refreshLenDropdowns();
    return changed;
  }

  // -----------------------------
  // Headers (year + dev) via GetDataset-like flow
  // key = ProjectName + OriginLength
  // -----------------------------

  function getCurrentOriginLength() {
    return runtime.datasetHeadersService.getCurrentOriginLength();
  }

  function getCurrentDevLength() {
    return runtime.datasetHeadersService.getCurrentDevLength();
  }

  function ensureHeadersForProject(project, options = {}) {
    return runtime.datasetHeadersService.ensureHeadersForProject(project, options);
  }

  function ensureDevHeadersForProject(project, options = {}) {
    return runtime.datasetHeadersService.ensureDevHeadersForProject(project, options);
  }


  Object.assign(runtime, {
    initializeDatasetId,
    readDatasetInputsFromQueryParams,
    normalizeDraftDataFormat,
    getProjectInstanceDraftDataFormat,
    numericDevelopmentLabels,
    resolveDevelopmentLabels,
    buildProjectInstanceDraftMask,
    buildProjectInstanceDraftModel,
    initializeProjectInstanceDraftModel,
    refreshProjectInstanceDraftModel,
    getLenDropdownIds,
    getLenDropdownElements,
    getLenDropdownActiveIndex,
    setLenDropdownActiveIndex,
    syncLenDropdownButtonLabel,
    renderLenDropdownOptions,
    refreshLenDropdowns,
    showLenDropdown,
    closeAllLenDropdowns,
    isLenSelectLocked,
    setLenSelectLock,
    setLenSelectValue,
    lenChoicesForStoredLength,
    lenChoicesForDisplayLength,
    lenOptionIsUnavailable,
    lenUnavailableReason,
    storedLenUnavailableReason,
    setLenSelectStoredLength,
    setLenSelectDisplayLength,
    chooseActiveLenDropdownOption,
    moveLenDropdownActiveOption,
    cycleLenSelect,
    wireLenDropdown,
    wireLenDropdowns,
    fillLenDropdowns,
    loadProjectsDropdown,
    showDatasetLoadingPopup,
    hideDatasetLoadingPopup,
    getTriInputs,
    resolveTriRequestInputs,
    buildTriRequestPayload,
    precheckArcRhoTriCsv,
    buildVecRequestPayload,
    precheckArcRhoVecCsv,
    getDatasetRunDataFormat,
    getTriInputsForStorage,
    loadDataset,
    savePatch,
    toggleBlanks,
    getValidDevelopmentLengthForOrigin,
    getValidOriginLengthForDevelopment,
    enforceDevLenRule,
    getCurrentOriginLength,
    getCurrentDevLength,
    ensureHeadersForProject,
    ensureDevHeadersForProject,
  });
}
