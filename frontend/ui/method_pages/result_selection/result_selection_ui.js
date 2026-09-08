(function () {
  const parts = window.ResultSelectionParts || (window.ResultSelectionParts = {});

  parts.installUi = function installUi(ctx) {
    with (ctx) {
      function numberOrNull(value) {
        if (value === null || value === undefined || value === "") return null;
        const n = Number(value);
        return Number.isFinite(n) ? n : null;
      }

      function positiveInt(value, fallback = DEFAULT_ORIGIN_LENGTH) {
        const n = Number.parseInt(String(value ?? ""), 10);
        return Number.isFinite(n) && n > 0 ? n : fallback;
      }

      function validOriginLength(value, fallback = DEFAULT_ORIGIN_LENGTH) {
        const n = positiveInt(value, fallback);
        return VALID_ORIGIN_LENGTHS.includes(n) ? n : fallback;
      }

      function validSourceOriginLength(value) {
        const n = validOriginLength(value, 0);
        return VALID_ORIGIN_LENGTHS.includes(n) ? n : null;
      }

      function isEngineSource(source) {
        return norm(source?.sourceKind || source?.source_kind) === "engine";
      }

      function isVectorSource(source) {
        return norm(source?.dataFormat || source?.data_format) === "vector";
      }

      function datasetTypeCategoryForName(name) {
        const key = norm(name);
        if (!key) return "";
        const item = datasetTypeItems.find((entry) => norm(entry?.name) === key);
        return text(item?.category);
      }

      function getOutputCategory() {
        const fromOutputType = datasetTypeCategoryForName(els.outputTypeInput?.value);
        return fromOutputType || state.outputCategory || "";
      }

      function syncOutputCategory() {
        state.outputCategory = getOutputCategory();
        return state.outputCategory;
      }

      function sourceCategory(source) {
        return text(source?.category || source?.dataset_category || datasetTypeCategoryForName(source?.datasetType || source?.dataset_type || source?.name));
      }

      function matchesOutputCategory(source) {
        const outputCategory = getOutputCategory();
        if (!outputCategory) return true;
        return norm(sourceCategory(source)) === norm(outputCategory);
      }

      function nonNegativeInt(value, fallback = 0) {
        const n = Number.parseInt(String(value ?? ""), 10);
        return Number.isFinite(n) && n >= 0 ? n : fallback;
      }

      function statisticDecimalPlacesValue(value, fallback = 1) {
        return Math.max(0, Math.min(8, nonNegativeInt(value, fallback)));
      }

      function syncStatisticDecimalInputs(source = "details") {
        const sourceEl = source === "method" && els.methodStatisticDecimalsInput
          ? els.methodStatisticDecimalsInput
          : els.statisticDecimalsInput;
        const fallback = statisticDecimalPlacesValue(els.statisticDecimalsInput?.value, 1);
        const next = String(statisticDecimalPlacesValue(sourceEl?.value, fallback));
        if (els.statisticDecimalsInput && els.statisticDecimalsInput.value !== next) {
          els.statisticDecimalsInput.value = next;
        }
        if (els.methodStatisticDecimalsInput && els.methodStatisticDecimalsInput.value !== next) {
          els.methodStatisticDecimalsInput.value = next;
        }
        return next;
      }

      function getHostApi() {
        if (window.ADAHost) return window.ADAHost;
        try {
          let w = window.parent;
          while (w && w !== window) {
            if (w.ADAHost) return w.ADAHost;
            if (w === w.parent) break;
            w = w.parent;
          }
        } catch {}
        return null;
      }

      function postStatus(message, tone = "") {
        try {
          window.parent?.postMessage({ type: "arcrho:status", text: String(message || ""), ...(tone ? { tone } : {}) }, "*");
        } catch {}
      }

      function postDirty(dirty, force = false) {
        const next = !!dirty;
        if (!force && isDirty === next) return;
        isDirty = next;
        updateTabbedPageSaveControls({
          saveButton: els.saveBtn,
          cancelButton: els.cancelBtn,
          dirty: next,
        });
        if (state.loadBlocked && els.saveBtn) els.saveBtn.disabled = true;
        try {
          window.parent?.postMessage({ type: "arcrho:dataset-dirty", inst, dirty: next }, "*");
        } catch {}
      }

      function markDirty() {
        if (programmatic) return;
        if (state.loadBlocked) {
          postStatus("Reload the Result Selection successfully before editing or saving.", "error");
          if (els.saveBtn) els.saveBtn.disabled = true;
          return;
        }
        postDirty(true);
        postResultSelectionDependencyPreview();
      }

      function syncLoadBlockedControls() {
        const blocked = !!state.loadBlocked;
        const controls = [
          els.nameInput,
          els.outputTypeInput,
          els.outputTypeBtn,
          els.originLengthInput,
          els.originLengthButton,
          ...(els.ratioBasisInputs || []),
          els.ratioBasisAddButton,
          els.showRatiosPctInput,
          els.statisticDecimalsInput,
          els.methodStatisticDecimalsInput,
          els.methodStatisticDecimalsUp,
          els.methodStatisticDecimalsDown,
          els.showWeightsInput,
          els.weightDisplayButton,
          els.activeRatioBasisButton,
          els.saveBtn,
          els.notesInput,
          ...Array.from(els.methodGrid?.querySelectorAll?.("input, button") || []),
          ...Array.from(els.resultsGrid?.querySelectorAll?.("input, button") || []),
        ];
        for (const control of controls) {
          if (control) control.disabled = blocked;
        }
      }

      function sourceMessageNames(message = {}) {
        const names = Array.isArray(message.names) ? message.names : [message.datasetName, message.datasetTypeName, message.name];
        return new Set(names.map((value) => norm(value)).filter(Boolean));
      }

      function sourceMessageMatchesContext(message = {}) {
        const project = text(message.project);
        const reservingClass = text(message.reservingClass || message.reserving_class);
        if (project && norm(project) !== norm(state.project)) return false;
        if (reservingClass && norm(reservingClass) !== norm(state.reservingClass)) return false;
        return true;
      }

      function reportMatchesCurrentContext(report) {
        const contexts = resultSelectionUpdateContexts(report);
        if (!contexts.length) return false;
        return contexts.some((context) => {
          const project = text(context.project);
          const reservingClass = text(context.reservingClass);
          if (project && norm(project) !== norm(state.project)) return false;
          if (reservingClass && norm(reservingClass) !== norm(state.reservingClass)) return false;
          return !!(project || reservingClass);
        });
      }

      function recordPersistedMethodDependencies(payload = {}) {
        const details = payload?.details_tab || {};
        const method = payload?.method_tab || {};
        const names = [
          ...(Array.isArray(method.loaded_datasets)
            ? method.loaded_datasets.map((item) => item?.name)
            : []),
          ...(Array.isArray(details.ratio_basis_datasets)
            ? details.ratio_basis_datasets
            : []),
        ];
        state.persistedDependencyNames = new Set(names.map(norm).filter(Boolean));
      }

      function sourceMessageMatchesSource(message, source) {
        if (!source) return false;
        const names = sourceMessageNames(message);
        return [source.name, source.datasetType, source.dataset_type]
          .some((value) => names.has(norm(value)));
      }

      function sourceMessageMatchesRatioBasis(message) {
        const names = sourceMessageNames(message);
        return getRatioBasisNames().some((value) => names.has(norm(value)));
      }

      function dependencyPreviewKey(message = {}) {
        const messageInst = text(message.inst);
        const names = Array.from(sourceMessageNames(message)).sort();
        return `${messageInst}::${names.join("|")}`;
      }

      function removeDependencyPreview(message = {}) {
        const key = dependencyPreviewKey(message);
        let removed = state.dependencyPreviews.delete(key);
        if (!removed && text(message.inst)) {
          const clearedNames = sourceMessageNames(message);
          for (const [candidateKey, candidate] of state.dependencyPreviews) {
            if (text(candidate?.inst) !== text(message.inst)) continue;
            const candidateNames = sourceMessageNames(candidate);
            if (!Array.from(candidateNames).some((name) => clearedNames.has(name))) continue;
            state.dependencyPreviews.delete(candidateKey);
            removed = true;
          }
        }
        state.hasDependencyPreview = state.dependencyPreviews.size > 0;
        return removed;
      }

      function noteDependencyEvent() {
        state.dependencyEventSeq = Number(state.dependencyEventSeq || 0) + 1;
        return state.dependencyEventSeq;
      }

      function calculatedUpdateAffectsCurrentResultSelection(message = {}) {
        if (!sourceMessageMatchesContext(message)) return false;
        if (!reportMatchesCurrentContext(message.report)) return false;
        const currentName = norm(getDetails().name);
        if (!currentName) return false;
        return Array.from(resultSelectionUpdateNames(message.report)).some(
          (name) => norm(name) === currentName,
        );
      }

      function normalizePreviewValues(values) {
        return Array.isArray(values) ? values.map(numberOrNull) : [];
      }

      function buildResultSelectionDependencySourceMessage(type, reason = "") {
        const details = getDetails();
        const payload = {
          type,
          inst,
          project: state.project,
          reservingClass: state.reservingClass,
          datasetName: details.name,
          datasetTypeName: details.outputType,
          names: [details.name, details.outputType].filter(Boolean),
          methodType: "Result Selection",
          sourceKind: "result_selection",
          dataFormat: "Vector",
          reason,
        };
        if (type === "arcrho:dependency-source-preview") {
          // Live previews remain available while a newly selected Ratio Basis is
          // loading. Strict completeness validation still runs on save.
          payload.values = selectedUltimateVector();
          payload.originLabels = state.originLabels.map(String);
        }
        if (type === "arcrho:dependency-source-cleared") {
          // Set only by a save that enqueued an Engine propagation job;
          // Project Instance defers the downstream preview clear until the
          // job reaches a terminal status.
          payload.propagationJobId = String(state.pendingPropagationJobId || "").trim();
          state.pendingPropagationJobId = "";
        }
        return payload;
      }

      function postResultSelectionDependencySourceMessage(type, reason = "") {
        const message = buildResultSelectionDependencySourceMessage(type, reason);
        if (!message.names.length) return;
        try {
          window.parent?.postMessage(message, "*");
        } catch {}
      }

      function postResultSelectionDependencyPreview() {
        postResultSelectionDependencySourceMessage("arcrho:dependency-source-preview", "dirty");
        state.dependencyPreviewPublished = true;
      }

      function clearResultSelectionDependencyPreview(reason = "") {
        if (!state.dependencyPreviewPublished) return;
        postResultSelectionDependencySourceMessage("arcrho:dependency-source-cleared", reason || "clean");
        state.dependencyPreviewPublished = false;
      }

      async function reloadSourcesMatchingMessages(messages = []) {
        const scopedMessages = messages.filter(sourceMessageMatchesContext);
        if (!scopedMessages.length) return false;
        const matches = state.sources
          .map((source, index) => ({ source, index }))
          .filter(({ source }) => scopedMessages.some((message) => sourceMessageMatchesSource(message, source)));
        if (!matches.length) return false;
        const reloaded = await mapWithConcurrency(
          matches,
          SOURCE_LOAD_CONCURRENCY,
          async ({ source, index }) => {
            const record = cachedRows.find((row) => norm(row.name) === norm(source.name)) || null;
            const built = await buildSourceFromRecord(
              record || { name: source.name },
              { ...source, values: [] },
            );
            if (!built || built.unavailable) {
              throw new Error(`The cleared local source '${source.name}' could not be reloaded from disk.`);
            }
            return { source, index, built };
          },
        );
        for (const { source, index, built } of reloaded) {
          const currentIndex = state.sources.indexOf(source);
          const targetIndex = currentIndex >= 0
            ? currentIndex
            : state.sources.findIndex((item) => norm(item?.name) === norm(source.name));
          if (targetIndex >= 0) state.sources[targetIndex] = built;
          else if (index < state.sources.length) state.sources[index] = built;
        }
        renderMethodGrid();
        return true;
      }

      async function reloadSourcesMatchingMessage(message) {
        return reloadSourcesMatchingMessages([message]);
      }

      async function refreshPersistedValuesFromDisk(reason = "dependency update") {
        const refreshSeq = Number(state.persistedRefreshSeq || 0) + 1;
        state.persistedRefreshSeq = refreshSeq;
        const result = await fetchPersistedResultSelection(true);
        if (state.persistedRefreshSeq !== refreshSeq) return false;
        if (!result.method_exists || !result.method) {
          throw new Error("Saved Result Selection method is missing.");
        }
        recordPersistedMethodDependencies(result.method);
        if (!isDirty) {
          applyOutputSidecar(result.sidecar);
          await applyPayload(result.method);
          state.methodRevision = String(result.method_revision || "");
          state.loadBlocked = false;
          syncLoadBlockedControls();
          markClean();
        } else {
          const persistedOriginLength = validOriginLength(result.method?.details_tab?.origin_length, 0);
          if (persistedOriginLength && persistedOriginLength !== getDetails().originLength) {
            postStatus(
              `Result Selection dependency refresh was deferred because the local Origin Length is unsaved (${getDetails().originLength}).`,
              "warn",
            );
            return false;
          }
          const persistedSources = new Map(
            (Array.isArray(result.method?.method_tab?.loaded_datasets)
              ? result.method.method_tab.loaded_datasets
              : [])
              .map((item) => [norm(item?.name), buildSourceFromPersisted(item)])
              .filter(([, item]) => !!item),
          );
          state.sources = state.sources.map((source) => {
            const persisted = persistedSources.get(norm(source.name));
            if (!persisted) return source;
            return {
              ...source,
              datasetType: persisted.datasetType || source.datasetType,
              dataFormat: persisted.dataFormat || source.dataFormat,
              originLength: persisted.originLength || source.originLength,
              methodType: persisted.methodType || source.methodType,
              category: persisted.category || source.category,
              sourceKind: persisted.sourceKind || source.sourceKind,
              values: persisted.values.slice(),
            };
          });
          const persistedBasisSets = normalizeRatioBasisValueSets(
            result.method?.method_tab?.ratio_basis_values,
            result.method?.details_tab?.ratio_basis_datasets,
          );
          const persistedBasisByName = new Map(
            persistedBasisSets.map((item) => [norm(item.name), item]),
          );
          state.ratioBasisValueSets = normalizeRatioBasisValueSets(
            getRatioBasisNames().map((name) => (
              persistedBasisByName.get(norm(name))
              || state.ratioBasisValueSets.find((item) => norm(item?.name) === norm(name))
              || { name, values: [] }
            )),
            getRatioBasisNames(),
          );
          state.ratioBasisValues = ratioBasisValuesForName(
            state.ratioBasisValueSets,
            getActiveRatioBasisName(),
          );
          state.methodRevision = String(result.method_revision || "");
          state.needsReview = Number(result.sidecar?.status) === 2;
          renderMethodGrid();
        }
        reapplyActiveDependencyPreviews();
        postStatus(
          state.needsReview
            ? `Result Selection still needs review after ${reason}.`
            : `Result Selection values refreshed after ${reason}.`,
          state.needsReview ? "warn" : "",
        );
        return true;
      }

      function queueDependencyClear(message, options = {}) {
        const key = dependencyPreviewKey(message);
        const existing = state.pendingDependencyClearMessages.find(
          (item) => dependencyPreviewKey(item.message) === key,
        );
        if (existing) {
          existing.reloadLocalSource ||= !!options.reloadLocalSource;
          existing.reloadLocalBasis ||= !!options.reloadLocalBasis;
          existing.message = { ...existing.message, ...message };
          existing.generation = Number(state.dependencyEventSeq || 0);
          return;
        }
        state.pendingDependencyClearMessages.push({
          message: { ...message },
          reloadLocalSource: !!options.reloadLocalSource,
          reloadLocalBasis: !!options.reloadLocalBasis,
          generation: Number(state.dependencyEventSeq || 0),
        });
      }

      async function restoreLocalOnlyDependencies(clears = []) {
        const sourceMessages = clears
          .filter((item) => item.reloadLocalSource)
          .map((item) => item.message);
        if (sourceMessages.length) {
          const reloaded = await reloadSourcesMatchingMessages(sourceMessages);
          if (!reloaded) throw new Error("A cleared local source could not be reloaded from disk.");
        }

        const basisNames = new Set();
        for (const item of clears) {
          if (!item.reloadLocalBasis) continue;
          for (const name of sourceMessageNames(item.message)) basisNames.add(name);
        }
        if (!basisNames.size) return;
        state.ratioBasisValueSets = state.ratioBasisValueSets.filter(
          (item) => !basisNames.has(norm(item?.name)),
        );
        state.ratioBasisValues = ratioBasisValuesForName(
          state.ratioBasisValueSets,
          getActiveRatioBasisName(),
        );
        const refreshed = await refreshMissingRatioBasisValues();
        if (!refreshed) throw new Error("A cleared local Ratio Basis could not be reloaded from disk.");
      }

      async function flushPersistedValuesRefresh() {
        if (state.initialLoadPending || state.persistedMutationInFlight) return false;
        if (state.persistedRefreshTimer != null) {
          window.clearTimeout(state.persistedRefreshTimer);
          state.persistedRefreshTimer = null;
        }
        if (state.dependencyRefreshPromise) return state.dependencyRefreshPromise;

        let retryQueuedRefresh = false;
        const refreshPromise = (async () => {
          while (state.persistedRefreshReason || state.pendingDependencyClearMessages.length) {
            const refreshGeneration = Number(state.dependencyEventSeq || 0);
            const clears = state.pendingDependencyClearMessages
              .filter((item) => Number(item.generation || 0) <= refreshGeneration)
              .map((item) => ({ ...item, message: { ...item.message } }));
            const refreshReason = state.persistedRefreshReason
              || text(clears.at(-1)?.message?.reason)
              || "dependency update";
            state.persistedRefreshReason = "";
            try {
              await restoreLocalOnlyDependencies(clears);
              const refreshed = await refreshPersistedValuesFromDisk(refreshReason);
              if (!refreshed) {
                throw new Error("The latest persisted Result Selection refresh was superseded or deferred.");
              }
            } catch (err) {
              state.persistedRefreshReason ||= refreshReason;
              retryQueuedRefresh = Number(state.dependencyEventSeq || 0) > refreshGeneration
                || state.pendingDependencyClearMessages.some(
                  (item) => Number(item.generation || 0) > refreshGeneration,
                );
              throw err;
            }
            state.pendingDependencyClearMessages = state.pendingDependencyClearMessages.filter(
              (item) => Number(item.generation || 0) > refreshGeneration,
            );
          }
          return true;
        })();
        state.dependencyRefreshPromise = refreshPromise;
        try {
          const refreshed = await refreshPromise;
          if (state.dependencyRefreshPromise === refreshPromise) {
            state.dependencyRestorePending = false;
            state.dependencyRestoreError = "";
          }
          return refreshed;
        } catch (err) {
          if (state.dependencyRefreshPromise === refreshPromise) {
            state.dependencyRestorePending = true;
            state.dependencyRestoreError = retryQueuedRefresh
              ? ""
              : String(err?.message || err || "Dependency restore failed.");
          }
          throw err;
        } finally {
          if (state.dependencyRefreshPromise === refreshPromise) {
            state.dependencyRefreshPromise = null;
            if (retryQueuedRefresh) armPersistedValuesRefresh();
          }
        }
      }

      function armPersistedValuesRefresh() {
        if (state.initialLoadPending || state.persistedMutationInFlight || state.dependencyRefreshPromise) return;
        if (state.persistedRefreshTimer != null) window.clearTimeout(state.persistedRefreshTimer);
        state.persistedRefreshTimer = window.setTimeout(() => {
          state.persistedRefreshTimer = null;
          flushPersistedValuesRefresh()
            .catch((err) => postStatus(`Result Selection refresh failed: ${err?.message || err}`, "error"));
        }, 25);
      }

      function schedulePersistedValuesRefresh(reason = "dependency update") {
        state.persistedRefreshReason = text(reason) || "dependency update";
        state.dependencyRestorePending = true;
        state.dependencyRestoreError = "";
        armPersistedValuesRefresh();
      }

      function scheduleDependencyClearRestore(message, options = {}) {
        queueDependencyClear(message, options);
        schedulePersistedValuesRefresh(message.reason || "dependency update");
      }

      function resumePersistedValuesRefresh() {
        if (state.persistedRefreshReason || state.pendingDependencyClearMessages.length) {
          armPersistedValuesRefresh();
        }
      }

      function applyDependencySourcePreview(message, options = {}) {
        if (!sourceMessageMatchesContext(message)) return false;
        const values = normalizePreviewValues(message.values);
        if (!values.length) return false;
        let changed = false;
        for (const source of state.sources) {
          if (!sourceMessageMatchesSource(message, source)) continue;
          source.values = values.slice();
          source.dataFormat = text(message.dataFormat || message.data_format || source.dataFormat || "Vector");
          source.sourceKind = text(message.sourceKind || message.source_kind || source.sourceKind);
          source.methodType = text(message.methodType || message.method_type || source.methodType);
          source.originLength = validSourceOriginLength(message.originLength || message.origin_length) || source.originLength;
          if (!Array.isArray(source.weights)) source.weights = [];
          while (source.weights.length < source.values.length) source.weights.push(0);
          changed = true;
        }
        if (changed) {
          if (options.store !== false) {
            state.dependencyPreviews.set(dependencyPreviewKey(message), { ...message });
          }
          state.hasDependencyPreview = state.dependencyPreviews.size > 0;
          renderMethodGrid();
          if (options.publish !== false) postResultSelectionDependencyPreview();
        }
        return changed;
      }

      function reapplyActiveDependencyPreviews() {
        for (const preview of state.dependencyPreviews.values()) {
          applyDependencySourcePreview(preview, { store: false, publish: false });
        }
        state.hasDependencyPreview = state.dependencyPreviews.size > 0;
        if (state.hasDependencyPreview || isDirty) postResultSelectionDependencyPreview();
        return state.hasDependencyPreview;
      }

      function withProgrammatic(fn) {
        programmatic = true;
        try {
          return fn();
        } finally {
          programmatic = false;
        }
      }

      function dropdownOptions(menu) {
        return Array.from(menu?.querySelectorAll?.(".rsDropdownOption") || []);
      }

      function getDropdownValue(menu) {
        const selected = dropdownOptions(menu).find((option) => option.getAttribute("aria-selected") === "true");
        return text(selected?.dataset?.value);
      }

      function setDropdownOpen(dropdown, button, open) {
        if (!dropdown || !button || button.disabled) return;
        dropdown.classList.toggle("open", !!open);
        button.setAttribute("aria-expanded", open ? "true" : "false");
      }

      function closeDropdown(dropdown, button) {
        if (!dropdown || !button) return;
        dropdown.classList.remove("open");
        button.setAttribute("aria-expanded", "false");
      }

      function closeAllDropdowns(except = null) {
        const pairs = [
          [els.weightDisplayDropdown, els.weightDisplayButton],
          [els.originLengthDropdown, els.originLengthButton],
          [els.activeRatioBasisDropdown, els.activeRatioBasisButton],
        ];
        for (const [dropdown, button] of pairs) {
          if (!dropdown || dropdown === except) continue;
          closeDropdown(dropdown, button);
        }
      }

      function makeDropdownOption(value, label, selected = false) {
        const option = document.createElement("button");
        option.className = "rsDropdownOption";
        option.type = "button";
        option.setAttribute("role", "option");
        option.dataset.value = text(value);
        option.textContent = text(label);
        option.setAttribute("aria-selected", selected ? "true" : "false");
        return option;
      }

      function syncDropdownValue(menu, labelEl, value, fallbackLabel = "") {
        const options = dropdownOptions(menu);
        const wanted = text(value);
        let selected = null;
        for (const option of options) {
          const isSelected = text(option.dataset.value) === wanted;
          option.setAttribute("aria-selected", isSelected ? "true" : "false");
          if (isSelected) selected = option;
        }
        if (!selected && options.length) {
          selected = options[0];
          selected.setAttribute("aria-selected", "true");
        }
        if (labelEl) labelEl.textContent = selected?.textContent || fallbackLabel;
        return text(selected?.dataset?.value);
      }

      function wireDropdown(dropdown, button, menu, onSelect) {
        if (!dropdown || !button || !menu) return;
        button.addEventListener("click", (event) => {
          event.preventDefault();
          const nextOpen = !dropdown.classList.contains("open");
          closeAllDropdowns(dropdown);
          setDropdownOpen(dropdown, button, nextOpen);
          if (nextOpen) {
            const selected = dropdownOptions(menu).find((option) => option.getAttribute("aria-selected") === "true");
            selected?.focus?.({ preventScroll: true });
          }
        });
        menu.addEventListener("click", (event) => {
          const option = event.target?.closest?.(".rsDropdownOption");
          if (!option) return;
          event.preventDefault();
          const value = text(option.dataset.value);
          onSelect?.(value);
          closeDropdown(dropdown, button);
          button.focus?.({ preventScroll: true });
        });
        dropdown.addEventListener("keydown", (event) => {
          const key = event.key;
          if (key === "Escape") {
            event.preventDefault();
            closeDropdown(dropdown, button);
            button.focus?.({ preventScroll: true });
            return;
          }
          if (key !== "ArrowDown" && key !== "ArrowUp" && key !== "Enter" && key !== " ") return;
          const options = dropdownOptions(menu);
          if (!options.length) return;
          if (!dropdown.classList.contains("open")) {
            event.preventDefault();
            closeAllDropdowns(dropdown);
            setDropdownOpen(dropdown, button, true);
            const selected = options.find((option) => option.getAttribute("aria-selected") === "true") || options[0];
            selected.focus?.({ preventScroll: true });
            return;
          }
          const activeIndex = Math.max(0, options.indexOf(document.activeElement));
          if (key === "ArrowDown" || key === "ArrowUp") {
            event.preventDefault();
            const delta = key === "ArrowDown" ? 1 : -1;
            const nextIndex = (activeIndex + delta + options.length) % options.length;
            options[nextIndex]?.focus?.({ preventScroll: true });
            return;
          }
          if (document.activeElement?.classList?.contains("rsDropdownOption")) {
            event.preventDefault();
            document.activeElement.click();
          }
        });
      }

      function uniqueRatioBasisNames(names) {
        const out = [];
        const seen = new Set();
        for (const value of Array.isArray(names) ? names : []) {
          const name = text(value);
          const key = norm(name);
          if (!name || !key || seen.has(key)) continue;
          seen.add(key);
          out.push(name);
          if (out.length >= MAX_RATIO_BASIS_COUNT) break;
        }
        return out;
      }

      function getRatioBasisNames() {
        return uniqueRatioBasisNames((els.ratioBasisInputs || []).map((input) => input?.value));
      }

      function setRatioBasisNames(names) {
        const normalized = uniqueRatioBasisNames(names);
        (els.ratioBasisInputs || []).forEach((input, index) => {
          if (input) input.value = normalized[index] || "";
        });
        return normalized;
      }

      function matchRatioBasisName(value, names = getRatioBasisNames()) {
        const key = norm(value);
        if (!key) return "";
        return names.find((name) => norm(name) === key) || "";
      }

      function normalizeRatioBasisDetails(details = {}) {
        const fromList = Array.isArray(details.ratio_basis_datasets)
          ? details.ratio_basis_datasets
          : [];
        const fallback = text(details.ratio_basis_dataset || details.ratio_basis);
        const names = uniqueRatioBasisNames(fromList.length ? fromList : [fallback]);
        const active = text(details.active_ratio_basis_dataset || fallback);
        if (active && !matchRatioBasisName(active, names) && names.length < MAX_RATIO_BASIS_COUNT) {
          names.push(active);
        }
        return {
          names: uniqueRatioBasisNames(names),
          active: matchRatioBasisName(active, names) || names[0] || "",
        };
      }

      function renderRatioBasisPills(names = getRatioBasisNames()) {
        const list = els.ratioBasisList;
        if (!list) return;
        list.replaceChildren();

        names.forEach((name, index) => {
          const token = document.createElement("span");
          token.className = "rsRatioBasisToken";
          token.setAttribute("role", "listitem");
          token.setAttribute("draggable", "true");
          token.dataset.ratioBasisIndex = String(index);

          const openButton = document.createElement("button");
          openButton.className = "rsRatioBasisOpen";
          openButton.type = "button";
          openButton.dataset.ratioBasisOpenIndex = String(index);
          openButton.setAttribute("aria-label", `Open dataset ${name}`);

          const label = document.createElement("span");
          label.className = "rsRatioBasisTokenLabel";
          label.textContent = name;
          openButton.appendChild(label);
          token.appendChild(openButton);

          const removeButton = document.createElement("button");
          removeButton.className = "rsRatioBasisRemove";
          removeButton.type = "button";
          removeButton.dataset.ratioBasisRemoveIndex = String(index);
          removeButton.setAttribute("aria-label", `Remove dataset ${name}`);
          removeButton.innerHTML = "<svg viewBox=\"0 0 24 24\" aria-hidden=\"true\"><path d=\"M6 6l12 12M18 6L6 18\"></path></svg>";
          token.appendChild(removeButton);
          list.appendChild(token);
        });

        if (els.ratioBasisAddButton) {
          const atLimit = names.length >= MAX_RATIO_BASIS_COUNT;
          els.ratioBasisAddButton.disabled = atLimit;
          els.ratioBasisAddButton.hidden = atLimit;
        }
      }

      function closeRatioBasisContextMenu() {
        if (!els.ratioBasisContextMenu) return;
        els.ratioBasisContextMenu.classList.remove("open");
        els.ratioBasisContextMenu.setAttribute("aria-hidden", "true");
        delete els.ratioBasisContextMenu.dataset.ratioBasisIndex;
      }

      function openRatioBasisContextMenu(event, index) {
        const names = getRatioBasisNames();
        if (!els.ratioBasisContextMenu || !names[index]) return;
        event.preventDefault();
        event.stopPropagation();
        closeCellContextMenu();
        closeSourceContextMenu();
        els.ratioBasisContextMenu.dataset.ratioBasisIndex = String(index);
        els.ratioBasisContextMenu.classList.add("open");
        els.ratioBasisContextMenu.setAttribute("aria-hidden", "false");
        positionContextMenu(els.ratioBasisContextMenu, event.clientX, event.clientY);
      }

      function resetRatioBasisDragState() {
        state.ratioBasisDragIndex = null;
        els.ratioBasisPicker?.classList.remove("rsRatioBasisDragActive", "rsRatioBasisDragOutside");
        els.ratioBasisList?.querySelector(".rsRatioBasisDragging")?.classList.remove("rsRatioBasisDragging");
      }

      async function openRatioBasisDataset(index) {
        const name = getRatioBasisNames()[index];
        if (!name) return;
        await ensureDatasetCatalogLoaded();
        const record = cachedRows.find((row) => norm(row?.name) === norm(name)) || null;
        const requestId = `rs_open_ratio_basis_${Date.now()}_${Math.random().toString(36).slice(2)}`;
        const onMessage = (event) => {
          const msg = event.data || {};
          if (msg.type !== "arcrho:automation-command-result" || msg.requestId !== requestId) return;
          window.removeEventListener("message", onMessage);
          if (msg.ok === false) postStatus(`Open ratio basis dataset failed: ${msg.error || "Unknown error."}`, "error");
        };
        window.addEventListener("message", onMessage);
        window.setTimeout(() => window.removeEventListener("message", onMessage), 10000);
        try {
          window.parent?.postMessage({
            type: "arcrho:automation-open-dataset",
            requestId,
            args: {
              datasetName: name,
              datasetTypeName: text(record?.datasetTypeName || record?.datasetType || name),
              methodType: text(record?.methodType),
              readOnly: !!record?.readOnly,
            },
          }, "*");
        } catch (err) {
          window.removeEventListener("message", onMessage);
          postStatus(`Open ratio basis dataset failed: ${err?.message || err}`, "error");
        }
      }

      function removeRatioBasisAt(index) {
        const names = getRatioBasisNames();
        if (!Number.isInteger(index) || index < 0 || index >= names.length) return;
        const previousActive = state.activeRatioBasisName;
        names.splice(index, 1);
        setRatioBasisNames(names);
        syncRatioBasisSelector();
        state.ratioBasisValueSets = normalizeRatioBasisValueSets(state.ratioBasisValueSets, names);
        if (previousActive !== state.activeRatioBasisName) state.ratioBasisValues = [];
        markDirty();
        void useOrRefreshRatioBasisValues();
      }

      function syncRatioBasisSelector() {
        const menu = els.activeRatioBasisMenu;
        const button = els.activeRatioBasisButton;
        const names = getRatioBasisNames();
        const active = matchRatioBasisName(state.activeRatioBasisName, names) || names[0] || "";
        state.activeRatioBasisName = active;
        renderRatioBasisPills(names);
        if (!menu) return active;
        const previous = getDropdownValue(menu);
        menu.replaceChildren();
        if (!names.length) {
          menu.appendChild(makeDropdownOption("", "No basis", true));
          if (button) {
            button.disabled = true;
            button.title = "No ratio basis";
          }
          syncDropdownValue(menu, els.activeRatioBasisLabel, "", "No basis");
          closeDropdown(els.activeRatioBasisDropdown, button);
          return "";
        }
        for (const name of names) {
          menu.appendChild(makeDropdownOption(name, name, norm(name) === norm(active || names[0])));
        }
        if (button) button.disabled = false;
        const selected = syncDropdownValue(menu, els.activeRatioBasisLabel, active || names[0], names[0]);
        if (button && previous !== selected) {
          button.title = selected ? `Active ratio basis: ${selected}` : "No ratio basis";
        }
        return selected;
      }

      function getActiveRatioBasisName() {
        const selectValue = getDropdownValue(els.activeRatioBasisMenu);
        const names = getRatioBasisNames();
        const selected = matchRatioBasisName(selectValue, names)
          || matchRatioBasisName(state.activeRatioBasisName, names)
          || names[0]
          || "";
        state.activeRatioBasisName = selected;
        return selected;
      }

      function getDetails() {
        const ratioBases = getRatioBasisNames();
        const ratioBasis = getActiveRatioBasisName();
        const outputCategory = syncOutputCategory();
        return {
          name: text(els.nameInput.value),
          outputType: text(els.outputTypeInput.value),
          outputCategory,
          originLength: validOriginLength(els.originLengthInput.value),
          ratioBasis,
          ratioBases,
          showRatiosAsPercentages: !!els.showRatiosPctInput.checked,
          statisticDecimalPlaces: statisticDecimalPlacesValue(els.statisticDecimalsInput.value, 1),
          showWeights: !!els.showWeightsInput.checked,
        };
      }

      function getResultSelectionDisplayName() {
        return getDetails().name || "Result Selection";
      }

      function showCloseConfirm(reason = "close") {
        closeCellContextMenu();
        closeSourceContextMenu();
        return closeConfirm.confirm({ reason });
      }

      function requestConfirmedClose() {
        clearResultSelectionDependencyPreview("close-discard");
        postDirty(false, true);
        requestTabbedPageWindowClose({
          messageType: "arcrho:dataset-close-confirmed",
          inst,
        });
      }

      function normalizeRsTab(tab) {
        const key = norm(tab);
        return ALLOWED_RS_TABS.has(key) ? key : "details";
      }

      function getRsPageId(tab) {
        const next = normalizeRsTab(tab);
        return `rs${next[0].toUpperCase()}${next.slice(1)}Page`;
      }

      function isRsFlexPage(page) {
        return ["rsDetailsPage", "rsMethodPage", "rsChartPage", "rsResultsPage"].includes(page?.id);
      }

      function syncRsPageState(tab) {
        const activePageId = getRsPageId(tab);
        document.querySelectorAll(".rsPage").forEach((page) => {
          const active = page.id === activePageId;
          const floating = page.classList.contains("rsTabFloatingPage");
          page.classList.toggle("active", active);
          if (floating) {
            page.style.display = isRsFlexPage(page) ? "flex" : "block";
          } else if (!active) {
            page.style.display = "none";
          } else if (isRsFlexPage(page)) {
            page.style.display = "flex";
          } else {
            page.style.display = "block";
          }
        });
      }

      function refreshAuditLogFromSidecar() {
        void loadOutputSidecarSettings({ auditOnly: true }).catch((err) => {
          console.warn("Result Selection audit log refresh failed:", err);
        });
      }

      function onRsTabChanged(tab, previousTab) {
        const next = normalizeRsTab(tab);
        state.activeTab = next;
        syncRsPageState(next);
        try {
          window.parent?.postMessage({ type: "arcrho:result-selection-tab-changed", inst, tab: next }, "*");
        } catch {}
        if (next === "audit" && previousTab !== null) refreshAuditLogFromSidecar();
        if (next === "chart") window.requestAnimationFrame(() => ctx.refreshResultSelectionChart?.());
      }

      function setTab(tab) {
        const next = normalizeRsTab(tab);
        if (rsTabSystem) {
          if (rsTabSystem.getCurrentTab?.() !== next) rsTabSystem.setActive(next);
          else syncRsPageState(next);
          return;
        }
        state.activeTab = next;
        document.querySelectorAll(".rsTab").forEach((btn) => btn.classList.toggle("active", btn.dataset.page === next));
        document.querySelectorAll(".rsPage").forEach((page) => {
          const active = page.id === getRsPageId(next);
          page.classList.toggle("active", active);
          page.style.display = active ? (isRsFlexPage(page) ? "flex" : "block") : "none";
        });
        try {
          window.parent?.postMessage({ type: "arcrho:result-selection-tab-changed", inst, tab: next }, "*");
        } catch {}
      }

      function refreshRsFloatingTabLayout(tabId) {
        window.requestAnimationFrame(() => {
          if (tabId === "method") renderMethodGrid();
          if (tabId === "chart") ctx.refreshResultSelectionChart?.();
          if (tabId === "results") renderResultsGrid();
          if (tabId === "audit") refreshAuditLogFromSidecar();
        });
      }

      function wireRsGridScrollbarActivity() {
        document.querySelectorAll(".rsGridHost").forEach((host) => {
          if (host.__arcRhoScrollbarActivityWired) return;
          host.__arcRhoScrollbarActivityWired = true;

          let idleTimer = null;
          const syncScrollbarHover = (event) => {
            const rect = host.getBoundingClientRect();
            const verticalScrollbarWidth = Math.max(0, host.offsetWidth - host.clientWidth);
            const horizontalScrollbarHeight = Math.max(0, host.offsetHeight - host.clientHeight);
            const hasVerticalScrollbar = host.scrollHeight > host.clientHeight && verticalScrollbarWidth > 0;
            const hasHorizontalScrollbar = host.scrollWidth > host.clientWidth && horizontalScrollbarHeight > 0;
            const nearVerticalScrollbar = hasVerticalScrollbar
              && event.clientX >= rect.right - Math.max(verticalScrollbarWidth, 16);
            const nearHorizontalScrollbar = hasHorizontalScrollbar
              && event.clientY >= rect.bottom - Math.max(horizontalScrollbarHeight, 16);

            host.classList.toggle("isScrollbarHover", nearVerticalScrollbar || nearHorizontalScrollbar);
          };

          host.addEventListener("scroll", () => {
            host.classList.add("isScrolling");
            if (idleTimer) clearTimeout(idleTimer);
            idleTimer = setTimeout(() => {
              host.classList.remove("isScrolling");
            }, 550);
          }, { passive: true });
          host.addEventListener("pointermove", syncScrollbarHover, { passive: true });
          host.addEventListener("pointerleave", () => {
            host.classList.remove("isScrollbarHover");
          }, { passive: true });
        });
      }

      function wireEvents() {
        [els.nameInput, els.outputTypeInput, els.originLengthInput, els.showRatiosPctInput, els.showWeightsInput].forEach((el) => {
          el?.addEventListener("input", () => {
            markDirty();
            if (el === els.outputTypeInput) {
              syncOutputCategory();
              pruneLoadedSourcesForOutputCategory();
            }
            if (el === els.originLengthInput) {
              state.sidecarOriginLength = null;
              state.sidecarOriginLabels = [];
              setOriginLabels([], getDetails().originLength);
            }
            renderMethodGrid();
          });
          el?.addEventListener("change", () => {
            markDirty();
            if (el === els.outputTypeInput) {
              syncOutputCategory();
              pruneLoadedSourcesForOutputCategory();
            }
            if (el === els.originLengthInput) {
              state.sidecarOriginLength = null;
              state.sidecarOriginLabels = [];
              setOriginLabels([], getDetails().originLength);
              void (async () => {
                try {
                  await ensureDatasetCatalogLoaded();
                  await refreshOriginLabels({ render: false });
                  await reloadSourcesForCurrentOriginLength({ render: false });
                  state.ratioBasisValueSets = [];
                  if (getActiveRatioBasisName()) await refreshAllRatioBasisValues();
                  else renderMethodGrid();
                } catch (err) {
                  postStatus(`Origin length reload failed: ${err?.message || err}`, "error");
                  renderMethodGrid();
                }
              })();
              return;
            }
            renderMethodGrid();
          });
        });
        els.methodStatisticDecimalsInput?.addEventListener("input", () => {
          syncStatisticDecimalInputs("method");
          markDirty();
          renderMethodGrid();
        });
        els.methodStatisticDecimalsInput?.addEventListener("change", () => {
          syncStatisticDecimalInputs("method");
          markDirty();
          renderMethodGrid();
        });
        function stepStatisticDecimals(delta) {
          const target = els.methodStatisticDecimalsInput;
          const current = statisticDecimalPlacesValue(target?.value, 1);
          const next = String(statisticDecimalPlacesValue(current + delta, current));
          if (target) target.value = next;
          syncStatisticDecimalInputs("method");
          markDirty();
          renderMethodGrid();
        }
        els.methodStatisticDecimalsUp?.addEventListener("click", () => stepStatisticDecimals(1));
        els.methodStatisticDecimalsDown?.addEventListener("click", () => stepStatisticDecimals(-1));
        wireDropdown(els.weightDisplayDropdown, els.weightDisplayButton, els.weightDisplayMenu, (value) => {
          const next = text(value) === "effective";
          if (state.showEffectiveWeights === next) return;
          state.showEffectiveWeights = next;
          renderMethodGrid();
        });
        wireDropdown(els.originLengthDropdown, els.originLengthButton, els.originLengthMenu, (value) => {
          const next = validOriginLength(value, 0);
          if (!next || !els.originLengthInput || text(els.originLengthInput.value) === String(next)) return;
          els.originLengthInput.value = String(next);
          syncOriginLengthDropdownOptions();
          els.originLengthInput.dispatchEvent(new Event("input", { bubbles: true }));
          els.originLengthInput.dispatchEvent(new Event("change", { bubbles: true }));
        });
        els.cellContextMenu?.addEventListener("click", (event) => {
          const action = event.target?.closest?.("[data-rs-cell-action]")?.dataset?.rsCellAction || "";
          const table = els.cellContextMenu.dataset.table || "method";
          if (table === "results") {
            if (action === "copy-values") {
              void copyHighlightedResultsValues().catch((err) => postStatus(`Copy failed: ${err?.message || err}`, "error"));
            } else if (action === "remove-highlights") {
              removeResultsHighlights();
            }
            if (action) {
              closeCellContextMenu();
              focusResultsGrid();
            }
            return;
          }
          if (action === "copy-values") {
            void copyHighlightedMethodValues().catch((err) => postStatus(`Copy failed: ${err?.message || err}`, "error"));
            closeCellContextMenu();
          } else if (action === "paste-values") {
            void pasteHighlightedMethodValues().catch((err) => postStatus(`Paste failed: ${err?.message || err}`, "error"));
          } else if (action === "remove-highlights") {
            removeMethodHighlights();
            closeCellContextMenu();
          } else if (action === "revert-ultimate") {
            if (revertHighlightedUltimateValues()) {
              markDirty();
              renderMethodGrid();
            }
            closeCellContextMenu();
          } else if (action === "revert-all-ultimate") {
            if (revertAllUltimateValues()) {
              markDirty();
              renderMethodGrid();
            }
            closeCellContextMenu();
          }
          if (action) focusMethodGrid();
        });
        els.sourceContextMenu?.addEventListener("click", (event) => {
          const action = event.target?.closest?.("[data-rs-source-action]")?.dataset?.rsSourceAction || "";
          if (!action) return;
          const sourceIndex = sourceContextIndex();
          const anchor = {
            left: Number(els.sourceContextMenu.dataset.anchorLeft) || 8,
            bottom: Number(els.sourceContextMenu.dataset.anchorTop) || 8,
          };
          closeSourceContextMenu();
          if (action === "view-edit") {
            void viewOrEditSourceDataset(sourceIndex)
              .catch((err) => postStatus(`Open source dataset failed: ${err?.message || err}`, "error"));
          } else if (action === "add") {
            void openAddSourcePicker(anchor).catch((err) => postStatus(`Source picker failed: ${err?.message || err}`, "error"));
          } else if (action === "delete") {
            removeSourceAt(sourceIndex);
          }
        });
        document.addEventListener("mousedown", (event) => {
          if (els.cellContextMenu?.contains(event.target)) return;
          if (els.sourceContextMenu?.contains(event.target)) return;
          closeCellContextMenu();
          closeSourceContextMenu();
        }, true);
        document.addEventListener("keydown", (event) => {
          if (event.key === "Escape") {
            if (state.activeTab === "results" && normalizedResultsHighlight()) {
              removeResultsHighlights();
              event.preventDefault();
            } else if (normalizedMethodHighlight()) {
              removeMethodHighlights();
              event.preventDefault();
            } else if (normalizedResultsHighlight()) {
              removeResultsHighlights();
              event.preventDefault();
            } else {
              resetWeightEditSession();
            }
            closeCellContextMenu();
            closeSourceContextMenu();
            closeRatioBasisContextMenu();
            return;
          }
          if (handleMethodHighlightArrowKey(event)) return;
          if (handleResultsHighlightArrowKey(event)) return;
          if (
            (event.ctrlKey || event.metaKey)
            && event.key?.toLowerCase?.() === "c"
            && !isTextEntryTarget(event.target)
          ) {
            if (state.activeTab === "method" && normalizedMethodHighlight()) {
              event.preventDefault();
              void copyHighlightedMethodValues().catch((err) => postStatus(`Copy failed: ${err?.message || err}`, "error"));
              return;
            }
            if (state.activeTab === "results" && normalizedResultsHighlight()) {
              event.preventDefault();
              void copyHighlightedResultsValues().catch((err) => postStatus(`Copy failed: ${err?.message || err}`, "error"));
              return;
            }
          }
          if (
            (event.ctrlKey || event.metaKey)
            && event.key?.toLowerCase?.() === "v"
            && state.activeTab === "method"
            && normalizedMethodHighlight()
            && !isTextEntryTarget(event.target)
          ) {
            event.preventDefault();
            void pasteHighlightedMethodValues().catch((err) => postStatus(`Paste failed: ${err?.message || err}`, "error"));
            return;
          }
          if (
            state.activeTab === "method"
            && (event.key === "Delete" || event.key === "Backspace")
            && normalizedMethodHighlight()
            && !isTextEntryTarget(event.target)
          ) {
            let changed = false;
            if (highlightedHasWeightTargets()) changed = applyHighlightedWeightValue(0);
            else if (highlightedHasUltimateCells()) changed = revertHighlightedUltimateValues();
            if (changed) {
              event.preventDefault();
              markDirty();
              renderMethodGrid();
            }
            return;
          }
          if (
            state.activeTab === "method"
            && normalizedMethodHighlight()
            && !event.ctrlKey
            && !event.metaKey
            && !event.altKey
            && /^[0-9.]$/.test(event.key || "")
          ) {
            if (isTextEntryTarget(event.target)) return;
            if (applyHighlightedWeightKey(event.key)) {
              event.preventDefault();
              markDirty();
              renderMethodGrid();
            }
          }
        });
        (els.ratioBasisInputs || []).forEach((input) => {
          input?.addEventListener("change", () => {
            markDirty();
            const previousActive = state.activeRatioBasisName;
            syncRatioBasisSelector();
            if (previousActive !== state.activeRatioBasisName) state.ratioBasisValues = [];
            void refreshMissingRatioBasisValues().catch((err) => {
              postStatus(`Ratio Basis load failed: ${err?.message || err}`, "error");
            });
          });
        });
        els.ratioBasisAddButton?.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          void openRatioBasisDatasetPicker();
        });
        els.ratioBasisList?.addEventListener("click", (event) => {
          const removeButton = event.target.closest?.("button[data-ratio-basis-remove-index]");
          if (removeButton) {
            event.preventDefault();
            event.stopPropagation();
            removeRatioBasisAt(Number.parseInt(removeButton.dataset.ratioBasisRemoveIndex || "", 10));
            return;
          }
          const button = event.target.closest?.("button[data-ratio-basis-open-index]");
          if (!button) return;
          event.preventDefault();
          event.stopPropagation();
          void openRatioBasisDataset(Number.parseInt(button.dataset.ratioBasisOpenIndex || "", 10))
            .catch((err) => postStatus(`Ratio Basis open failed: ${err?.message || err}`, "error"));
        });
        els.ratioBasisList?.addEventListener("contextmenu", (event) => {
          const token = event.target.closest?.("[data-ratio-basis-index]");
          if (!token) return;
          openRatioBasisContextMenu(event, Number.parseInt(token.dataset.ratioBasisIndex || "", 10));
        });
        els.ratioBasisList?.addEventListener("dragstart", (event) => {
          const token = event.target.closest?.("[data-ratio-basis-index]");
          const index = Number.parseInt(token?.dataset.ratioBasisIndex || "", 10);
          const names = getRatioBasisNames();
          if (!token || !Number.isInteger(index) || !names[index]) return;
          closeRatioBasisContextMenu();
          state.ratioBasisDragIndex = index;
          token.classList.add("rsRatioBasisDragging");
          els.ratioBasisPicker?.classList.add("rsRatioBasisDragActive");
          if (event.dataTransfer) {
            event.dataTransfer.effectAllowed = "move";
            event.dataTransfer.setData("text/plain", names[index]);
          }
        });
        document.addEventListener("dragover", (event) => {
          if (!Number.isInteger(state.ratioBasisDragIndex)) return;
          event.preventDefault();
          const insidePicker = event.target.closest?.(".rsRatioBasisPicker");
          els.ratioBasisPicker?.classList.toggle("rsRatioBasisDragOutside", !insidePicker);
          if (event.dataTransfer) event.dataTransfer.dropEffect = insidePicker ? "none" : "move";
        });
        document.addEventListener("drop", (event) => {
          if (!Number.isInteger(state.ratioBasisDragIndex)) return;
          event.preventDefault();
          const index = state.ratioBasisDragIndex;
          const insidePicker = event.target.closest?.(".rsRatioBasisPicker");
          resetRatioBasisDragState();
          if (!insidePicker) removeRatioBasisAt(index);
        });
        document.addEventListener("dragend", resetRatioBasisDragState);
        els.ratioBasisContextMenu?.addEventListener("click", (event) => {
          if (!event.target.closest?.('[data-rs-ratio-basis-action="delete"]')) return;
          const index = Number.parseInt(els.ratioBasisContextMenu.dataset.ratioBasisIndex || "", 10);
          closeRatioBasisContextMenu();
          removeRatioBasisAt(index);
        });
        wireDropdown(els.activeRatioBasisDropdown, els.activeRatioBasisButton, els.activeRatioBasisMenu, (value) => {
          state.activeRatioBasisName = text(value);
          syncRatioBasisSelector();
          markDirty();
          void useOrRefreshRatioBasisValues();
        });
        document.addEventListener("mousedown", (event) => {
          const target = event.target;
          if (els.ratioBasisContextMenu?.contains?.(target)) return;
          closeRatioBasisContextMenu();
          if (target instanceof Node && (
            els.weightDisplayDropdown?.contains?.(target)
            || els.originLengthDropdown?.contains?.(target)
            || els.activeRatioBasisDropdown?.contains?.(target)
          )) return;
          closeAllDropdowns();
        });
        els.outputTypeBtn?.addEventListener("click", async () => {
          if (!state.project) {
            postStatus("Select a project before choosing an output vector.", "warn");
            return;
          }
          await ensureDatasetCatalogLoaded();
          await openDatasetNamePicker({
            projectName: state.project,
            initialName: els.outputTypeInput?.value || "",
            anchorElement: els.outputTypeInput || els.outputTypeBtn,
            title: "Select Output Vector",
            allowedDataFormats: ["Vector"],
            includeCalculated: true,
            emptyMessage: "No output vectors found (Vector).",
            setStatus: (message) => {
              const msg = text(message);
              if (msg) postStatus(msg, "warn");
            },
            onError: (err) => {
              console.error("Failed to open Result Selection output picker:", err);
              postStatus(`Error loading output vector names: ${String(err?.message || err)}`, "error");
            },
            onSelect: (name, item) => {
              const selected = text(name);
              if (!selected) return;
              els.outputTypeInput.value = selected;
              state.outputCategory = text(item?.category || datasetTypeCategoryForName(selected) || state.outputCategory);
              const removed = pruneLoadedSourcesForOutputCategory();
              markDirty();
              renderMethodGrid();
              if (removed) postStatus(`Removed ${removed} loaded source${removed === 1 ? "" : "s"} from a different Category.`, "warn");
            },
          });
        });
        els.saveBtn?.addEventListener("click", () => {
          saveResultSelection()
            // A save keeps the window open; only Cancel and a confirmed dirty
            // close dismiss it.
            .then(async (saved) => {
              if (saved?.ok && saved?.propagationClean) {
                await showSavedDependentsNotice(saved.refreshedDatasets, { linkWarnings: saved.linkWarnings });
              }
            })
            .catch((err) => postStatus(`Result Selection save failed: ${err?.message || err}`, "error"));
        });
        els.cancelBtn?.addEventListener("click", async () => {
          if (!isDirty) {
            requestConfirmedClose();
            return;
          }
          const discard = await showCloseConfirm("close");
          if (!discard) return;
          try {
            await restoreCleanState();
            requestConfirmedClose();
          } catch (err) {
            postStatus(`Result Selection restore failed: ${err?.message || err}`, "error");
          }
        });
        window.addEventListener("message", (event) => {
          const msg = event.data || {};
          if (msg.type === "arcrho:dataset-save" || msg.type === "arcrho:result-selection-save") {
            saveResultSelection()
              .then(async (saved) => {
                if (saved?.ok && saved?.propagationClean) {
                  await showSavedDependentsNotice(saved.refreshedDatasets, { linkWarnings: saved.linkWarnings });
                }
              })
              .catch((err) => postStatus(`Result Selection save failed: ${err?.message || err}`, "error"));
            return;
          }
          if (msg.type === "arcrho:dependency-source-preview") {
            if (applyDependencySourcePreview(msg)) noteDependencyEvent();
            return;
          }
          if (msg.type === "arcrho:dependency-source-cleared") {
            if (!sourceMessageMatchesContext(msg)) return;
            const directMatch = state.sources.some((source) => sourceMessageMatchesSource(msg, source));
            const basisMatch = sourceMessageMatchesRatioBasis(msg);
            const clearedNames = sourceMessageNames(msg);
            const reloadLocalSource = state.sources.some(
              (source) => sourceMessageMatchesSource(msg, source)
                && !state.persistedDependencyNames.has(norm(source.name)),
            );
            const reloadLocalBasis = getRatioBasisNames().some(
              (name) => clearedNames.has(norm(name))
                && !state.persistedDependencyNames.has(norm(name)),
            );
            const previewRemoved = removeDependencyPreview(msg);
            if (!directMatch && !basisMatch && !previewRemoved) return;
            noteDependencyEvent();
            scheduleDependencyClearRestore(msg, { reloadLocalSource, reloadLocalBasis });
            return;
          }
          if (msg.type === "arcrho:calculated-datasets-updated") {
            if (!calculatedUpdateAffectsCurrentResultSelection(msg)) return;
            noteDependencyEvent();
            schedulePersistedValuesRefresh(msg.source || "calculated dataset update");
          }
        });
        window.__arcrho_request_close = () => {
          if (!isDirty) return false;
          if (closeConfirm.isOpen) return true;
          void (async () => {
            const close = await showCloseConfirm("close");
            if (close) requestConfirmedClose();
          })();
          return true;
        };
        window.__arcrho_consume_close_shortcut = window.__arcrho_request_close;
      }

      function initTabs() {
        rsTabSystem = createTabbedPage(document.body, {
          tabs: RS_TAB_DEFS,
          cssPrefix: "rs",
          initialTab: state.activeTab,
          injectTabBar: false,
          onTabChange: onRsTabChanged,
        });
        window.rsTabSystem = rsTabSystem;
        wireTabPopoutWindows({
          cssPrefix: "rs",
          tabs: RS_TAB_DEFS,
          tabSystem: () => window.rsTabSystem,
          onPopoutTab: refreshRsFloatingTabLayout,
          onDockTab: refreshRsFloatingTabLayout,
          onFocusTab: refreshRsFloatingTabLayout,
        });
      }

      async function init() {
        withProgrammatic(() => {
          els.nameInput.value = text(params.get("name") || params.get("dataset_name"));
          els.outputTypeInput.value = text(params.get("output_type") || params.get("dataset_type") || els.nameInput.value);
          els.originLengthInput.value = String(validOriginLength(params.get("origin_length"), DEFAULT_ORIGIN_LENGTH));
          state.outputCategory = text(params.get("category"));
        });
        applyTabbedPageSaveBar(els.saveBar);
        initTabs();
        syncOriginLengthDropdownOptions();
        wireEvents();
        syncRatioBasisSelector();
        wireRsGridScrollbarActivity();
        wireNotes();
        let loadError = null;
        const loaded = await tryLoadExistingMethod().catch((err) => {
          loadError = err;
          state.loadBlocked = true;
          syncLoadBlockedControls();
          postStatus(`Result Selection load failed: ${err?.message || err}`, "error");
          return false;
        });
        state.initialLoadPending = false;
        if (loaded) {
          if (state.persistedRefreshReason || state.pendingDependencyClearMessages.length) {
            try {
              await flushPersistedValuesRefresh();
            } catch (err) {
              postStatus(`Result Selection dependency restore failed: ${err?.message || err}`, "error");
            }
          }
        } else {
          state.pendingDependencyClearMessages.length = 0;
          state.dependencyRestorePending = false;
          state.persistedRefreshReason = "";
        }
        let originLabelError = "";
        if (!loaded && !loadError) {
          await ensureDatasetCatalogLoaded().catch((err) => postStatus(`Cached dataset lookup failed: ${err?.message || err}`, "error"));
        }
        if (!loaded && !loadError && !state.originLabels.length) {
          try {
            await refreshOriginLabels({ render: false });
          } catch (err) {
            originLabelError = String(err?.message || err || "Origin labels are unavailable.");
            renderMethodGrid();
          }
        }
        if (!loaded && !loadError) {
          await initializeDefaultSources().catch((err) => postStatus(`Default source load failed: ${err?.message || err}`, "error"));
          renderMethodGrid();
        }
        setTab(state.activeTab);
        markClean();
        void ctx.refreshDetailsDependencies?.();
        if (loadError) {
          syncLoadBlockedControls();
          return;
        }
        if (originLabelError && !state.originLabels.length) {
          postStatus(`Origin labels unavailable: ${originLabelError}`, "error");
        } else if (state.needsReview) {
          postStatus("Result Selection loaded, but its saved dependency refresh needs review.", "warn");
        } else {
          postStatus("Result Selection ready.");
        }
      }

      return {
        numberOrNull,
        positiveInt,
        validOriginLength,
        validSourceOriginLength,
        isEngineSource,
        isVectorSource,
        datasetTypeCategoryForName,
        getOutputCategory,
        syncOutputCategory,
        sourceCategory,
        matchesOutputCategory,
        nonNegativeInt,
        statisticDecimalPlacesValue,
        syncStatisticDecimalInputs,
        getHostApi,
        postStatus,
        postDirty,
        markDirty,
        sourceMessageNames,
        sourceMessageMatchesContext,
        sourceMessageMatchesSource,
        syncLoadBlockedControls,
        normalizePreviewValues,
        buildResultSelectionDependencySourceMessage,
        postResultSelectionDependencySourceMessage,
        postResultSelectionDependencyPreview,
        clearResultSelectionDependencyPreview,
        reloadSourcesMatchingMessage,
        refreshPersistedValuesFromDisk,
        recordPersistedMethodDependencies,
        flushPersistedValuesRefresh,
        schedulePersistedValuesRefresh,
        scheduleDependencyClearRestore,
        resumePersistedValuesRefresh,
        applyDependencySourcePreview,
        reapplyActiveDependencyPreviews,
        withProgrammatic,
        dropdownOptions,
        getDropdownValue,
        setDropdownOpen,
        closeDropdown,
        closeAllDropdowns,
        makeDropdownOption,
        syncDropdownValue,
        wireDropdown,
        uniqueRatioBasisNames,
        getRatioBasisNames,
        setRatioBasisNames,
        matchRatioBasisName,
        normalizeRatioBasisDetails,
        renderRatioBasisPills,
        syncRatioBasisSelector,
        getActiveRatioBasisName,
        getDetails,
        getResultSelectionDisplayName,
        showCloseConfirm,
        requestConfirmedClose,
        normalizeRsTab,
        getRsPageId,
        syncRsPageState,
        refreshAuditLogFromSidecar,
        onRsTabChanged,
        setTab,
        refreshRsFloatingTabLayout,
        wireRsGridScrollbarActivity,
        wireEvents,
        initTabs,
        init
      };
    }
  };
})();
