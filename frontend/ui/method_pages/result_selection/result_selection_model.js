(function () {
  const parts = window.ResultSelectionParts || (window.ResultSelectionParts = {});

  parts.installModel = function installModel(ctx) {
    with (ctx) {
      // The open-window change watch is host-provided; installs without one
      // (tests, embedded hosts) run with an inert stand-in.
      const rsWatch = ctx.rsObjectChangeWatch
        || { ensure() {}, pause() {}, resume() {}, stop() {} };
      // The saving animation is host-provided for the same reason; installs
      // without a document run with an inert stand-in.
      const saveProgress = ctx.rsSaveProgress
        || {
          run: (work) => work({
            writing() {},
            finish() {},
          }),
        };
      function buildPayload() {
        const details = getDetails();
        return buildResultSelectionMethodPayload({
          details,
          originLabels: Array.from({ length: getRowCount() }, (_, index) => originLabel(index)),
          showWeights: details.showWeights,
          sources: state.sources,
          ratioBasisValueSets: state.ratioBasisValueSets,
          calculatedUltimate: calculatedUltimateVector(),
          selectedUltimate: selectedUltimateVector(),
          ultimateOverrides: serializedUltimateOverrides(),
          lastModified: new Date().toISOString(),
        });
      }

      async function applyPayload(payload) {
        const data = payload && typeof payload === "object" ? payload : {};
        const details = data.details_tab || {};
        const method = data.method_tab || {};
        const ratioBasisDetails = normalizeRatioBasisDetails(details);
        withProgrammatic(() => {
          els.nameInput.value = text(details.name || els.nameInput.value);
          els.outputTypeInput.value = text(details.output_type || els.outputTypeInput.value);
          els.originLengthInput.value = String(validOriginLength(details.origin_length || els.originLengthInput.value));
          if (state.sidecarOriginLength) els.originLengthInput.value = String(state.sidecarOriginLength);
          (els.ratioBasisInputs || []).forEach((input, index) => {
            if (input) input.value = ratioBasisDetails.names[index] || "";
          });
          state.activeRatioBasisName = ratioBasisDetails.active;
          syncRatioBasisSelector();
          els.showRatiosPctInput.checked = details.show_ratios_as_percentages !== false;
          els.statisticDecimalsInput.value = String(Math.max(0, Math.min(8, nonNegativeInt(details.statistic_decimal_places, 1))));
          syncStatisticDecimalInputs();
          els.showWeightsInput.checked = method.show_weights !== false;
        });

        state.sources = (Array.isArray(method.loaded_datasets) ? method.loaded_datasets : [])
          .map((source) => buildSourceFromPersisted(source))
          .filter(Boolean);
        state.ratioBasisValueSets = normalizeRatioBasisValueSets(
          method.ratio_basis_values,
          ratioBasisDetails.names,
        );
        state.ratioBasisValues = ratioBasisValuesForName(
          state.ratioBasisValueSets,
          getActiveRatioBasisName(),
        );

        const methodOriginLabels = Array.isArray(method.origin_labels) ? method.origin_labels.map(String) : [];
        if (methodOriginLabels.length && !shouldRejectOriginLabels(getDetails().originLength, methodOriginLabels)) {
          setOriginLabels(methodOriginLabels, getDetails().originLength);
        } else if (state.sidecarOriginLabels.length && !shouldRejectOriginLabels(getDetails().originLength, state.sidecarOriginLabels)) {
          setOriginLabels(state.sidecarOriginLabels, getDetails().originLength);
        } else {
          setOriginLabels([], getDetails().originLength);
        }
        state.ultimateOverrides = normalizeUltimateOverrides(method.ultimate_overrides, getRowCount());
        renderMethodGrid();
      }

      function applyOutputSidecar(sidecar, options = {}) {
        const payload = sidecar && typeof sidecar === "object" ? sidecar : {};
        if (payload.exists === false) {
          auditLogView.clear();
          return false;
        }
        auditLogView.render(payload.audit_log);
        state.needsReview = Number(payload.status) === 2;
        if (options.auditOnly === true) return true;
        setNotesText(String(payload.notes ?? ""));
        const originLength = validOriginLength(payload.origin_length, 0);
        const labels = Array.isArray(payload.origin_labels) ? payload.origin_labels.map(String) : [];
        const resolvedLabels = shouldRejectOriginLabels(originLength, labels) ? [] : labels;
        state.sidecarOriginLength = originLength || null;
        state.sidecarOriginLabels = resolvedLabels;
        if (originLength) applyOriginLength(originLength);
        if (resolvedLabels.length) setOriginLabels(resolvedLabels, originLength || getDetails().originLength);
        return true;
      }

      async function fetchPersistedResultSelection(includeMethod = true) {
        const resp = await fetch("/result-selection/load", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            project_name: state.project,
            reserving_class: state.reservingClass,
            method_name: text(els.nameInput.value),
            include_method: includeMethod,
          }),
        });
        const payload = await resp.json().catch(() => ({}));
        if (!resp.ok || payload?.ok === false) {
          throw new Error(payload?.detail || payload?.error || `Result Selection load failed (${resp.status}).`);
        }
        return payload;
      }

      async function tryLoadExistingMethod() {
        if (!text(els.nameInput.value)) return false;
        const loadSeq = invalidatePersistedRefresh();
        const result = await fetchPersistedResultSelection(true);
        if (state.persistedRefreshSeq !== loadSeq) return true;
        applyOutputSidecar(result.sidecar);
        if (!result.method_exists || !result.method) return false;
        await applyPayload(result.method);
        recordPersistedMethodDependencies(result.method);
        state.methodRevision = String(result.method_revision || "");
        state.loadBlocked = false;
        rsWatch.ensure({
          projectName: state.project,
          reservingClass: state.reservingClass,
          methodName: getDetails().name,
          outputDataset: getDetails().name,
          selfWriteStamp: result?.sidecar?.updated_at,
        });
        const basisNames = getRatioBasisNames();
        const storedBasisNames = new Set(
          (Array.isArray(result.method?.method_tab?.ratio_basis_values)
            ? result.method.method_tab.ratio_basis_values
            : [])
            .map((item) => norm(item?.name))
            .filter(Boolean),
        );
        if (basisNames.some((name) => !storedBasisNames.has(norm(name)))) {
          postStatus("This legacy Result Selection has no persisted Ratio Basis values. Save it once to upgrade the method JSON.", "warn");
        } else {
          postStatus(`Loaded Result Selection: ${getDetails().name}`);
        }
        return true;
      }

      function snapshotPayload() {
        return JSON.stringify({ method: buildPayload(), notes: els.notesInput?.value || "" });
      }

      function markClean() {
        cleanSnapshot = snapshotPayload();
        notesController.markClean();
        clearResultSelectionDependencyPreview("clean");
        postDirty(false, true);
      }

      function invalidatePersistedRefresh() {
        state.persistedRefreshSeq = Number(state.persistedRefreshSeq || 0) + 1;
        if (state.persistedRefreshTimer != null) {
          window.clearTimeout(state.persistedRefreshTimer);
          state.persistedRefreshTimer = null;
        }
        state.persistedRefreshReason = "";
        return state.persistedRefreshSeq;
      }

      function assertPersistedMutationReady(mutation = null) {
        if (state.loadBlocked) {
          throw new Error("Reload the Result Selection successfully before saving.");
        }
        if (state.initialLoadPending) {
          throw new Error("Wait for the Result Selection to finish loading before saving.");
        }
        if (state.dependencyRestoreError) {
          throw new Error(`Resolve the upstream dependency restore error before saving: ${state.dependencyRestoreError}`);
        }
        if (state.dependencyRestorePending) {
          throw new Error("Wait for the upstream dependency restore to finish before saving Result Selection.");
        }
        if (state.hasDependencyPreview) {
          throw new Error("Save or discard the upstream dependency preview before saving Result Selection.");
        }
        if (mutation && state.persistedMutationInFlight !== mutation.id) {
          throw new Error("Another Result Selection save superseded this operation.");
        }
        if (mutation && Number(state.dependencyEventSeq || 0) !== mutation.dependencyEventSeq) {
          throw new Error("An upstream dependency changed while Result Selection was preparing to save; wait for refresh and try again.");
        }
      }

      function beginPersistedMutation() {
        assertPersistedMutationReady();
        if (state.persistedMutationInFlight) {
          throw new Error("Another Result Selection save is already in progress.");
        }
        const mutation = {
          id: Number(state.persistedMutationSeq || 0) + 1,
          dependencyEventSeq: Number(state.dependencyEventSeq || 0),
        };
        state.persistedMutationSeq = mutation.id;
        state.persistedMutationInFlight = mutation.id;
        return mutation;
      }

      function finishPersistedMutation(mutation) {
        if (mutation && state.persistedMutationInFlight === mutation.id) {
          state.persistedMutationInFlight = 0;
        }
        resumePersistedValuesRefresh();
      }

      function reconcilePersistedMutation(mutation, reason) {
        reapplyActiveDependencyPreviews();
        const dependencyChanged = !!mutation
          && Number(state.dependencyEventSeq || 0) !== mutation.dependencyEventSeq;
        if (dependencyChanged && !state.hasDependencyPreview && !state.dependencyRestorePending) {
          schedulePersistedValuesRefresh(reason);
        }
      }

      function trackPersistedPropagation(payload, progress = null) {
        return trackSavePropagation(payload?.propagation, {
          onStatus: (message, statusOptions) => {
            progress?.setMessage?.(message, statusOptions);
            postStatus(message, statusOptions?.tone === "warn" ? "warn" : "");
          },
          onComplete: () => {
            try {
              window.parent?.postMessage({ type: "arcrho:project-instance-refresh-datasets" }, "*");
            } catch {}
          },
        });
      }

      async function saveResultSelection() {
        return saveProgress.run((progress) => runResultSelectionSave(progress));
      }

      /** 32-hex save-job identity, unique per save attempt. */
      function newHostedSaveRequestId() {
        try {
          if (typeof crypto !== "undefined" && crypto.randomUUID) {
            return crypto.randomUUID().replace(/-/g, "");
          }
        } catch {}
        let id = "";
        while (id.length < 32) id += Math.floor(Math.random() * 16).toString(16);
        return id;
      }

      async function runResultSelectionSave(progress) {
        const details = getDetails();
        if (!details.name || !details.outputType) {
          postStatus("Result Selection save requires Name and Output Type.", "error");
          return { ok: false };
        }
        const mutation = beginPersistedMutation();
        rsWatch.pause();
        try {
          await refreshOriginLabels({ render: false });
          assertPersistedMutationReady(mutation);
          const method = buildPayload();
          const saveBody = {
            project_name: state.project,
            reserving_class: state.reservingClass,
            method,
            notes: els.notesInput?.value || "",
            expected_revision: state.methodRevision,
            // The save job's identity is chosen client-side so the saving
            // card can follow the Engine's dependent walk live (one row per
            // refreshed object) while this request is still in flight.
            client_request_id: newHostedSaveRequestId(),
          };
          progress.writing();
          progress.trackHostedSave?.(saveBody.client_request_id);
          const resp = await fetch("/result-selection/save", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(saveBody),
          });
          const payload = await resp.json().catch(() => ({}));
          if (!resp.ok || payload?.ok === false) {
            const message = String(payload?.detail || payload?.error || `Result Selection save failed (${resp.status}).`);
            if (isEngineUnavailableSaveError({ status: resp.status, message })) {
              // The save was refused before anything was written; unsaved
              // work stays in this window. Drop the spinner first so the
              // message box cannot open behind it.
              progress.finish();
              void showPageMessageBox({ title: "ArcRho Engine Unavailable", message, tone: "warn" });
            }
            throw new Error(message);
          }
          invalidateOutputSidecarLoad();
          auditLogView.render(payload?.sidecar?.audit_log);
          state.methodRevision = String(payload.method_revision || "");
          recordPersistedMethodDependencies(payload.method || method);
          state.needsReview = false;
          state.pendingPropagationJobId = String(payload?.propagation?.job_id || "").trim();
          markClean();
          rsWatch.ensure({
            projectName: state.project,
            reservingClass: state.reservingClass,
            methodName: getDetails().name,
            outputDataset: getDetails().name,
            selfWriteStamp: payload?.sidecar?.updated_at,
          });
          reconcilePersistedMutation(mutation, "dependency update during Result Selection save");
          // A save rewrites the graph on both sides, so the Details rows are
          // stale until they are re-read.
          void ctx.refreshDetailsDependencies?.();
          try {
            window.parent?.postMessage({ type: "arcrho:project-instance-refresh-datasets" }, "*");
          } catch {}
          const aggregateCount = Array.isArray(payload.aggregated_csv_paths) ? payload.aggregated_csv_paths.length : 0;
          if (payload.propagation_ok === false) {
            // The walk ran; some dependents declined or failed. The server
            // names them, so show that instead of a generic scheduling line.
            const walkReason = String(payload?.propagation?.message || "").trim();
            postStatus(
              `Result Selection saved, but some dependent updates did not complete: ${walkReason || details.name}`,
              "warn",
            );
          } else if (payload.index_ok === false) {
            postStatus(`Result Selection saved, but the dataset index could not be refreshed: ${payload.index_error || details.name}`, "warn");
          } else {
            postStatus(`Result Selection saved: ${details.name}${aggregateCount ? ` (+${aggregateCount} aggregated)` : ""}`);
          }
          // Hold the saving card open through the dependent walk so the user
          // sees each live update; a null outcome (failed or stalled walk)
          // keeps the window open with the dataset table as the failure
          // surface.
          const propagationOutcome = await trackPersistedPropagation(payload, progress);
          // The save and its dependent walk are done; drop the spinner before
          // the review dialog.
          progress.finish();
          await showMethodSaveReviewWarning(payload, {
            instanceId: inst,
            projectName: state.project,
            reservingClass: state.reservingClass,
          });
          return {
            ...payload,
            propagationClean: propagationOutcome !== null,
            refreshedDatasets: propagationOutcome?.refreshed_datasets || [],
            linkWarnings: propagationOutcome?.link_warnings || [],
          };
        } finally {
          rsWatch.resume();
          finishPersistedMutation(mutation);
        }
      }

      function setNotesText(value) {
        notesController.setValue(String(value ?? ""), { markClean: true });
      }

      function wireNotes() {
        return notesController;
      }

      return {
        buildPayload,
        applyPayload,
        applyOutputSidecar,
        fetchPersistedResultSelection,
        tryLoadExistingMethod,
        snapshotPayload,
        markClean,
        saveResultSelection,
        beginPersistedMutation,
        finishPersistedMutation,
        assertPersistedMutationReady,
        setNotesText,
        wireNotes
      };
    }
  };
})();
