import {
  getDfmIsDirty,
  getEffectiveDevLabelsForModel,
  getResolvedProjectName,
  getResolvedReservingClass,
  markDfmDirty,
  getRatioHeaderLabels,
  state,
} from "/ui/method_pages/dfm/dfm_state.js";
import {
  applyDfmOwnedPatchPayload,
  saveRatioSelectionPattern,
} from "/ui/method_pages/dfm/dfm_persistence.js?v=20260907a";
import {
  confirmDfmRpcBridgeAction,
  createDfmRpcBridgeDialog,
  createDfmRpcBridgeMessageBox,
} from "/ui/method_pages/dfm/dfm_rpc_bridge_dialog.js?v=20260514c";

let syncInFlight = false;

function textValue(id) {
  return String(document.getElementById(id)?.value || "").trim();
}

function numberValue(id, fallback) {
  const raw = Number.parseInt(textValue(id), 10);
  return Number.isFinite(raw) ? raw : fallback;
}

function buildRequestPayload() {
  return {
    project_name: getResolvedProjectName() || textValue("projectSelect"),
    reserving_class: getResolvedReservingClass() || textValue("pathInput"),
    method_name: textValue("dfmMethodName"),
    output_vector: textValue("dfmOutputVector"),
    input_triangle: textValue("triInput"),
    origin_length: numberValue("originLenSelect", 12),
    development_length: numberValue("devLenSelect", 12),
    decimal_places: numberValue("decimalPlaces", 4),
    timeout_sec: 8.0,
  };
}

function validatePayload(payload) {
  const missing = [];
  if (!payload.project_name) missing.push("Project");
  if (!payload.reserving_class) missing.push("Segment");
  if (!payload.method_name) missing.push("Name");
  if (!payload.output_vector) missing.push("Output Type");
  if (!payload.input_triangle) missing.push("Input Triangle");
  if (!payload.origin_length) missing.push("Origin Length");
  if (!payload.development_length) missing.push("Development Length");
  return missing;
}

async function postJson(url, payload) {
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data?.detail || data?.message || `Request failed: ${resp.status}`);
  }
  return data;
}

async function cleanupRemoteTmp(payload) {
  if (!payload) return null;
  try {
    return await postJson("/dfm/rpc-bridge/cleanup", payload);
  } catch (err) {
    postStatus(`DFM sync cleanup failed: ${String(err?.message || err)}`, "warn");
    return null;
  }
}

function postStatus(text, tone = "") {
  window.parent.postMessage({ type: "arcrho:status", text, ...(tone ? { tone } : {}) }, "*");
}

function formatApplyResultMessage(data) {
  const missing = Array.isArray(data?.sync_report?.missing_components)
    ? data.sync_report.missing_components
    : [];
  if (!missing.length) return { text: "Local updated.", tone: "ok" };
  const lines = [
    "Local updated, but these RPC components were missing and could not be synced:",
    ...missing.map((name) => `- ${name}`),
  ];
  return { text: lines.join("\n"), tone: "warn" };
}

function buildCurrentPatternLabelFallbacks() {
  const model = state?.model || {};
  const originLabels = Array.isArray(model.origin_labels)
    ? model.origin_labels.map((label) => String(label ?? ""))
    : [];
  const ratioLabels = getRatioHeaderLabels(getEffectiveDevLabelsForModel(model));
  const developmentLabels = ratioLabels.map((label, index) => {
    const text = String(label ?? "");
    if (index === ratioLabels.length - 1) return text || "Ult";
    return text ? `(${index + 1}) ${text}` : `(${index + 1})`;
  });
  return {
    origin_labels: originLabels,
    development_labels: developmentLabels,
  };
}

function cleanText(value) {
  return String(value ?? "").trim();
}

function jsonTab(payload, key) {
  const value = payload && typeof payload === "object" && !Array.isArray(payload)
    ? payload[key]
    : null;
  return value && typeof value === "object" && !Array.isArray(value) ? value : {};
}

function normalizeBinaryCell(value, missingValue = 0) {
  if (value === 1 || value === true || value === "1" || value === "true" || value === "True") return 1;
  if (missingValue === 2 && (value === 2 || value === "2")) return 2;
  return 0;
}

function extractPatternSnapshot(payload) {
  const ratiosTab = jsonTab(payload, "ratios_tab");
  const ratioTriangle = jsonTab(ratiosTab, "ratio_triangle");
  const dataTab = jsonTab(payload, "data_tab");
  const pattern = ratioTriangle.excluded;
  const originLabels = Array.isArray(ratioTriangle["origin_labels"])
    ? ratioTriangle["origin_labels"]
    : dataTab["origin_labels"];
  const developmentLabels = Array.isArray(ratioTriangle["development_labels"])
    ? ratioTriangle["development_labels"]
    : dataTab["development_labels"];
  const previewOriginLabels = Array.isArray(originLabels) ? originLabels.map(cleanText) : [];
  const previewDevelopmentLabels = Array.isArray(developmentLabels) ? developmentLabels.map(cleanText) : [];
  if (!Array.isArray(pattern)) {
    return {
      exists: false,
      rows: 0,
      columns: 0,
      selected_count: 0,
      preview: [],
      origin_labels: [],
      development_labels: [],
    };
  }
  let columns = 0;
  let selectedCount = 0;
  const preview = pattern.map((row) => {
    if (!Array.isArray(row)) return [];
    columns = Math.max(columns, row.length);
    return row.map((cell) => {
      const value = normalizeBinaryCell(cell, 2);
      if (value === 1) selectedCount += 1;
      return value;
    });
  });
  return {
    exists: true,
    rows: pattern.length,
    columns,
    selected_count: selectedCount,
    preview,
    origin_labels: previewOriginLabels,
    development_labels: previewDevelopmentLabels,
  };
}

function extractAverageFormulaSnapshot(payload) {
  const ratiosTab = jsonTab(payload, "ratios_tab");
  const formulaPayload = jsonTab(ratiosTab, "average_formulas");
  const ratioTriangle = jsonTab(ratiosTab, "ratio_triangle");
  const dataTab = jsonTab(payload, "data_tab");
  const labels = Array.isArray(formulaPayload.label) ? formulaPayload.label : [];
  const selected = Array.isArray(formulaPayload.selected) ? formulaPayload.selected : null;
  const developmentLabels = Array.isArray(ratioTriangle["development_labels"])
    ? ratioTriangle["development_labels"]
    : dataTab["development_labels"];
  if (!selected) {
    return {
      exists: false,
      rows: 0,
      columns: 0,
      selected_count: 0,
      preview: [],
      formula_labels: labels.map(cleanText),
      development_labels: Array.isArray(developmentLabels) ? developmentLabels.map(cleanText) : [],
    };
  }
  let columns = 0;
  let selectedCount = 0;
  const preview = selected.map((row) => {
    if (!Array.isArray(row)) return [];
    columns = Math.max(columns, row.length);
    return row.map((cell) => {
      const value = normalizeBinaryCell(cell);
      if (value === 1) selectedCount += 1;
      return value;
    });
  });
  return {
    exists: true,
    rows: Math.max(preview.length, labels.length),
    columns,
    selected_count: selectedCount,
    preview,
    formula_labels: labels.map(cleanText),
    development_labels: Array.isArray(developmentLabels) ? developmentLabels.map(cleanText) : [],
  };
}

function extractCellNotesSnapshot(payload) {
  const ratiosTab = jsonTab(payload, "ratios_tab");
  const cellNotes = jsonTab(ratiosTab, "cell_notes");
  const entries = [];
  Object.entries(cellNotes).forEach(([tableKey, tableNotes]) => {
    if (!tableNotes || typeof tableNotes !== "object" || Array.isArray(tableNotes)) return;
    Object.entries(tableNotes).forEach(([rowLabel, rowNotes]) => {
      if (!rowNotes || typeof rowNotes !== "object" || Array.isArray(rowNotes)) return;
      Object.entries(rowNotes).forEach(([colLabel, note]) => {
        const text = cleanText(note);
        if (!text) return;
        entries.push({
          table: cleanText(tableKey),
          row: cleanText(rowLabel),
          column: cleanText(colLabel),
          note: text,
        });
      });
    });
  });
  entries.sort((a, b) => (
    `${a.table}\t${a.row}\t${a.column}\t${a.note}`.localeCompare(`${b.table}\t${b.row}\t${b.column}\t${b.note}`)
  ));
  return {
    exists: entries.length > 0,
    count: entries.length,
    entries: entries.slice(0, 50),
    truncated: entries.length > 50,
  };
}

function extractMethodNotesSnapshot(payload) {
  const metadata = jsonTab(payload, "method_metadata");
  if (!Object.prototype.hasOwnProperty.call(metadata, "method_notes")) {
    return { exists: false, text: "" };
  }
  return { exists: true, text: String(metadata["method_notes"] ?? "") };
}

function buildJsonSnapshot(payload) {
  const safePayload = payload && typeof payload === "object" && !Array.isArray(payload) ? payload : {};
  const formulaPayload = jsonTab(jsonTab(safePayload, "ratios_tab"), "average_formulas");
  const formulas = Array.isArray(formulaPayload.label) ? formulaPayload.label : [];
  return {
    available: !!Object.keys(safePayload).length,
    error: "",
    ratio_pattern: extractPatternSnapshot(safePayload),
    average_formula_pattern: extractAverageFormulaSnapshot(safePayload),
    cell_notes: extractCellNotesSnapshot(safePayload),
    method_notes: extractMethodNotesSnapshot(safePayload),
    average_formulas: formulas.map((item) => String(item)),
    last_modified: cleanText(jsonTab(safePayload, "method_metadata")["last_modified"]),
  };
}

function buildApprovalMeta(payload, fallbackLabel, timestamp) {
  const metadataTime = cleanText(jsonTab(payload, "method_metadata")["last_modified"]);
  return {
    exists: !!payload && typeof payload === "object" && !Array.isArray(payload),
    last_modified: metadataTime || fallbackLabel,
    last_modified_timestamp: timestamp,
  };
}

function buildAgentApprovalComparison(originalJson, proposedJson) {
  const nowSeconds = Date.now() / 1000;
  return {
    ok: true,
    status: "approval_pending",
    comparison: "approval_pending",
    local: buildApprovalMeta(originalJson, "Current DFM tab", nowSeconds),
    remote: buildApprovalMeta(proposedJson, "Pending ArcBot edit", nowSeconds + 1),
    labels: {
      local: "ArcRho - Current",
      remote: "ArcBot - Proposed",
    },
    actions: {
      local: "reject-agent-edit",
      remote: "accept-agent-edit",
    },
    snapshots: {
      local: buildJsonSnapshot(originalJson),
      remote: buildJsonSnapshot(proposedJson),
    },
  };
}

export function reviewArcBotDfmEditApproval(options = {}) {
  return new Promise((resolve) => {
    let settled = false;
    const finish = (payload) => {
      if (settled) return;
      settled = true;
      resolve(payload);
    };
    const originalJson = options?.originalJson;
    const proposedJson = options?.proposedJson;
    if (!proposedJson || typeof proposedJson !== "object" || Array.isArray(proposedJson)) {
      finish({ ok: false, error: "ArcBot did not provide a valid proposed DFM method." });
      return;
    }
    const dialog = createDfmRpcBridgeDialog({
      onClose: (reason) => {
        if (reason === "primary-action") return;
        finish({ ok: true, accepted: false, message: "DFM edit approval was closed. No changes were applied." });
      },
    });
    dialog.setComparison(buildAgentApprovalComparison(originalJson, proposedJson), {
      labelFallbacks: buildCurrentPatternLabelFallbacks(),
      onPrimary: async (action) => {
        if (action === "reject-agent-edit") {
          dialog.close("primary-action");
          postStatus("ArcBot DFM edit rejected.");
          finish({ ok: true, accepted: false, message: "DFM edit was rejected. No changes were applied." });
          return;
        }
        if (action !== "accept-agent-edit") return;
        dialog.setBusy(true);
        const statusDialog = createDfmRpcBridgeMessageBox("Applying approved ArcBot DFM edit...", "", {
          title: "ArcBot DFM Edit",
        });
        statusDialog.setBusy(true);
        dialog.close("primary-action");
        try {
          const applied = await applyDfmOwnedPatchPayload(proposedJson, { reason: "arcbot-approval" });
          if (!applied?.ok) {
            statusDialog.setMessage("Could not apply the approved DFM edit to this tab.", "error");
            finish({ ok: false, error: "Could not apply the approved DFM edit to this tab." });
            return;
          }
          statusDialog.setWaiting("Saving approved DFM method...");
          const saved = await saveRatioSelectionPattern(false, { showReviewWarning: false });
          if (!saved?.ok) {
            markDfmDirty();
            const saveError = String(saved?.error || "").trim();
            statusDialog.setMessage(
              saveError
                ? `Applied in the app, but final JSON save failed: ${saveError} Save the DFM before closing.`
                : "Applied in the app, but final JSON save failed. Save the DFM before closing.",
              "warn",
            );
            finish({
              ok: false,
              error: saveError
                ? `Approved DFM edit applied in the app, but final JSON save failed: ${saveError}`
                : "Approved DFM edit applied in the app, but final JSON save failed.",
            });
            return;
          }
          const message = options?.reply || "Applied the approved DFM edit.";
          statusDialog.setMessage("Approved DFM edit applied.", "ok");
          postStatus("ArcBot DFM edit approved and saved.");
          finish({ ok: true, accepted: true, message });
        } catch (err) {
          const message = String(err?.message || err || "Approved DFM edit failed.");
          statusDialog.setMessage(message, "error");
          postStatus(`ArcBot DFM edit failed: ${message}`, "warn");
          finish({ ok: false, error: message });
        } finally {
          statusDialog.setBusy(false);
          dialog.setBusy(false);
        }
      },
    });
  });
}

async function ensureSavedBeforeSync(dialog) {
  if (!getDfmIsDirty()) return true;
  const shouldSave = window.confirm("This DFM tab has unsaved edits. Save and proceed with sync?");
  if (!shouldSave) return false;
  dialog.setWaiting("Saving current DFM method before sync...");
  const result = await saveRatioSelectionPattern(false, { showReviewWarning: false });
  if (!result?.ok) {
    dialog.setMessage(result?.error ? `Save failed: ${result.error}` : "Save was canceled. Sync stopped.", "error");
    return false;
  }
  return true;
}

async function refreshComparison(dialog, payload) {
  dialog.setBusy(true);
  try {
    const data = await postJson("/dfm/rpc-bridge/compare", payload);
    dialog.setComparison(data, {
      labelFallbacks: buildCurrentPatternLabelFallbacks(),
      onRefresh: () => refreshComparison(dialog, payload),
      onPrimary: (action) => runPrimaryAction(dialog, payload, action),
    });
  } catch (err) {
    dialog.setMessage(String(err?.message || err), "error");
  } finally {
    dialog.setBusy(false);
  }
}

async function runPrimaryAction(dialog, payload, action) {
  let actionPayload = payload;
  if (action === "update-remote") {
    const confirmed = await confirmDfmRpcBridgeAction(
      "This action will write the selected DFM settings to the RPC server. Continue?",
      { title: "Confirm Remote Update" },
    );
    if (!confirmed) return;
    actionPayload = { ...payload, rpc_server_write_confirmed: true };
  }

  dialog.setBusy(true);
  const statusDialog = createDfmRpcBridgeMessageBox("Preparing selected DFM version action...");
  statusDialog.setBusy(true);
  dialog.close("primary-action");
  try {
    if (action === "update-local") {
      statusDialog.setWaiting("Updating local DFM JSON from remote...");
      const data = await postJson("/dfm/rpc-bridge/apply", payload);
      const applied = await applyDfmOwnedPatchPayload(data?.payload, { reason: "rpc-update-local" });
      if (!applied?.ok) {
        statusDialog.setMessage("Updated, but could not reload this tab.", "error");
        postStatus("DFM sync: local JSON updated, but tab apply failed.", "warn");
        return;
      }
      markDfmDirty();
      statusDialog.setWaiting("Saving recalculated local DFM JSON...");
      const saved = await saveRatioSelectionPattern(false, { showReviewWarning: false });
      if (!saved?.ok) {
        const saveError = String(saved?.error || "").trim();
        statusDialog.setMessage(
          saveError
            ? `Local updated in app, but final JSON save failed: ${saveError} Save the DFM before closing.`
            : "Local updated in app, but final JSON save failed. Save the DFM before closing.",
          "warn",
        );
        postStatus(
          saveError
            ? `DFM sync: local app data updated, but final JSON save failed: ${saveError}`
            : "DFM sync: local app data updated, but final JSON save failed.",
          "warn",
        );
        return;
      }
      const resultMessage = formatApplyResultMessage(data);
      statusDialog.setMessage(resultMessage.text, resultMessage.tone);
      postStatus(
        resultMessage.tone === "warn"
          ? "DFM sync: local DFM JSON updated from remote with missing RPC components."
          : "DFM sync: local DFM JSON updated from remote.",
        resultMessage.tone === "warn" ? "warn" : "",
      );
      return;
    }
    if (action === "keep-local") {
      statusDialog.setWaiting("Keeping local DFM JSON and removing remote RPC JSON...");
      const data = await postJson("/dfm/rpc-bridge/keep-local", payload);
      const message = data?.ok ? "No changes made on local." : (data?.message || "Keep local failed.");
      statusDialog.setMessage(message, data?.ok ? "ok" : "error");
      postStatus(`DFM sync: ${message}`, data?.ok ? "" : "warn");
      return;
    }
    if (action === "update-remote") {
      statusDialog.setWaiting("Sending SyncDFM request and waiting for remote result...");
      const data = await postJson("/dfm/rpc-bridge/update-remote", actionPayload);
      const message = data?.ok ? "Remote database updated" : (data?.message || "Remote update failed.");
      statusDialog.setMessage(message, data?.ok ? "ok" : "error");
      postStatus(`DFM sync: ${message}`, data?.ok ? "" : "warn");
      return;
    }
  } catch (err) {
    statusDialog.setMessage(String(err?.message || err), "error");
    postStatus(`DFM sync failed: ${String(err?.message || err)}`, "warn");
  } finally {
    statusDialog.setBusy(false);
  }
}

export async function startDfmRpcBridgeSync(buttonEl = null) {
  if (syncInFlight) return;
  syncInFlight = true;
  if (buttonEl) buttonEl.disabled = true;
  let cleanupPayload = null;
  let dialogClosed = false;
  let cleanupAfterClose = null;
  const cleanupAfterUserClose = () => {
    if (!cleanupPayload || cleanupAfterClose) return;
    cleanupAfterClose = cleanupRemoteTmp(cleanupPayload).finally(() => {
      cleanupAfterClose = null;
    });
    window.setTimeout(() => {
      if (cleanupPayload) cleanupRemoteTmp(cleanupPayload);
    }, 10000);
  };
  const dialog = createDfmRpcBridgeDialog({
    onClose: (reason) => {
      dialogClosed = true;
      if (reason === "primary-action") return;
      cleanupAfterUserClose();
    },
  });
  dialog.setWaiting("Preparing DFM RPC bridge sync...");
  try {
    const saved = await ensureSavedBeforeSync(dialog);
    if (!saved) return;

    const payload = buildRequestPayload();
    const missing = validatePayload(payload);
    if (missing.length) {
      dialog.setMessage(`Complete these Details fields before syncing: ${missing.join(", ")}.`, "error");
      return;
    }

    cleanupPayload = payload;
    dialog.setWaiting("Sending DFM request and waiting for remote JSON...");
    const data = await postJson("/dfm/rpc-bridge/sync", payload);
    if (dialogClosed) {
      await cleanupRemoteTmp(payload);
      return;
    }
    if (!data?.ok && data?.status === "timeout") {
      dialog.setMessage("Timed out waiting for remote DFM JSON. Use Refresh if the remote file appears later.", "warn");
      postStatus("DFM sync timed out waiting for remote JSON.", "warn");
      return;
    }
    dialog.setComparison(data, {
      labelFallbacks: buildCurrentPatternLabelFallbacks(),
      onRefresh: () => refreshComparison(dialog, payload),
      onPrimary: (action) => runPrimaryAction(dialog, payload, action),
    });
  } catch (err) {
    dialog.setMessage(String(err?.message || err), "error");
    postStatus(`DFM sync failed: ${String(err?.message || err)}`, "warn");
  } finally {
    syncInFlight = false;
    if (buttonEl) buttonEl.disabled = false;
  }
}
