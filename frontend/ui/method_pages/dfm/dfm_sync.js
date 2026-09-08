/*
===============================================================================
DFM Sync - BroadcastChannel cross-window synchronization
===============================================================================
*/
import {
  state,
  ratioStrikeSet,
  selectedSummaryByCol,
  ratioSyncSourceId,
  ratioSyncChannelName,
  getRatioSyncChannel,
  setRatioSyncChannel,
  getRatioSyncMuted,
  setRatioSyncMuted,
  getEffectiveDevLabelsForModel,
  getRatioHeaderLabels,
  ensureDefaultSummarySelectionForColumns,
  isRatiosTabVisible,
} from "/ui/method_pages/dfm/dfm_state.js";
import { renderRatioTable, setNotifyRatioStateChanged } from "/ui/method_pages/dfm/dfm_ratios_tab.js?v=20260907b";
import { renderResultsTable } from "/ui/method_pages/dfm/dfm_results_tab.js?v=20260907a";

function getRatioSyncPayload() {
  return {
    type: "ratio-sync-state",
    source: ratioSyncSourceId,
    ts: Date.now(),
    strikes: Array.from(ratioStrikeSet),
    selected: Array.from(selectedSummaryByCol.entries()),
  };
}

function applyRatioSyncPayload(payload) {
  if (!payload || payload.source === ratioSyncSourceId) return;
  if (!Array.isArray(payload.strikes) || !Array.isArray(payload.selected)) return;

  setRatioSyncMuted(true);
  try {
    ratioStrikeSet.clear();
    payload.strikes.forEach((key) => {
      if (typeof key === "string") ratioStrikeSet.add(key);
    });
    selectedSummaryByCol.clear();
    payload.selected.forEach((entry) => {
      if (!Array.isArray(entry) || entry.length < 2) return;
      const col = Number(entry[0]);
      const rowId = String(entry[1] || "");
      if (!Number.isFinite(col) || !rowId) return;
      selectedSummaryByCol.set(col, rowId);
    });
  } finally {
    setRatioSyncMuted(false);
  }

  const model = state.model;
  if (model) {
    const devs = getEffectiveDevLabelsForModel(model);
    const colCount = getRatioHeaderLabels(devs).length;
    ensureDefaultSummarySelectionForColumns(colCount);
  }

  if (isRatiosTabVisible()) renderRatioTable();
  if (document.getElementById("resultsWrap")) renderResultsTable();
}

export function notifyRatioStateChanged() {
  if (getRatioSyncMuted()) return;
  const ch = getRatioSyncChannel();
  if (ch) {
    ch.postMessage(getRatioSyncPayload());
  }
}

export function requestRatioStateSync() {
  const ch = getRatioSyncChannel();
  if (!ch) return;
  ch.postMessage({
    type: "ratio-sync-request",
    source: ratioSyncSourceId,
  });
}

export function wireRatioSyncChannel() {
  if (!window.BroadcastChannel || getRatioSyncChannel()) return;
  let ch;
  try {
    ch = new BroadcastChannel(ratioSyncChannelName);
  } catch {
    ch = null;
  }
  if (!ch) return;
  setRatioSyncChannel(ch);
  ch.addEventListener("message", (e) => {
    const data = e?.data;
    if (!data || data.source === ratioSyncSourceId) return;
    if (data.type === "ratio-sync-request") {
      notifyRatioStateChanged();
      return;
    }
    if (data.type === "ratio-sync-state") {
      applyRatioSyncPayload(data);
    }
  });

  // Inject notifyRatioStateChanged into ratios tab to avoid circular dependency
  setNotifyRatioStateChanged(notifyRatioStateChanged);
}
