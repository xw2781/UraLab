import { createLinksTab } from "/ui/shared/tabs/links/links_tab.js?v=20260907a";
import {
  breakDfmExternalLinks,
  getDfmExternalLinkRecords,
  refreshAllExcelLinks,
} from "/ui/method_pages/dfm/dfm_ratios_summary_table.js?v=20260903a";

let dfmLinksController = null;
let linksChangedListener = null;
let dfmExcelFreshnessState = null;

function postDfmLinksStatus(message, tone = "") {
  const text = String(message || "").trim();
  if (!text) return;
  window.parent?.postMessage?.({
    type: "arcrho:status",
    text,
    ...(tone ? { tone } : {}),
  }, "*");
}

export function refreshDfmLinks() {
  return dfmLinksController?.refresh?.() || Promise.resolve(false);
}

function renderDfmExcelFreshnessWarning() {
  if (!dfmLinksController) return;
  const staleCount = Number(dfmExcelFreshnessState?.staleCount || 0);
  const unverifiedCount = Number(dfmExcelFreshnessState?.unverifiedCount || 0);
  const invalidCount = Number(dfmExcelFreshnessState?.invalidCount || 0);
  if (!staleCount && !unverifiedCount && !invalidCount) {
    dfmLinksController.clearWarning?.();
    return;
  }
  const parts = [];
  if (invalidCount) parts.push(`${invalidCount} broken reference${invalidCount === 1 ? "" : "s"}`);
  if (staleCount) parts.push(`${staleCount} stale linked value${staleCount === 1 ? "" : "s"}`);
  if (unverifiedCount) parts.push(`${unverifiedCount} unverified linked value${unverifiedCount === 1 ? "" : "s"}`);
  dfmLinksController.setWarning?.(
    invalidCount ? "Excel links need attention" : "Saved Excel values may be out of date",
    `${parts.join(" and ")}. Stored values remain active until you choose Refresh.`,
  );
}

export function setDfmExcelFreshnessState(nextState) {
  dfmExcelFreshnessState = nextState && typeof nextState === "object" ? { ...nextState } : null;
  renderDfmExcelFreshnessWarning();
}

export function initDfmLinks() {
  if (dfmLinksController) return dfmLinksController;
  const container = document.getElementById("dfmLinksMount");
  if (!container) return null;
  dfmLinksController = createLinksTab({
    container,
    ariaLabel: "DFM external links",
    emptyDescription: "Excel links used by User Entry cells in the Ratios tab will appear here.",
    noun: "external links",
    getLinks: () => getDfmExternalLinkRecords(),
    onRefreshLinks: async (records) => {
      const result = await refreshAllExcelLinks({
        source: "links-tab",
        sourceIds: records.map((record) => record?.id).filter(Boolean),
      });
      if (result?.ok !== false && !Number(result?.failedCount || 0)) {
        setDfmExcelFreshnessState(null);
      }
      return result;
    },
    onBreakLinks: (records) => breakDfmExternalLinks(
      records.map((record) => record?.id).filter(Boolean),
    ),
    onStatus: postDfmLinksStatus,
  });
  renderDfmExcelFreshnessWarning();
  linksChangedListener = () => {
    void refreshDfmLinks();
  };
  window.addEventListener("arcrho:dfm-links-changed", linksChangedListener);
  void refreshDfmLinks();
  return dfmLinksController;
}

export function destroyDfmLinks() {
  if (linksChangedListener) {
    window.removeEventListener("arcrho:dfm-links-changed", linksChangedListener);
    linksChangedListener = null;
  }
  dfmLinksController?.destroy?.();
  dfmLinksController = null;
}
