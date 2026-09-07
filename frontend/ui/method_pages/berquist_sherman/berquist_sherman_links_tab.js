/*
===============================================================================
Berquist Sherman Links tab
Mounts the shared Links table for the Excel references the User Value rows of
the Case Reserve Adequacy "Avg. Selections" grids read. The page owns the link
records, the refresh, the break, and the dirty state; this module owns the
table and the freshness advisory above it, in the same words the DFM Links tab
uses so a stale workbook reads the same on every method page.
===============================================================================
*/
import { createLinksTab } from "/ui/shared/tabs/links/links_tab.js?v=20260907a";

function plural(count, noun) {
  return `${count} ${noun}${count === 1 ? "" : "s"}`;
}

export function excelFreshnessWarning(freshness) {
  const staleCount = Number(freshness?.staleCount || 0);
  const unverifiedCount = Number(freshness?.unverifiedCount || 0);
  const invalidCount = Number(freshness?.invalidCount || 0);
  if (!staleCount && !unverifiedCount && !invalidCount) return null;
  const parts = [];
  if (invalidCount) parts.push(plural(invalidCount, "broken reference"));
  if (staleCount) parts.push(plural(staleCount, "stale linked value"));
  if (unverifiedCount) parts.push(plural(unverifiedCount, "unverified linked value"));
  return {
    title: invalidCount ? "Excel links need attention" : "Saved Excel values may be out of date",
    detail: `${parts.join(" and ")}. Stored values remain active until you choose Refresh.`,
  };
}

export function createBerquistShermanLinksTab({
  container,
  displayLabel,
  getRecords,
  onRefresh,
  onBreak,
  onStatus,
}) {
  if (!container) return null;
  const controller = createLinksTab({
    container,
    ariaLabel: `${displayLabel} external links`,
    emptyDescription: "Excel links used by the User Value rows in the Avg. Selections view will appear here.",
    noun: "external links",
    getLinks: () => getRecords(),
    onRefreshLinks: (records) => onRefresh(records.map((record) => record?.id).filter(Boolean)),
    onBreakLinks: (records) => onBreak(records.map((record) => record?.id).filter(Boolean)),
    onStatus,
  });
  return {
    refresh: () => controller.refresh(),
    setFreshness(freshness) {
      const warning = excelFreshnessWarning(freshness);
      if (warning) controller.setWarning(warning.title, warning.detail);
      else controller.clearWarning();
    },
    destroy: () => controller.destroy(),
  };
}
