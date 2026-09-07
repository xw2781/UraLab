import { mountDatasetViewer } from "/ui/dataset_viewer/dataset_viewer_view.js?v=20260906b";
import { configureDataTabHost } from "/ui/shared/tabs/data/data_tab_context.js";
import { configureDataTabChart } from "/ui/shared/tabs/data/data_tab_chart_port.js";
import { configureDataTabNotes } from "/ui/shared/tabs/data/data_tab_notes_port.js";
import { configureDataTabPageHost } from "/ui/shared/tabs/data/data_tab_page_host_port.js";
import { configureDataTabAudit } from "/ui/shared/tabs/data/data_tab_audit_port.js";
import { configureDataTabCloseConfirm } from "/ui/shared/tabs/data/data_tab_close_port.js";
import { createPageCloseConfirm } from "/ui/shared/components/close_confirm/close_confirm.js";
import { createAuditLogView } from "/ui/shared/tabs/audit_log/audit_log_view.js?v=20260714c";
import {
  formatSidecarAuditEventDate,
  normalizeSidecarAuditEntries,
} from "/ui/shared/tabs/audit_log/sidecar_audit_entries.js?v=20260714c";
import {
  applyTabbedPageSaveBar,
  createTabbedPage,
} from "/ui/shared/tabbed_page/tabbed_page.js?v=20260816a";
import { wireTabPopoutWindows } from "/ui/shared/tabbed_page/tab_popout_window.js?v=20260722a";
import {
  redrawDatasetChartSafely,
  renderDatasetChart,
} from "/ui/dataset_viewer/tabs/dataset_chart_tab.js?v=20260805a";
import { wireDatasetNotesEditor } from "/ui/dataset_viewer/tabs/dataset_notes_tab.js?v=20260715a";
import { createLinksTab } from "/ui/shared/tabs/links/links_tab.js?v=20260901c";
import { configureDataTabLinks } from "/ui/shared/tabs/data/data_tab_links_port.js";
import { configureDataTabChangeWatch } from "/ui/shared/tabs/data/data_tab_change_watch_port.js?v=20260806a";
import {
  createObjectChangeWatch,
  showObjectUpdatedAlert,
  wireSamePropagationScopePause,
} from "/ui/shared/services/object_change_watch.js?v=20260820a";
import { showPageMessageBox } from "/ui/shared/components/message_box/message_box.js?v=20260831a";
import { state as sharedDatasetState } from "/ui/shared/dataset/dataset_state.js";
import { DATASET_VIEWER_TAB_DEFS as DATASET_VIEWER_TABS } from "/ui/shared/tabs/window_tab_catalog.js?v=20260903a";

function mountDatasetViewerTabs({
  initialTab,
  onDetailsActivated,
  onChartActivated,
  wireDataTabTopBarToggle,
} = {}) {
  const handleChartLayout = (tabId) => {
    if (tabId === "chart") onChartActivated?.();
  };
  const tabSystem = createTabbedPage(document.body, {
    tabs: DATASET_VIEWER_TABS,
    cssPrefix: "ds",
    initialTab,
    injectTabBar: false,
    onTabChange: (tabId) => {
      if (tabId === "details") onDetailsActivated?.();
      if (tabId === "chart") onChartActivated?.();
    },
  });
  applyTabbedPageSaveBar(document.getElementById("datasetSaveBar"));
  window.dsTabSystem = tabSystem;
  wireDataTabTopBarToggle?.(tabSystem);
  wireTabPopoutWindows({
    cssPrefix: "ds",
    tabs: DATASET_VIEWER_TABS,
    tabSystem: () => window.dsTabSystem,
    onPopoutTab: handleChartLayout,
    onDockTab: handleChartLayout,
    onFocusTab: handleChartLayout,
    onLayout: handleChartLayout,
  });
  return tabSystem;
}

function wireDatasetChangeWatch() {
  const params = new URLSearchParams(window.location.search);
  const projectName = (params.get("project") || "").trim();
  const reservingClass = (params.get("path") || "").trim();
  const instanceName = (params.get("instance_name") || params.get("tri") || "").trim();
  const isDurableInstance = params.get("temporary_view") !== "1"
    && params.get("draft_instance") !== "1";
  if (!projectName || !reservingClass || !instanceName || !isDurableInstance) return;
  const changeWatch = createObjectChangeWatch({
    identity: {
      project_name: projectName,
      reserving_class: reservingClass,
      kind: "dataset",
      name: instanceName,
    },
    onChange: (attribution) => {
      void showObjectUpdatedAlert({
        showMessageBox: showPageMessageBox,
        attribution,
        isDirty: () => sharedDatasetState.dirty.size > 0,
        onBlockedRefresh: () => {
          window.parent?.postMessage?.({
            type: "arcrho:status",
            text: "Unsaved grid changes block the refresh. Save or discard them, then reopen the window.",
            tone: "warn",
          }, "*");
        },
      });
    },
  });
  wireSamePropagationScopePause({
    watch: changeWatch,
    getProject: () => projectName,
    getReservingClass: () => reservingClass,
  });
  let watchStarted = false;
  configureDataTabChangeWatch({
    onMutationStarted: () => changeWatch.pause(),
    onMutationEnded: () => { void changeWatch.resume(); },
    onDurableDatasetState: () => {
      changeWatch.noteSelfWrite(sharedDatasetState.sidecarUpdatedAt);
      if (!watchStarted) {
        watchStarted = true;
        changeWatch.start();
        return;
      }
      void changeWatch.rebase();
    },
  });
}

wireDatasetChangeWatch();
mountDatasetViewer(document.getElementById("datasetRoot"));
configureDataTabAudit(createAuditLogView({
  container: document.getElementById("datasetAuditLogMount"),
  ariaLabel: "Dataset audit log",
  emptyDescription: "Dataset changes will appear here after the first save.",
  normalizeEntries: normalizeSidecarAuditEntries,
  formatEventDate: formatSidecarAuditEventDate,
}));
configureDataTabCloseConfirm(createPageCloseConfirm({ subject: "dataset" }));
configureDataTabHost("dataset_viewer");
configureDataTabChart({
  renderChart: renderDatasetChart,
  redrawChartSafely: redrawDatasetChartSafely,
});
configureDataTabNotes({ mountNotes: wireDatasetNotesEditor });
configureDataTabPageHost(mountDatasetViewerTabs);

const datasetDataTab = await import(
  "/ui/shared/tabs/data/data_tab_controller.js?v=20260906b"
);

const postLinksStatus = (message, tone = "") => {
  if (!message) return;
  window.parent?.postMessage?.({
    type: "arcrho:status",
    text: message,
    ...(tone ? { tone } : {}),
  }, "*");
};

// One table lists every link the Data tab holds - Excel, ArcRho, and formula -
// and the Data tab routes each row back to the controller that owns it.
const datasetLinksTab = createLinksTab({
  container: document.getElementById("datasetLinksMount"),
  ariaLabel: "Dataset links",
  emptyDescription: "Links used by editable cells in the Data tab will appear here.",
  getLinks: () => datasetDataTab.getDatasetLinkRecords(),
  onRefreshLinks: (records) => datasetDataTab.refreshDatasetLinkRecords(records),
  onBreakLinks: (records) => datasetDataTab.breakDatasetLinks(records),
  onOpenDataset: (record) => {
    const params = new URLSearchParams(window.location.search);
    window.parent?.postMessage?.({
      type: "arcrho:project-instance-open-dependent-dataset",
      datasetName: record?.datasetName,
      reservingClass: (params.get("path") || "").trim(),
      projectName: (params.get("project") || "").trim(),
    }, "*");
    return { ok: true, message: `Opening dataset ${record?.datasetName || ""}...` };
  },
  onStatus: postLinksStatus,
});

configureDataTabLinks({ refresh: () => datasetLinksTab.refresh() });

window.ADA_DATASET_READY = datasetDataTab.bootDatasetDataTab();

// The window stays blank until the opening tab has its content, so the grid,
// the top row and the tab frame all appear together. The rethrow keeps a boot
// failure reported exactly as it was before.
void window.ADA_DATASET_READY.then(
  () => window.arcrhoRevealPage?.(),
  (err) => {
    window.arcrhoRevealPage?.();
    throw err;
  },
);
