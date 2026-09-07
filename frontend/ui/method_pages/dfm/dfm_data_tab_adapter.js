import { configureDataTabHost } from "/ui/shared/tabs/data/data_tab_context.js";
import { configureDataTabHostPublisher } from "/ui/shared/tabs/data/data_tab_host_port.js";

function publishDfmInputHelpers(dependencies) {
  const {
    getResolvedProjectValue,
    getResolvedReservingClassValue,
    getDisplayProjectValue,
    getDisplayReservingClassValue,
    getDisplayTriValue,
    isInputDefaultBound,
  } = dependencies;

  window.ADA_GET_DFM_INPUTS = () => ({
    resolved: {
      project: getResolvedProjectValue(),
      reservingClass: getResolvedReservingClassValue(),
      tri: getDisplayTriValue(),
    },
    display: {
      project: getDisplayProjectValue(),
      reservingClass: getDisplayReservingClassValue(),
      tri: getDisplayTriValue(),
    },
    defaults: {
      projectDefault: isInputDefaultBound(document.getElementById("projectSelect")),
      reservingClassDefault: isInputDefaultBound(document.getElementById("pathInput")),
    },
  });
}

configureDataTabHost("dfm");
configureDataTabHostPublisher(publishDfmInputHelpers);

const { bootDatasetDataTab } = await import(
  "/ui/shared/tabs/data/data_tab_controller.js?v=20260907e"
);

window.ADA_DATASET_READY = bootDatasetDataTab();
