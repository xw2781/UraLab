import { syncDetailsLabelWidth } from "/ui/shared/tabs/details/details_form_layout.js?v=20260820b";
import { applyHostFixedDetailsFields } from "/ui/shared/tabs/details/details_host_fields.js?v=20260820b";
import { attachArcrhoTooltip } from "/ui/shared/components/tooltip/tooltip.js?v=20260812a";

export function mountDatasetViewer(container) {
  if (!container) return null;
  if (container.querySelector("#topFrame")) return container;
  const wrapper = document.createElement("div");
  wrapper.innerHTML = `<!-- Tab bar -->
  <div class="dsTabBar tabbedPageTabBar">
    <button class="dsTab tabbedPageTab" data-page="details" type="button">Details</button>
    <button class="dsTab tabbedPageTab active" data-page="data" type="button">Data</button>
    <button class="dsTab tabbedPageTab" data-page="chart" type="button">Chart</button>
    <button class="dsTab tabbedPageTab" data-page="notes" type="button">Notes</button>
    <button class="dsTab tabbedPageTab" data-page="links" type="button">Links</button>
    <button class="dsTab tabbedPageTab" data-page="auditLog" type="button">Audit Log</button>
  </div>

  <!-- Details tab page: identity first, then what this dataset is computed from
       and what consumes it. Name is the first row and Dataset Type the second,
       so every Details tab in the app opens the same way. -->
  <div id="dsDetailsPage" data-page="details" class="arDetailsRoot" style="display:none;">
    <div id="topFrame">
      <div class="arDetailsSection">
        <div class="arDetailsGrid">
          <div class="dsDetailLabel arDetailsLabel">
            <label for="dsDetailName">Name : </label>
          </div>
          <div class="dsDetailInput arDetailsField">
            <div class="dsDetailNameWrap">
              <input id="dsDetailName" class="arDetailsControl" autocomplete="off" />
              <span id="dsDetailNameWarning" class="dsDetailNameWarning" role="tooltip" aria-live="polite" hidden></span>
            </div>
          </div>

          <div class="dsDetailLabel arDetailsLabel">
            <label for="triInput">Dataset Type : </label>
          </div>
          <div class="dsDetailInput arDetailsField">
            <div class="datasetSelectWrap">
              <input id="triInput" class="arDetailsControl" autocomplete="off" />
              <button id="datasetTreeBtn" type="button" class="datasetTreeBtn" title="Browse dataset types" aria-label="Browse dataset types">...</button>
              <div id="datasetDropdown" class="datasetDropdown"></div>
            </div>
          </div>

          <div class="dsDetailLabel arDetailsLabel" data-details-field="project">
            <label for="projectSelect">Project Name : </label>
          </div>
          <div class="dsDetailInput arDetailsField" data-details-field="project">
            <div class="projectSelectWrap">
              <input id="projectSelect" class="arDetailsControl" autocomplete="off" />
              <button id="projectTreeBtn" type="button" class="projectTreeBtn" title="Browse project folders" aria-label="Browse project folders">
                ...
              </button>
              <div id="projectDropdown" class="projectDropdown"></div>
            </div>
          </div>

          <div class="dsDetailLabel arDetailsLabel">
            <label for="pathInput">Segment : </label>
          </div>
          <div class="dsDetailInput arDetailsField">
            <input id="pathInput" class="arDetailsControl" readonly />
          </div>
        </div>
      </div>

      <div class="arDetailsSection">
        <div class="arDetailsGrid">
          <div class="dsDetailLabel arDetailsLabel">
            <label id="dsFormulaLabel" for="dsDetailFormulaBox">Formula : </label>
          </div>
          <div class="dsDetailInput arDetailsField">
            <div id="dsDetailFormulaBox" class="arDetailsFormulaBox" role="group" aria-labelledby="dsFormulaLabel"></div>
            <textarea id="dsDetailFormula" autocomplete="off" readonly rows="1" tabindex="-1" aria-hidden="true"></textarea>
          </div>

          <div class="dsDetailLabel arDetailsLabel">
            <label id="dsPrecedentsTitle">Precedents : </label>
          </div>
          <div class="dsDetailInput arDetailsField">
            <div class="arDetailsChipBox">
              <div id="dsPrecedentsList" class="arDetailsChipList" aria-live="polite"></div>
            </div>
          </div>

          <div class="dsDetailLabel arDetailsLabel">
            <label id="dsDependentsTitle">Dependents : </label>
          </div>
          <div class="dsDetailInput arDetailsField">
            <div class="arDetailsChipBox">
              <div id="dsDependentsList" class="arDetailsChipList" aria-live="polite"></div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- Data tab page: parameter strip + formula bar + triangle table -->
  <div id="dsDataPage" data-page="data">
    <!-- parameter strip -->
    <div class="topRow">
      <div class="panel" id="datasetTopBar">
        <div class="topbar-grid">
          <!-- Col 1: dataset display and orientation controls -->
          <div class="topbar-left" style="grid-column: 1; grid-row: 1 / span 2;">
            <label class="chk"><span>Cumulative:</span> <input id="cumulativeChk" type="checkbox" checked /></label>
            <label class="chk"><span>Transposed:</span> <input id="transposedChk" type="checkbox" /></label>
            <div class="timeModeFrame" role="group" aria-label="Time mode">
              <label class="rad">
                <input type="radio" name="timeMode" value="development" checked />
                <span>Development</span>
              </label>
              <label class="rad">
                <input type="radio" name="timeMode" value="calendar" />
                <span>Calendar</span>
              </label>
            </div>
          </div>

          <!-- Col 2: Labels -->
          <div class="topbar-label-stack" style="grid-column: 2; grid-row: 1 / span 2;">
            <div class="topbar-label"><span class="lbl">Origin Length:</span></div>
            <div class="topbar-label"><span class="lbl">Development Length:</span></div>
          </div>

          <!-- Col 3: Inputs -->
          <div class="topbar-input-stack" style="grid-column: 3; grid-row: 1 / span 2;">
            <div class="topbar-input">
              <div id="originLenWrap" class="lenSelectWrap">
                <button
                  id="originLenDisplay"
                  class="lenSelectDisplay"
                  type="button"
                  aria-haspopup="listbox"
                  aria-expanded="false"
                  aria-controls="originLenDropdown"
                >
                  <span class="lenSelectValue">12</span>
                  <span class="lenSelectCaret" aria-hidden="true"></span>
                </button>
                <div id="originLenDropdown" class="datasetDropdown lenDropdown" role="listbox" aria-label="Origin Length options"></div>
                <select id="originLenSelect" class="lenSelectNative" tabindex="-1" aria-hidden="true"></select>
              </div>
              <span class="lenStoredLabel">Stored at:</span>
              <div id="originStoredLenWrap" class="lenSelectWrap">
                <button
                  id="originStoredLenDisplay"
                  class="lenSelectDisplay"
                  type="button"
                  aria-haspopup="listbox"
                  aria-expanded="false"
                  aria-controls="originStoredLenDropdown"
                >
                  <span class="lenSelectValue">12</span>
                  <span class="lenSelectCaret" aria-hidden="true"></span>
                </button>
                <div id="originStoredLenDropdown" class="datasetDropdown lenDropdown" role="listbox" aria-label="Origin Stored at options"></div>
                <select id="originStoredLenSelect" class="lenSelectNative" tabindex="-1" aria-hidden="true"></select>
              </div>
            </div>
            <div class="topbar-input">
              <div id="devLenWrap" class="lenSelectWrap">
                <button
                  id="devLenDisplay"
                  class="lenSelectDisplay"
                  type="button"
                  aria-haspopup="listbox"
                  aria-expanded="false"
                  aria-controls="devLenDropdown"
                >
                  <span class="lenSelectValue">12</span>
                  <span class="lenSelectCaret" aria-hidden="true"></span>
                </button>
                <div id="devLenDropdown" class="datasetDropdown lenDropdown" role="listbox" aria-label="Development Length options"></div>
                <select id="devLenSelect" class="lenSelectNative" tabindex="-1" aria-hidden="true"></select>
              </div>
              <span class="lenStoredLabel">Stored at:</span>
              <div id="devStoredLenWrap" class="lenSelectWrap">
                <button
                  id="devStoredLenDisplay"
                  class="lenSelectDisplay"
                  type="button"
                  aria-haspopup="listbox"
                  aria-expanded="false"
                  aria-controls="devStoredLenDropdown"
                >
                  <span class="lenSelectValue">12</span>
                  <span class="lenSelectCaret" aria-hidden="true"></span>
                </button>
                <div id="devStoredLenDropdown" class="datasetDropdown lenDropdown" role="listbox" aria-label="Development Stored at options"></div>
                <select id="devStoredLenSelect" class="lenSelectNative" tabindex="-1" aria-hidden="true"></select>
              </div>
            </div>
          </div>

          <!-- Col 4: Number formatting labels -->
          <div class="topbar-format-label-stack" style="grid-column: 4; grid-row: 1 / span 2;">
            <div class="topbar-label"><span class="lbl">Number Format:</span></div>
            <div class="topbar-label"><span class="lbl">Decimal Places:</span></div>
          </div>

          <!-- Col 5: Number formatting inputs -->
          <div class="topbar-format-input-stack" style="grid-column: 5; grid-row: 1 / span 2;">
            <div class="topbar-input">
              <div id="numberFormatWrap" class="arNumberFormatField">
                <input id="numberFormatSelect" type="text" value="0,000" aria-label="Number Format" aria-controls="numberFormatDropdown" aria-expanded="false" autocomplete="off" />
                <button id="numberFormatDropdownBtn" class="arNumberFormatToggle" type="button" aria-label="Show Number Format presets" aria-controls="numberFormatDropdown" aria-expanded="false">
                  <span class="arNumberFormatCaret" aria-hidden="true"></span>
                </button>
                <div id="numberFormatDropdown" class="datasetDropdown arNumberFormatMenu" role="listbox" aria-label="Number Format presets"></div>
              </div>
            </div>
            <div class="topbar-input">
              <div id="decimalPlacesWrap" class="decimalPlacesWrap">
                <input id="decimalPlaces" type="number" min="0" max="6" value="1" aria-label="Decimal Places" />
                <div class="decimalPlacesStepper">
                  <button id="decimalPlacesUpBtn" class="decimalPlacesStepBtn" type="button" aria-label="Increase Decimal Places">
                    <span class="datasetStepperCaret datasetStepperCaretUp" aria-hidden="true"></span>
                  </button>
                  <button id="decimalPlacesDownBtn" class="decimalPlacesStepBtn" type="button" aria-label="Decrease Decimal Places">
                    <span class="datasetStepperCaret" aria-hidden="true"></span>
                  </button>
                </div>
              </div>
            </div>
          </div>

        </div>
      </div>
      <button
        id="clearCacheReloadBtn"
        type="button"
        aria-label="Clear cache and reload current dataset"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
          <path d="M20 6v5h-5"></path>
          <path d="M18.6 9A7 7 0 1 0 19 15"></path>
        </svg>
      </button>
    </div>

    <!-- Triangle -->
    <div class="panel" id="triPanel">
      <div id="tableWrapHost">
        <div id="tableWrap"></div>
      </div>
    </div>
  </div>

  <!-- Chart tab page -->
  <div id="dsChartPage" data-page="chart" style="display:none;">
    <div class="right">
      <div class="panel" id="chartPanel">
        <div class="panelInner">
          <div class="chartHeader">
            <span class="small"><b id="chartTitle">Development Curves</b></span>
            <div class="chartToggle" id="chartModeToggle">
              <button class="chartToggleBtn active" data-mode="byCol" title="By Column (Dev Period)">By Column</button>
              <button class="chartToggleBtn" data-mode="byRow" title="By Row (Origin)">By Row</button>
            </div>
          </div>
          <div class="chartRow">
            <div class="chartCanvasWrap">
              <canvas id="devChart"></canvas>
            </div>
            <aside class="arChartLegendPanel" aria-labelledby="devChartLegendTitle">
              <div class="arChartLegendHeader">
                <h3 class="arChartLegendTitle" id="devChartLegendTitle">Series</h3>
                <span class="arChartLegendCount" id="devChartLegendCount"></span>
              </div>
              <div class="arChartLegendList" id="devChartLegend" role="group" aria-label="Chart series visibility"></div>
            </aside>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- Notes tab page -->
  <div id="dsNotesPage" data-page="notes" style="display:none;">
    <div class="dsNotesEditorWrap">
      <div id="datasetNotesMount"></div>
      <div class="dsNotesToolbar" id="dsNotesToolbar">
        <div class="dsNotesActions">
          <span id="dsNotesSaveState" class="small dsNotesSaveState">Not saved</span>
        </div>
      </div>
    </div>
  </div>

  <!-- Links tab page: one table for Excel, ArcRho, and formula links -->
  <div id="dsLinksPage" data-page="links" style="display:none;">
    <div id="datasetLinksMount"></div>
  </div>

  <!-- Audit Log tab page -->
  <div id="dsAuditLogPage" data-page="auditLog" style="display:none;">
    <div id="datasetAuditLogMount"></div>
  </div>

  <div id="datasetSaveBar" class="datasetSaveBar" hidden>
    <button id="datasetSaveBtn" class="datasetPrimaryBtn" type="button">Save</button>
    <button id="datasetCancelBtn" class="datasetSecondaryBtn" type="button">Cancel</button>
  </div>

  <div id="hiddenControls" style="display:none;">
    <div class="small" id="dsMeta"></div>
    <button id="saveBtn">Save</button>
    <button id="toggleBlankBtn">Show blanks</button>
    <pre id="log"></pre>
  </div>

  <div id="ctxMenu" class="ctx-menu" role="menu" style="display:none;">
    <div class="ctx-menu-inner">
      <button class="ctx-item" data-action="copy_value">Copy values</button>
      <button class="ctx-item" data-action="paste">Paste</button>
      <button class="ctx-item" data-action="remove_highlights">Remove Highlights</button>
      <div class="ctx-sep"></div>
      <button class="ctx-item" data-action="toggle_subtotal">Show/Hide subtotal</button>
      <div class="ctx-sep"></div>
      <button class="ctx-item" data-action="export_data">Export data</button>
    </div>
  </div>

  <!-- Same-folder JS entrypoint (no /static) -->
  <!--  -->`;
  // The Data tab is the panel this markup pre-selects, but the tab system does
  // not choose the real one until the Data tab's first reads finish. Apply the
  // requested tab to the fragment first, so the window paints on it once
  // instead of opening on Data and jumping.
  window.arcrhoApplyInitialTabbedPage?.(wrapper);
  while (wrapper.firstElementChild) {
    container.appendChild(wrapper.firstElementChild);
  }
  attachArcrhoTooltip(
    container.querySelector("#clearCacheReloadBtn"),
    "Clear cache and reload current dataset",
  );
  syncDatasetDetailsLabelWidth(container);
  wireTableScrollbarActivity(container);
  return container;
}

function syncDatasetDetailsLabelWidth(container) {
  const detailsPage = container?.querySelector?.("#dsDetailsPage");
  applyHostFixedDetailsFields({ root: detailsPage });
  syncDetailsLabelWidth({
    root: detailsPage,
    labelSelector: ".arDetailsLabel",
  });
}

function wireTableScrollbarActivity(container) {
  const wraps = Array.from(container?.querySelectorAll?.("#tableWrap") || []);
  wraps.forEach((wrap) => {
    if (wrap.__arcRhoScrollbarActivityWired) return;
    wrap.__arcRhoScrollbarActivityWired = true;

    let idleTimer = null;
    const syncScrollbarHover = (event) => {
      const rect = wrap.getBoundingClientRect();
      const verticalScrollbarWidth = Math.max(0, wrap.offsetWidth - wrap.clientWidth);
      const horizontalScrollbarHeight = Math.max(0, wrap.offsetHeight - wrap.clientHeight);
      const hasVerticalScrollbar = wrap.scrollHeight > wrap.clientHeight && verticalScrollbarWidth > 0;
      const hasHorizontalScrollbar = wrap.scrollWidth > wrap.clientWidth && horizontalScrollbarHeight > 0;
      const nearVerticalScrollbar = hasVerticalScrollbar
        && event.clientX >= rect.right - Math.max(verticalScrollbarWidth, 16);
      const nearHorizontalScrollbar = hasHorizontalScrollbar
        && event.clientY >= rect.bottom - Math.max(horizontalScrollbarHeight, 16);

      wrap.classList.toggle("isScrollbarHover", nearVerticalScrollbar || nearHorizontalScrollbar);
    };

    wrap.addEventListener("scroll", () => {
      wrap.classList.add("isScrolling");
      if (idleTimer) clearTimeout(idleTimer);
      idleTimer = setTimeout(() => {
        wrap.classList.remove("isScrolling");
      }, 550);
    }, { passive: true });
    wrap.addEventListener("pointermove", syncScrollbarHover, { passive: true });
    wrap.addEventListener("pointerleave", () => {
      wrap.classList.remove("isScrollbarHover");
    }, { passive: true });
  });
}
