import assert from "node:assert/strict";
import { access, readFile, readdir } from "node:fs/promises";
import test from "node:test";

const frontendRoot = new URL("../", import.meta.url);

async function source(relativePath) {
  return readFile(new URL(relativePath, frontendRoot), "utf8");
}

async function runtimeSources(relativeDirectory) {
  const sources = [];

  async function visit(directoryUrl) {
    const entries = await readdir(directoryUrl, { withFileTypes: true });
    for (const entry of entries) {
      const entryUrl = new URL(entry.name, directoryUrl);
      if (entry.isDirectory()) {
        await visit(new URL(`${entry.name}/`, directoryUrl));
      } else if (/\.(?:html|js)$/u.test(entry.name)) {
        sources.push(await readFile(entryUrl, "utf8"));
      }
    }
  }

  await visit(new URL(relativeDirectory, frontendRoot));
  return sources.join("\n");
}

test("shared tab surfaces live in feature-neutral logical groups", async () => {
  const requiredFiles = [
    "ui/shared/tabbed_page/tabbed_page.js",
    "ui/shared/tabbed_page/tabbed_page.css",
    "ui/shared/tabbed_page/tab_popout_window.js",
    "ui/shared/tabs/details/details_form_layout.js",
    "ui/shared/tabs/details/details_form_layout.css",
    "ui/shared/tabs/notes/notes_tab.js",
    "ui/shared/tabs/notes/notes_tab.css",
    "ui/shared/tabs/links/links_tab.js",
    "ui/shared/tabs/links/links_tab.css",
    "ui/shared/tabs/audit_log/audit_log_view.js",
    "ui/shared/tabs/audit_log/audit_log.css",
    "ui/shared/tabs/data/data_tab_controller.js",
    "ui/shared/tabs/data/data_tab_host_controller.js",
    "ui/shared/tabs/data/data_tab_details_controller.js",
    "ui/shared/tabs/data/data_tab_inputs_controller.js",
    "ui/shared/tabs/data/data_tab_preferences_controller.js",
    "ui/shared/tabs/data/data_tab_request_controller.js",
    "ui/shared/tabs/data/data_tab_persistence_controller.js",
    "ui/shared/tabs/data/data_tab_controls.js",
    "ui/shared/tabs/data/data_tab_dom.js",
    "ui/shared/tabs/data/data_tab_context.js",
    "ui/shared/tabs/data/data_tab_links_port.js",
    "ui/shared/tabs/data/dataset_grid_view.js",
    "ui/shared/tabs/data/dataset_grid_interactions.js",
    "ui/shared/tabs/data/data_tab.css",
    "ui/shared/components/workspace/workspace.css",
    "ui/shared/components/pickers/dataset_name_picker.js",
    "ui/shared/components/context_menu/context_menu.js",
    "ui/shared/components/spreadsheet/spreadsheet_table.js",
    "ui/shared/components/spreadsheet/spreadsheet_table.css",
    "ui/shared/dataset/dataset_api.js",
    "ui/shared/dataset/dataset_external_links.js",
    "ui/shared/dataset/dataset_state.js",
    "ui/shared/dataset/dataset_origin_labels.js",
    "ui/shared/dataset/dataset_types_source.js",
    "ui/dataset_viewer/dataset_viewer.html",
    "ui/dataset_viewer/dataset_viewer_main.js",
    "ui/dataset_viewer/dataset_viewer_view.js",
    "ui/dataset_viewer/dataset_viewer.css",
    "ui/method_pages/dfm/dfm.html",
    "ui/method_pages/dfm/dfm_data_tab_adapter.js",
    "ui/method_pages/dfm/dfm_external_links_model.js",
    "ui/method_pages/dfm/dfm_links_tab.js",
    "ui/method_pages/bornhuetter_ferguson/bornhuetter_ferguson.html",
    "ui/method_pages/cape_cod/cape_cod.html",
    "ui/method_pages/result_selection/result_selection.html",
  ];

  await Promise.all(requiredFiles.map((path) => access(new URL(path, frontendRoot))));
});

test("Dataset Viewer and method pages consume feature-neutral shared styling", async () => {
  const [datasetHtml, dfmHtml, bornhuetterFergusonHtml, capeCodHtml, resultSelectionHtml] = await Promise.all([
    source("ui/dataset_viewer/dataset_viewer.html"),
    source("ui/method_pages/dfm/dfm.html"),
    source("ui/method_pages/bornhuetter_ferguson/bornhuetter_ferguson.html"),
    source("ui/method_pages/cape_cod/cape_cod.html"),
    source("ui/method_pages/result_selection/result_selection.html"),
  ]);

  for (const html of [datasetHtml, dfmHtml, bornhuetterFergusonHtml, capeCodHtml, resultSelectionHtml]) {
    assert.match(html, /\/ui\/shared\/tabbed_page\/tabbed_page\.css/u);
  }
  assert.match(capeCodHtml, /\/ui\/shared\/components\/spreadsheet\/spreadsheet_table\.css/u);
  for (const html of [datasetHtml, dfmHtml]) {
    assert.match(html, /\/ui\/shared\/components\/workspace\/workspace\.css/u);
    assert.match(html, /\/ui\/shared\/components\/spreadsheet\/spreadsheet_table\.css/u);
    assert.match(html, /\/ui\/shared\/tabs\/data\/data_tab\.css/u);
    assert.match(html, /\/ui\/shared\/tabs\/links\/links_tab\.css/u);
  }
  assert.match(datasetHtml, /\/ui\/dataset_viewer\/dataset_viewer\.css/u);
  assert.match(resultSelectionHtml, /\/ui\/shared\/styles\/scrollbars\.css/u);
  assert.match(resultSelectionHtml, /\/ui\/shared\/components\/spreadsheet\/spreadsheet_table\.css/u);
  assert.doesNotMatch(resultSelectionHtml, /\/ui\/shared\/tabs\/data\/data_tab\.css/u);
});

test("DSV and DFM place reusable Links tabs immediately after Notes", async () => {
  const [
    datasetView,
    datasetMain,
    dfmHtml,
    dfmConfig,
    tabCatalog,
    dfmLinks,
    dfmSummary,
    dataController,
  ] = await Promise.all([
    source("ui/dataset_viewer/dataset_viewer_view.js"),
    source("ui/dataset_viewer/dataset_viewer_main.js"),
    source("ui/method_pages/dfm/dfm.html"),
    source("ui/method_pages/dfm/dfm_tab_config.js"),
    source("ui/shared/tabs/window_tab_catalog.js"),
    source("ui/method_pages/dfm/dfm_links_tab.js"),
    Promise.all([
      source("ui/method_pages/dfm/dfm_ratios_summary_table.js"),
      runtimeSources("ui/method_pages/dfm/ratios_summary/"),
    ]).then((parts) => parts.join("\n")),
    runtimeSources("ui/shared/tabs/data/"),
  ]);

  for (const sourceText of [datasetView, datasetMain, dfmHtml]) {
    assert.match(sourceText, /notes[\s\S]*links[\s\S]*(?:auditLog|audit)/u);
  }
  // Both tab lists now live in the shared window tab catalog, so the order is
  // asserted there and the DFM page module only has to read from it.
  assert.match(dfmConfig, /window_tab_catalog\.js/u);
  for (const listName of ["DATASET_VIEWER_TAB_DEFS", "DFM_TAB_DEFS", "BERQUIST_SHERMAN_TAB_DEFS"]) {
    const block = tabCatalog.split(`export const ${listName}`)[1]?.split("]);")[0] || "";
    assert.match(block, /notes[\s\S]*links[\s\S]*(?:auditLog|audit)/u, `${listName} keeps Links after Notes`);
  }
  assert.match(datasetMain, /shared\/tabs\/links\/links_tab\.js/u);
  assert.match(dfmLinks, /shared\/tabs\/links\/links_tab\.js/u);
  // Berquist Sherman mounts the same shared table for its User Value links.
  const [bsHtml, bsLinks] = await Promise.all([
    source("ui/method_pages/berquist_sherman/berquist_sherman.html"),
    source("ui/method_pages/berquist_sherman/berquist_sherman_links_tab.js"),
  ]);
  assert.match(bsHtml, /notes[\s\S]*links[\s\S]*audit/u);
  assert.match(bsHtml, /\/ui\/shared\/tabs\/links\/links_tab\.css/u);
  assert.match(bsLinks, /shared\/tabs\/links\/links_tab\.js/u);
  assert.match(dfmSummary, /shared\/integrations\/excel_reference\.js/u);
  assert.match(dfmSummary, /\bgetDfmExternalLinkRecords\b/u);
  assert.match(dfmSummary, /\bbreakDfmExternalLink\b/u);
  assert.match(dataController, /shared\/dataset\/dataset_external_links\.js/u);
  assert.match(dataController, /external_links:\s*runtime\.datasetExternalLinks\.serialize\(\)/u);
});

test("DSV exposes cache refresh as an accessible SVG action in the Data top row", async () => {
  const [datasetView, dataControls] = await Promise.all([
    source("ui/dataset_viewer/dataset_viewer_view.js"),
    source("ui/shared/tabs/data/data_tab_controls.js"),
  ]);

  assert.match(
    datasetView,
    /class="topRow"[\s\S]*id="datasetTopBar"[\s\S]*id="clearCacheReloadBtn"[\s\S]*<svg[\s\S]*<\/svg>[\s\S]*<\/button>/u,
  );
  assert.match(datasetView, /aria-label="Clear cache and reload current dataset"/u);
  assert.match(datasetView, /attachArcrhoTooltip\([\s\S]*#clearCacheReloadBtn[\s\S]*Clear cache and reload current dataset/u);
  assert.match(dataControls, /clearCacheReloadBtn[\s\S]*runArcRhoTri\(\{ clearCache: true/u);
});

test("DSV and DFM reach the current shared Data validation runtime", async () => {
  const [datasetHtml, datasetMain, dfmHtml, dfmAdapter, dataController] = await Promise.all([
    source("ui/dataset_viewer/dataset_viewer.html"),
    source("ui/dataset_viewer/dataset_viewer_main.js"),
    source("ui/method_pages/dfm/dfm.html"),
    source("ui/method_pages/dfm/dfm_data_tab_adapter.js"),
    source("ui/shared/tabs/data/data_tab_controller.js"),
  ]);

  assert.match(datasetHtml, /dataset_viewer_main\.js\?v=20260907c/u);
  assert.match(datasetMain, /data_tab_controller\.js\?v=20260907e/u);
  assert.match(dfmHtml, /dfm_data_tab_adapter\.js\?v=20260907c/u);
  assert.match(dfmAdapter, /data_tab_controller\.js\?v=20260907e/u);
  assert.match(dataController, /data_tab_controls\.js\?v=20260907a/u);
  assert.match(dataController, /data_tab_inputs_controller\.js\?v=20260906b/u);
  assert.match(dataController, /data_tab_request_controller\.js\?v=20260907b/u);
});

test("method pages and shared runtime do not depend on Dataset feature assets", async () => {
  const [
    bornhuetterFergusonSources,
    capeCodSources,
    dfmSources,
    resultSelectionSources,
    sharedDatasetSources,
    sharedDataTabSources,
  ] = await Promise.all([
    runtimeSources("ui/method_pages/bornhuetter_ferguson/"),
    runtimeSources("ui/method_pages/cape_cod/"),
    runtimeSources("ui/method_pages/dfm/"),
    runtimeSources("ui/method_pages/result_selection/"),
    runtimeSources("ui/shared/dataset/"),
    runtimeSources("ui/shared/tabs/data/"),
  ]);

  assert.doesNotMatch(bornhuetterFergusonSources, /\/ui\/dataset_viewer\//u);
  assert.doesNotMatch(capeCodSources, /\/ui\/dataset_viewer\//u);
  assert.doesNotMatch(dfmSources, /\/ui\/dataset_viewer\//u);
  assert.doesNotMatch(resultSelectionSources, /\/ui\/dataset_viewer\//u);
  assert.doesNotMatch(sharedDatasetSources, /\/ui\/dataset_viewer\//u);
  assert.doesNotMatch(sharedDataTabSources, /\/ui\/dataset_viewer\//u);
});

test("DSV and every method page consume the shared Notes and Details surfaces", async () => {
  const featureConsumerGroups = await Promise.all([
    Promise.all([
      source("ui/dataset_viewer/dataset_viewer_view.js"),
      source("ui/dataset_viewer/tabs/dataset_notes_tab.js"),
    ]),
    Promise.all([
      source("ui/method_pages/dfm/dfm_tabs_orchestrator.js"),
      source("ui/method_pages/dfm/dfm_notes_tab.js"),
    ]),
    Promise.all([source("ui/method_pages/bornhuetter_ferguson/bornhuetter_ferguson_main.js")]),
    Promise.all([source("ui/method_pages/cape_cod/cape_cod_main.js")]),
    Promise.all([source("ui/method_pages/result_selection/result_selection_main.js")]),
  ]);

  for (const consumerGroup of featureConsumerGroups) {
    const combined = consumerGroup.join("\n");
    assert.match(combined, /shared\/tabs\/notes\/notes_tab\.js/u);
    assert.match(combined, /shared\/tabs\/details\/details_form_layout\.js/u);
    assert.doesNotMatch(combined, /notes_editor_interactions\.js/u);
    assert.doesNotMatch(combined, /shared\/details_form_layout\.js/u);
  }
});

test("shared Notes renders the canonical full toolbar and preserves the shell bridge", async () => {
  const notesSource = await source("ui/shared/tabs/notes/notes_tab.js");

  for (const control of ["Font family", "Font size", "Text color", "Bold", "Italic", "Underline", "Strikethrough"]) {
    assert.match(notesSource, new RegExp(control, "u"));
  }
  assert.match(notesSource, /type:\s*"arcrho:open-path"/u);
  assert.match(notesSource, /"arcrho:open-path-result"/u);
  assert.match(notesSource, /destroy\(\)/u);
});

test("legacy top-level shared entry points are removed", async () => {
  const removedFiles = [
    "ui/shared/tabbed_page.js",
    "ui/shared/tab_popout_window.js",
    "ui/shared/details_form_layout.js",
    "ui/shared/notes_editor_interactions.js",
    "ui/shared/dataset_audit_log.js",
    "ui/shared/dataset_audit_log.css",
    "ui/shared/method_page/method_page.css",
    "ui/shared/method_page/data_surface/data_surface.css",
    "ui/shared/method_page/data_surface/dataset_main.js",
    "ui/dataset/dataset_shared.css",
    "ui/dataset/dataset_main.js",
    "ui/dataset/dataset_name_picker.js",
    "ui/dataset/dataset_origin_labels.js",
    "ui/dataset/dataset_types_source.js",
    "ui/dataset/dataset_viewer.html",
    "ui/dfm/dfm.html",
    "ui/bornhuetter_ferguson/bornhuetter_ferguson.html",
    "ui/result_selection/result_selection.html",
  ];

  for (const path of removedFiles) {
    await assert.rejects(access(new URL(path, frontendRoot)));
  }
});
