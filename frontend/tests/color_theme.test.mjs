import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import test from "node:test";
import vm from "node:vm";

const read = (path) => readFileSync(new URL(path, import.meta.url), "utf8");

const cssHexToken = (css, name) => {
  const match = css.match(new RegExp(`${name}:\\s*(#[0-9a-fA-F]{6})\\b`));
  assert.ok(match, `theme defines a hex value for ${name}`);
  return match[1];
};

const relativeLuminance = (hex) => {
  const channels = hex.slice(1).match(/.{2}/g).map((value) => Number.parseInt(value, 16) / 255);
  const linear = channels.map((value) => (value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4));
  return linear[0] * 0.2126 + linear[1] * 0.7152 + linear[2] * 0.0722;
};

const contrastRatio = (foreground, background) => {
  const lighter = Math.max(relativeLuminance(foreground), relativeLuminance(background));
  const darker = Math.min(relativeLuminance(foreground), relativeLuminance(background));
  return (lighter + 0.05) / (darker + 0.05);
};

const declarationsFor = (css, selectorFragment) => [...css.matchAll(/([^{}]+)\{([^{}]*)\}/g)]
  .filter((match) => match[1].includes(selectorFragment))
  .map((match) => match[2])
  .join("\n");

const THEMED_DOCUMENTS = [
  "../ui/index.html",
  "../ui/file_explorer/file_explorer.html",
  "../ui/dataset_viewer/dataset_viewer.html",
  "../ui/method_pages/dfm/dfm.html",
  "../ui/method_pages/bornhuetter_ferguson/bornhuetter_ferguson.html",
  "../ui/method_pages/cape_cod/cape_cod.html",
  "../ui/method_pages/berquist_sherman/berquist_sherman.html",
  "../ui/method_pages/result_selection/result_selection.html",
  "../ui/workflow/workflow.html",
  "../ui/project_instance/project_instance.html",
  "../ui/project_instance/excel_links_window.html",
  "../ui/shared/components/review_table/review_table_window.html",
  "../ui/project_settings/project_settings.html",
  "../ui/shell/browsing_history.html",
  "../ui/agent_guide/agent_guide.html",
  "../ui/task_designer/task_designer.html",
  "../ui/arcode/index.html",
  "../ui/arcode/main.html",
  "../ui/arcode/code-editor/index.html",
  "../ui/arcode/notebook-editor/index.html",
  "../ui/arcode/snowflake-console/index.html",
  "../ui/arcode/sql-server-console/index.html",
];

test("every runtime frontend document bootstraps the shared theme before loading separated theme sheets", () => {
  for (const path of THEMED_DOCUMENTS) {
    const html = read(path);
    const bootstrap = html.indexOf("/ui/shared/services/color_theme.js");
    const firstStylesheet = html.indexOf("rel=\"stylesheet\"");
    const light = html.indexOf("/ui/shared/styles/themes/light.css");
    const dark = html.indexOf("/ui/shared/styles/themes/dark.css");
    const highContrast = html.indexOf("/ui/shared/styles/themes/high_contrast.css?v=20260811c");
    const endHead = html.indexOf("</head>");
    assert.ok(bootstrap >= 0, `${path} loads the shared bootstrap`);
    assert.ok(firstStylesheet < 0 || bootstrap < firstStylesheet, `${path} applies theme state before visual CSS`);
    assert.ok(light > bootstrap && dark > light && highContrast > dark, `${path} loads light, dark, then high contrast theme ownership`);
    assert.ok(endHead > highContrast, `${path} loads theme sheets inside the head`);
  }
});

test("light values remain explicit, high contrast reuses them, and dark values stay isolated", () => {
  const light = read("../ui/shared/styles/themes/light.css");
  const dark = read("../ui/shared/styles/themes/dark.css");
  const highContrast = read("../ui/shared/styles/themes/high_contrast.css");

  assert.match(light, /:root\[data-arcrho-theme="light"\]/);
  assert.match(light, /:root\[data-arcrho-theme="high-contrast"\]/);
  assert.match(light, /--ar-native-window-background:\s*#ffffff/);
  assert.match(light, /--ar-color-surface:\s*#ffffff/);
  assert.match(light, /--ar-color-text:\s*#1f2937/);
  assert.match(light, /--ar-color-accent:\s*#2b6df6/);
  assert.match(light, /--ar-color-scrollbar-track:\s*#f1f3f5/);
  assert.match(light, /--ar-chart-dataset-status-text:\s*#000000/);
  assert.match(light, /--ar-chart-dfm-empty-text:\s*#555555/);
  assert.match(light, /--ar-chart-dfm-point-border:\s*#94a3b8/);
  assert.doesNotMatch(light, /data-arcrho-theme="dark"/);

  assert.match(highContrast, /:root\[data-arcrho-theme="high-contrast"\]/);
  assert.match(highContrast, /color-scheme:\s*light/);
  assert.match(highContrast, /--ar-spreadsheet-label-text:\s*#000000/);
  assert.match(highContrast, /--ar-spreadsheet-selection-text:\s*#000000/);
  assert.match(highContrast, /#tableWrap/);
  assert.match(highContrast, /\.pi-table/);
  assert.match(highContrast, /\.taskDesignerTable/);
  assert.match(highContrast, /color:\s*#000000\s*!important/);
  assert.deepEqual(
    [...highContrast.matchAll(/(--ar-[\w-]+)\s*:/g)].map((match) => match[1]),
    ["--ar-spreadsheet-label-text", "--ar-spreadsheet-selection-text"],
    "High Contrast only overrides spreadsheet text tokens",
  );
  assert.doesNotMatch(highContrast, /(?:background|border|fill|stroke)\s*:/);
  assert.doesNotMatch(highContrast, /--ar-color-/);

  const excludedRatioDeclarations = declarationsFor(highContrast, "#ratioWrap td.ratioCell.strike");
  assert.match(excludedRatioDeclarations, /color:\s*#b000c2\s*!important/);
  assert.match(excludedRatioDeclarations, /text-decoration-color:\s*#b000c2\s*!important/);

  assert.match(dark, /:root\[data-arcrho-theme="dark"\]/);
  assert.match(dark, /color-scheme:\s*dark/);
  assert.match(dark, /--ar-color-canvas:\s*#282c34/);
  assert.match(dark, /--ar-color-text:\s*#abb2bf/);
  assert.match(dark, /--ar-color-surface-muted:\s*#21252b/);
  assert.match(dark, /--ar-color-border-focus:\s*#528bff/);
  assert.match(dark, /--ar-monaco-editor-background:\s*#282c34/);
  assert.match(dark, /--ar-monaco-syntax-keyword:\s*#c678dd/);
  assert.match(dark, /--ar-monaco-syntax-string:\s*#98c379/);
  assert.doesNotMatch(dark, /:root\[data-arcrho-theme="light"\]/);

  const lightTokens = [...light.matchAll(/(--ar-[\w-]+)\s*:/g)].map((match) => match[1]);
  const darkTokens = new Set([...dark.matchAll(/(--ar-[\w-]+)\s*:/g)].map((match) => match[1]));
  for (const token of lightTokens) assert.ok(darkTokens.has(token), `dark theme defines ${token}`);
});

test("dark theme text tokens keep readable contrast on operational surfaces", () => {
  const dark = read("../ui/shared/styles/themes/dark.css");
  const pairs = [
    ["--ar-color-text", "--ar-color-canvas"],
    ["--ar-color-text", "--ar-color-surface"],
    ["--ar-color-text-muted", "--ar-color-surface"],
    ["--ar-color-text-subtle", "--ar-color-surface"],
    ["--ar-color-accent", "--ar-color-accent-soft"],
    ["--ar-color-success", "--ar-color-success-soft"],
    ["--ar-color-warning", "--ar-color-warning-soft"],
    ["--ar-color-danger", "--ar-color-danger-soft"],
  ];

  for (const [foregroundName, backgroundName] of pairs) {
    const foreground = cssHexToken(dark, foregroundName);
    const background = cssHexToken(dark, backgroundName);
    const ratio = contrastRatio(foreground, background);
    assert.ok(ratio >= 4.5, `${foregroundName} on ${backgroundName} has ${ratio.toFixed(2)}:1 contrast`);
  }
});

test("Dark Home and shared spreadsheet tables distinguish panels, labels, and values", () => {
  const dark = read("../ui/shared/styles/themes/dark.css");
  const homeWelcome = declarationsFor(dark, ".homeWelcomePanel");
  const homeCard = declarationsFor(dark, ".home .card");
  const valueCell = declarationsFor(dark, ".arSpreadsheetTable td:not");
  const headerCell = declarationsFor(dark, ".arSpreadsheetTable thead th:not");
  const labelCell = declarationsFor(dark, ".arSpreadsheetTable :is(tbody th, tbody td:first-child)");

  assert.match(homeWelcome, /background-color:\s*var\(--ar-color-input\)/);
  assert.match(homeCard, /background-color:\s*var\(--ar-color-canvas-subtle\)/);
  assert.match(homeCard, /border-color:\s*var\(--ar-color-border-strong\)/);
  assert.match(valueCell, /background-color:\s*var\(--ar-color-surface\)/);
  assert.match(headerCell, /background-color:\s*var\(--ar-spreadsheet-header-fill\)/);
  assert.match(labelCell, /background-color:\s*var\(--ar-spreadsheet-label-fill\)/);
  assert.match(dark, /--ar-spreadsheet-header-fill:\s*var\(--ar-color-input\)/);
  assert.match(dark, /--ar-spreadsheet-label-fill:\s*var\(--ar-color-input\)/);
});

test("Dark tabbed pages keep tab frames and the selected seam visible", () => {
  const dark = read("../ui/shared/styles/themes/dark.css");
  const tabBar = declarationsFor(dark, ".tabbedPageTabBar");
  const tabBarSeam = declarationsFor(dark, ".tabbedPageTabBar::after");
  const tab = declarationsFor(dark, ".tabbedPageTab");
  const selectedTab = declarationsFor(dark, ".tabbedPageTab.active, .tabbedPageTab[aria-selected");
  const selectedSeam = declarationsFor(dark, ".tabbedPageTab.active::after");

  assert.match(tabBar, /border-color:\s*var\(--ar-color-border-strong\)/);
  assert.match(tabBarSeam, /background-color:\s*var\(--ar-color-border-strong\)/);
  assert.match(tab, /border-color:\s*var\(--ar-color-border-strong\)/);
  assert.match(tab, /background-color:\s*var\(--ar-color-canvas\)/);
  assert.match(selectedTab, /background-color:\s*var\(--ar-color-surface-raised\)/);
  assert.match(selectedSeam, /background-color:\s*var\(--ar-color-surface-raised\)/);
});

test("Details fields and the shared dependency surface stay readable in Dark mode", () => {
  const dark = read("../ui/shared/styles/themes/dark.css");
  const control = declarationsFor(dark, "#dsDetailsPage .arDetailsControl");
  const readonly = declarationsFor(dark, "#dsDetailsPage :is(.arDetailsControl[readonly]");
  // The chips, the formula field, and the tooltip are one surface the Dataset
  // Viewer and all five method pages share, so Dark mode moves its tokens
  // rather than repainting each page.
  const surface = declarationsFor(dark, ":root[data-arcrho-theme=\"dark\"] .arDetailsRoot");
  const tooltip = declarationsFor(dark, ".arDetailsFormulaTooltip");

  assert.match(control, /background-color:\s*var\(--ar-color-input\)/);
  assert.match(control, /color:\s*var\(--ar-color-text-strong\)/);
  assert.match(readonly, /background-color:\s*var\(--ar-color-surface-muted\)/);
  assert.match(readonly, /color:\s*var\(--ar-color-text-muted\)/);
  assert.match(surface, /--ar-details-chip-box-border:\s*var\(--ar-color-border-strong\)/);
  assert.match(surface, /--ar-details-chip-background:\s*var\(--ar-color-surface-muted\)/);
  // The formula carries no chip or token shape any more: a source name is body
  // text and an operator is quiet punctuation - colour only, no fill.
  assert.match(surface, /--ar-details-formula-operator:\s*var\(--ar-color-text-muted\)/);
  assert.match(tooltip, /background-color:\s*var\(--ar-color-surface-raised\)/);
  // No page may repaint the shared surface behind the tokens' back.
  assert.doesNotMatch(dark, /dsFormulaComponent|dsDependentLink|dsDatasetChipBox/u);
});

test("DSV tab pages use the same muted outer frame as the tab strip in Dark mode", () => {
  const dark = read("../ui/shared/styles/themes/dark.css");
  const pageHost = declarationsFor(dark, ":is(#dsDetailsPage, #dsDataPage, #dsChartPage, #dsNotesPage, #dsLinksPage, #dsAuditLogPage)");

  assert.match(pageHost, /border-color:\s*var\(--ar-color-border-strong\)/);
});

test("Dark-mode table value cells use the shared lighter surface", () => {
  const dark = read("../ui/shared/styles/themes/dark.css");
  const valueCellSelectors = [
    "#tableWrap td",
    "#ratioWrap td:not",
    ".dfmDevelopedCurveSegmentTable td",
    ".bfMethodTable td:not",
    ".pi-table td",
    ".pi-number-formats-table td",
    ".pi-add-picker-table td",
    ".field-mapping-table, .dataset-types-table, .dpr-rules-table) td",
    ".taskDesignerTable td",
    ".sfTable td",
  ];

  for (const selector of valueCellSelectors) {
    const declarations = declarationsFor(dark, selector);
    assert.match(declarations, /background(?:-color)?:\s*var\(--ar-color-surface\)/, `${selector} uses the shared table value fill`);
  }
  assert.match(declarationsFor(dark, ".gc-table td"), /background-color:\s*var\(--ar-color-surface\)/);
  assert.match(dark, /--ar-color-table-header:\s*var\(--ar-color-input\)/);
});

test("Arcode code surfaces use the shared Atom One Dark editor tokens", () => {
  const dark = read("../ui/shared/styles/themes/dark.css");
  const notebookCode = declarationsFor(dark, ".sc-markdown-render code");
  const taskDesignerCode = declarationsFor(dark, ".taskDesignerDetailPre");

  for (const declarations of [notebookCode, taskDesignerCode]) {
    assert.match(declarations, /background-color:\s*var\(--ar-monaco-editor-background\)/);
    assert.match(declarations, /color:\s*var\(--ar-monaco-editor-foreground\)/);
  }
});

test("Arcode Dark explorer uses quiet themed surfaces and a visible resize state", () => {
  const dark = read("../ui/shared/styles/themes/dark.css");
  const chrome = read("../ui/arcode/shared/chrome.css");
  const shellCss = read("../ui/arcode/main.css");
  const sidebar = declarationsFor(dark, ".arcodeHomeSidebar");
  const entry = declarationsFor(dark, ".arcodeExplorerEntry");
  const supportedEntry = declarationsFor(dark, ".arcodeExplorerEntry.file.supported");

  assert.match(sidebar, /background-color:\s*var\(--ar-color-surface-muted\)/);
  assert.match(entry, /background-color:\s*transparent/);
  assert.match(entry, /border-color:\s*transparent/);
  assert.match(entry, /color:\s*var\(--ar-color-text\)/);
  assert.match(supportedEntry, /color:\s*var\(--ar-color-text\)/);

  // The seam and its resize state are token-driven, so Dark themes the shared
  // Arcode chrome tokens once instead of restyling the resizer per surface.
  assert.match(chrome, /--ark-seam:\s*1px/);
  assert.match(shellCss, /\.arcodeExplorerResizer\s*\{[^}]*background:\s*var\(--ark-border\)/);
  assert.match(
    shellCss,
    /\.arcodeExplorerResizer:hover,[\s\S]{0,160}\{[^}]*background:\s*var\(--ark-accent\)/,
  );
  assert.match(dark, /--ark-border:\s*var\(--ar-color-border\)/);
  assert.match(dark, /--ark-accent:\s*var\(--ar-color-border-focus\)/);
  assert.doesNotMatch(dark, /\.arcodeExplorerResizer/);
});

test("Arcode Dark notebook keeps its canvas, TOC, split panel, and code cells visually distinct", () => {
  const dark = read("../ui/shared/styles/themes/dark.css");
  const cellsArea = declarationsFor(dark, ".sc-cells-area");
  const sidebar = declarationsFor(dark, ".sc-sidebar-content");
  const tocItem = declarationsFor(dark, ".sc-toc-item");
  const tocFold = declarationsFor(dark, ".sc-toc-fold");
  const splitHandle = declarationsFor(dark, ".sc-split-resize-handle");
  const cell = declarationsFor(dark, ".sc-cell");
  const cellSide = declarationsFor(dark, ".sc-cell-side");
  const outputSide = declarationsFor(dark, ".sc-cell.code .sc-cell-output-side-placeholder");
  const inputFrame = declarationsFor(dark, ".sc-cell-input-frame");

  assert.match(cellsArea, /background-color:\s*var\(--sc-cells-panel-bg\)/);
  assert.match(sidebar, /background-color:\s*var\(--ar-color-surface-muted\)/);
  assert.match(tocItem, /background-color:\s*transparent/);
  assert.match(tocItem, /color:\s*var\(--ar-color-text\)/);
  assert.match(tocFold, /border-color:\s*transparent/);
  assert.match(splitHandle, /background-color:\s*var\(--ar-color-canvas\)/);
  assert.match(cell, /border-color:\s*var\(--ar-color-border-strong\)/);
  assert.match(cellSide, /background:\s*var\(--ar-color-surface-muted\)/);
  assert.match(outputSide, /background:\s*var\(--ar-color-surface-muted\)/);
  assert.match(inputFrame, /background-color:\s*var\(--ar-monaco-editor-background\)/);
});

test("dark titlebar window controls use themed surfaces and state colors", () => {
  const dark = read("../ui/shared/styles/themes/dark.css");
  const button = declarationsFor(dark, ".titlebarBtn");
  const icon = declarationsFor(dark, ".titlebarIcon");
  const standardHover = declarationsFor(dark, "#titlebarMinBtn");
  const closeHover = declarationsFor(dark, "#titlebarCloseBtn");

  assert.match(button, /background-color:\s*var\(--ar-color-surface-raised\)/);
  assert.match(button, /border-color:\s*var\(--ar-color-border-strong\)/);
  assert.match(icon, /stroke:\s*currentColor/);
  assert.match(standardHover, /background-color:\s*var\(--ar-color-accent-soft\)/);
  assert.match(standardHover, /border-color:\s*var\(--ar-color-border-focus\)/);
  assert.match(closeHover, /background-color:\s*#9f2940/);
  assert.match(closeHover, /border-color:\s*var\(--ar-color-danger\)/);
  assert.match(closeHover, /color:\s*#ffffff/);
});

test("Project Settings and Project Instance override light-only child paint in Dark mode", () => {
  const dark = read("../ui/shared/styles/themes/dark.css");
  const requiredRules = [
    [".ptree-label", /color:\s*var\(--ar-color-text\)/],
    [".pi-table-header-cell", /background:\s*var\(--ar-color-table-header\)/],
    [".pi-window-titlebar", /background:\s*var\(--ar-popout-header\)/],
    [".pi-window-titlebar-icon", /stroke:\s*currentColor/],
    [".pi-prefs-header", /background-color:\s*var\(--ar-color-surface-muted\)/],
    [".pi-prefs-nav-item.is-active", /background:\s*var\(--ar-color-accent-soft\)/],
    [".pi-prefs-tabchip", /background:\s*var\(--ar-color-input\)/],
    [".tree-project", /color:\s*var\(--ar-color-text\)/],
    [".ribbon-label", /color:\s*var\(--ar-color-text-muted\)/],
    [".summary-header", /color:\s*var\(--ar-color-text-strong\)/],
    ["#datasetTypesErrorTitle", /color:\s*var\(--ar-color-text\)/],
    [".rct-formula-calculated-icon", /color:\s*var\(--ar-color-accent-strong\)/],
    [".dpr-token-menu-tick", /color:\s*var\(--ar-color-accent-strong\)/],
    [".dpr-editor", /--dpr-ink:\s*inherit/],
  ];

  for (const [selector, expectedDeclaration] of requiredRules) {
    const declarations = declarationsFor(dark, selector);
    assert.ok(declarations, `dark theme contains ${selector}`);
    assert.match(declarations, expectedDeclaration, `${selector} uses shared dark-theme paint`);
  }
  assert.doesNotMatch(dark, /\.pi-prefs-titlebar/);

  const datasetTypesCss = read("../ui/project_settings/project_settings_dataset_types.css");
  const datasetTypesJs = read("../ui/project_settings/project_settings_dataset_types.js");
  const projectSettingsHtml = read("../ui/project_settings/project_settings.html");
  assert.match(datasetTypesCss, /\.datasetTypesRecalcOverlay\s*\{/);
  assert.doesNotMatch(datasetTypesJs, /datasetTypesRecalcDialogStyles|createElement\("style"\)/);
  assert.match(projectSettingsHtml, /project_settings_dataset_types\.css\?v=20260821pstree1/);
});

test("theme runtime validates, persists per user, applies, notifies frames, and updates Monaco live", async () => {
  const source = read("../ui/shared/services/color_theme.js");
  const attributes = new Map();
  const storage = new Map();
  const listeners = new Map();
  const posted = [];
  const monacoThemes = [];
  const monacoDefinitions = new Map();
  const events = [];
  const nativeBackgrounds = [];
  const savedHostThemes = [];
  const root = {
    getAttribute: (name) => attributes.get(name) || null,
    setAttribute: (name, value) => attributes.set(name, value),
  };
  const atomOneDarkCss = {
    "--ar-monaco-editor-background": "#282c34",
    "--ar-monaco-editor-foreground": "#abb2bf",
    "--ar-monaco-editor-active-foreground": "#d7dae0",
    "--ar-monaco-editor-cursor": "#528bff",
    "--ar-monaco-editor-line-highlight": "#99bbff0a",
    "--ar-monaco-editor-selection": "#3e4451",
    "--ar-monaco-editor-selection-highlight": "#3e445166",
    "--ar-monaco-editor-find-match-highlight": "#528bff3d",
    "--ar-monaco-editor-indent-guide": "#abb2bf26",
    "--ar-monaco-editor-indent-guide-active": "#626772",
    "--ar-monaco-editor-line-number": "#636d83",
    "--ar-monaco-editor-widget-background": "#21252b",
    "--ar-monaco-editor-widget-border": "#3a3f4b",
    "--ar-monaco-editor-list-selection": "#2c313a",
    "--ar-monaco-editor-list-hover": "#2c313a66",
    "--ar-monaco-editor-scrollbar": "#4e566680",
    "--ar-monaco-editor-scrollbar-hover": "#5a637580",
    "--ar-monaco-editor-scrollbar-active": "#747d9180",
    "--ar-monaco-syntax-comment": "#5c6370",
    "--ar-monaco-syntax-keyword": "#c678dd",
    "--ar-monaco-syntax-number": "#d19a66",
    "--ar-monaco-syntax-string": "#98c379",
    "--ar-monaco-syntax-type": "#e5c07b",
    "--ar-monaco-syntax-function": "#61afef",
    "--ar-monaco-syntax-variable": "#e06c75",
    "--ar-monaco-syntax-cyan": "#56b6c2",
    "--ar-monaco-syntax-interpolation": "#be5046",
    "--ar-monaco-syntax-invalid": "#e05252",
  };
  const getComputedStyle = () => ({
    getPropertyValue: (name) => {
      if (name === "--ar-native-window-background") {
        return attributes.get("data-arcrho-theme") === "dark" ? "#282c34" : "#ffffff";
      }
      return atomOneDarkCss[name] || "";
    },
  });
  const context = {
    CustomEvent: class CustomEvent {
      constructor(type, init = {}) { this.type = type; this.detail = init.detail; }
    },
    document: {
      documentElement: root,
      readyState: "complete",
      querySelectorAll: () => [{ contentWindow: { postMessage: (message) => posted.push(message) } }],
    },
    localStorage: {
      getItem: (key) => storage.get(key) || null,
      setItem: (key, value) => storage.set(key, value),
    },
    location: { search: "" },
    monaco: {
      editor: {
        defineTheme: (name, definition) => monacoDefinitions.set(name, definition),
        setTheme: (theme) => monacoThemes.push(theme),
      },
    },
    getComputedStyle,
    ADAHost: {
      setWindowBackgroundColor: (color) => nativeBackgrounds.push(color),
      loadColorThemePreference: async () => ({ exists: false, theme: "light" }),
      saveColorThemePreference: async (theme) => {
        savedHostThemes.push(theme);
        return { ok: true, theme };
      },
    },
    addEventListener: (type, handler) => listeners.set(type, handler),
    dispatchEvent: (event) => events.push(event),
  };
  context.window = context;
  context.top = context;

  vm.runInNewContext(source, context, { filename: "color_theme.js" });
  assert.equal(attributes.get("data-arcrho-theme"), "light");
  assert.equal(context.ArcRhoColorTheme.normalizeTheme("unsupported"), "light");
  assert.equal(context.ArcRhoColorTheme.normalizeTheme("high-contrast"), "high-contrast");
  assert.equal(context.ArcRhoColorTheme.getMonacoTheme("dark"), "arcrho-atom-one-dark");
  assert.equal(context.ArcRhoColorTheme.getMonacoTheme("high-contrast"), "vs");

  context.ArcRhoColorTheme.setTheme("dark");
  await Promise.resolve();
  await Promise.resolve();
  assert.equal(storage.get("arcrho_color_theme"), "dark");
  assert.equal(savedHostThemes.at(-1), "dark");
  assert.equal(attributes.get("data-arcrho-theme"), "dark");
  assert.equal(monacoThemes.at(-1), "arcrho-atom-one-dark");
  assert.equal(posted.at(-1)?.type, "arcrho:set-color-theme");
  assert.equal(posted.at(-1)?.theme, "dark");
  assert.equal(events.at(-1).type, "arcrho:color-theme-changed");
  assert.equal(nativeBackgrounds.at(-1), "#282c34");
  const atomOneDarkTheme = monacoDefinitions.get("arcrho-atom-one-dark");
  assert.ok(atomOneDarkTheme, "defines the shared Atom One Dark Monaco theme");
  assert.equal(atomOneDarkTheme.colors["editor.background"], "#282c34");
  assert.equal(atomOneDarkTheme.colors["editor.foreground"], "#abb2bf");
  assert.equal(atomOneDarkTheme.colors["editor.selectionBackground"], "#3e4451");
  assert.equal(atomOneDarkTheme.rules.find((rule) => rule.token === "keyword")?.foreground, "c678dd");
  assert.equal(atomOneDarkTheme.rules.find((rule) => rule.token === "string")?.foreground, "98c379");

  context.ArcRhoColorTheme.setTheme("high-contrast");
  await Promise.resolve();
  await Promise.resolve();
  assert.equal(storage.get("arcrho_color_theme"), "high-contrast");
  assert.equal(savedHostThemes.at(-1), "high-contrast");
  assert.equal(attributes.get("data-arcrho-theme"), "high-contrast");
  assert.equal(monacoThemes.at(-1), "vs");
  assert.equal(posted.at(-1)?.theme, "high-contrast");

  listeners.get("message")?.({ data: { type: "arcrho:set-color-theme", theme: "light" } });
  assert.equal(attributes.get("data-arcrho-theme"), "light");
  assert.equal(monacoThemes.at(-1), "vs");

  const childNativeBackgrounds = [];
  const childContext = {
    CustomEvent: context.CustomEvent,
    document: {
      documentElement: root,
      readyState: "complete",
      querySelectorAll: () => [],
    },
    localStorage: context.localStorage,
    location: { search: "" },
    getComputedStyle: context.getComputedStyle,
    ADAHost: { setWindowBackgroundColor: (color) => childNativeBackgrounds.push(color) },
    addEventListener: () => {},
    dispatchEvent: () => {},
  };
  childContext.window = childContext;
  childContext.top = {};
  vm.runInNewContext(source, childContext, { filename: "color_theme_child.js" });
  assert.deepEqual(childNativeBackgrounds, [], "child frames do not overwrite native BrowserWindow paint");
});

test("theme runtime restores the Electron user preference after renderer storage is cleared", async () => {
  const source = read("../ui/shared/services/color_theme.js");
  const attributes = new Map();
  const storage = new Map();
  const context = {
    CustomEvent: class CustomEvent {
      constructor(type, init = {}) { this.type = type; this.detail = init.detail; }
    },
    URLSearchParams,
    document: {
      documentElement: {
        getAttribute: (name) => attributes.get(name) || null,
        setAttribute: (name, value) => attributes.set(name, value),
      },
      readyState: "complete",
      querySelectorAll: () => [],
    },
    localStorage: {
      getItem: (key) => storage.get(key) || null,
      setItem: (key, value) => storage.set(key, value),
    },
    location: { search: "" },
    getComputedStyle: () => ({ getPropertyValue: () => "#151a22" }),
    ADAHost: {
      loadColorThemePreference: async () => ({ exists: true, theme: "dark" }),
      saveColorThemePreference: async () => ({ ok: true, theme: "dark" }),
      setWindowBackgroundColor: () => {},
    },
    addEventListener: () => {},
    dispatchEvent: () => {},
  };
  context.window = context;
  context.top = context;

  vm.runInNewContext(source, context, { filename: "color_theme_restore.js" });
  assert.equal(attributes.get("data-arcrho-theme"), "light", "renderer cache starts from its fallback");
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(attributes.get("data-arcrho-theme"), "dark", "host preference restores the selected theme");
  assert.equal(storage.get("arcrho_color_theme"), "dark", "renderer cache is rebuilt from the host preference");
});

test("ArcRho and standalone Arcode keep accessible theme menus without topbar toggles", () => {
  const shellHtml = read("../ui/index.html");
  const shellPreferences = read("../ui/shell/shell_preferences.js");
  const shellHotkeys = read("../ui/shell/shell_hotkeys.js");
  const shellMenus = read("../ui/shell/shell_menus.js");
  const iframeHost = read("../ui/shell/iframe_host.js");
  const arcodeHtml = read("../ui/arcode/main.html");
  const arcodeMain = read("../ui/arcode/main.js");

  for (const html of [shellHtml, arcodeHtml]) {
    assert.match(html, /data-action="color-theme-light"[^>]*role="menuitemradio"/);
    assert.match(html, /data-action="color-theme-dark"[^>]*role="menuitemradio"/);
    assert.match(html, /data-action="color-theme-high-contrast"[^>]*role="menuitemradio"/);
    assert.match(html, /data-color-theme-value="high-contrast"[^>]*tabindex="-1"/);
    assert.match(html, /data-color-theme-trigger[^>]*tabindex="0"/);
    assert.match(html, /data-color-theme-menu[^>]*aria-haspopup="menu"/);
    assert.match(html, /data-color-theme-value="light"[^>]*tabindex="-1"/);
    assert.doesNotMatch(html, /arThemeToggle|theme_toggle\.css|color-theme\.svg/);
  }
  assert.match(shellPreferences, /api\?\.setTheme\?\./);
  assert.match(shellPreferences, /type:\s*messageType,\s*theme:\s*normalized/);
  assert.doesNotMatch(shellPreferences, /initColorThemeToggle|colorThemeToggle/);
  assert.match(shellHotkeys, /ArcRhoColorTheme\?\.THEMES/);
  assert.match(shellHotkeys, /themes\.length/);
  assert.match(shellMenus, /data-color-theme-value/);
  assert.match(iframeHost, /postMessage\(\{ type: messageType, theme \}/);
  assert.match(arcodeMain, /ArcRhoColorTheme\?\.setTheme/);
  assert.match(arcodeMain, /action\.startsWith\("color-theme-"\)/);
  assert.doesNotMatch(arcodeMain, /initColorThemeToggle|updateColorThemeToggleUI|arcodeColorThemeToggle/);
  assert.match(arcodeMain, /\.menu\[aria-expanded="true"\][^\n]*setAttribute\("aria-expanded", "false"\)/);
  assert.equal(existsSync(new URL("../ui/shared/styles/theme_toggle.css", import.meta.url)), false);
  assert.equal(existsSync(new URL("../ui/shared/icons/color-theme.svg", import.meta.url)), false);

  const runtime = read("../ui/shared/services/color_theme.js");
  assert.match(runtime, /wireThemeMenus/);
  assert.match(runtime, /\["Enter", " ", "ArrowDown"\]/);
  assert.match(runtime, /event\.key === "ArrowUp"/);
  assert.match(runtime, /event\.key === "ArrowLeft" \|\| event\.key === "Escape"/);
});

test("shell submenu indicators use the shared SVG chevron instead of text glyphs", () => {
  const shellHtml = read("../ui/index.html");
  const shellCss = read("../ui/shell/shell.css");
  const arcodeHtml = read("../ui/arcode/main.html");
  const arcodeCss = read("../ui/arcode/main.css");
  const icon = read("../ui/shared/icons/chevron-right.svg");

  for (const html of [shellHtml, arcodeHtml]) {
    assert.match(html, /class="menuSubmenuIcon"/);
    assert.match(html, /chevron-right\.svg\?v=20260722a#chevron-right/);
  }
  assert.match(icon, /<symbol id="chevron-right"/);
  assert.match(icon, /stroke="currentColor"/);
  assert.doesNotMatch(shellCss, /hasSubmenu::after/);
  assert.doesNotMatch(shellCss, /content:\s*">"/);
  assert.doesNotMatch(arcodeHtml, /class="menuArrow"/);
  assert.doesNotMatch(arcodeCss, /\.menuArrow/);
});

test("all Monaco owners choose the shared initial theme and Electron accepts computed theme paint", () => {
  // The code editor and both SQL editors create their Monaco editor through
  // the one editor framework, so it is the owner checked for all three.
  const editorFramework = read("../ui/arcode/shared/editor_framework.js");
  const notebookEditor = read("../ui/arcode/notebook-editor/core.js");
  const editorOwners = [editorFramework, notebookEditor];
  for (const owner of editorOwners) {
    assert.match(owner, /ArcRhoColorTheme\?\.getMonacoTheme\?\.\(\) \|\| "vs"/);
    assert.doesNotMatch(owner, /theme:\s*"vs"/);
  }
  assert.match(editorFramework, /const monacoTheme = window\.ArcRhoColorTheme\?\.getMonacoTheme\?\.\(\) \|\| "vs";[\s\S]*theme:\s*monacoTheme/);
  assert.match(notebookEditor, /monacoReady = true;[\s\S]*EDITOR_OPTIONS\.theme = window\.ArcRhoColorTheme\?\.getMonacoTheme\?\.\(\) \|\| "vs";/);

  const preload = read("../electron/preload.js");
  const main = read("../electron/main.js");
  assert.match(preload, /setWindowBackgroundColor/);
  assert.match(preload, /loadColorThemePreference/);
  assert.match(preload, /saveColorThemePreference/);
  assert.match(main, /ipcMain\.handle\("window-set-background-color"/);
  assert.match(main, /ipcMain\.handle\("color-theme-preference-load"/);
  assert.match(main, /ipcMain\.handle\("color-theme-preference-save"/);
  assert.match(main, /color_theme:\s*theme/);
  assert.match(main, /buildArcRhoUrl\(\{ uiVersion \}\)/);
  assert.match(main, /normalizeColorThemePreference\(payload\?\.colorTheme\)/);
  assert.match(read("../ui/shell/app_lifecycle.js"), /colorTheme:\s*shell\.getColorTheme\?\.\(\)/);
  assert.match(read("../ui/arcode/main.js"), /colorTheme:\s*window\.ArcRhoColorTheme\?\.getTheme\?\.\(\)/);
  assert.match(main, /getIpcWindow\(event\)/);
  assert.match(main, /setBackgroundColor\(color\)/);
  assert.match(main, /saveCachedWindowBackgroundColor\(color\)/);
  assert.doesNotMatch(main, /saveCachedWindowBackgroundColor\(DEFAULT_WINDOW_BACKGROUND_COLOR\)/);

  const light = read("../ui/shared/styles/themes/light.css");
  const runtime = read("../ui/shared/services/color_theme.js");
  assert.match(light, /data-arcrho-app="arcode"[\s\S]*--ar-native-window-background:\s*#f7f8fa/);
  assert.match(runtime, /global\.top && global\.top !== global/);
});

test("the startup splash mirrors the renderer-derived persisted theme without changing Light defaults", () => {
  const splash = read("../ui/splash.html");
  const dark = read("../ui/shared/styles/themes/dark.css");
  const main = read("../electron/main.js");

  const bootstrap = splash.indexOf("data-arcrho-theme");
  const inlineStyles = splash.indexOf("<style>");
  assert.ok(bootstrap >= 0 && bootstrap < inlineStyles, "splash theme state is set before first paint styles");
  assert.match(splash, /themes\.has\(requestedTheme\) \? requestedTheme : "light"/);
  assert.match(splash, /background:\s*#f8f9fc/);
  assert.match(splash, /\.\/shared\/styles\/themes\/light\.css\?v=20260817d/);
  assert.match(splash, /\.\/shared\/styles\/themes\/dark\.css\?v=20260903a/);
  assert.match(splash, /\.\/shared\/styles\/themes\/high_contrast\.css\?v=20260811c/);
  assert.match(dark, /\.startupSplash/);
  assert.match(dark, /\.splash-container\s*\{[^}]*width:\s*292px[^}]*border:\s*1px solid var\(--ar-color-border\)[^}]*border-radius:\s*6px/s);
  assert.match(dark, /\.logo-icon img\s*\{[^}]*width:\s*88px[^}]*height:\s*88px/s);
  assert.match(dark, /\.logo-icon\s*\{[^}]*animation:\s*none/s);
  assert.match(dark, /:is\(\.orb, \.scan-line\)\s*\{[^}]*display:\s*none/s);
  assert.match(main, /theme:\s*startupTheme/);
  assert.match(main, /backgroundColor:\s*startupBackgroundColor/);
});

test("DFM Ratios dark mode keeps exclusions visible and selected averages restrained", () => {
  const dark = read("../ui/shared/styles/themes/dark.css");
  const dfm = read("../ui/method_pages/dfm/dfm.html");
  const excludedDeclarations = declarationsFor(dark, "#ratioWrap td.ratioCell.strike");
  const selectedAverageDeclarations = declarationsFor(dark, "#ratioWrap td.summaryCell.ratioSelectedCell");

  assert.match(excludedDeclarations, /color:\s*#c58bd8/);
  assert.match(excludedDeclarations, /text-decoration-color:\s*#c58bd8/);
  assert.match(selectedAverageDeclarations, /background-color:\s*#526331/);
  assert.match(selectedAverageDeclarations, /color:\s*#edf4d5/);
  assert.ok(contrastRatio("#c58bd8", "#282c34") >= 4.5, "excluded ratios remain readable on the table surface");
  assert.ok(contrastRatio("#edf4d5", "#526331") >= 4.5, "selected average text remains readable on its fill");
  assert.match(dfm, /themes\/dark\.css\?v=20260903a/);
});

test("changed theme and chart owners are reached through current cache-version chains", () => {
  const expectedReferences = [
    ["../ui/dataset_viewer/dataset_viewer.html", "dataset_viewer_main.js?v=20260907e"],
    ["../ui/dataset_viewer/dataset_viewer_main.js", "dataset_viewer_view.js?v=20260907a"],
    ["../ui/dataset_viewer/dataset_viewer_main.js", "dataset_chart_tab.js?v=20260907a"],
    ["../ui/dataset_viewer/tabs/dataset_chart_tab.js", "dataset_chart_renderer.js?v=20260724a"],
    ["../ui/method_pages/bornhuetter_ferguson/bornhuetter_ferguson.html", "bornhuetter_ferguson_main.js?v=20260830a"],
    ["../ui/method_pages/cape_cod/cape_cod.html", "cape_cod_main.js?v=20260830a"],
    ["../ui/method_pages/result_selection/result_selection.html", "result_selection_main.js?v=20260908a"],
    ["../ui/method_pages/dfm/dfm.html", "dfm_main.js?v=20260907a"],
    ["../ui/project_settings/project_settings.html", "project_settings.js?v=20260903live1"],
    ["../ui/project_settings/project_settings.js", "project_settings_dataset_types.js?v=20260901dup1"],
    ["../ui/arcode/code-editor/index.html", "code-editor/index.js?v=20260818a"],
    ["../ui/arcode/code-editor/index.js", "shared/editor_framework.js?v=20260818a"],
    ["../ui/arcode/notebook-editor/index.html", "notebook-editor/core.js?v=20260816b"],
    ["../ui/arcode/snowflake-console/index.html", "snowflake-console/index.js?v=20260818a"],
    ["../ui/arcode/snowflake-console/index.js", "shared/sql_mode.js?v=20260818a"],
    ["../ui/arcode/sql-server-console/index.html", "sql-server-console/index.js?v=20260818a"],
    ["../ui/arcode/sql-server-console/index.js", "shared/sql_mode.js?v=20260818a"],
    ["../ui/arcode/shared/sql_mode.js", "./editor_framework.js?v=20260818a"],
    ["../ui/arcode/shared/sql_mode.js", "./sql_engines.js?v=20260818a"],
    ["../ui/arcode/main.html", "main.js?v=20260818a"],
    ["../ui/arcode/main.js", "database-connections/dialog.js?v=20260818a"],
    ["../ui/arcode/database-connections/dialog.js", "../shared/sql_engines.js?v=20260818a"],
  ];
  for (const [path, reference] of expectedReferences) {
    assert.ok(read(path).includes(reference), `${path} loads ${reference}`);
  }

  const iframeHost = read("../ui/shell/iframe_host.js");
  assert.match(iframeHost, /workflow\.html\?\$\{params\.toString\(\)\}/);
  assert.match(iframeHost, /params\.set\("v", uiVersionParam\)/);
});

test("large page styles are maintained as feature CSS instead of inline blocks", () => {
  const extractedPages = [
    ["../ui/index.html", "/ui/shell/shell.css"],
    ["../ui/workflow/workflow.html", "/ui/workflow/workflow.css"],
    ["../ui/project_instance/project_instance.html", "/ui/project_instance/project_instance.css"],
    ["../ui/method_pages/dfm/dfm.html", "/ui/method_pages/dfm/dfm.css"],
    ["../ui/method_pages/bornhuetter_ferguson/bornhuetter_ferguson.html", "/ui/method_pages/bornhuetter_ferguson/bornhuetter_ferguson.css"],
    ["../ui/method_pages/cape_cod/cape_cod.html", "/ui/method_pages/cape_cod/cape_cod.css"],
    ["../ui/method_pages/result_selection/result_selection.html", "/ui/method_pages/result_selection/result_selection.css"],
    ["../ui/shell/browsing_history.html", "/ui/shell/browsing_history.css"],
    ["../ui/agent_guide/agent_guide.html", "/ui/agent_guide/agent_guide.css"],
  ];
  for (const [path, stylesheet] of extractedPages) {
    const html = read(path);
    assert.match(html, new RegExp(stylesheet.replaceAll("/", "\\/")));
    assert.doesNotMatch(html, /<style>/);
  }
});
