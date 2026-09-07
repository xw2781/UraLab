import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const componentUrl = new URL("../ui/shared/tabs/links/links_tab.js", import.meta.url);
const componentSource = await readFile(componentUrl, "utf8");
const stylesheetUrl = new URL("../ui/shared/tabs/links/links_tab.css", import.meta.url);
const stylesheetSource = await readFile(stylesheetUrl, "utf8");
const spreadsheetStylesheetSource = await readFile(
  new URL("../ui/shared/components/spreadsheet/spreadsheet_table.css", import.meta.url),
  "utf8",
);
// The real positioner needs a layout engine; record the placement request instead.
const contextMenuStubSource = `
  export const openedMenus = [];
  export function openContextMenu(menu, opts = {}) {
    if (!menu) return;
    menu.style.display = "block";
    openedMenus.push({ menu, opts });
  }
`;
const contextMenuStubUrl = `data:text/javascript;base64,${Buffer.from(contextMenuStubSource).toString("base64")}`;
const openPathStubSource = `
  export function openPathThroughDesktopHost() {
    return Promise.resolve({ ok: true });
  }
  export function revealPathThroughDesktopHost() {
    return Promise.resolve({ ok: true });
  }
  export function copyTextToClipboard() {
    return Promise.resolve({ ok: true });
  }
`;
const openPathStubUrl = `data:text/javascript;base64,${Buffer.from(openPathStubSource).toString("base64")}`;
const testableComponentSource = componentSource
  .replace(
    '"/ui/shared/components/context_menu/context_menu.js"',
    JSON.stringify(contextMenuStubUrl),
  )
  .replace(
    '"/ui/shared/integrations/open_path.js?v=20260907b"',
    JSON.stringify(openPathStubUrl),
  );
const linksTab = await import(
  `data:text/javascript;base64,${Buffer.from(testableComponentSource).toString("base64")}`
);

class FakeClassList {
  constructor(owner) {
    this.owner = owner;
  }

  tokens() {
    return new Set(String(this.owner.className || "").split(/\s+/u).filter(Boolean));
  }

  write(tokens) {
    this.owner.className = Array.from(tokens).join(" ");
  }

  add(...names) {
    const tokens = this.tokens();
    names.forEach((name) => tokens.add(name));
    this.write(tokens);
  }

  remove(...names) {
    const tokens = this.tokens();
    names.forEach((name) => tokens.delete(name));
    this.write(tokens);
  }

  contains(name) {
    return this.tokens().has(name);
  }

  toggle(name, force) {
    const tokens = this.tokens();
    const enabled = force === undefined ? !tokens.has(name) : Boolean(force);
    if (enabled) tokens.add(name);
    else tokens.delete(name);
    this.write(tokens);
    return enabled;
  }
}

class FakeElement {
  constructor(tagName, ownerDocument) {
    this.tagName = String(tagName).toUpperCase();
    this.ownerDocument = ownerDocument;
    this.children = [];
    this.parentElement = null;
    this.attributes = new Map();
    this.listeners = new Map();
    this.className = "";
    this.classList = new FakeClassList(this);
    this.style = {};
    this.dataset = {};
    this.textContent = "";
    this.hidden = false;
    this.disabled = false;
    this.title = "";
    this.tabIndex = -1;
    this.offsetWidth = 0;
    this.clientWidth = 0;
    this.offsetHeight = 0;
    this.clientHeight = 0;
    this.scrollWidth = 0;
    this.scrollHeight = 0;
  }

  appendChild(child) {
    child.parentElement = this;
    this.children.push(child);
    return child;
  }

  replaceChildren(...children) {
    this.children.forEach((child) => {
      child.parentElement = null;
    });
    this.children = [];
    children.forEach((child) => this.appendChild(child));
  }

  setAttribute(name, value) {
    this.attributes.set(name, String(value));
    if (name === "class") this.className = String(value);
  }

  getAttribute(name) {
    if (this.attributes.has(name)) return this.attributes.get(name);
    if (name === "href" && this.href !== undefined) return String(this.href);
    return null;
  }

  removeAttribute(name) {
    this.attributes.delete(name);
  }

  toggleAttribute(name, force) {
    const enabled = force === undefined ? !this.attributes.has(name) : Boolean(force);
    if (enabled) this.setAttribute(name, "");
    else this.removeAttribute(name);
    return enabled;
  }

  addEventListener(type, listener) {
    if (!this.listeners.has(type)) this.listeners.set(type, new Set());
    this.listeners.get(type).add(listener);
  }

  removeEventListener(type, listener) {
    this.listeners.get(type)?.delete(listener);
  }

  getBoundingClientRect() {
    return { right: 0, bottom: 0 };
  }

  async dispatch(type, eventInit = {}) {
    const event = {
      type,
      currentTarget: this,
      target: this,
      ctrlKey: false,
      metaKey: false,
      shiftKey: false,
      preventDefault() {},
      stopPropagation() {},
      ...eventInit,
    };
    const results = Array.from(this.listeners.get(type) || []).map((listener) => listener(event));
    await Promise.all(results);
  }

  async click(eventInit = {}) {
    if (this.disabled) return;
    await this.dispatch("click", eventInit);
  }

  contains(candidate) {
    if (candidate === this) return true;
    return this.children.some((child) => child.contains(candidate));
  }

  remove() {
    if (!this.parentElement) return;
    const siblings = this.parentElement.children;
    const index = siblings.indexOf(this);
    if (index >= 0) siblings.splice(index, 1);
    this.parentElement = null;
  }
}

class FakeDocument {
  constructor() {
    this.defaultView = globalThis;
    this.documentElement = new FakeElement("html", this);
    this.head = new FakeElement("head", this);
    this.body = new FakeElement("body", this);
    this.documentElement.appendChild(this.head);
    this.documentElement.appendChild(this.body);
  }

  createElement(tagName) {
    return new FakeElement(tagName, this);
  }

  createElementNS(_namespace, tagName) {
    return this.createElement(tagName);
  }

  allElements() {
    const elements = [];
    const visit = (element) => {
      elements.push(element);
      element.children.forEach(visit);
    };
    visit(this.documentElement);
    return elements;
  }

  getElementById(id) {
    return this.allElements().find((element) => element.id === id) || null;
  }

  querySelectorAll(selector) {
    if (selector !== 'link[rel="stylesheet"]') return [];
    return this.allElements().filter(
      (element) => element.tagName === "LINK" && element.rel === "stylesheet",
    );
  }
}

function descendants(root) {
  const elements = [];
  const visit = (element) => {
    elements.push(element);
    element.children.forEach(visit);
  };
  visit(root);
  return elements;
}

function byTag(root, tagName) {
  const normalized = String(tagName).toUpperCase();
  return descendants(root).filter((element) => element.tagName === normalized);
}

function byClass(root, className) {
  return descendants(root).filter((element) => element.classList.contains(className));
}

/** Every rendered row, in page order across the ArcRho and Excel sections. */
function allRows(root) {
  return byTag(root, "tbody").flatMap((body) => body.children);
}

function renderedText(element) {
  return [element.textContent, ...element.children.map(renderedText)].join("");
}

function menuOf(documentRef) {
  return byClass(documentRef.body, "arExternalLinksMenu")[0];
}

function menuItem(documentRef, action) {
  return descendants(menuOf(documentRef)).find((element) => element.dataset?.action === action);
}

function isMenuOpen(documentRef) {
  return menuOf(documentRef).style.display === "block";
}

function visibleMenuLabels(documentRef) {
  return descendants(menuOf(documentRef))
    .filter((element) => element.classList.contains("ctx-item") && !element.hidden)
    .map(renderedText);
}

/** Right-clicks the split below the rows, which never changes the selection. */
async function openHostMenu(documentRef, container) {
  const splitHost = byClass(container.children[0], "arLinksSplit")[0];
  await splitHost.dispatch("contextmenu", { clientX: 40, clientY: 60 });
  return splitHost;
}

async function runMenuAction(documentRef, container, action) {
  await openHostMenu(documentRef, container);
  await menuItem(documentRef, action).click();
}

function setup(options = {}) {
  const documentRef = new FakeDocument();
  const container = documentRef.createElement("div");
  documentRef.body.appendChild(container);
  const controller = linksTab.createLinksTab({
    container,
    documentRef,
    ariaLabel: "Dataset links",
    emptyDescription: "No cells are linked.",
    noun: "external links",
    getLinks: options.getLinks || (() => []),
    onRefreshLinks: options.onRefreshLinks || (() => ({ ok: true })),
    onBreakLinks: options.onBreakLinks || (() => ({ ok: true })),
    onOpenWorkbook: options.onOpenWorkbook || (() => ({ ok: true })),
    onRevealWorkbook: options.onRevealWorkbook || (() => ({ ok: true })),
    onCopyPath: options.onCopyPath || (() => ({ ok: true })),
    onOpenDataset: options.onOpenDataset || null,
    onStatus: options.onStatus,
  });
  return { documentRef, container, controller };
}

const sampleLink = {
  id: "link-1",
  workbookPath: "C:\\Claims\\Quarterly Book.xlsx",
  worksheet: "Paid Loss",
  address: "A1:C2",
  destination: "2024 / 12m",
  value: "125.4 ...",
  affectedCellCount: 6,
};

const sampleLinks = [
  sampleLink,
  { ...sampleLink, id: "link-2", address: "D4", destination: "2025 / 12m", value: "90" },
  { ...sampleLink, id: "link-3", address: "E5", destination: "2026 / 12m", value: "80" },
  { ...sampleLink, id: "link-4", address: "F6", destination: "2027 / 12m", value: "70" },
];

const internalLink = {
  id: "internal-1",
  sourceKind: "internal",
  datasetName: "C 82 - Prior Qtr Selected",
  sourceRange: "1:7",
  destination: "2017 / Value + 6 more",
  value: "14802...",
  affectedCellCount: 7,
};

const formulaLink = {
  id: "formula-1#component-0",
  sourceKind: "formula",
  formula: "=[C 82 - Prior Qtr Selected][1:7] * 2",
  datasetName: "C 82 - Prior Qtr Selected",
  destination: "2017 / Value + 6 more",
  value: "29604...",
  affectedCellCount: 7,
};

test("renders one compact row per link and offers the bulk actions through the context menu", async () => {
  const { documentRef, container, controller } = setup({ getLinks: () => [sampleLink] });

  assert.equal(await controller.refresh(), true);
  assert.equal(documentRef.querySelectorAll('link[rel="stylesheet"]').length, 1);
  assert.equal(container.classList.contains("arExternalLinksMount"), true);

  const root = container.children[0];
  assert.equal(byClass(root, "arExternalLinksToolbar").length, 0);
  assert.equal(byTag(root, "button").length, 0);

  const menu = menuOf(documentRef);
  assert.equal(menu.parentElement, documentRef.body);
  assert.equal(menu.getAttribute("role"), "menu");
  assert.equal(isMenuOpen(documentRef), false);
  assert.deepEqual(
    byTag(menu, "button").map(renderedText),
    [
      "Open workbook",
      "Open workbook as Read-Only",
      "Copy file path",
      "Open file location",
      "Open source dataset",
      "Refresh selected",
      "Break selected",
      "Refresh all",
      "Break all",
    ],
  );

  // Nothing selected: only the all-scope entries, and no separator above them.
  await openHostMenu(documentRef, container);
  assert.equal(isMenuOpen(documentRef), true);
  assert.deepEqual(visibleMenuLabels(documentRef), ["Refresh all", "Break all"]);
  assert.equal(byClass(menu, "ctx-sep")[0].hidden, true);

  // Two framed panes in page order, ArcRho above Excel, with the divider
  // between them; both stay in view even while one of them has no rows.
  const splitHost = byClass(root, "arLinksSplit")[0];
  assert.equal(splitHost.hidden, false);
  const panes = byClass(root, "arExternalLinksScroll");
  assert.equal(panes.length, 2);
  assert.deepEqual(panes.map((pane) => pane.dataset.linkKind), ["internal", "excel"]);
  const divider = byClass(root, "arLinksDivider")[0];
  assert.deepEqual(splitHost.children, [panes[0], divider, panes[1]]);
  assert.equal(divider.getAttribute("role"), "separator");
  const sections = byClass(root, "arLinksSection");
  assert.deepEqual(sections.map((section) => section.parentElement), panes);
  assert.deepEqual(sections.map((section) => section.dataset.linkKind), ["internal", "excel"]);
  // Each section is headed by a stamp in its kind's colour that reads as the
  // kind's name; it is the only place the kind is written.
  const stamps = sections.map((section) => byClass(section, "arLinksSectionStamp")[0]);
  assert.deepEqual(stamps.map(renderedText), ["ArcRho", "Excel"]);
  assert.deepEqual(
    stamps.map((stamp) => stamp.className),
    [
      "arLinksSectionStamp arLinksKind arLinksKind-internal",
      "arLinksSectionStamp arLinksKind arLinksKind-excel",
    ],
  );
  assert.equal(byClass(root, "arLinksSectionTitle").length, 0);
  assert.deepEqual(sections.map((section) => section.hidden), [false, false]);
  assert.deepEqual(byTag(sections[0], "tbody")[0].children, []);

  const internalTable = byTag(sections[0], "table")[0];
  assert.deepEqual(
    byTag(internalTable, "th").map(renderedText),
    ["Source Dataset", "Cell Range", "Destination", "Total Cells"],
  );

  const table = byTag(sections[1], "table")[0];
  assert.deepEqual(
    byTag(table, "th").map(renderedText),
    ["Source Workbook", "Cell Address", "Destination", "Total Cells"],
  );
  assert.equal(table.getAttribute("role"), "grid");
  assert.equal(table.getAttribute("aria-multiselectable"), "true");
  assert.equal(table.getAttribute("aria-label"), "Dataset links: Excel Links");
  // Every column has its explicit width, and every table is their sum, so
  // the sections' columns always line up.
  for (const candidate of byTag(root, "table")) {
    assert.deepEqual(byTag(candidate, "col").map((col) => col.style.width), ["260px", "200px", "200px", "64px"]);
    assert.equal(candidate.style.width, "724px");
    assert.equal(candidate.style.minWidth, "724px");
  }

  const row = allRows(root)[0];
  const cells = row.children;
  assert.equal(row.parentElement, byTag(table, "tbody")[0]);
  assert.equal(cells.length, 4);
  assert.equal(row.getAttribute("aria-selected"), "false");
  assert.equal(row.dataset.linkKind, "excel");
  assert.equal(byClass(cells[0], "arLinksKind").length, 0);
  assert.equal(byClass(cells[0], "arLinksCellText")[0].textContent, "Quarterly Book.xlsx");
  assert.equal(renderedText(cells[1]), "Paid Loss!A1:C2");
  assert.equal(renderedText(cells[2]), "2024 / 12m");
  assert.equal(renderedText(cells[3]), "6");
  assert.equal(byClass(root, "arExternalLinksState")[0].hidden, true);
});

test("ArcRho links sit in the top section, Excel below, and a formula's rows join the sections of what it reads", async () => {
  const opened = [];
  const openedWorkbooks = [];
  // The page lists a formula once per source it reads; this is the workbook
  // row of a formula that also reads a dataset.
  const workbookFormulaLink = {
    id: "formula-2#component-1",
    sourceKind: "formula",
    formula: "=[C 82 - Prior Qtr Selected][1:7] * 'C:\\Claims\\[Quarterly Book.xlsx]Paid Loss'!A1",
    workbookPath: "C:\\Claims\\Quarterly Book.xlsx",
    destination: "2017 / Value + 6 more",
    affectedCellCount: 7,
  };
  const { documentRef, container, controller } = setup({
    getLinks: () => [sampleLink, internalLink, formulaLink, workbookFormulaLink],
    onOpenDataset: (record) => {
      opened.push(record.datasetName);
      return { ok: true };
    },
    onOpenWorkbook: (path) => {
      openedWorkbooks.push(path);
      return { ok: true };
    },
  });
  await controller.refresh();

  const root = container.children[0];
  const sections = byClass(root, "arLinksSection");
  assert.deepEqual(sections.map((section) => section.hidden), [false, false]);
  // Rows follow the section order on the page, not the order the page gave
  // them: a formula's dataset row files under ArcRho and its workbook row
  // under Excel, each after the plain links of its section.
  const rows = allRows(root);
  assert.deepEqual(rows.map((row) => row.dataset.linkKind), ["internal", "internal", "excel", "excel"]);
  assert.deepEqual(
    rows.map((row) => row.parentElement),
    [0, 0, 1, 1].map((index) => byTag(sections[index], "tbody")[0]),
  );
  // The section stamp names the kind, so no row carries a chip of its own.
  for (const row of rows) assert.equal(byClass(row.children[0], "arLinksKind").length, 0);
  assert.equal(byClass(rows[0].children[0], "arLinksCellText")[0].textContent, "C 82 - Prior Qtr Selected");
  assert.equal(renderedText(rows[0].children[1]), "[1:7]");
  assert.equal(renderedText(rows[0].children[2]), internalLink.destination);
  assert.equal(byClass(rows[1].children[0], "arLinksCellText")[0].textContent, "C 82 - Prior Qtr Selected");
  assert.equal(renderedText(rows[1].children[1]), formulaLink.formula);
  assert.equal(renderedText(rows[1].children[3]), "7");
  assert.equal(renderedText(rows[2].children[0]), "Quarterly Book.xlsx");
  assert.equal(renderedText(rows[2].children[1]), "Paid Loss!A1:C2");
  assert.equal(renderedText(rows[3].children[0]), "Quarterly Book.xlsx");
  assert.equal(renderedText(rows[3].children[1]), workbookFormulaLink.formula);

  await rows[0].dispatch("contextmenu", { clientX: 12, clientY: 24 });
  assert.deepEqual(visibleMenuLabels(documentRef), [
    "Open source dataset",
    "Refresh selected",
    "Break selected",
    "Refresh all",
    "Break all",
  ]);
  await menuItem(documentRef, "open-dataset").click();
  assert.deepEqual(opened, ["C 82 - Prior Qtr Selected"]);

  // A formula's rows open the source each one stands for, like a plain link.
  await rows[1].dispatch("contextmenu", { clientX: 12, clientY: 24 });
  assert.deepEqual(visibleMenuLabels(documentRef), [
    "Open source dataset",
    "Refresh selected",
    "Break selected",
    "Refresh all",
    "Break all",
  ]);
  await menuItem(documentRef, "open-dataset").click();
  assert.deepEqual(opened, ["C 82 - Prior Qtr Selected", "C 82 - Prior Qtr Selected"]);
  await rows[3].dispatch("contextmenu", { clientX: 12, clientY: 24 });
  assert.deepEqual(visibleMenuLabels(documentRef), [
    "Open workbook",
    "Open workbook as Read-Only",
    "Copy file path",
    "Open file location",
    "Refresh selected",
    "Break selected",
    "Refresh all",
    "Break all",
  ]);
  await menuItem(documentRef, "open-workbook").click();
  assert.deepEqual(openedWorkbooks, ["C:\\Claims\\Quarterly Book.xlsx"]);
});

test("plain, Ctrl, Meta, and Shift clicks provide accessible multi-row selection", async () => {
  let refreshedRecords = [];
  const { documentRef, container, controller } = setup({
    getLinks: () => sampleLinks,
    onRefreshLinks: (records) => {
      refreshedRecords = records;
      return { ok: true };
    },
  });
  await controller.refresh();

  const root = container.children[0];
  const rows = allRows(root);

  await rows[0].click();
  assert.deepEqual(rows.map((row) => row.getAttribute("aria-selected")), ["true", "false", "false", "false"]);
  await openHostMenu(documentRef, container);
  assert.deepEqual(
    visibleMenuLabels(documentRef),
    ["Refresh selected", "Break selected", "Refresh all", "Break all"],
  );
  assert.equal(byClass(menuOf(documentRef), "ctx-sep")[1].hidden, false);

  await rows[2].click({ ctrlKey: true });
  assert.deepEqual(rows.map((row) => row.getAttribute("aria-selected")), ["true", "false", "true", "false"]);

  await rows[3].click({ shiftKey: true });
  assert.deepEqual(rows.map((row) => row.getAttribute("aria-selected")), ["false", "false", "true", "true"]);

  await rows[1].click({ metaKey: true });
  assert.deepEqual(rows.map((row) => row.getAttribute("aria-selected")), ["false", "true", "true", "true"]);
  await runMenuAction(documentRef, container, "refresh-selected");
  assert.deepEqual(refreshedRecords.map((record) => record.id), ["link-2", "link-3", "link-4"]);
});

test("the all-scope entries stay available and ignore the current selection", async () => {
  let refreshedRecords = [];
  let brokenRecords = [];
  const { documentRef, container, controller } = setup({
    getLinks: () => sampleLinks,
    onRefreshLinks: (links) => {
      refreshedRecords = links;
      return { ok: true };
    },
    onBreakLinks: (links) => {
      brokenRecords = links;
      return { ok: true };
    },
  });
  await controller.refresh();

  const rows = allRows(container.children[0]);
  await rows[1].click();

  await runMenuAction(documentRef, container, "refresh-all");
  assert.deepEqual(refreshedRecords.map((record) => record.id), sampleLinks.map((record) => record.id));

  await runMenuAction(documentRef, container, "break-all");
  assert.deepEqual(brokenRecords.map((record) => record.id), sampleLinks.map((record) => record.id));

  // The selection survives an all-scope action.
  const rerenderedRows = allRows(container.children[0]);
  assert.deepEqual(
    rerenderedRows.map((row) => row.getAttribute("aria-selected")),
    ["false", "true", "false", "false"],
  );
});

test("right-clicking a row keeps an existing selection but claims an unselected row", async () => {
  let refreshedRecords = [];
  const { documentRef, container, controller } = setup({
    getLinks: () => sampleLinks,
    onRefreshLinks: (records) => {
      refreshedRecords = records;
      return { ok: true };
    },
  });
  await controller.refresh();

  const rows = allRows(container.children[0]);
  await rows[0].click();
  await rows[1].click({ ctrlKey: true });

  // Inside the selection: the selection survives and the action stays scoped to it.
  await rows[1].dispatch("contextmenu", { clientX: 12, clientY: 24 });
  assert.equal(isMenuOpen(documentRef), true);
  assert.deepEqual(rows.map((row) => row.getAttribute("aria-selected")), ["true", "true", "false", "false"]);
  assert.equal(menuItem(documentRef, "break-selected").hidden, false);
  await menuItem(documentRef, "refresh-selected").click();
  assert.deepEqual(refreshedRecords.map((record) => record.id), ["link-1", "link-2"]);
  assert.equal(isMenuOpen(documentRef), false);

  // Outside the selection: the target row becomes the whole selection first.
  const refreshedRows = allRows(container.children[0]);
  await refreshedRows[3].dispatch("contextmenu", { clientX: 12, clientY: 24 });
  assert.deepEqual(
    refreshedRows.map((row) => row.getAttribute("aria-selected")),
    ["false", "false", "false", "true"],
  );
  await menuItem(documentRef, "refresh-selected").click();
  assert.deepEqual(refreshedRecords.map((record) => record.id), ["link-4"]);
});

test("row context menu opens its workbook normally or read-only from the top entries", async () => {
  const opened = [];
  const statuses = [];
  const { documentRef, container, controller } = setup({
    getLinks: () => sampleLinks,
    onOpenWorkbook: (path, options) => {
      opened.push({ path, ...options });
      return { ok: true };
    },
    onStatus: (message, tone) => statuses.push({ message, tone }),
  });
  await controller.refresh();

  const rows = allRows(container.children[0]);
  await rows[1].dispatch("contextmenu", { clientX: 12, clientY: 24 });
  assert.deepEqual(visibleMenuLabels(documentRef), [
    "Open workbook",
    "Open workbook as Read-Only",
    "Copy file path",
    "Open file location",
    "Refresh selected",
    "Break selected",
    "Refresh all",
    "Break all",
  ]);
  await menuItem(documentRef, "open-workbook").click();

  await rows[1].dispatch("contextmenu", { clientX: 12, clientY: 24 });
  await menuItem(documentRef, "open-workbook-read-only").click();

  assert.deepEqual(opened, [
    { path: sampleLinks[1].workbookPath, readOnly: false },
    { path: sampleLinks[1].workbookPath, readOnly: true },
  ]);
  assert.deepEqual(statuses, [
    { message: "Workbook opened.", tone: "success" },
    { message: "Workbook opened read-only.", tone: "success" },
  ]);
});

test("an Excel row copies its workbook path and opens the folder holding it", async () => {
  const copied = [];
  const revealed = [];
  const statuses = [];
  const { documentRef, container, controller } = setup({
    getLinks: () => sampleLinks,
    onCopyPath: (text) => {
      copied.push(text);
      return { ok: true };
    },
    onRevealWorkbook: (path) => {
      revealed.push(path);
      return { ok: true };
    },
    onStatus: (message, tone) => statuses.push({ message, tone }),
  });
  await controller.refresh();

  const rows = allRows(container.children[0]);
  await rows[1].dispatch("contextmenu", { clientX: 12, clientY: 24 });
  await menuItem(documentRef, "copy-workbook-path").click();

  await rows[1].dispatch("contextmenu", { clientX: 12, clientY: 24 });
  await menuItem(documentRef, "open-workbook-location").click();

  assert.deepEqual(copied, [sampleLinks[1].workbookPath]);
  assert.deepEqual(revealed, [sampleLinks[1].workbookPath]);
  assert.deepEqual(statuses, [
    { message: "File path copied.", tone: "success" },
    { message: "File location opened.", tone: "success" },
  ]);
});

test("selection is retained by link id and pruned when records disappear", async () => {
  let records = sampleLinks;
  const { documentRef, container, controller } = setup({ getLinks: () => records });
  await controller.refresh();
  let rows = allRows(container.children[0]);
  await rows[0].click();
  await rows[2].click({ ctrlKey: true });

  records = [sampleLinks[0], sampleLinks[1]];
  await controller.refresh();

  rows = allRows(container.children[0]);
  assert.deepEqual(rows.map((row) => row.getAttribute("aria-selected")), ["true", "false"]);
  await openHostMenu(documentRef, container);
  assert.equal(menuItem(documentRef, "refresh-selected").hidden, false);

  records = [sampleLinks[1]];
  await controller.refresh();
  rows = allRows(container.children[0]);
  assert.equal(rows[0].getAttribute("aria-selected"), "false");
  await openHostMenu(documentRef, container);
  assert.deepEqual(visibleMenuLabels(documentRef), ["Refresh all", "Break all"]);
});

test("Refresh all uses the bulk callback, exposes busy state, and reports success", async () => {
  let resolveRefresh;
  let refreshCount = 0;
  let received = [];
  const statuses = [];
  const { documentRef, container, controller } = setup({
    getLinks: () => {
      refreshCount += 1;
      return sampleLinks;
    },
    onRefreshLinks: (records) => {
      received = records;
      return new Promise((resolve) => {
        resolveRefresh = resolve;
      });
    },
    onStatus: (message, tone) => statuses.push({ message, tone }),
  });
  await controller.refresh();

  const root = container.children[0];
  await openHostMenu(documentRef, container);
  const pendingClick = menuItem(documentRef, "refresh-all").click();
  assert.equal(isMenuOpen(documentRef), false);
  assert.equal(root.getAttribute("aria-busy"), "true");
  assert.match(renderedText(byClass(root, "arExternalLinksState")[0]), /Refreshing external links/u);
  assert.deepEqual(received.map((record) => record.id), sampleLinks.map((record) => record.id));

  // A right-click while an action runs must not reopen the menu.
  await openHostMenu(documentRef, container);
  assert.equal(isMenuOpen(documentRef), false);

  resolveRefresh({ ok: true, message: "Workbook values refreshed." });
  await pendingClick;

  assert.equal(refreshCount, 2);
  assert.equal(root.getAttribute("aria-busy"), "false");
  await openHostMenu(documentRef, container);
  assert.equal(isMenuOpen(documentRef), true);
  assert.deepEqual(visibleMenuLabels(documentRef), ["Refresh all", "Break all"]);
  assert.deepEqual(statuses.at(-1), { message: "Workbook values refreshed.", tone: "success" });
});

test("Break all excludes read-only records and hides the break action when none remain", async () => {
  const readOnlyLink = { ...sampleLinks[3], readOnly: true };
  let records = [...sampleLinks.slice(0, 2), readOnlyLink];
  let received = [];
  const { documentRef, container, controller } = setup({
    getLinks: () => records,
    onBreakLinks: (links) => {
      received = links;
      records = [readOnlyLink];
      return { ok: true, message: "Link values are now hard-coded." };
    },
  });
  await controller.refresh();

  const root = container.children[0];
  await runMenuAction(documentRef, container, "break-all");

  assert.deepEqual(received.map((record) => record.id), ["link-1", "link-2"]);
  assert.equal(allRows(root).length, 1);

  await openHostMenu(documentRef, container);
  assert.equal(isMenuOpen(documentRef), true);
  assert.deepEqual(visibleMenuLabels(documentRef), ["Refresh all"]);
});

test("failed bulk actions retain rows, selection, and restored controls", async () => {
  let resolveBreak;
  const statuses = [];
  const { documentRef, container, controller } = setup({
    getLinks: () => sampleLinks,
    onBreakLinks: () => new Promise((resolve) => {
      resolveBreak = resolve;
    }),
    onStatus: (message, tone) => statuses.push({ message, tone }),
  });
  await controller.refresh();

  const root = container.children[0];
  const rows = allRows(root);
  await rows[1].click();
  await openHostMenu(documentRef, container);
  const pendingClick = menuItem(documentRef, "break-selected").click();
  assert.equal(isMenuOpen(documentRef), false);
  assert.match(renderedText(byClass(root, "arExternalLinksState")[0]), /Breaking external links/u);

  resolveBreak({ ok: false, error: "Workbook is locked." });
  await pendingClick;

  assert.equal(allRows(root).length, 4);
  assert.equal(rows[1].getAttribute("aria-selected"), "true");
  assert.equal(root.getAttribute("aria-busy"), "false");
  await openHostMenu(documentRef, container);
  assert.equal(menuItem(documentRef, "break-selected").hidden, false);
  const state = byClass(root, "arExternalLinksState")[0];
  assert.equal(state.hidden, false);
  assert.equal(state.getAttribute("role"), "alert");
  assert.equal(state.classList.contains("hasRows"), true);
  assert.match(renderedText(state), /Workbook is locked\./u);
  assert.deepEqual(statuses.at(-1), { message: "Workbook is locked.", tone: "error" });
});

test("partial refresh failures are reported instead of shown as success", async () => {
  const statuses = [];
  const { documentRef, container, controller } = setup({
    getLinks: () => [sampleLink],
    onRefreshLinks: () => ({ linkedCellCount: 2, failedCount: 2 }),
    onStatus: (message, tone) => statuses.push({ message, tone }),
  });
  await controller.refresh();

  const root = container.children[0];
  await runMenuAction(documentRef, container, "refresh-all");

  assert.match(renderedText(byClass(root, "arExternalLinksState")[0]), /2 linked values could not be refreshed/u);
  assert.deepEqual(statuses.at(-1), {
    message: "2 linked values could not be refreshed.",
    tone: "error",
  });
});

test("explicit loading and error states retain rendered rows and destroy cleanly", async () => {
  const { documentRef, container, controller } = setup({ getLinks: () => [sampleLink] });
  await controller.refresh();
  const root = container.children[0];
  const menu = menuOf(documentRef);

  controller.setLoading("Refreshing workbook values...");
  assert.equal(root.getAttribute("aria-busy"), "true");
  await openHostMenu(documentRef, container);
  assert.equal(isMenuOpen(documentRef), false);
  assert.equal(allRows(root).length, 1);
  assert.match(renderedText(byClass(root, "arExternalLinksState")[0]), /Refreshing workbook values/u);

  controller.setError("Excel is unavailable.");
  assert.equal(root.getAttribute("aria-busy"), "false");
  await openHostMenu(documentRef, container);
  assert.equal(isMenuOpen(documentRef), true);
  assert.equal(allRows(root).length, 1);
  assert.match(renderedText(byClass(root, "arExternalLinksState")[0]), /Excel is unavailable\./u);

  controller.destroy();
  assert.equal(container.children.length, 0);
  assert.equal(container.classList.contains("arExternalLinksMount"), false);
  assert.equal(menu.parentElement, null);
  assert.equal(menu.style.display, "none");
  assert.equal(await controller.refresh(), false);
});

test("persistent warnings survive link refreshes until explicitly cleared", async () => {
  const { container, controller } = setup({ getLinks: () => [sampleLink] });
  await controller.refresh();
  const root = container.children[0];
  const state = byClass(root, "arExternalLinksState")[0];

  controller.setWarning(
    "Saved Excel values may be out of date",
    "2 stale linked values. Stored values remain active.",
  );
  assert.equal(state.classList.contains("isWarning"), true);
  assert.match(renderedText(state), /2 stale linked values/u);

  await controller.refresh();
  assert.equal(state.classList.contains("isWarning"), true);
  assert.match(renderedText(state), /Saved Excel values may be out of date/u);

  controller.clearWarning();
  assert.equal(state.hidden, true);
});

test("dragging a header edge resizes only that column and the table follows", async () => {
  const { container, controller } = setup({ getLinks: () => [sampleLink] });
  await controller.refresh();
  const table = byTag(container.children[0], "table")[0];
  const referenceHeader = byTag(table, "th").find((header) => header.dataset.colKey === "reference");
  const resizer = byClass(referenceHeader, "arLinksColResizer")[0];

  await resizer.dispatch("pointerdown", { clientX: 100, pointerId: 1 });
  await resizer.dispatch("pointermove", { clientX: 160 });
  assert.equal(container.children[0].classList.contains("isResizingColumn"), true);
  await resizer.dispatch("pointerup", { clientX: 160 });

  assert.equal(controller.getColumnWidth("reference"), 260);
  assert.equal(controller.getColumnWidth("source"), 260);
  // Every section's table follows the shared width.
  for (const candidate of byTag(container.children[0], "table")) {
    assert.equal(byTag(candidate, "col")[1].style.width, "260px");
    assert.equal(candidate.style.width, "784px");
  }
  assert.equal(container.children[0].classList.contains("isResizingColumn"), false);

  // A column cannot shrink below its minimum, can grow far past its default,
  // and a double-click restores the default.
  await resizer.dispatch("pointerdown", { clientX: 100, pointerId: 2 });
  await resizer.dispatch("pointermove", { clientX: -900 });
  await resizer.dispatch("pointerup", { clientX: -900 });
  assert.equal(controller.getColumnWidth("reference"), 90);
  await resizer.dispatch("pointerdown", { clientX: 100, pointerId: 3 });
  await resizer.dispatch("pointermove", { clientX: 2600 });
  await resizer.dispatch("pointerup", { clientX: 2600 });
  assert.equal(controller.getColumnWidth("reference"), 2590);
  await resizer.dispatch("dblclick");
  assert.equal(controller.getColumnWidth("reference"), 200);
  assert.equal(table.style.width, "724px");
});

test("the shared columns re-fit whenever the host resizes, and a short table keeps its right edge", async () => {
  const observers = [];
  class FakeResizeObserver {
    constructor(callback) {
      this.callback = callback;
      observers.push(this);
    }

    observe(target) {
      (this.targets ||= []).push(target);
    }

    disconnect() {
      this.disconnected = true;
    }
  }
  globalThis.ResizeObserver = FakeResizeObserver;
  try {
    const { container, controller } = setup({ getLinks: () => [sampleLink, internalLink] });
    await controller.refresh();
    const root = container.children[0];
    const panes = byClass(root, "arExternalLinksScroll");
    const [scrollHost] = panes;
    const tables = byTag(root, "table");
    assert.equal(panes.length, 2);
    assert.equal(observers.length, 1);
    assert.deepEqual(observers[0].targets, panes);

    // The panes gain width: the defaults stretch to the width the narrower
    // pane leaves once its own scrollbar has taken its lane, so neither pane
    // gains a horizontal scrollbar from the fit.
    panes[0].clientWidth = 1006;
    panes[1].clientWidth = 986;
    observers[0].callback();
    assert.equal(controller.getColumnWidth("source"), 354);
    assert.equal(controller.getColumnWidth("cells"), 88);
    for (const table of tables) assert.equal(table.style.width, "986px");
    assert.equal(scrollHost.classList.contains("isTableShort"), false);

    // The panes shrink: the defaults scale back down, never below themselves.
    for (const pane of panes) pane.clientWidth = 600;
    observers[0].callback();
    for (const table of tables) assert.equal(table.style.width, "724px");
    assert.equal(scrollHost.classList.contains("isTableShort"), false);

    // A dragged column ends the fitting; the tables keep their width and a
    // wider host only asks the last column for its own right edge.
    const resizer = byClass(byTag(tables[0], "th")[1], "arLinksColResizer")[0];
    await resizer.dispatch("pointerdown", { clientX: 100, pointerId: 1 });
    await resizer.dispatch("pointermove", { clientX: 60 });
    await resizer.dispatch("pointerup", { clientX: 60 });
    for (const pane of panes) pane.clientWidth = 1000;
    observers[0].callback();
    for (const table of tables) assert.equal(table.style.width, "684px");
    for (const pane of panes) assert.equal(pane.classList.contains("isTableShort"), true);

    // Restoring the default hands the column back to the fit, which resumes.
    await resizer.dispatch("dblclick");
    assert.equal(controller.getColumnWidth("reference"), 276);
    for (const table of tables) assert.equal(table.style.width, "1000px");
    assert.equal(scrollHost.classList.contains("isTableShort"), false);
    controller.destroy();
    assert.equal(observers[0].disconnected, true);
  } finally {
    delete globalThis.ResizeObserver;
  }
});

test("the divider starts at the middle, drags between the panes, and double-clicks back", async () => {
  const { container, controller } = setup({ getLinks: () => [sampleLink, internalLink] });
  await controller.refresh();
  const root = container.children[0];
  const splitHost = byClass(root, "arLinksSplit")[0];
  const panes = byClass(root, "arExternalLinksScroll");
  const divider = byClass(root, "arLinksDivider")[0];
  // Until it is dragged the stylesheet's equal growth places it; the module
  // writes nothing.
  assert.equal(controller.getSplitRatio(), 0.5);
  for (const pane of panes) assert.equal(pane.style.flexGrow, undefined);

  splitHost.getBoundingClientRect = () => ({ top: 100, height: 407 });
  divider.offsetHeight = 7;
  await divider.dispatch("pointerdown", { clientY: 303, pointerId: 1 });
  assert.equal(root.classList.contains("isResizingSplit"), true);
  // The divider follows the pointer from where it was grabbed: 100px up over
  // a 400px track (the split less the divider) is a quarter of the way.
  await divider.dispatch("pointermove", { clientY: 203 });
  assert.equal(controller.getSplitRatio(), 0.25);
  assert.equal(panes[0].style.flexGrow, "0.25");
  assert.equal(panes[1].style.flexGrow, "0.75");
  // Neither pane may be dragged shorter than its stamp and header row.
  await divider.dispatch("pointermove", { clientY: 100 });
  assert.equal(controller.getSplitRatio(), 0.16);
  await divider.dispatch("pointermove", { clientY: 900 });
  assert.equal(controller.getSplitRatio(), 0.84);
  await divider.dispatch("pointerup", { clientY: 900 });
  assert.equal(root.classList.contains("isResizingSplit"), false);

  // The place survives a re-mount within the page session, like a dragged
  // column width, and a double-click hands it back to the middle.
  controller.destroy();
  const remounted = setup({ getLinks: () => [sampleLink, internalLink] });
  await remounted.controller.refresh();
  const remountedPanes = byClass(remounted.container.children[0], "arExternalLinksScroll");
  assert.equal(remounted.controller.getSplitRatio(), 0.84);
  assert.equal(remountedPanes[0].style.flexGrow, "0.84");
  await byClass(remounted.container.children[0], "arLinksDivider")[0].dispatch("dblclick");
  assert.equal(remounted.controller.getSplitRatio(), 0.5);
  for (const pane of remountedPanes) assert.equal(pane.style.flexGrow, "");
  remounted.controller.destroy();
});

test("shared styling keeps the compact framed tables and colours the kind badges by link", () => {
  assert.doesNotMatch(stylesheetSource, /arExternalLinksToolbar/u);
  // One stamp per kind heads its section, each in the colour that kind's cells
  // wear in the Data tab grid, and a table narrower than its frame draws its
  // own right edge on the last column.
  assert.doesNotMatch(stylesheetSource, /arLinksSectionTitle/u);
  assert.match(stylesheetSource, /\.arLinksSectionStamp \{[^}]*position: sticky;/u);
  assert.match(stylesheetSource, /--ar-links-excel: var\(--ar-spreadsheet-excel-link-border, #217346\);/u);
  assert.match(stylesheetSource, /--ar-links-internal: var\(--ar-spreadsheet-internal-link-border, #2b6df6\);/u);
  assert.match(stylesheetSource, /\.arLinksKind-excel \{\s*--ar-links-kind: var\(--ar-links-excel\);/u);
  assert.match(stylesheetSource, /\.arLinksKind-internal \{\s*--ar-links-kind: var\(--ar-links-internal\);/u);
  // The framed scrollbar matches the Data and Audit Log tabs: a 20px lane with
  // step arrows, on the same chrome tokens the dark theme re-points.
  assert.match(stylesheetSource, /\.arExternalLinksScroll::-webkit-scrollbar \{\s*width: 20px;\s*height: 20px;/u);
  assert.match(stylesheetSource, /\.arExternalLinksScroll::-webkit-scrollbar-button \{\s*display: block;/u);
  assert.match(stylesheetSource, /--ar-links-scrollbar-chrome: #f1f3f5;/u);
  assert.match(
    stylesheetSource,
    /\.isTableShort \.arExternalLinksTable td:last-child \{\s*border-right: 1px solid #e2e8f0;/u,
  );
  // The kind badge reads as the name is written (ArcRho, not ARCRHO), and no
  // cell carries a tooltip.
  assert.doesNotMatch(stylesheetSource, /text-transform:\s*uppercase/u);
  assert.doesNotMatch(componentSource, /tooltip/iu);
  assert.doesNotMatch(stylesheetSource, /arLinksCell-value/u);
  // The last row must keep its bottom rule so the table does not look unfinished
  // when the rows are shorter than the framed scroll host.
  assert.doesNotMatch(stylesheetSource, /tbody tr:last-child td\s*\{[^}]*border-bottom:\s*0;/su);
  const cellRule = stylesheetSource.match(
    /\.arExternalLinksTable th,\s*\.arExternalLinksTable td\s*\{([^}]*)\}/u,
  )?.[1] || "";
  assert.match(cellRule, /border-bottom:\s*1px solid #e2e8f0;/u);
  assert.match(cellRule, /white-space:\s*nowrap;/u);
  assert.match(stylesheetSource, /border-collapse:\s*separate;/u);
  assert.match(stylesheetSource, /table-layout:\s*fixed;/u);
  assert.match(stylesheetSource, /position:\s*sticky;/u);
  assert.match(stylesheetSource, /height:\s*31px;/u);
  assert.match(stylesheetSource, /tbody tr\[aria-selected="true"\]/u);
  assert.match(stylesheetSource, /\.arLinksColResizer \{[^}]*cursor:\s*col-resize;/u);
  // With no Formula section there is no formula stamp, so no formula token.
  assert.doesNotMatch(stylesheetSource, /--ar-links-formula|arLinksKind-formula/u);
  assert.doesNotMatch(stylesheetSource, /td\.arExternalLinkCell\s*\{/u);
  assert.doesNotMatch(stylesheetSource, /td\.arExternalLinkAnchor::after/u);
  const arrayRule = spreadsheetStylesheetSource.match(
    /\.arSpreadsheetTable td\.arArrayFormulaCell\s*\{([^}]*)\}/u,
  )?.[1] || "";
  assert.match(arrayRule, /box-shadow:/u);
  assert.match(arrayRule, /transition:\s*box-shadow 150ms ease, border-color 150ms ease;/u);
  for (const edge of ["Top", "Right", "Bottom", "Left"]) {
    assert.match(
      spreadsheetStylesheetSource,
      new RegExp(`td\\.arArrayFormulaEdge${edge}\\s*\\{`, "u"),
    );
  }
  assert.match(spreadsheetStylesheetSource, /#2b6df6/u);
  assert.match(spreadsheetStylesheetSource, /rgba\(43, 109, 246, 0\.28\)/u);
  assert.match(stylesheetSource, /\.arExternalLinksState\.isWarning/u);
  assert.match(stylesheetSource, /color:\s*#b45309;/u);
});
