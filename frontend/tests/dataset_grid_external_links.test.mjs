import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const dataUrl = (source) => `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;
const referenceSource = await readFile(
  new URL("../ui/shared/integrations/excel_reference.js", import.meta.url),
  "utf8",
);
const spreadsheetStubUrl = dataUrl(`
  export function createSpreadsheetTableController() {
    return {
      applyDom() {}, clear() {}, copy() {}, move() { return false; },
      prepareContextCell() {}, selectCell() {}, selectColumn() {}, selectRow() {},
      selection() { return { ranges: [] }; }, setRange() {},
    };
  }
  export function getTopLeftRangeCell(ranges) { return ranges?.[0] ? { r: ranges[0].r0, c: ranges[0].c0 } : null; }
  export function normalizeRange(r0, c0, r1, c1) { return { r0, c0, r1, c1 }; }
`);
const viewStubUrl = dataUrl(`
  export function getDatasetGridSelectionLayout() { return { maxRow: 0, maxCol: 0 }; }
  export function getDisplayDatasetModel() { return globalThis.__arTestDisplayModel; }
  export function setDatasetGridEditConfig(config) { globalThis.__arTestGridEditConfig = config; }
`);
const formulaHoverStubUrl = dataUrl(`
  export function createFormulaHoverEditor(options) {
    globalThis.__arTestFormulaHoverOptions = options;
    const controller = {
      attached: [], openCalls: [],
      attach(cell, context, attachOptions) { this.attached.push({ cell, context, attachOptions }); return true; },
      open(cell, context, options) { this.openCalls.push({ cell, context, options }); return true; },
    };
    globalThis.__arTestFormulaHover = controller;
    return controller;
  }
`);

const internalReferenceUrl = dataUrl(await readFile(
  new URL("../ui/shared/dataset/dataset_internal_reference.js", import.meta.url),
  "utf8",
));
const datasetFormulaUrl = dataUrl((await readFile(
  new URL("../ui/shared/dataset/dataset_formula.js", import.meta.url),
  "utf8",
))
  .replace('"/ui/shared/integrations/excel_reference.js?v=20260715a"', JSON.stringify(dataUrl(referenceSource)))
  .replace('"/ui/shared/dataset/dataset_internal_reference.js?v=20260830a"', JSON.stringify(internalReferenceUrl)));

// The read-only refusal opens a page message box rather than a status line.
const messageBoxStubUrl = dataUrl(`
  export function showPageMessageBox(options) {
    globalThis.__arTestMessageBoxes = globalThis.__arTestMessageBoxes || [];
    globalThis.__arTestMessageBoxes.push(options);
    return Promise.resolve();
  }
`);

let interactionSource = await readFile(
  new URL("../ui/shared/tabs/data/dataset_grid_interactions.js", import.meta.url),
  "utf8",
);
const rawInteractionSource = interactionSource;
const dataTabControllerSource = await readFile(
  new URL("../ui/shared/tabs/data/data_tab_controller.js", import.meta.url),
  "utf8",
);
const gridViewImportPattern = /["']([^"']*\/dataset_grid_view\.js\?v=[^"']+)["']/u;

test("the grid renderer and interactions share one module instance", () => {
  const controllerImport = dataTabControllerSource.match(gridViewImportPattern)?.[1];
  const interactionsImport = rawInteractionSource.match(gridViewImportPattern)?.[1];

  assert.ok(controllerImport, "Data-tab controller must import the grid view");
  assert.equal(interactionsImport, controllerImport);
});

interactionSource = interactionSource
  .replace(
    '"/ui/shared/components/spreadsheet/spreadsheet_table.js?v=20260715a"',
    JSON.stringify(spreadsheetStubUrl),
  )
  .replace(
    '"/ui/shared/tabs/data/dataset_grid_view.js?v=20260907c"',
    JSON.stringify(viewStubUrl),
  )
  .replace(
    '"/ui/shared/integrations/excel_reference.js?v=20260715a"',
    JSON.stringify(dataUrl(referenceSource)),
  )
  .replace(
    '"/ui/shared/components/formula_hover/formula_hover.js?v=20260907a"',
    JSON.stringify(formulaHoverStubUrl),
  )
  .replace(
    '"/ui/shared/dataset/dataset_internal_reference.js?v=20260830a"',
    JSON.stringify(internalReferenceUrl),
  )
  .replace(
    '"/ui/shared/dataset/dataset_formula.js?v=20260830a"',
    JSON.stringify(datasetFormulaUrl),
  )
  .replace(
    '"/ui/shared/components/message_box/message_box.js?v=20260831a"',
    JSON.stringify(messageBoxStubUrl),
  );
const interactions = await import(dataUrl(interactionSource));

const EXCEL_REFERENCE = "='C:\\Data\\[Book.xlsx]Sheet 1'!A1:B2";

function classList() {
  const values = new Set();
  return {
    contains: (name) => values.has(name),
    toggle(name, force) {
      if (force) values.add(name);
      else values.delete(name);
    },
  };
}

function setup({
  commitExternalReference,
  decorateExternalLinkCell,
  getExternalLinkCellInfo,
  model,
} = {}) {
  const previousWindow = globalThis.window;
  const previousDocument = globalThis.document;
  const previousAnimationFrame = globalThis.requestAnimationFrame;
  const listeners = new Map();
  const postedMessages = [];
  globalThis.window = {
    parent: { postMessage(message) { postedMessages.push(message); } },
  };
  const tableWrap = {
    addEventListener() {},
    querySelector() { return null; },
  };
  globalThis.document = {
    activeElement: null,
    addEventListener(type, listener) {
      if (!listeners.has(type)) listeners.set(type, []);
      listeners.get(type).push(listener);
    },
    getElementById(id) {
      if (id === "transposedChk") return { checked: false };
      if (id === "tableWrap") return tableWrap;
      return null;
    },
    querySelector() { return null; },
  };
  globalThis.requestAnimationFrame = (callback) => callback();

  const state = {
    model: model || {
      origin_labels: ["2024"],
      dev_labels: ["12m"],
      values: [[5]],
      mask: [[true]],
    },
    dirty: new Map(),
    activeCell: { r: 0, c: 0 },
    selectionAnchor: { r: 0, c: 0 },
    selRanges: [{ r0: 0, c0: 0, r1: 0, c1: 0 }],
    showSubtotal: true,
  };
  globalThis.__arTestDisplayModel = state.model;
  const calls = {
    renders: 0,
    updates: 0,
    statuses: [],
    hardCoded: [],
    externalRequestCancellations: 0,
    settingsRefreshes: 0,
    events: [],
    postedMessages,
  };
  interactions.wireDatasetGridInteractions({
    state,
    renderTable: () => {
      calls.renders += 1;
      calls.events.push("render");
    },
    notifyDatasetUpdated: () => { calls.updates += 1; },
    refreshDatasetSettingsDirty: () => { calls.settingsRefreshes += 1; },
    setStatus: (message) => calls.statuses.push(message),
    commitExternalReference,
    cancelExternalReference: () => {
      calls.externalRequestCancellations += 1;
      calls.events.push("cancel");
    },
    hardCodeExternalLinkCells: (cells) => calls.hardCoded.push(cells),
    decorateExternalLinkCell,
    getExternalLinkCellInfo,
  });

  const cleanup = () => {
    globalThis.window = previousWindow;
    globalThis.document = previousDocument;
    globalThis.requestAnimationFrame = previousAnimationFrame;
    delete globalThis.__arTestDisplayModel;
    delete globalThis.__arTestGridEditConfig;
    delete globalThis.__arTestFormulaHover;
    delete globalThis.__arTestFormulaHoverOptions;
  };
  return {
    state,
    calls,
    config: globalThis.__arTestGridEditConfig,
    formulaHover: globalThis.__arTestFormulaHover,
    formulaHoverOptions: globalThis.__arTestFormulaHoverOptions,
    listeners,
    cleanup,
  };
}

function editInput() {
  const attributes = new Map();
  return {
    classList: classList(),
    isConnected: true,
    readOnly: false,
    focused: false,
    blur() {},
    focus() { this.focused = true; },
    removeAttribute(name) { attributes.delete(name); },
    setAttribute(name, value) { attributes.set(name, String(value)); },
    getAttribute(name) { return attributes.get(name) ?? null; },
  };
}

test("starting a DSV cell edit invalidates an in-flight Excel refresh", () => {
  const context = setup();
  try {
    context.config.onCellFocus(0, 0);

    assert.equal(context.calls.externalRequestCancellations, 1);
  } finally {
    context.cleanup();
  }
});

test("the grid context menu toggles the persisted subtotal setting", async () => {
  const context = setup();
  try {
    await context.config.onContextAction("toggle_subtotal");

    assert.equal(context.state.showSubtotal, false);
    assert.deepEqual(context.state.selRanges, []);
    assert.equal(context.state.activeCell, null);
    assert.equal(context.calls.renders, 1);
    assert.equal(context.calls.settingsRefreshes, 1);
  } finally {
    context.cleanup();
  }
});

test("keyboard edit initiation invalidates Excel refresh before rendering the editor", () => {
  const context = setup();
  try {
    const keydown = context.listeners.get("keydown")?.at(-1);
    assert.equal(typeof keydown, "function");

    keydown({
      key: "=",
      target: null,
      ctrlKey: false,
      metaKey: false,
      altKey: false,
      preventDefault() {},
    });

    assert.equal(context.calls.externalRequestCancellations, 1);
    assert.equal(context.calls.events[0], "cancel");
    assert.ok(context.calls.renders >= 1);
  } finally {
    context.cleanup();
  }
});

test("Excel-reference drafts do not overwrite the numeric DSV model before commit", () => {
  const context = setup();
  try {
    context.state.editingCell = { r: 0, c: 0 };
    context.config.onCellInput(0, 0, EXCEL_REFERENCE, editInput(), { dataset: {} });

    assert.equal(context.state.model.values[0][0], 5);
    assert.equal(context.state.dirty.size, 0);
    assert.equal(context.state.editingCell.pendingExternalReference, EXCEL_REFERENCE);
  } finally {
    context.cleanup();
  }
});

test("successful Excel-reference commits render only after the asynchronous load succeeds", async () => {
  const requests = [];
  const context = setup({
    commitExternalReference: async (request) => {
      requests.push(request);
      return { ok: true, affectedCellCount: 4 };
    },
  });
  try {
    context.state.editingCell = { r: 0, c: 0 };
    const input = editInput();
    await context.config.onCellCommit(0, 0, EXCEL_REFERENCE, input, { dataset: {} });

    assert.deepEqual(requests, [{ displayRow: 0, displayColumn: 0, reference: EXCEL_REFERENCE }]);
    assert.equal(context.state.editingCell, null);
    assert.equal(input.readOnly, false);
    assert.equal(input.getAttribute("aria-busy"), null);
    assert.equal(context.calls.renders, 1);
    assert.equal(context.calls.updates, 1);
    assert.deepEqual(context.calls.postedMessages, []);
    assert.match(context.calls.statuses.at(-1), /Linked 4 dataset cells/u);
  } finally {
    context.cleanup();
  }
});

test("failed Excel-reference commits retain the draft and do not render partial values", async () => {
  const context = setup({
    commitExternalReference: async () => ({ ok: false, error: "Workbook unavailable." }),
  });
  try {
    const edit = { r: 0, c: 0 };
    context.state.editingCell = edit;
    const input = editInput();
    await context.config.onCellCommit(0, 0, EXCEL_REFERENCE, input, { dataset: {} });

    assert.equal(context.state.editingCell, edit);
    assert.equal(input.focused, true);
    assert.equal(context.calls.renders, 0);
    assert.equal(context.calls.updates, 0);
    assert.equal(context.calls.statuses.at(-1), "Workbook unavailable.");
  } finally {
    context.cleanup();
  }
});

test("numeric commits break an overlapping DSV link before hard-coding the value", async () => {
  const context = setup();
  try {
    context.state.editingCell = { r: 0, c: 0 };
    const input = editInput();
    const cell = { dataset: {} };
    await context.config.onCellCommit(0, 0, "7", input, cell);

    assert.deepEqual(context.calls.hardCoded, [[{ r: 0, c: 0 }]]);
    assert.equal(context.state.model.values[0][0], 7);
    assert.equal(context.state.dirty.get("0,0"), 7);
    assert.equal(context.state.editingCell, null);
  } finally {
    context.cleanup();
  }
});

test("pasting an Excel reference updates the unsaved DSV without refreshing the PI dataset table", async () => {
  const requests = [];
  const context = setup({
    commitExternalReference: async (request) => {
      requests.push(request);
      return { ok: true, affectedCellCount: 1 };
    },
  });
  try {
    const paste = context.listeners.get("paste")?.at(-1);
    let prevented = false;
    paste({
      target: null,
      clipboardData: { getData: () => EXCEL_REFERENCE },
      preventDefault() { prevented = true; },
    });
    await new Promise((resolve) => setImmediate(resolve));

    assert.equal(prevented, true);
    assert.deepEqual(requests, [{ displayRow: 0, displayColumn: 0, reference: EXCEL_REFERENCE }]);
    assert.equal(context.calls.renders, 1);
    assert.equal(context.calls.updates, 1);
    assert.deepEqual(context.calls.postedMessages, []);
  } finally {
    context.cleanup();
  }
});

test("large matrix paste clips values outside the editable triangle", () => {
  const context = setup({
    model: {
      origin_labels: ["2023", "2024"],
      dev_labels: ["12m", "24m"],
      values: [[10, 20], [30, 40]],
      mask: [[true, true], [true, false]],
    },
  });
  try {
    const paste = context.listeners.get("paste")?.at(-1);
    let prevented = false;
    paste({
      target: null,
      clipboardData: { getData: () => "1\t2\t3\n4\t5\t6\n7\t8\t9" },
      preventDefault() { prevented = true; },
    });

    assert.equal(prevented, true);
    assert.deepEqual(context.state.model.values, [[1, 2], [4, 40]]);
    assert.deepEqual(context.state.selRanges, [{ r0: 0, c0: 0, r1: 1, c1: 1 }]);
    assert.equal(context.state.dirty.size, 3);
    assert.deepEqual(context.calls.postedMessages, []);
  } finally {
    context.cleanup();
  }
});

test("linked cells attach the reusable hover editor and commit edits at the link anchor", async () => {
  const nextReference = "='C:\\Data\\[Book.xlsx]Sheet 1'!C3:D4";
  const linkInfo = {
    reference: EXCEL_REFERENCE,
    anchorDisplayRow: 0,
    anchorDisplayColumn: 0,
  };
  const decorated = [];
  const requests = [];
  const context = setup({
    decorateExternalLinkCell: (cell, row, column) => decorated.push({ cell, row, column }),
    getExternalLinkCellInfo: () => linkInfo,
    commitExternalReference: async (request) => {
      requests.push(request);
      return { ok: true, affectedCellCount: 4 };
    },
  });
  try {
    const cell = {};
    context.config.decorateCell(cell, 0, 0);

    assert.deepEqual(decorated, [{ cell, row: 0, column: 0 }]);
    assert.equal(context.formulaHover.attached.length, 1);
    assert.deepEqual(context.formulaHover.attached[0].context, {
      ...linkInfo,
      formula: EXCEL_REFERENCE,
      readOnly: false,
    });
    assert.equal(typeof context.formulaHover.attached[0].attachOptions.resolveAnchor, "function");
    assert.equal(typeof context.formulaHover.attached[0].attachOptions.positionRect, "function");

    const result = await context.formulaHoverOptions.onCommit({
      formula: nextReference,
      context: linkInfo,
    });

    assert.equal(result.ok, true);
    assert.deepEqual(requests, [{
      displayRow: 0,
      displayColumn: 0,
      reference: nextReference,
    }]);
    assert.equal(context.calls.externalRequestCancellations, 1);
    assert.equal(context.calls.renders, 1);
    assert.equal(context.calls.updates, 1);
    assert.deepEqual(context.calls.postedMessages, []);
  } finally {
    context.cleanup();
  }
});
