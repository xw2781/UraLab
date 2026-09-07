import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const referenceSource = await readFile(
  new URL("../ui/shared/integrations/excel_reference.js", import.meta.url),
  "utf8",
);
const referenceUrl = `data:text/javascript;base64,${Buffer.from(referenceSource).toString("base64")}`;
const excelApiStubUrl = `data:text/javascript;base64,${Buffer.from(
  "export async function readExcelCellsBatch(){ return { ok: false, results: [] }; } export async function validateExcelLinksBatch(){ return { ok: false, results: [], workbooks: [] }; } export async function readExcelFileMtimesBatch(){ return { ok: false, results: [] }; }",
).toString("base64")}`;
let controllerSource = await readFile(
  new URL("../ui/shared/dataset/dataset_external_links.js", import.meta.url),
  "utf8",
);
controllerSource = controllerSource
  .replace('"/ui/shared/integrations/excel_api.js?v=20260819a"', JSON.stringify(excelApiStubUrl))
  .replace(
    '"/ui/shared/integrations/excel_reference.js?v=20260715a"',
    JSON.stringify(referenceUrl),
  );
const externalLinks = await import(
  `data:text/javascript;base64,${Buffer.from(controllerSource).toString("base64")}`
);

const REF = "='C:\\Data\\[Book.xlsx]Sheet 1'!A1:B2";

function model2x2() {
  return {
    origin_labels: ["2024", "2025"],
    dev_labels: ["12m", "24m"],
    values: [[1, 2], [3, 4]],
    mask: [[true, true], [true, true]],
  };
}

function decoratedCell() {
  const classes = new Set();
  return {
    classes,
    dataset: {},
    classList: {
      contains: (name) => classes.has(name),
      remove: (...names) => names.forEach((name) => classes.delete(name)),
      toggle: (name, force) => {
        if (force) classes.add(name);
        else classes.delete(name);
        return !!force;
      },
    },
    removeAttribute() {},
  };
}

function arrayOutlineClasses(cell) {
  return Array.from(cell.classes)
    .filter((name) => name.startsWith("arArrayFormula"))
    .sort();
}

test("normalizes link metadata without merging separate consumers", () => {
  const normalized = externalLinks.normalizeDatasetExternalLinks([
    {
      reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!$a$1",
      target_cells: [{ row: 0, column: 0 }, { row: 0, column: 0 }],
    },
    {
      reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1",
      target_cells: [{ row: 1, column: 0 }],
    },
    {
      reference: "='C:\\Data\\[Other.xlsx]Sheet 2'!B2",
      target_cells: [{ row: 0, column: 0 }],
    },
    {
      reference: "='C:\\Data\\[Other.xlsx]Sheet 2'!A1:B1",
      target_cells: [{ row: 1, column: 1 }],
    },
  ]);

  assert.equal(normalized.length, 2);
  assert.equal(normalized[0].reference, "='C:\\Data\\[Book.xlsx]Sheet 1'!A1");
  assert.deepEqual(normalized[0].target_cells, [
    { row: 0, column: 0, source_cell: "A1" },
  ]);
  assert.deepEqual(normalized[1].target_cells, [
    { row: 1, column: 0, source_cell: "A1" },
  ]);
});

test("accepts clipped source mappings and rejects mixed or out-of-range mappings", () => {
  const reference = "='C:\\Data\\[Book.xlsx]Sheet 1'!A1:C3";
  const normalized = externalLinks.normalizeDatasetExternalLinks([{
    reference,
    target_cells: [
      { row: 0, column: 0, source_cell: "$A$1" },
      { row: 1, column: 0, source_cell: "A2" },
    ],
  }]);
  assert.deepEqual(normalized[0].target_cells, [
    { row: 0, column: 0, source_cell: "A1" },
    { row: 1, column: 0, source_cell: "A2" },
  ]);
  assert.deepEqual(externalLinks.normalizeDatasetExternalLinks([{
    reference,
    target_cells: [
      { row: 0, column: 0, source_cell: "A1" },
      { row: 1, column: 0 },
    ],
  }]), []);
  assert.deepEqual(externalLinks.normalizeDatasetExternalLinks([{
    reference,
    target_cells: [{ row: 0, column: 0, source_cell: "D4" }],
  }]), []);
  assert.deepEqual(externalLinks.normalizeDatasetExternalLinks([{
    reference,
    target_cells: [
      { row: 0, column: 0, source_cell: "A1" },
      { row: 0, column: 0, source_cell: "B1" },
    ],
  }]), []);
});

test("clips normal and transposed destinations to editable cells", () => {
  const model = model2x2();
  const normal = externalLinks.buildDatasetExternalLinkTargets({
    model,
    startRow: 0,
    startColumn: 0,
    rowCount: 2,
    columnCount: 1,
  });
  assert.equal(normal.ok, true);
  assert.deepEqual(
    normal.targets.map(({ row, column }) => ({ row, column })),
    [{ row: 0, column: 0 }, { row: 1, column: 0 }],
  );

  const transposed = externalLinks.buildDatasetExternalLinkTargets({
    model,
    transposed: true,
    startRow: 0,
    startColumn: 0,
    rowCount: 1,
    columnCount: 2,
  });
  assert.equal(transposed.ok, true);
  assert.deepEqual(
    transposed.targets.map(({ row, column }) => ({ row, column })),
    [{ row: 0, column: 0 }, { row: 1, column: 0 }],
  );

  model.mask[1][0] = false;
  const clipped = externalLinks.buildDatasetExternalLinkTargets({
    model,
    startRow: 0,
    startColumn: 0,
    rowCount: 4,
    columnCount: 3,
  });
  assert.equal(clipped.ok, true);
  assert.deepEqual(
    clipped.targets.map(({ row, column, rowOffset, columnOffset }) => ({
      row,
      column,
      rowOffset,
      columnOffset,
    })),
    [
      { row: 0, column: 0, rowOffset: 0, columnOffset: 0 },
      { row: 0, column: 1, rowOffset: 0, columnOffset: 1 },
      { row: 1, column: 1, rowOffset: 1, columnOffset: 1 },
    ],
  );
  assert.equal(clipped.ignoredCellCount, 9);
});

test("commits a linked range as numeric values plus separate metadata", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: async (items) => ({
      ok: true,
      results: items.map((_item, index) => ({ ok: true, value: [10, 0, -2, 40][index] })),
    }),
  });
  controller.load([]);

  const result = await controller.commitReference({
    displayRow: 0,
    displayColumn: 0,
    reference: REF,
  });

  assert.equal(result.ok, true);
  assert.deepEqual(state.model.values, [[10, 0], [-2, 40]]);
  assert.equal(state.dirty.size, 4);
  assert.equal(controller.isDirty(), true);
  assert.deepEqual(controller.serialize()[0].target_cells, [
    { row: 0, column: 0, source_cell: "A1" },
    { row: 0, column: 1, source_cell: "B1" },
    { row: 1, column: 0, source_cell: "A2" },
    { row: 1, column: 1, source_cell: "B2" },
  ]);
});

test("commits transposed ranges in displayed row-major order", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    isTransposed: () => true,
    readCellsBatch: async () => ({
      ok: true,
      results: [10, 20, 30, 40].map((value) => ({ ok: true, value })),
    }),
  });
  controller.load([]);

  const result = await controller.commitReference({
    displayRow: 0,
    displayColumn: 0,
    reference: REF,
  });

  assert.equal(result.ok, true);
  assert.deepEqual(state.model.values, [[10, 30], [20, 40]]);
  assert.deepEqual(controller.serialize()[0].target_cells, [
    { row: 0, column: 0, source_cell: "A1" },
    { row: 1, column: 0, source_cell: "B1" },
    { row: 0, column: 1, source_cell: "A2" },
    { row: 1, column: 1, source_cell: "B2" },
  ]);
});

test("commits only the in-grid triangle portion of a large Excel range", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  state.model.mask[1][1] = false;
  const readItems = [];
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: async (items) => {
      readItems.push(...items);
      return {
        ok: true,
        results: items.map((_item, index) => ({ ok: true, value: 10 + index })),
      };
    },
  });
  controller.load([]);

  const result = await controller.commitReference({
    displayRow: 0,
    displayColumn: 0,
    reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1:D4",
  });

  assert.equal(result.ok, true);
  assert.equal(result.affectedCellCount, 3);
  assert.deepEqual(readItems.map((item) => item.cell), ["A1", "B1", "A2"]);
  assert.deepEqual(state.model.values, [[10, 11], [12, 4]]);
  assert.deepEqual(controller.serialize()[0].target_cells, [
    { row: 0, column: 0, source_cell: "A1" },
    { row: 0, column: 1, source_cell: "B1" },
    { row: 1, column: 0, source_cell: "A2" },
  ]);
});

test("clips a full-sheet Excel reference without materializing its source cells", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const readItems = [];
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: async (items) => {
      readItems.push(...items);
      return {
        ok: true,
        results: items.map((_item, index) => ({ ok: true, value: 20 + index })),
      };
    },
  });
  controller.load([]);

  const result = await controller.commitReference({
    displayRow: 0,
    displayColumn: 0,
    reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1:XFD1048576",
  });

  assert.equal(result.ok, true);
  assert.equal(result.affectedCellCount, 4);
  assert.deepEqual(readItems.map((item) => item.cell), ["A1", "B1", "A2", "B2"]);
  assert.equal(controller.listRecords()[0].value, "20...");
  assert.equal(externalLinks.normalizeDatasetExternalLinks([{
    reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1:XFD1048576",
    target_cells: [{ row: 0, column: 0, source_cell: "XFD1048576" }],
  }]).length, 1);
});

test("range Values previews keep an ellipsis when clipping leaves one target", () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({ state });
  controller.load([{
    reference: REF,
    target_cells: [{ row: 0, column: 0, source_cell: "A1" }],
  }]);

  assert.equal(controller.listRecords()[0].value, "1...");
});

test("failed range reads leave every value and link unchanged", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: async () => ({
      ok: true,
      results: [
        { ok: true, value: 10 },
        { ok: false, error: "not numeric" },
        { ok: true, value: 30 },
        { ok: true, value: 40 },
      ],
    }),
  });
  controller.load([]);

  const result = await controller.commitReference({
    displayRow: 0,
    displayColumn: 0,
    reference: REF,
  });

  assert.equal(result.ok, false);
  assert.deepEqual(state.model.values, [[1, 2], [3, 4]]);
  assert.deepEqual(controller.serialize(), []);
  assert.equal(state.dirty.size, 0);
});

test("blank Excel cells commit as null values without rejecting the linked range", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: async () => ({
      ok: true,
      results: [
        { ok: true, value: 10 },
        { ok: true, value: "" },
        { ok: true, value: 30 },
        { ok: true, value: 40 },
      ],
    }),
  });
  controller.load([]);

  const result = await controller.commitReference({
    displayRow: 0,
    displayColumn: 0,
    reference: REF,
  });

  assert.equal(result.ok, true);
  assert.deepEqual(state.model.values, [[10, null], [30, 40]]);
  assert.equal(state.dirty.get("0,1"), null);
  assert.equal(controller.serialize().length, 1);
});

test("refresh replaces a numeric linked value with null when Excel is blank", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  state.model.values[0][0] = 0;
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: async () => ({
      ok: true,
      results: [
        { ok: true, value: null },
        { ok: true, value: 2 },
        { ok: true, value: 3 },
        { ok: true, value: 4 },
      ],
    }),
  });
  controller.load([{
    reference: REF,
    target_cells: [
      { row: 0, column: 0 },
      { row: 0, column: 1 },
      { row: 1, column: 0 },
      { row: 1, column: 1 },
    ],
  }]);

  const result = await controller.refreshAll();

  assert.deepEqual(result, { linkedCellCount: 4, changedCount: 1, failedCount: 0, failures: [] });
  assert.deepEqual(state.model.values, [[null, 2], [3, 4]]);
  assert.equal(state.dirty.get("0,0"), null);
});

test("breaking a grouped source preserves values and hard-codes all consumers", () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({ state });
  controller.load([
    { reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1", target_cells: [{ row: 0, column: 0 }] },
    { reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1", target_cells: [{ row: 1, column: 1 }] },
  ]);
  const before = structuredClone(state.model.values);
  const [record] = controller.listRecords();

  const result = controller.breakLink(record.id);

  assert.equal(result.ok, true);
  assert.equal(result.affectedCellCount, 2);
  assert.deepEqual(state.model.values, before);
  assert.deepEqual(controller.serialize(), []);
  assert.equal(controller.isDirty(), true);
  controller.restoreSaved();
  assert.equal(controller.isDirty(), false);
  assert.equal(controller.serialize().length, 2);
});

test("refresh applies each range atomically and marks only changed cells", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: async () => ({
      ok: true,
      results: [
        { ok: true, value: 1 },
        { ok: true, value: 20 },
        { ok: true, value: 3 },
        { ok: true, value: 40 },
      ],
    }),
  });
  controller.load([{
    reference: REF,
    target_cells: [
      { row: 0, column: 0 },
      { row: 0, column: 1 },
      { row: 1, column: 0 },
      { row: 1, column: 1 },
    ],
  }]);

  const result = await controller.refreshAll();

  assert.deepEqual(result, { linkedCellCount: 4, changedCount: 2, failedCount: 0, failures: [] });
  assert.deepEqual(state.model.values, [[1, 20], [3, 40]]);
  assert.deepEqual(Array.from(state.dirty.keys()), ["0,1", "1,1"]);
  assert.equal(controller.isDirty(), false);
});

test("accepted freshness refresh marks equal linked values dirty for a durable save", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: async () => ({
      ok: true,
      results: [1, 2, 3, 4].map((value) => ({ ok: true, value })),
    }),
  });
  controller.load([{
    reference: REF,
    target_cells: [
      { row: 0, column: 0 },
      { row: 0, column: 1 },
      { row: 1, column: 0 },
      { row: 1, column: 1 },
    ],
  }]);

  const result = await controller.refreshAll(null, { markRefreshedCellsDirty: true });

  assert.deepEqual(result, { linkedCellCount: 4, changedCount: 0, failedCount: 0, failures: [] });
  assert.deepEqual(state.model.values, [[1, 2], [3, 4]]);
  assert.deepEqual(Array.from(state.dirty.keys()), ["0,0", "0,1", "1,0", "1,1"]);
});

test("refresh does not apply result payloads from a failed batch response", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: async () => ({
      ok: false,
      results: [10, 20, 30, 40].map((value) => ({ ok: true, value })),
    }),
  });
  controller.load([{
    reference: REF,
    target_cells: [
      { row: 0, column: 0 },
      { row: 0, column: 1 },
      { row: 1, column: 0 },
      { row: 1, column: 1 },
    ],
  }]);

  const result = await controller.refreshAll();

  assert.deepEqual(result, {
    linkedCellCount: 4,
    changedCount: 0,
    failedCount: 4,
    failures: [],
    error: "Excel refresh failed.",
  });
  assert.deepEqual(state.model.values, [[1, 2], [3, 4]]);
  assert.equal(state.dirty.size, 0);
});

test("refresh honors clipped source mappings and selected link groups", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const readItems = [];
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: async (items) => {
      readItems.push(...items);
      return {
        ok: true,
        results: items.map((item) => ({ ok: true, value: item.cell === "D4" ? 44 : 11 })),
      };
    },
  });
  controller.load([
    {
      reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1:D4",
      target_cells: [
        { row: 0, column: 0, source_cell: "A1" },
        { row: 1, column: 1, source_cell: "$D$4" },
      ],
    },
    {
      reference: "='C:\\Data\\[Other.xlsx]Sheet 2'!C3",
      target_cells: [{ row: 0, column: 1, source_cell: "C3" }],
    },
  ]);
  const records = controller.listRecords();

  const result = await controller.refreshAll([records[0].id]);

  assert.deepEqual(result, { linkedCellCount: 2, changedCount: 2, failedCount: 0, failures: [] });
  assert.deepEqual(readItems.map((item) => item.cell), ["A1", "D4"]);
  assert.deepEqual(state.model.values, [[11, 2], [3, 44]]);
  assert.equal(records[0].value, "1...");
});

test("exposes the range anchor for linked-cell formula editing", () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    isTransposed: () => true,
  });
  controller.load([{
    reference: REF,
    target_cells: [
      { row: 0, column: 0, source_cell: "A1" },
      { row: 0, column: 1, source_cell: "B1" },
    ],
  }]);

  assert.deepEqual(controller.getCellLinkInfo(1, 0), {
    id: "c:\\data\\book.xlsx\u001fsheet 1\u001fA1:B2",
    reference: REF,
    sourceCell: "B1",
    anchorDisplayRow: 0,
    anchorDisplayColumn: 0,
  });
});

test("frames a clipped Dataset array formula as the rectangle its reference covers", () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({ state });
  controller.load([{
    reference: REF,
    target_cells: [
      { row: 0, column: 0, source_cell: "A1" },
      { row: 0, column: 1, source_cell: "B1" },
      { row: 1, column: 0, source_cell: "A2" },
    ],
  }]);
  const topLeft = decoratedCell();
  const topRight = decoratedCell();
  const bottomLeft = decoratedCell();
  const unlinked = decoratedCell();

  controller.decorateCell(topLeft, 0, 0);
  controller.decorateCell(topRight, 0, 1);
  controller.decorateCell(bottomLeft, 1, 0);
  controller.decorateCell(unlinked, 1, 1);

  assert.deepEqual(arrayOutlineClasses(topLeft), [
    "arArrayFormulaCell",
    "arArrayFormulaEdgeLeft",
    "arArrayFormulaEdgeTop",
  ]);
  assert.deepEqual(arrayOutlineClasses(topRight), [
    "arArrayFormulaCell",
    "arArrayFormulaEdgeRight",
    "arArrayFormulaEdgeTop",
  ]);
  assert.deepEqual(arrayOutlineClasses(bottomLeft), [
    "arArrayFormulaCell",
    "arArrayFormulaEdgeBottom",
    "arArrayFormulaEdgeLeft",
  ]);
  // The mask kept a value out of the fourth corner, but the corner is still
  // part of the range, so it closes the frame instead of leaving a staircase.
  assert.deepEqual(arrayOutlineClasses(unlinked), [
    "arArrayFormulaCell",
    "arArrayFormulaEdgeBottom",
    "arArrayFormulaEdgeRight",
  ]);
  assert.equal(unlinked.classes.has("arExternalLinkCell"), false);

  controller.load([{
    reference: REF,
    target_cells: [{ row: 0, column: 0, source_cell: "A1" }],
  }]);
  controller.decorateCell(topLeft, 0, 0);
  assert.deepEqual(arrayOutlineClasses(topLeft), [
    "arArrayFormulaCell",
    "arArrayFormulaEdgeBottom",
    "arArrayFormulaEdgeLeft",
    "arArrayFormulaEdgeRight",
    "arArrayFormulaEdgeTop",
  ]);

  controller.load([{
    reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1",
    target_cells: [{ row: 0, column: 0, source_cell: "A1" }],
  }]);
  controller.decorateCell(topLeft, 0, 0);
  assert.equal(topLeft.classes.has("arExternalLinkCell"), true);
  assert.deepEqual(arrayOutlineClasses(topLeft), []);
});

test("rotates Dataset array-formula perimeter edges in Transposed mode", () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    isTransposed: () => true,
  });
  controller.load([{
    reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1:B1",
    target_cells: [
      { row: 0, column: 0, source_cell: "A1" },
      { row: 0, column: 1, source_cell: "B1" },
    ],
  }]);
  const top = decoratedCell();
  const bottom = decoratedCell();

  controller.decorateCell(top, 0, 0);
  controller.decorateCell(bottom, 1, 0);

  assert.deepEqual(arrayOutlineClasses(top), [
    "arArrayFormulaCell",
    "arArrayFormulaEdgeLeft",
    "arArrayFormulaEdgeRight",
    "arArrayFormulaEdgeTop",
  ]);
  assert.deepEqual(arrayOutlineClasses(bottom), [
    "arArrayFormulaCell",
    "arArrayFormulaEdgeBottom",
    "arArrayFormulaEdgeLeft",
    "arArrayFormulaEdgeRight",
  ]);
});

test("bulk break removes selected source groups once", () => {
  let inventoryChanges = 0;
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    onInventoryChanged: () => { inventoryChanges += 1; },
  });
  controller.load([
    { reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1", target_cells: [{ row: 0, column: 0 }] },
    { reference: "='C:\\Data\\[Other.xlsx]Sheet 2'!B2", target_cells: [{ row: 1, column: 1 }] },
  ]);
  const ids = controller.listRecords().map((record) => record.id);

  const result = controller.breakLinks(ids);

  assert.equal(result.ok, true);
  assert.equal(result.affectedCellCount, 2);
  assert.equal(controller.serialize().length, 0);
  assert.equal(inventoryChanges, 2);
});

test("hard-coding a target invalidates an unresolved Excel commit", async () => {
  let resolveRead;
  const readResult = new Promise((resolve) => { resolveRead = resolve; });
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: () => readResult,
  });
  controller.load([]);

  const pendingCommit = controller.commitReference({
    displayRow: 0,
    displayColumn: 0,
    reference: REF,
  });
  controller.hardCodeTargetCells([{ row: 0, column: 0 }]);
  resolveRead({
    ok: true,
    results: [1, 2, 3, 4].map((value) => ({ ok: true, value: value * 10 })),
  });
  const result = await pendingCommit;

  assert.equal(result.stale, true);
  assert.deepEqual(state.model.values, [[1, 2], [3, 4]]);
  assert.deepEqual(controller.serialize(), []);
});

test("breaking a link invalidates its unresolved refresh", async () => {
  let resolveRead;
  const readResult = new Promise((resolve) => { resolveRead = resolve; });
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: () => readResult,
  });
  controller.load([{
    reference: REF,
    target_cells: [
      { row: 0, column: 0 },
      { row: 0, column: 1 },
      { row: 1, column: 0 },
      { row: 1, column: 1 },
    ],
  }]);

  const pendingRefresh = controller.refreshAll();
  const [record] = controller.listRecords();
  assert.equal(controller.breakLink(record.id).ok, true);
  resolveRead({
    ok: true,
    results: [10, 20, 30, 40].map((value) => ({ ok: true, value })),
  });
  const result = await pendingRefresh;

  assert.equal(result.stale, true);
  assert.deepEqual(state.model.values, [[1, 2], [3, 4]]);
  assert.deepEqual(controller.serialize(), []);
});

const LINKED_WORKBOOKS = [
  { reference: "='C:\\Data\\[Older.xlsx]Sheet 1'!A1", target_cells: [{ row: 0, column: 0 }] },
  { reference: "='C:\\Data\\[Older.xlsx]Sheet 2'!B2", target_cells: [{ row: 0, column: 1 }] },
  { reference: "='C:\\Data\\[Newer.xlsx]Sheet 1'!A1", target_cells: [{ row: 1, column: 0 }] },
  { reference: "='C:\\Data\\[Missing.xlsx]Sheet 1'!A1", target_cells: [{ row: 1, column: 1 }] },
];

test("validates every linked cell and reports newer workbooks in one pass", async () => {
  const requests = [];
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    validateLinksBatch: async (items) => {
      requests.push(items);
      return {
        ok: true,
        results: [
          { ok: true, value: 5 },
          { ok: true, value: 6 },
          { ok: true, value: 7 },
          { ok: false, error: "Not numeric: '#REF!'" },
        ],
        workbooks: [
          { ok: true, path: "C:\\Data\\Older.xlsx", mtime: 99 },
          { ok: true, path: "C:\\Data\\Newer.xlsx", mtime: 101 },
          { ok: false, path: "C:\\Data\\Missing.xlsx", error: "Unavailable" },
        ],
      };
    },
  });
  controller.load(LINKED_WORKBOOKS);

  const result = await controller.validateLinks(100);

  // One request, every stored source cell in it, in link order.
  assert.equal(requests.length, 1);
  assert.deepEqual(
    requests[0].map((item) => `${item.book_path}|${item.sheet}|${item.cell}`),
    [
      "C:\\Data\\Older.xlsx|Sheet 1|A1",
      "C:\\Data\\Older.xlsx|Sheet 2|B2",
      "C:\\Data\\Newer.xlsx|Sheet 1|A1",
      "C:\\Data\\Missing.xlsx|Sheet 1|A1",
    ],
  );
  assert.equal(result.newerWorkbookCount, 1);
  assert.deepEqual(result.newerWorkbooks, [{ path: "C:\\Data\\Newer.xlsx", mtime: 101 }]);
  assert.equal(result.unverifiedWorkbookCount, 1);
  assert.equal(result.failedCellCount, 1);
  const [failure] = result.failures;
  assert.equal(failure.workbookPath, "C:\\Data\\Missing.xlsx");
  assert.equal(failure.worksheet, "Sheet 1");
  assert.equal(failure.sourceCell, "A1");
  assert.equal(failure.destination, "2025 / 24m");
  assert.equal(failure.error, "Not numeric: '#REF!'");
  // Validation never touches the stored values.
  assert.deepEqual(state.model.values, [[1, 2], [3, 4]]);
  assert.equal(state.dirty.size, 0);
});

test("a transport failure is not reported as a broken reference", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    validateLinksBatch: async () => ({ ok: false, error: "Network error" }),
  });
  controller.load(LINKED_WORKBOOKS);

  const result = await controller.validateLinks(100);

  assert.equal(result.ok, false);
  assert.equal(result.error, "Network error");
  assert.deepEqual(result.failures, []);
  assert.deepEqual(controller.getLinkFailures(), []);
});

test("a refresh that hits a broken reference keeps the saved value and marks the cell", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: async () => ({
      ok: true,
      results: [
        { ok: true, value: 10 },
        { ok: false, error: "Not numeric: '#REF!'" },
      ],
    }),
  });
  controller.load([
    { reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1", target_cells: [{ row: 0, column: 0 }] },
    { reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!B1", target_cells: [{ row: 0, column: 1 }] },
  ]);

  const result = await controller.refreshAll();

  assert.equal(result.changedCount, 1);
  assert.equal(result.failedCount, 1);
  // The broken cell keeps its saved value; the readable one refreshes.
  assert.deepEqual(state.model.values, [[10, 2], [3, 4]]);
  assert.equal(result.failures.length, 1);
  assert.equal(result.failures[0].sourceCell, "B1");
  assert.equal(result.failures[0].destination, "2024 / 24m");
  assert.deepEqual(controller.getLinkFailures().map((item) => item.sourceCell), ["B1"]);

  const refreshed = decoratedCell();
  const broken = decoratedCell();
  controller.decorateCell(refreshed, 0, 0);
  controller.decorateCell(broken, 0, 1);
  assert.equal(refreshed.classList.contains("arExternalLinkErrorCell"), false);
  assert.equal(broken.classList.contains("arExternalLinkErrorCell"), true);
});

test("a link that no longer parses or no longer has a dataset cell is reported by name", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    validateLinksBatch: async () => ({ ok: true, results: [], workbooks: [] }),
  });
  controller.load([
    { reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1", target_cells: [{ row: 0, column: 0 }] },
  ]);
  // The grid shrank underneath a saved link.
  state.model.mask[0][0] = false;

  const result = await controller.validateLinks(100);

  assert.equal(result.failedCellCount, 1);
  assert.match(result.failures[0].error, /no longer part of this dataset/);
});

test("breaking a broken link clears its reference failure", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    validateLinksBatch: async () => ({
      ok: true,
      results: [{ ok: false, error: "Sheet not found: Sheet 1" }],
      workbooks: [{ ok: true, path: "C:\\Data\\Book.xlsx", mtime: 10 }],
    }),
  });
  controller.load([
    { reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1", target_cells: [{ row: 0, column: 0 }] },
  ]);

  await controller.validateLinks(100);
  assert.equal(controller.getLinkFailures().length, 1);

  const [record] = controller.listRecords();
  assert.equal(controller.breakLink(record.id).ok, true);

  assert.deepEqual(controller.getLinkFailures(), []);
  const cell = decoratedCell();
  controller.decorateCell(cell, 0, 0);
  assert.equal(cell.classList.contains("arExternalLinkErrorCell"), false);
});

test("blank source cells refresh as nulls and never read as broken references", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: async () => ({
      ok: true,
      results: [
        { ok: true, value: 10 },
        { ok: true, value: null },
        { ok: true, value: null },
        { ok: true, value: 40 },
      ],
    }),
  });
  controller.load([{
    reference: REF,
    target_cells: [
      { row: 0, column: 0 },
      { row: 0, column: 1 },
      { row: 1, column: 0 },
      { row: 1, column: 1 },
    ],
  }]);

  const result = await controller.refreshAll();

  // A blank is a value, not a failure: the range still applies whole.
  assert.equal(result.failedCount, 0);
  assert.deepEqual(result.failures, []);
  assert.deepEqual(state.model.values, [[10, null], [null, 40]]);
  assert.deepEqual(controller.getLinkFailures(), []);
  const blank = decoratedCell();
  controller.decorateCell(blank, 0, 1);
  assert.equal(blank.classList.contains("arExternalLinkErrorCell"), false);
});

test("the grid shows a blank linked value as a muted zero rather than an empty cell", async () => {
  const gridView = await readFile(
    new URL("../ui/shared/tabs/data/dataset_grid_view.js", import.meta.url),
    "utf8",
  );
  const dataTabCss = await readFile(
    new URL("../ui/shared/tabs/data/data_tab.css", import.meta.url),
    "utf8",
  );
  assert.match(gridView, /const displayNullAsZero = isEditable && v == null;/u);
  assert.match(gridView, /td\.textContent = formatCellValue\(displayNullAsZero \? 0 : v\);/u);
  assert.match(dataTabCss, /#tableWrap td\.dsNullValue[\s\S]*?color:\s*#7a858f/u);
});

test("a blank cell inside a linked rectangle keeps that rectangle's edge", async () => {
  const dataTabCss = await readFile(
    new URL("../ui/shared/tabs/data/data_tab.css", import.meta.url),
    "utf8",
  );
  // A blank drops every grid line, so each edge has to be restored by a rule
  // naming three classes: two would lose to the border the last column keeps.
  ["Top", "Right", "Bottom", "Left"].forEach((side) => {
    assert.match(
      dataTabCss,
      new RegExp(
        `#tableWrap td\\.na\\.arArrayFormulaCell\\.arArrayFormulaEdge${side} \\{\\s*`
        + `border-${side.toLowerCase()}: 1px solid var\\(--ar-array-formula-border\\) !important;`,
        "u",
      ),
    );
  });
});

test("the shared grid stylesheet paints a failed link red in both themes", async () => {
  const spreadsheetCss = await readFile(
    new URL("../ui/shared/components/spreadsheet/spreadsheet_table.css", import.meta.url),
    "utf8",
  );
  const darkCss = await readFile(
    new URL("../ui/shared/styles/themes/dark.css", import.meta.url),
    "utf8",
  );
  assert.match(
    spreadsheetCss,
    /\.arSpreadsheetTable td\.arExternalLinkErrorCell,\s*\.arSpreadsheetTable td\.arInternalLinkErrorCell,\s*\.arSpreadsheetTable td\.arFormulaLinkErrorCell \{\s*color: var\(--ar-spreadsheet-link-error-text\);/u,
  );
  assert.match(spreadsheetCss, /--ar-spreadsheet-link-error-text: #b91c1c;/u);
  assert.match(
    spreadsheetCss,
    /:root\[data-arcrho-theme="dark"\] \.arSpreadsheetTable td\.arFormulaLinkErrorCell \{\s*color: var\(--ar-color-danger\);/u,
  );
  // The Dark rule for a plain cell is more specific than the base rule, so the
  // deviation has to be declared in the theme or the red is lost.
  assert.match(
    darkCss,
    /\[data-arcrho-theme="dark"\] \.arSpreadsheetTable td\.arExternalLinkErrorCell,\s*:root\[data-arcrho-theme="dark"\] \.arSpreadsheetTable td\.arInternalLinkErrorCell \{\s*color: var\(--ar-color-danger\);/u,
  );
});

test("a view off the lengths the links were read at holds the whole link inventory still", async () => {
  // A saved link names a cell of the grid its dataset was displayed at. Shown
  // at other lengths, every cell on screen stands for other cells entirely, so
  // painting, naming, or releasing one would land on the wrong cell - and every
  // stored link would report itself missing. The links themselves are untouched.
  const state = { model: model2x2(), dirty: new Map() };
  let atLinkedShape = true;
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    isAtLinkedShape: () => atLinkedShape,
    readCellsBatch: async (items) => ({
      ok: true,
      results: items.map(() => ({ ok: true, value: 7 })),
    }),
  });
  controller.load([{ reference: REF, target_cells: [
    { row: 0, column: 0, source_cell: "A1" },
    { row: 0, column: 1, source_cell: "B1" },
    { row: 1, column: 0, source_cell: "A2" },
    { row: 1, column: 1, source_cell: "B2" },
  ] }]);

  assert.ok(controller.hasLinks());
  assert.ok(controller.getCellLinkInfo(0, 0));
  const onShape = controller.listRecords()[0];
  assert.equal(onShape.destination, "2024~2025 / 12m~24m");

  atLinkedShape = false;
  assert.equal(controller.getCellLinkInfo(0, 0), null);
  const painted = decoratedCell();
  controller.decorateCell(painted, 0, 0);
  assert.equal(painted.classes.size, 0);
  assert.equal(controller.hardCodeTargetCells([{ row: 0, column: 0 }]), 0);
  // The link is still there to come back to, and the columns read off the grid
  // stay blank rather than quoting a cell the link never named.
  assert.ok(controller.hasLinks());
  const offShape = controller.listRecords()[0];
  assert.equal(offShape.value, "");
  assert.equal(offShape.destination, "Data");
  assert.equal(offShape.affectedCellCount, 4);

  atLinkedShape = true;
  assert.ok(controller.getCellLinkInfo(0, 0));
});

test("the workbook probe states each linked workbook once and reports only newer ones", async () => {
  // The cheap half of validation: no cell is read, one entry per distinct
  // workbook, and only a workbook saved after the dataset's own file counts.
  const state = { model: model2x2(), dirty: new Map() };
  const asked = [];
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readCellsBatch: async () => {
      throw new Error("A freshness probe must not read workbook cells.");
    },
    readFileMtimesBatch: async (bookPaths) => {
      asked.push(bookPaths);
      return { ok: true, results: bookPaths.map(() => ({ ok: true, path: bookPaths[0], mtime: 200 })) };
    },
  });
  controller.load([
    { reference: REF, target_cells: [
      { row: 0, column: 0, source_cell: "A1" },
      { row: 0, column: 1, source_cell: "B1" },
      { row: 1, column: 0, source_cell: "A2" },
      { row: 1, column: 1, source_cell: "B2" },
    ] },
  ]);

  const newer = await controller.findNewerWorkbooks(100);
  assert.equal(asked.length, 1);
  assert.deepEqual(asked[0], [String.raw`C:\Data\Book.xlsx`]);
  assert.equal(newer.ok, true);
  assert.equal(newer.newerWorkbooks.length, 1);

  const unchanged = await controller.findNewerWorkbooks(300);
  assert.equal(unchanged.ok, true);
  assert.equal(unchanged.newerWorkbooks.length, 0);
});

test("a workbook that cannot be stated is unverified, never newer", async () => {
  const state = { model: model2x2(), dirty: new Map() };
  const controller = externalLinks.createDatasetExternalLinksController({
    state,
    readFileMtimesBatch: async (bookPaths) => ({
      ok: true,
      results: bookPaths.map(() => ({ ok: false, path: "", error: "File not found" })),
    }),
  });
  controller.load([{ reference: REF, target_cells: [
    { row: 0, column: 0, source_cell: "A1" },
    { row: 0, column: 1, source_cell: "B1" },
    { row: 1, column: 0, source_cell: "A2" },
    { row: 1, column: 1, source_cell: "B2" },
  ] }]);

  const result = await controller.findNewerWorkbooks(100);
  assert.equal(result.ok, true);
  assert.equal(result.newerWorkbooks.length, 0);
  assert.equal(result.unverifiedWorkbookCount, 1);
});
