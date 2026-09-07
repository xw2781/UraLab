import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const dataUrl = (source) => `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;
const read = (path) => readFile(new URL(path, import.meta.url), "utf8");

const referenceUrl = dataUrl(await read("../ui/shared/integrations/excel_reference.js"));
const excelApiStubUrl = dataUrl(
  "export async function readExcelCellsBatch(){ return { ok: false, results: [] }; } export async function validateExcelLinksBatch(){ return { ok: false, results: [], workbooks: [] }; } export async function readExcelFileMtimesBatch(){ return { ok: false, results: [] }; }",
);
const externalUrl = dataUrl((await read("../ui/shared/dataset/dataset_external_links.js"))
  .replace('"/ui/shared/integrations/excel_api.js?v=20260819a"', JSON.stringify(excelApiStubUrl))
  .replace('"/ui/shared/integrations/excel_reference.js?v=20260715a"', JSON.stringify(referenceUrl)));
const internalReferenceUrl = dataUrl(await read("../ui/shared/dataset/dataset_internal_reference.js"));
const formulaUrl = dataUrl((await read("../ui/shared/dataset/dataset_formula.js"))
  .replace('"/ui/shared/integrations/excel_reference.js?v=20260715a"', JSON.stringify(referenceUrl))
  .replace('"/ui/shared/dataset/dataset_internal_reference.js?v=20260830a"', JSON.stringify(internalReferenceUrl)));
const controllerSource = (await read("../ui/shared/dataset/dataset_formula_links.js"))
  .replace('"/ui/shared/integrations/excel_api.js?v=20260819a"', JSON.stringify(excelApiStubUrl))
  .replace('"/ui/shared/integrations/excel_reference.js?v=20260715a"', JSON.stringify(referenceUrl))
  .replace('"/ui/shared/dataset/dataset_external_links.js?v=20260907b"', JSON.stringify(externalUrl))
  .replace('"/ui/shared/dataset/dataset_formula.js?v=20260830a"', JSON.stringify(formulaUrl))
  .replace('"/ui/shared/dataset/dataset_internal_reference.js?v=20260830a"', JSON.stringify(internalReferenceUrl));
const formulaLinks = await import(dataUrl(controllerSource));

const FORMULA = "=[C 82 - Prior Qtr Selected][1:2] * 2";
const LINK = {
  formula: FORMULA,
  target_cells: [
    { row: 0, column: 0, result_row: 0, result_column: 0 },
    { row: 1, column: 0, result_row: 1, result_column: 0 },
  ],
};

function vectorModel() {
  return {
    origin_labels: ["2017", "2018", "2019", "2020"],
    dev_labels: ["Value"],
    values: [[1], [2], [3], [4]],
    mask: [[true], [true], [true], [true]],
  };
}

function resolvedVector(values) {
  return {
    ok: true,
    status: 200,
    data: {
      ok: true,
      results: [{
        row_start: 0,
        column_start: 0,
        row_count: values.length,
        column_count: 1,
        cells: values.map((value, row) => ({ row, column: 0, value })),
      }],
    },
  };
}

function controllerWith({ model = vectorModel(), resolve = [], excel = [], claimed = [] } = {}) {
  const state = { model, dirty: new Map() };
  const resolveCalls = [];
  const excelCalls = [];
  const controller = formulaLinks.createDatasetFormulaLinksController({
    state,
    resolveReferences: async (references) => {
      resolveCalls.push(references);
      const result = typeof resolve === "function" ? resolve(references) : resolve.shift();
      if (result instanceof Error) throw result;
      return result ?? { ok: true, status: 200, data: { ok: true, results: [] } };
    },
    readCellsBatch: async (items) => {
      excelCalls.push(items);
      return typeof excel === "function" ? excel(items) : excel.shift();
    },
    onTargetsClaimed: (cells) => claimed.push(...cells),
  });
  return { state, controller, resolveCalls, excelCalls, claimed };
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
  };
}

test("normalizes formula links, deduplicates, and enforces one owner per target cell", () => {
  const normalized = formulaLinks.normalizeDatasetFormulaLinks([
    { formula: " [ C 82 - Prior Qtr Selected ][ 1 : 2 ]*2 ", target_cells: LINK.target_cells },
    LINK,
    { formula: "=2 * 3", target_cells: [{ row: 5, column: 0, result_row: 0, result_column: 0 }] },
    { formula: "=[C 84][1] +", target_cells: [{ row: 6, column: 0, result_row: 0, result_column: 0 }] },
    { formula: "=[C 84][1] + 1", target_cells: [{ row: 0, column: 0, result_row: 0, result_column: 0 }] },
    { formula: "=[C 84][1] + 1", target_cells: [{ row: 7, column: 0, result_row: -1, result_column: 0 }] },
  ]);
  assert.deepEqual(normalized, [LINK]);
});

test("commitReference resolves every source once, calculates, spills, and claims the cells", async () => {
  const { state, controller, resolveCalls, claimed } = controllerWith({
    resolve: [resolvedVector([10, 20])],
  });

  const result = await controller.commitReference({
    displayRow: 2,
    displayColumn: 0,
    reference: "=[ C 82 - Prior Qtr Selected ][1:2]*2",
  });

  assert.equal(result.ok, true);
  assert.equal(result.affectedCellCount, 2);
  assert.match(result.message, /Calculated 2 dataset cells from the formula\./u);
  assert.deepEqual(resolveCalls, [["=[C 82 - Prior Qtr Selected][1:2]"]]);
  assert.equal(state.model.values[2][0], 20);
  assert.equal(state.model.values[3][0], 40);
  assert.equal(state.dirty.get("3,0"), 40);
  assert.deepEqual(claimed, [{ row: 2, column: 0 }, { row: 3, column: 0 }]);
  assert.deepEqual(controller.serialize(), [{
    formula: FORMULA,
    target_cells: [
      { row: 2, column: 0, result_row: 0, result_column: 0 },
      { row: 3, column: 0, result_row: 1, result_column: 0 },
    ],
  }]);
  assert.equal(controller.isDirty(), true);
});

test("a formula over an Excel range and a dataset range reads the workbook cells in one batch", async () => {
  const { state, controller, excelCalls } = controllerWith({
    resolve: [resolvedVector([1, 2])],
    excel: [{ ok: true, results: [{ ok: true, value: 100 }, { ok: true, value: "" }] }],
  });

  const result = await controller.commitReference({
    displayRow: 0,
    displayColumn: 0,
    reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1:A2 + [C 82 - Prior Qtr Selected][1:2]",
  });

  assert.equal(result.ok, true, result.error);
  assert.deepEqual(excelCalls, [[
    { book_path: "C:\\Data\\Book.xlsx", sheet: "Sheet 1", cell: "A1" },
    { book_path: "C:\\Data\\Book.xlsx", sheet: "Sheet 1", cell: "A2" },
  ]]);
  // A blank workbook cell counts as zero.
  assert.deepEqual([state.model.values[0][0], state.model.values[1][0]], [101, 2]);
  assert.equal(
    controller.serialize()[0].formula,
    "='C:\\Data\\[Book.xlsx]Sheet 1'!A1:A2 + [C 82 - Prior Qtr Selected][1:2]",
  );

  // The Links tab gets one row per source the formula reads, and breaking
  // through either row's id breaks the whole formula.
  const records = controller.listRecords();
  assert.deepEqual(
    records.map((record) => [record.id.split("#component-")[1], record.workbookPath, record.datasetName]),
    [["0", "C:\\Data\\Book.xlsx", undefined], ["1", undefined, "C 82 - Prior Qtr Selected"]],
  );
  // Each row's reference names only the component it stands for, not the
  // whole formula.
  assert.deepEqual(records.map((record) => record.reference), ["Sheet 1!A1:A2", "[1:2]"]);
  assert.equal(new Set(records.map((record) => record.id.split("#component-")[0])).size, 1);
  assert.equal(controller.breakLinks([records[1].id]).ok, true);
  assert.deepEqual(controller.serialize(), []);
});

test("commitReference leaves standalone references to the other controllers and names failures", async () => {
  const { state, controller } = controllerWith({
    resolve: [{ ok: false, status: 422, data: { detail: "Row index 9 is outside 'C 82' (1-4)." } }],
    excel: [{ ok: true, results: [{ ok: true, value: "n/a" }] }],
  });

  const standalone = await controller.commitReference({
    displayRow: 0,
    displayColumn: 0,
    reference: "=[C 82][1:2]",
  });
  assert.equal(standalone.handled, false);

  const invalid = await controller.commitReference({ displayRow: 0, displayColumn: 0, reference: "=1 +" });
  assert.equal(invalid.ok, false);
  assert.match(invalid.error, /ends before its last operand/u);

  const resolveFailure = await controller.commitReference({
    displayRow: 0,
    displayColumn: 0,
    reference: "=[C 82][9] * 2",
  });
  assert.equal(resolveFailure.ok, false);
  assert.match(resolveFailure.error, /outside 'C 82'/u);

  const excelFailure = await controller.commitReference({
    displayRow: 0,
    displayColumn: 0,
    reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!B2 * 2",
  });
  assert.equal(excelFailure.ok, false);
  assert.match(excelFailure.error, /B2: Excel returned a non-numeric value: n\/a/u);
  assert.equal(state.dirty.size, 0);
  assert.deepEqual(controller.serialize(), []);
});

test("decorates calculated cells in the formula colour and describes the link", () => {
  const { controller } = controllerWith();
  controller.load([LINK]);
  const top = decoratedCell();
  const bottom = decoratedCell();
  const plain = decoratedCell();
  controller.decorateCell(top, 0, 0);
  controller.decorateCell(bottom, 1, 0);
  controller.decorateCell(plain, 2, 0);

  assert.equal(top.classes.has("arFormulaLinkCell"), true);
  assert.equal(top.classes.has("arArrayFormulaCell"), true);
  assert.equal(top.classes.has("arArrayFormulaEdgeTop"), true);
  assert.equal(top.classes.has("arArrayFormulaEdgeBottom"), false);
  assert.equal(bottom.classes.has("arArrayFormulaEdgeBottom"), true);
  assert.equal(top.dataset.formulaLinkReference, FORMULA);
  assert.equal(plain.classes.has("arFormulaLinkCell"), false);

  const info = controller.getCellLinkInfo(1, 0);
  assert.equal(info.reference, FORMULA);
  assert.equal(info.sourceKind, "formula");
  assert.equal(info.anchorDisplayRow, 0);

  const [record] = controller.listRecords();
  assert.equal(record.sourceKind, "formula");
  assert.equal(record.formula, FORMULA);
  assert.match(record.id, /#component-0$/u);
  assert.equal(record.datasetName, "C 82 - Prior Qtr Selected");
  assert.equal(record.workbookPath, undefined);
  assert.equal(record.destination, "2017~2018");
  assert.equal(record.affectedCellCount, 2);
});

test("refreshAll re-evaluates whole-link and records failures per link", async () => {
  const { state, controller } = controllerWith({ resolve: () => resolvedVector([100, 200]) });
  controller.load([LINK]);
  const refreshed = await controller.refreshAll();
  assert.equal(refreshed.changedCount, 2);
  assert.equal(refreshed.failedCount, 0);
  assert.deepEqual([state.model.values[0][0], state.model.values[1][0]], [200, 400]);
  assert.equal(controller.getLinkFailures().length, 0);

  const failing = controllerWith({
    resolve: [{ ok: false, status: 422, data: { detail: "Dataset 'C 82' was not found." } }],
  });
  failing.controller.load([LINK]);
  const failure = await failing.controller.refreshAll();
  assert.equal(failure.failedCount, 2);
  assert.equal(failing.state.model.values[0][0], 1);
  assert.equal(failing.controller.getLinkFailures().length, 2);
  assert.match(failing.controller.getLinkFailures()[0].error, /not found/u);

  // A result that shrank below the range it fills is refused rather than half-applied.
  const shrunk = controllerWith({ resolve: () => resolvedVector([5]) });
  shrunk.controller.load([LINK]);
  const short = await shrunk.controller.refreshAll();
  assert.equal(short.failedCount, 2);
  assert.match(shrunk.controller.getLinkFailures()[0].error, /smaller than the range/u);
});

test("break and hard-code remove ownership and mark the model dirty state", () => {
  const { controller } = controllerWith();
  controller.load([LINK]);
  const record = controller.listRecords()[0];
  const broken = controller.breakLinks([record.id]);
  assert.equal(broken.ok, true);
  assert.equal(broken.affectedCellCount, 2);
  assert.deepEqual(controller.serialize(), []);
  assert.equal(controller.isDirty(), true);

  controller.markClean([LINK]);
  assert.equal(controller.isDirty(), false);
  assert.equal(controller.hardCodeTargetCells([{ row: 1, column: 0 }]), 1);
  assert.deepEqual(controller.serialize(), []);
  controller.restoreSaved();
  assert.deepEqual(controller.serialize(), [LINK]);
});

const spreadsheetCss = await read("../ui/shared/components/spreadsheet/spreadsheet_table.css");

test("a calculated range wears the formula colour on its perimeter", () => {
  assert.match(
    spreadsheetCss,
    /td\.arArrayFormulaCell\.arFormulaLinkCell \{[^}]*--ar-array-formula-border: var\(--ar-spreadsheet-formula-link-border\)/u,
  );
  assert.match(spreadsheetCss, /--ar-spreadsheet-formula-link-border: #7c3aed;/u);
  assert.match(
    spreadsheetCss,
    /:root\[data-arcrho-theme="dark"\] \{[^}]*--ar-spreadsheet-formula-link-border: #b08cff;/u,
  );
  assert.match(spreadsheetCss, /td\.arFormulaLinkErrorCell \{/u);
});
