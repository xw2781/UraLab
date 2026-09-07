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

let externalSource = await readFile(
  new URL("../ui/shared/dataset/dataset_external_links.js", import.meta.url),
  "utf8",
);
externalSource = externalSource
  .replace('"/ui/shared/integrations/excel_api.js?v=20260819a"', JSON.stringify(excelApiStubUrl))
  .replace('"/ui/shared/integrations/excel_reference.js?v=20260715a"', JSON.stringify(referenceUrl));
const externalUrl = `data:text/javascript;base64,${Buffer.from(externalSource).toString("base64")}`;

const internalReferenceSource = await readFile(
  new URL("../ui/shared/dataset/dataset_internal_reference.js", import.meta.url),
  "utf8",
);
const internalReferenceUrl = `data:text/javascript;base64,${Buffer.from(internalReferenceSource).toString("base64")}`;

let controllerSource = await readFile(
  new URL("../ui/shared/dataset/dataset_internal_links.js", import.meta.url),
  "utf8",
);
controllerSource = controllerSource
  .replace('"/ui/shared/dataset/dataset_external_links.js?v=20260907b"', JSON.stringify(externalUrl))
  .replace('"/ui/shared/dataset/dataset_internal_reference.js?v=20260830a"', JSON.stringify(internalReferenceUrl));
const internalLinks = await import(
  `data:text/javascript;base64,${Buffer.from(controllerSource).toString("base64")}`
);

const LINK = {
  reference: "=[C 82 - Prior Qtr Selected][1:2]",
  target_cells: [
    { row: 0, column: 0, source_row: 0, source_column: 0 },
    { row: 1, column: 0, source_row: 1, source_column: 0 },
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

function resolvedVector(values, { rowStart = 0 } = {}) {
  return {
    reference: "=[C 82 - Prior Qtr Selected][1:2]",
    dataset_name: "C 82 - Prior Qtr Selected",
    data_format: "Vector",
    row_start: rowStart,
    column_start: 0,
    row_count: values.length,
    column_count: 1,
    cells: values.map((value, offset) => ({
      row: rowStart + offset,
      column: 0,
      row_label: String(2017 + rowStart + offset),
      col_label: "Value",
      value,
    })),
  };
}

function controllerWith({ model = vectorModel(), results = [], claimed = [] } = {}) {
  const state = { model, dirty: new Map() };
  const calls = [];
  const controller = internalLinks.createDatasetInternalLinksController({
    state,
    resolveReferences: async (references) => {
      calls.push(references);
      const result = typeof results === "function" ? results(references) : results.shift();
      if (result instanceof Error) throw result;
      return result ?? { ok: true, status: 200, data: { ok: true, results: [] } };
    },
    onTargetsClaimed: (cells) => claimed.push(...cells),
  });
  return { state, controller, calls, claimed };
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

test("normalizes links, deduplicates, and enforces one owner per target cell", () => {
  const normalized = internalLinks.normalizeDatasetInternalLinks([
    { reference: " [ C 82 - Prior Qtr Selected ][ 1 : 2 ] ", target_cells: LINK.target_cells },
    LINK,
    {
      reference: "=[C 84][1]",
      target_cells: [{ row: 0, column: 0, source_row: 0, source_column: 0 }],
    },
    {
      reference: "=[C 84][1] + 1",
      target_cells: [{ row: 5, column: 0, source_row: 0, source_column: 0 }],
    },
    {
      reference: "=[C 84][2]",
      target_cells: [{ row: 6, column: 0, source_row: -1, source_column: 0 }],
    },
  ]);
  assert.deepEqual(normalized, [LINK]);
});

test("commitReference resolves, spills the range, and claims the cells", async () => {
  const { state, controller, calls, claimed } = controllerWith({
    results: [{ ok: true, status: 200, data: { ok: true, results: [resolvedVector([10, 20])] } }],
  });

  const result = await controller.commitReference({
    displayRow: 2,
    displayColumn: 0,
    reference: "[ C 82 - Prior Qtr Selected ][1:2]",
  });

  assert.equal(result.ok, true);
  assert.equal(result.affectedCellCount, 2);
  assert.match(result.message, /Linked 2 dataset cells to C 82 - Prior Qtr Selected\./);
  assert.deepEqual(calls, [["=[C 82 - Prior Qtr Selected][1:2]"]]);
  assert.equal(state.model.values[2][0], 10);
  assert.equal(state.model.values[3][0], 20);
  assert.equal(state.dirty.get("2,0"), 10);
  assert.deepEqual(claimed, [{ row: 2, column: 0 }, { row: 3, column: 0 }]);
  assert.deepEqual(controller.serialize(), [{
    reference: "=[C 82 - Prior Qtr Selected][1:2]",
    target_cells: [
      { row: 2, column: 0, source_row: 0, source_column: 0 },
      { row: 3, column: 0, source_row: 1, source_column: 0 },
    ],
  }]);
  assert.equal(controller.isDirty(), true);
});

test("commitReference reports parse and resolve failures without changing cells", async () => {
  const { state, controller } = controllerWith({
    results: [{ ok: false, status: 422, data: { detail: "Row index 9 is outside 'C 82' (1-4)." } }],
  });

  const parseFailure = await controller.commitReference({
    displayRow: 0,
    displayColumn: 0,
    reference: "=[C 82][1] + 1",
  });
  assert.equal(parseFailure.ok, false);
  assert.match(parseFailure.error, /standalone reference/);

  const resolveFailure = await controller.commitReference({
    displayRow: 0,
    displayColumn: 0,
    reference: "=[C 82][9]",
  });
  assert.equal(resolveFailure.ok, false);
  assert.match(resolveFailure.error, /outside 'C 82'/);
  assert.equal(state.dirty.size, 0);
  assert.deepEqual(controller.serialize(), []);
});

test("decorates linked cells with internal classes and range edges", () => {
  const { controller } = controllerWith();
  controller.load([LINK]);
  const top = decoratedCell();
  const bottom = decoratedCell();
  const plain = decoratedCell();
  controller.decorateCell(top, 0, 0);
  controller.decorateCell(bottom, 1, 0);
  controller.decorateCell(plain, 2, 0);

  assert.equal(top.classes.has("arInternalLinkCell"), true);
  assert.equal(top.classes.has("arArrayFormulaCell"), true);
  assert.equal(top.classes.has("arArrayFormulaEdgeTop"), true);
  assert.equal(top.classes.has("arArrayFormulaEdgeBottom"), false);
  assert.equal(bottom.classes.has("arArrayFormulaEdgeBottom"), true);
  assert.equal(top.dataset.internalLinkReference, LINK.reference);
  assert.equal(plain.classes.has("arInternalLinkCell"), false);

  const info = controller.getCellLinkInfo(1, 0);
  assert.equal(info.reference, LINK.reference);
  assert.equal(info.datasetName, "C 82 - Prior Qtr Selected");
  assert.equal(info.anchorDisplayRow, 0);
  assert.equal(info.sourceKind, "internal");
});

test("lists one record per link with dataset name and source range", () => {
  const { controller } = controllerWith();
  controller.load([LINK]);
  const records = controller.listRecords();
  assert.equal(records.length, 1);
  assert.equal(records[0].datasetName, "C 82 - Prior Qtr Selected");
  assert.equal(records[0].sourceRange, "1:2");
  assert.equal(records[0].destination, "2017~2018");
  assert.equal(records[0].affectedCellCount, 2);
});

test("refreshAll applies new values whole-link and records failures per link", async () => {
  const { state, controller } = controllerWith({
    results: (references) => ({
      ok: true,
      status: 200,
      data: { ok: true, results: [resolvedVector([100, 200])] },
    }),
  });
  controller.load([LINK]);
  const refreshed = await controller.refreshAll();
  assert.equal(refreshed.changedCount, 2);
  assert.equal(refreshed.failedCount, 0);
  assert.equal(state.model.values[0][0], 100);
  assert.equal(state.model.values[1][0], 200);
  assert.equal(controller.getLinkFailures().length, 0);

  const failing = controllerWith({
    results: [{ ok: false, status: 422, data: { detail: "Dataset 'C 82' was not found." } }],
  });
  failing.controller.load([LINK]);
  const failure = await failing.controller.refreshAll();
  assert.equal(failure.failedCount, 2);
  assert.equal(failing.state.model.values[0][0], 1);
  assert.equal(failing.controller.getLinkFailures().length, 2);
  assert.match(failing.controller.getLinkFailures()[0].error, /not found/);
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

const spreadsheetCss = await readFile(
  new URL("../ui/shared/components/spreadsheet/spreadsheet_table.css", import.meta.url),
  "utf8",
);

test("a frame survives the pass of a controller that owns nothing", async () => {
  // The Data tab hands every cell to the Excel, ArcRho, and formula
  // controllers in turn. They share one set of perimeter classes, so a
  // controller owning no link must leave another's frame alone.
  const external = await import(externalUrl);
  const state = { model: vectorModel(), dirty: new Map() };
  const excelLinks = external.createDatasetExternalLinksController({ state });
  excelLinks.load([{
    reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1:A2",
    target_cells: [
      { row: 0, column: 0, source_cell: "A1" },
      { row: 1, column: 0, source_cell: "A2" },
    ],
  }]);
  const emptyInternal = internalLinks.createDatasetInternalLinksController({
    state,
    resolveReferences: async () => ({ ok: true, status: 200, data: { ok: true, results: [] } }),
  });
  const excelCell = decoratedCell();
  excelLinks.decorateCell(excelCell, 0, 0);
  emptyInternal.decorateCell(excelCell, 0, 0);

  assert.equal(excelCell.classes.has("arArrayFormulaCell"), true);
  assert.equal(excelCell.classes.has("arArrayFormulaEdgeTop"), true);
  assert.equal(excelCell.classes.has("arInternalLinkCell"), false);

  // The Excel pass comes first and equally has to leave an ArcRho frame alone.
  const { controller } = controllerWith();
  controller.load([LINK]);
  const emptyExcel = external.createDatasetExternalLinksController({
    state,
    isTransposed: () => false,
  });
  const arcRhoCell = decoratedCell();
  emptyExcel.decorateCell(arcRhoCell, 0, 0);
  controller.decorateCell(arcRhoCell, 0, 0);

  assert.equal(arcRhoCell.classes.has("arArrayFormulaCell"), true);
  assert.equal(arcRhoCell.classes.has("arArrayFormulaEdgeTop"), true);
  assert.equal(arcRhoCell.classes.has("arInternalLinkCell"), true);
});

test("a link is outlined in the color of where its values come from", () => {
  // Excel green is the default a linked range wears; a dataset link re-tints it.
  assert.match(
    spreadsheetCss,
    /td\.arArrayFormulaCell \{[^}]*--ar-array-formula-border: var\(--ar-spreadsheet-excel-link-border\)/u,
  );
  assert.match(
    spreadsheetCss,
    /td\.arArrayFormulaCell\.arInternalLinkCell \{[^}]*--ar-array-formula-border: var\(--ar-spreadsheet-internal-link-border\)/u,
  );
  assert.match(spreadsheetCss, /--ar-spreadsheet-excel-link-border: #217346;/u);
  assert.match(spreadsheetCss, /--ar-spreadsheet-internal-link-border: #2b6df6;/u);
  // Every edge takes its glow from the one token the two rules above set, so a
  // color can never be swapped on the border and missed on the glow.
  for (const edge of ["top", "right", "bottom", "left"]) {
    assert.ok(
      spreadsheetCss.includes(`--ar-array-formula-${edge}-glow: var(--ar-array-formula-glow);`),
      `the ${edge} edge glow follows the link's own color`,
    );
  }
  // Both kinds stay legible on a dark cell fill.
  assert.match(
    spreadsheetCss,
    /:root\[data-arcrho-theme="dark"\] \{[^}]*--ar-spreadsheet-excel-link-border: #3fbf7f;/u,
  );
});

test("a range offered to another window's formula is ringed by moving dashes", () => {
  assert.match(spreadsheetCss, /\.isReferencePickSource td\[data-r\]\[data-c\] \{\s*cursor: cell;/u);
  // Drawn as an overlay, so a theme's own cell background cannot swallow it.
  assert.match(
    spreadsheetCss,
    /td\.arReferencePickHover::before \{[^}]*box-shadow: inset 0 0 0 1px var\(--ar-spreadsheet-internal-link-border\)/u,
  );
  assert.match(
    spreadsheetCss,
    /td\.arReferencePickHover::before \{[^}]*background: var\(--ar-spreadsheet-reference-pick-hover-fill\)/u,
  );
  for (const edge of ["top", "right", "bottom", "left"]) {
    const name = `${edge[0].toUpperCase()}${edge.slice(1)}`;
    assert.ok(
      spreadsheetCss.includes(`td.arReferencePickEdge${name} {`),
      `the ${edge} edge of a picked range has a rule`,
    );
    assert.ok(
      spreadsheetCss.includes(`--ar-reference-pick-${edge}: repeating-linear-gradient(`),
      `the ${edge} edge of a picked range is drawn as dashes`,
    );
  }
  // The dash period divides both the cell width and the cell height, so the
  // dashes carry on unbroken from one cell into the next.
  assert.match(spreadsheetCss, /background-size: 4px 1px, 1px 4px, 4px 1px, 1px 4px;/u);
  assert.match(spreadsheetCss, /--ar-spreadsheet-cell-width: 100px;/u);
  assert.match(spreadsheetCss, /--ar-spreadsheet-cell-height: 20px;/u);
  assert.match(spreadsheetCss, /@keyframes arSpreadsheetReferencePickAnts \{/u);
  assert.match(spreadsheetCss, /prefers-reduced-motion: reduce[\s\S]*?animation: none;/u);
});
