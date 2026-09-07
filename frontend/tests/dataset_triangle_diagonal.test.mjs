// A triangle's cells stop on the project's calendar diagonal, and the Dataset
// window asks the project where that falls rather than laying the grid out on
// a rule of its own. The rule it used to use -- one column fewer per row --
// only ever matched a project whose origin and development periods are the
// same length, and let a 12/3 grid be typed into far past its own valuation
// date. A Vector has no diagonal and keeps every cell.
import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const requestControllerSource = await readFile(
  new URL("../ui/shared/tabs/data/data_tab_request_controller.js", import.meta.url),
  "utf8",
);

// The controller imports its siblings by their server-absolute `/ui/...`
// paths, which Node cannot resolve; none of them is reached on the path under
// test, so they are swapped for no-op stubs.
function importRequestController() {
  const stubbed = requestControllerSource.replace(
    /^import\s*\{([\s\S]*?)\}\s*from\s*"\/ui\/[^"]*";$/gmu,
    (_match, names) => `const {${names}} = __moduleStubs;`,
  );
  const source = `const __moduleStubs = new Proxy({}, { get: () => () => {} });\n${stubbed}`;
  return import(`data:text/javascript;base64,${Buffer.from(source).toString("base64")}`);
}

// Origins from 2017 through 2026 valued on 2026-05: ten annual rows over 38
// quarterly columns, each row four columns shorter than the one above it.
const ORIGIN_LABELS = Array.from({ length: 10 }, (_, index) => String(2017 + index));
const DEVELOPMENT_COUNT = 38;
const DIAGONAL = Array.from({ length: 10 }, (_, row) => (
  Array.from({ length: DEVELOPMENT_COUNT }, (_, column) => column < DEVELOPMENT_COUNT - 4 * row)
));
// The project's own development headers run past the valuation date.
const DEVELOPMENT_LABELS = Array.from({ length: 40 }, (_, index) => `${2 + 3 * index}m`);

function fakeSelect(value) {
  return { value: String(value), options: [{ value: String(value) }] };
}

async function createDraftRuntime({ dataFormat = "Triangle", devLen = 3 } = {}) {
  const elements = {
    originLenSelect: fakeSelect(12),
    devLenSelect: fakeSelect(devLen),
    triInput: { value: "Input Type" },
    dsMeta: { textContent: "" },
  };
  const originalDocument = globalThis.document;
  const originalFetch = globalThis.fetch;
  const requests = [];
  globalThis.document = {
    getElementById: (id) => elements[id] || null,
    querySelector: () => null,
    createElement: () => ({ style: {}, classList: { add() {}, toggle() {} } }),
    addEventListener() {},
  };
  globalThis.fetch = async (url) => {
    requests.push(String(url));
    return {
      ok: true,
      json: async () => ({
        ok: true,
        origin_count: DIAGONAL.length,
        development_count: DEVELOPMENT_COUNT,
        mask: DIAGONAL,
      }),
    };
  };

  const state = { dirty: new Map(), model: null, headerLabels: [], devHeaderLabels: [] };
  const runtime = {
    state,
    config: {},
    isTemporaryDatasetView: false,
    qs: new URLSearchParams(""),
    temporaryDatasetSessionId: "",
    LEN_DROPDOWN_CONFIG: {},
    statuses: [],
    readDatasetInputQueryValues: () => ({ project: "Project", path: "Class", tri: "Input Type", dataFormat }),
    normalizeReservingClassPath: (value) => value,
    normalizeBrowsingHistoryEntry: (entry) => entry,
    validateDatasetOriginLabels: () => ({ ok: true, labels: ORIGIN_LABELS }),
    getResolvedProjectValue: () => "Project",
    getResolvedReservingClassValue: () => "Class",
    getDatasetInstanceNameValue: () => "Instance",
    datasetHeadersService: {
      ensureHeadersForProject: async () => { state.headerLabels = ORIGIN_LABELS; },
      ensureDevHeadersForProject: async () => { state.devHeaderLabels = DEVELOPMENT_LABELS; },
    },
    updateDatasetSaveUi() {},
    notifyDatasetUpdated() {},
    renderTable() {},
    renderChart() {},
    setStatus(message) { runtime.statuses.push(String(message)); },
    createDatasetDependencyGuard: () => ({}),
    showProjectDropdown() {},
    showDatasetDropdown() {},
  };
  const { registerDataTabRequestController } = await importRequestController();
  registerDataTabRequestController(runtime);
  return {
    runtime,
    requests,
    restore: () => {
      globalThis.document = originalDocument;
      globalThis.fetch = originalFetch;
    },
  };
}

test("a new triangle takes its cells from the project's calendar diagonal", async () => {
  const { runtime, requests, restore } = await createDraftRuntime();
  try {
    assert.equal(await runtime.refreshProjectInstanceDraftModel(), true);

    // The shape is asked for at the lengths on screen, and it is the shape the
    // grid takes: the headers offer 40 columns, the triangle has 38.
    assert.deepEqual(requests, [
      "/datasets/triangle-shape?project_name=Project&origin_length=12&development_length=3",
    ]);
    const model = runtime.state.model;
    assert.equal(model.dev_labels.length, DEVELOPMENT_COUNT);
    assert.deepEqual(model.mask, DIAGONAL);

    // Four quarterly columns are dropped for each year of origin, not one, and
    // a cell past the diagonal carries no value for the save to write.
    assert.deepEqual(model.mask.map((row) => row.filter(Boolean).length), [38, 34, 30, 26, 22, 18, 14, 10, 6, 2]);
    assert.equal(model.values[1][33], 0);
    assert.equal(model.values[1][34], null);
  } finally {
    restore();
  }
});

test("a new vector keeps every cell and asks for no diagonal", async () => {
  const { runtime, requests, restore } = await createDraftRuntime({ dataFormat: "Vector", devLen: 12 });
  try {
    assert.equal(await runtime.refreshProjectInstanceDraftModel(), true);

    assert.deepEqual(requests, []);
    const model = runtime.state.model;
    assert.equal(model.dev_labels.length, 1);
    assert.deepEqual(model.mask, ORIGIN_LABELS.map(() => [true]));
  } finally {
    restore();
  }
});
