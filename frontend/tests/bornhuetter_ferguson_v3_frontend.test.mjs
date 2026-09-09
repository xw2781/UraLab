import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import {
  BORN_HUETTER_FERGUSON_JSON_FORMAT,
  buildBornhuetterFergusonMethodPayload,
  rebaseBornhuetterFergusonWeightsByOriginLabel,
  roundBornhuetterFergusonNumber,
} from "../ui/method_pages/bornhuetter_ferguson/bornhuetter_ferguson_json_contract.js";
import {
  loadBornhuetterFergusonMethod,
  saveBornhuetterFergusonMethod,
} from "../ui/method_pages/bornhuetter_ferguson/bornhuetter_ferguson_method_api.js";

const mainSource = await readFile(
  new URL("../ui/method_pages/bornhuetter_ferguson/bornhuetter_ferguson_main.js", import.meta.url),
  "utf8",
);

function functionSlice(source, startMarker, endMarker) {
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.notEqual(start, -1, `missing ${startMarker}`);
  assert.notEqual(end, -1, `missing ${endMarker}`);
  return source.slice(start, end);
}

function logicalMethod(overrides = {}) {
  return {
    details: {
      name: "BF Selection",
      outputType: "BF Ultimate",
      datasetCategory: "Claims",
      originLength: 12,
      latestDataset: "Paid Triangle",
      dfmDataset: "Paid DFM",
      statisticDecimalPlaces: 2,
    },
    originLabels: ["2025", "2026"],
    latestValues: [10.1234567, 20],
    dfmUltimateValues: [25, 40],
    priorSources: [{
      name: "Prior Ultimate",
      values: [30, 50],
      weights: [1, 0.5],
    }],
    percentageDeveloped: [0.4, 0.5],
    selectedPriorValues: [30, 50],
    newUltimate: [28, 45],
    showWeights: true,
    showEffectiveWeights: true,
    methodMetadata: {
      owned_revision: "owned-in-method",
      derived_revision: "derived-in-method",
    },
    lastModified: "2026-07-26T00:00:00Z",
    ...overrides,
  };
}

test("BF v3 payload is self-contained and preserves effective-weight and revision metadata", () => {
  const payload = buildBornhuetterFergusonMethodPayload(logicalMethod());

  assert.equal(payload.json_format, BORN_HUETTER_FERGUSON_JSON_FORMAT);
  assert.equal(payload.method_tab.show_effective_weights, true);
  assert.equal(payload.details_tab.dataset_category, "Claims");
  assert.deepEqual(payload.method_tab.latest_values, [10.1234567, 20]);
  assert.deepEqual(payload.method_tab.dfm_ultimate_values, [25, 40]);
  assert.deepEqual(payload.method_tab.prior_datasets, [{
    name: "Prior Ultimate",
    values: [30, 50],
    weights: [1, 0.5],
  }]);
  assert.deepEqual(payload.method_tab.percentage_developed, [0.4, 0.5]);
  assert.deepEqual(payload.method_tab.selected_prior_values, [30, 50]);
  assert.deepEqual(payload.method_tab.new_ultimate, [28, 45]);
  assert.equal(payload.method_metadata.owned_revision, "owned-in-method");
  assert.equal(payload.method_metadata.derived_revision, "derived-in-method");
});

test("BF numbers keep the precision they were observed with", () => {
  // The vector a BF reads is a DFM's, chained in full double precision, so the
  // copy is carried whole rather than projected onto six decimals.
  assert.equal(roundBornhuetterFergusonNumber(1.2345675), 1.2345675);
  assert.equal(roundBornhuetterFergusonNumber(-1.2345675), -1.2345675);
  assert.equal(roundBornhuetterFergusonNumber(null), null);
  assert.equal(roundBornhuetterFergusonNumber(""), null);
  assert.equal(roundBornhuetterFergusonNumber("nope"), null);
});

test("BF dirty refresh rebases local weights by origin label", () => {
  assert.deepEqual(rebaseBornhuetterFergusonWeightsByOriginLabel({
    localOriginLabels: ["2022", "2023"],
    localWeights: [0.25, 0.75],
    persistedOriginLabels: ["2023", "2024", "2025"],
    persistedWeights: [0.4, 0.6],
  }), [0.75, 0.6, 1]);

  const restore = functionSlice(
    mainSource,
    "function restoreLocalOwnedState",
    "async function applyPersistedAggregate",
  );
  assert.match(restore, /rebaseBornhuetterFergusonWeightsByOriginLabel/u);
  assert.match(restore, /localOriginLabels:\s*local\.originLabels/u);
  assert.match(restore, /persistedOriginLabels:\s*nextOriginLabels/u);
});

test("BF output calculation keeps the ultimate's fraction at six decimals", () => {
  // ResQ never rounds a BF ultimate to a whole number; everything reading the
  // BF output vector drifted from ResQ while the page did.
  const calculate = functionSlice(mainSource, "function calculateOutputs()", "function renderBfChart");
  assert.match(calculate, /roundBornhuetterFergusonNumber\(latest \+ \(1 - pct\) \* selectedPrior\)/u);
  assert.doesNotMatch(calculate, /WholeNumber|Math\.round/u);
});

test("BF aggregate API sends identity and revision-aware save requests", async () => {
  const requests = [];
  const previousFetch = globalThis.fetch;
  globalThis.fetch = async (path, init) => {
    requests.push({ path, body: JSON.parse(init.body) });
    return {
      ok: true,
      status: 200,
      json: async () => ({ ok: true, method: { json_format: BORN_HUETTER_FERGUSON_JSON_FORMAT } }),
    };
  };
  try {
    await loadBornhuetterFergusonMethod({
      project_name: "Project",
      reserving_class: "RC",
      method_name: "BF Selection",
    });
    await saveBornhuetterFergusonMethod({
      project_name: "Project",
      reserving_class: "RC",
      method: { json_format: BORN_HUETTER_FERGUSON_JSON_FORMAT },
      notes: "keep",
      expected_owned_revision: "owned",
      expected_derived_revision: "derived",
    });
  } finally {
    globalThis.fetch = previousFetch;
  }

  assert.deepEqual(requests, [{
    path: "/bornhuetter-ferguson/load",
    body: {
      project_name: "Project",
      reserving_class: "RC",
      method_name: "BF Selection",
    },
  }, {
    path: "/bornhuetter-ferguson/save",
    body: {
      project_name: "Project",
      reserving_class: "RC",
      method: { json_format: BORN_HUETTER_FERGUSON_JSON_FORMAT },
      notes: "keep",
      expected_owned_revision: "owned",
      expected_derived_revision: "derived",
    },
  }]);
});

test("existing BF open applies only the aggregate v3 method and sidecar", () => {
  const load = functionSlice(
    mainSource,
    "async function tryLoadExistingMethod()",
    "async function reloadPersistedBornhuetterFerguson",
  );
  const apply = functionSlice(
    mainSource,
    "async function applyPersistedAggregate",
    "async function fetchPersistedBornhuetterFerguson",
  );
  const init = functionSlice(mainSource, "async function init()", "void init()");

  assert.match(load, /fetchPersistedBornhuetterFerguson\(\)/u);
  assert.match(load, /applyPersistedAggregate\(result\)/u);
  assert.match(apply, /applyOutputSidecar\(result\?\.sidecar/u);
  assert.match(apply, /applyPayload\(method\)/u);
  assert.doesNotMatch(
    `${load}\n${apply}`,
    /loadCachedRows|loadConfiguredSourcePayload|refreshCalculations|refreshOriginLabels|readJsonFile|loadSidecar|workspace_paths|dataset\/cache\/load|dataset\/sidecar\/load/u,
  );
  assert.doesNotMatch(init, /loadCachedRows|loadSidecar|refreshCalculations/u);
  assert.match(init, /if \(loaded\) \{\s*postStatus\(`\$\{BF_METHOD_TYPE\} ready\.`\);\s*\} else if \(!loadError\)/u);
});

test("BF persisted apply uses stored vectors without dependency reads", () => {
  const apply = functionSlice(mainSource, "async function applyPayload(payload)", "function snapshotPayload");

  for (const field of [
    "latest_values",
    "dfm_ultimate_values",
    "prior_datasets",
    "percentage_developed",
    "selected_prior_values",
    "new_ultimate",
    "show_effective_weights",
  ]) {
    assert.match(apply, new RegExp(field, "u"));
  }
  assert.doesNotMatch(apply, /fetch\(|loadCachedRows|loadConfiguredSourcePayload|refreshCalculations|refreshOriginLabels/u);
});

test("BF dependency clear restores through aggregate load and save uses aggregate persistence", () => {
  const clear = functionSlice(
    mainSource,
    "async function clearDependencySourcePreview",
    "function renderPriorSourceList",
  );
  const save = functionSlice(
    mainSource,
    "async function saveBornhuetterFerguson()",
    "function setNotesText",
  );

  assert.match(clear, /reloadPersistedBornhuetterFerguson/u);
  assert.doesNotMatch(clear, /loadCachedRows|loadConfiguredSourcePayload|dataset\/cache\/load/u);
  assert.match(save, /saveBornhuetterFergusonMethod/u);
  assert.match(save, /expected_owned_revision:\s*state\.ownedRevision/u);
  assert.match(save, /expected_derived_revision:\s*state\.derivedRevision/u);
  assert.doesNotMatch(save, /readJsonFile|saveJsonFile|saveTextFile|loadCachedRows|loadSidecar|refreshCalculations|dataset\/sidecar\/save/u);
});
