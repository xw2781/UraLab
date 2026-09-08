import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import {
  RESULT_SELECTION_JSON_FORMAT,
  buildResultSelectionMethodPayload,
  normalizeRatioBasisValueSets,
  ratioBasisValuesForName,
  roundResultSelectionNumber,
} from "../ui/method_pages/result_selection/result_selection_json_contract.js";
import {
  hasResultSelectionUpdates,
  resultSelectionUpdateContexts,
  resultSelectionUpdateNames,
} from "../ui/shared/dataset/result_selection_update_report.js";


function logicalMethod(overrides = {}) {
  return {
    details: {
      name: "Selection",
      outputType: "Selected Ultimate",
      originLength: 12,
      ratioBasis: "Premium",
      ratioBases: [" Premium ", "premium", "Exposure"],
      showRatiosAsPercentages: true,
      statisticDecimalPlaces: 1,
    },
    originLabels: ["2025", "2026"],
    showWeights: true,
    sources: [{
      name: "Paid",
      datasetType: "Paid",
      dataFormat: "Vector",
      methodType: "None",
      category: "Loss",
      sourceKind: "input",
      values: [10.12345678, 20],
      weights: [1, 0],
    }],
    ratioBasisValueSets: [
      { name: "Exposure", values: [300, 400] },
      { name: "Premium", values: [100, null] },
    ],
    calculatedUltimate: [10.12345678, null],
    selectedUltimate: [10.12345678, 99],
    ultimateOverrides: [null, 99],
    lastModified: "2026-01-01T00:00:00Z",
    ...overrides,
  };
}


test("Result Selection v2 payload stores deterministic named Ratio Basis vectors", () => {
  const payload = buildResultSelectionMethodPayload(logicalMethod());

  assert.equal(payload.json_format, RESULT_SELECTION_JSON_FORMAT);
  assert.deepEqual(payload.details_tab.ratio_basis_datasets, ["Premium", "Exposure"]);
  assert.equal(payload.details_tab.active_ratio_basis_dataset, "Premium");
  assert.equal("ratio_basis" in payload.details_tab, false);
  assert.equal("ratio_basis_dataset" in payload.details_tab, false);
  assert.deepEqual(payload.method_tab.ratio_basis_values, [
    { name: "Premium", values: [100, null] },
    { name: "Exposure", values: [300, 400] },
  ]);
});

test("Result Selection refuses to save a configured basis without a complete vector", () => {
  assert.throws(
    () => buildResultSelectionMethodPayload(logicalMethod({
      details: {
        ...logicalMethod().details,
        ratioBases: ["Premium", "Exposure", "Missing"],
      },
    })),
    /Missing.*exactly 2 origin values/,
  );
});

test("Ratio Basis active values are derived by name without I/O", () => {
  const sets = normalizeRatioBasisValueSets([
    { name: "Exposure", values: [3, 4] },
    { name: "Premium", values: [1, 2] },
  ], ["Premium", "Exposure"]);

  assert.deepEqual(ratioBasisValuesForName(sets, " premium "), [1, 2]);
  assert.deepEqual(ratioBasisValuesForName(sets, "Exposure"), [3, 4]);
});

test("six-decimal Result Selection rounding is symmetric half-away-from-zero", () => {
  assert.equal(roundResultSelectionNumber(1.2345675), 1.234568);
  assert.equal(roundResultSelectionNumber(-1.2345675), -1.234568);
});

test("existing Result Selection apply uses persisted values without source or basis reloads", async () => {
  const source = await readFile(
    new URL("../ui/method_pages/result_selection/result_selection_model.js", import.meta.url),
    "utf8",
  );
  const applyBody = source.match(/async function applyPayload\(payload\) \{([\s\S]*?)\n      \}\n\n      function applyOutputSidecar/u)?.[1] || "";

  assert.match(applyBody, /buildSourceFromPersisted/);
  assert.match(applyBody, /ratioBasisValuesForName/);
  assert.doesNotMatch(applyBody, /buildSourceFromRecord|refreshRatioBasisValues|loadDatasetValues/);
});

test("calculated update reports identify only affected Result Selection outputs", () => {
  const report = {
    updated: [{ dataset_name: "Calculated Loss" }],
    result_selection_updates: {
      updated: [{ dataset_name: "Selection A" }],
      status_refreshed: [{ dataset_name: "Selection B" }],
      errors: [{ dataset_name: "Selection C", reason: "refresh failed" }],
    },
  };

  assert.deepEqual(
    Array.from(resultSelectionUpdateNames(report)),
    ["Selection A", "Selection B", "Selection C"],
  );
  assert.equal(hasResultSelectionUpdates(report), true);
  assert.equal(hasResultSelectionUpdates({ updated: [{ dataset_name: "Calculated Loss" }] }), false);
  assert.deepEqual(resultSelectionUpdateContexts({
    project_name: "Project A",
    reserving_class: "Class A",
    result_selection_updates: report.result_selection_updates,
  }), [{ project: "Project A", reservingClass: "Class A" }]);
});

test("persisted dependency refreshes use a request lease and filtered messages", async () => {
  const source = await readFile(
    new URL("../ui/method_pages/result_selection/result_selection_ui.js", import.meta.url),
    "utf8",
  );

  assert.match(source, /persistedRefreshSeq/);
  assert.match(source, /calculatedUpdateAffectsCurrentResultSelection/);
  assert.match(source, /sourceMessageMatchesRatioBasis/);
  assert.match(source, /dependencyPreviews\.size/);
  assert.match(source, /reportMatchesCurrentContext/);
  assert.match(source, /dependencyRestorePending/);
  assert.match(source, /reloadLocalBasis/);
});

test("Result Selection previews tolerate an incomplete Ratio Basis load", async () => {
  const source = await readFile(
    new URL("../ui/method_pages/result_selection/result_selection_ui.js", import.meta.url),
    "utf8",
  );

  assert.match(source, /payload\.values = selectedUltimateVector\(\);/);
  assert.doesNotMatch(source, /payload\.values = buildPayload\(\)\.method_tab\.selected_ultimate/);
});

test("Result Selection saves preserve newer dependency refreshes", async () => {
  const modelSource = await readFile(
    new URL("../ui/method_pages/result_selection/result_selection_model.js", import.meta.url),
    "utf8",
  );

  assert.match(modelSource, /await refreshOriginLabels\(\{ render: false \}\);\s*assertPersistedMutationReady\(mutation\);\s*const method = buildPayload\(\);/u);
  assert.doesNotMatch(modelSource, /^\s*invalidatePersistedRefresh\(\);/mu);
  assert.match(modelSource, /reconcilePersistedMutation\(mutation/u);
});

test("dependency clears coalesce and local-only reloads do not rebuild the dataset index", async () => {
  const source = await readFile(
    new URL("../ui/method_pages/result_selection/result_selection_ui.js", import.meta.url),
    "utf8",
  );
  const reloadBody = source.match(
    /async function reloadSourcesMatchingMessages\(messages = \[\]\) \{([\s\S]*?)\n      \}\n\n      async function reloadSourcesMatchingMessage/u,
  )?.[1] || "";

  assert.match(source, /dependencyRefreshPromise/);
  assert.match(source, /scheduleDependencyClearRestore/);
  assert.match(source, /if \(!refreshed\) \{\s*throw new Error/u);
  assert.doesNotMatch(reloadBody, /loadCachedRows\(true\)|refresh=true/u);
});
