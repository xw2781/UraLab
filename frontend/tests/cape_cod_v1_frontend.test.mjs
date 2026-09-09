import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import {
  CC_JSON_FORMAT,
  CC_METHOD_TYPE,
  CC_SOURCE_KIND,
  buildCapeCodMethodPayload,
  computeCapeCodUltimatesTriangle,
  fitCapeCodTrendRate,
  rebaseCapeCodTrendFactorOverridesByOriginLabel,
  roundCapeCodNumber,
  roundCapeCodRate,
} from "../ui/method_pages/cape_cod/cape_cod_json_contract.js";
import {
  loadCapeCodMethod,
  saveCapeCodMethod,
} from "../ui/method_pages/cape_cod/cape_cod_method_api.js";

const mainSource = await readFile(
  new URL("../ui/method_pages/cape_cod/cape_cod_main.js", import.meta.url),
  "utf8",
);
const fixture = JSON.parse(await readFile(
  new URL("../../python-api/tests/fixtures/resq_cape_cod_d53.json", import.meta.url),
  "utf8",
));

function functionSlice(source, startMarker, endMarker) {
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.notEqual(start, -1, `missing ${startMarker}`);
  assert.notEqual(end, -1, `missing ${endMarker}`);
  return source.slice(start, end);
}

function buildFixturePayload(overrides = {}) {
  return buildCapeCodMethodPayload({
    details: {
      name: fixture.method.name,
      outputType: fixture.method.name,
      datasetCategory: "Claims",
      originLength: fixture.method.origin_length,
      latestDataset: fixture.method.latest_dataset,
      exposureDataset: fixture.method.exposure_dataset,
      priorUltimateDataset: fixture.method.prior_ultimate_dataset,
      statisticDecimalPlaces: fixture.method.decimal_places,
    },
    originLabels: fixture.origin_labels,
    latestValues: fixture.latest_values,
    exposureValues: fixture.exposure_values,
    priorUltimateValues: fixture.prior_ultimate_values,
    priorUltimateMode: "latest_ultimates",
    trendRate: fixture.method.trend_rate,
    autoTrendFit: fixture.method.auto_trend_fit,
    decayFactor: fixture.method.decay_factor,
    scalingType: "percentage",
    alternativeUltimateCalculation: fixture.method.alternative_ultimate_calculation,
    trendFactorOverrides: null,
    methodMetadata: {},
    lastModified: "2026-08-04T00:00:00Z",
    ...overrides,
  });
}

function assertClose(actual, expected, tolerance, label) {
  if (expected === null || expected === undefined) {
    assert.equal(actual, null, `${label} should be blank`);
    return;
  }
  assert.ok(actual !== null && actual !== undefined, `${label} should not be blank`);
  const delta = Math.abs(Number(actual) - Number(expected));
  assert.ok(
    delta <= tolerance,
    `${label}: |${actual} - ${expected}| = ${delta} > ${tolerance}`,
  );
}

test("Cape Cod JS payload reproduces the ResQ-verified D 53 fixture columns", () => {
  const payload = buildFixturePayload();
  const method = payload.method_tab;
  const exposure = fixture.exposure_values;

  // Fitted trend rate (auto fit) must match the ResQ FitTrendRate result.
  assertClose(method.trend_rate, fixture.method.trend_rate, 1e-7, "trend_rate");

  const exposureScaled = new Set(["future_exposure_values", "future_latest_values"]);
  for (const [column, expectedValues] of Object.entries(fixture.expected)) {
    const actualValues = method[column];
    assert.ok(Array.isArray(actualValues), `missing derived column ${column}`);
    assert.equal(actualValues.length, expectedValues.length, `${column} length`);
    for (let index = 0; index < expectedValues.length; index += 1) {
      const expected = expectedValues[index];
      const tolerance = exposureScaled.has(column)
        ? 2e-6 * Math.max(1, Math.abs(Number(exposure[index]) || 0))
        : 2e-6 * Math.max(1, Math.abs(Number(expected) || 0));
      assertClose(actualValues[index], expected, tolerance, `${column}[${index}]`);
    }
  }
});

test("Cape Cod as-if ultimates triangle matches the ResQ UltimateTriangleValues fixture", () => {
  const payload = buildFixturePayload();
  const method = payload.method_tab;
  const triangle = computeCapeCodUltimatesTriangle({
    exposureValues: method.exposure_values,
    percentageDeveloped: method.percentage_developed,
    decayFactor: method.decay_factor,
    trendRate: method.trend_rate,
    alternativeUltimateCalculation: method.alternative_ultimate_calculation,
  }, fixture.latest_triangle);

  assert.ok(Array.isArray(triangle), "triangle should compute for the regular fixture");
  assert.equal(triangle.length, fixture.expected_ultimates_triangle.length);
  for (let origin = 0; origin < triangle.length; origin += 1) {
    const expectedRow = fixture.expected_ultimates_triangle[origin];
    assert.equal(triangle[origin].length, expectedRow.length, `triangle row ${origin} length`);
    const exposureTolerance = 2e-6 * Math.max(1, Math.abs(Number(fixture.exposure_values[origin]) || 0));
    for (let column = 0; column < expectedRow.length; column += 1) {
      const expected = expectedRow[column];
      const tolerance = Math.max(
        exposureTolerance,
        2e-6 * Math.max(1, Math.abs(Number(expected) || 0)),
      );
      assertClose(triangle[origin][column], expected, tolerance, `triangle[${origin}][${column}]`);
    }
  }
});

test("Cape Cod v1 payload is self-contained with canonical identity labels", () => {
  const payload = buildFixturePayload({
    methodMetadata: {
      owned_revision: "owned-in-method",
      derived_revision: "derived-in-method",
    },
  });

  assert.equal(payload.json_format, CC_JSON_FORMAT);
  assert.equal(CC_JSON_FORMAT, "arcrho-cape-cod-v4");
  assert.equal(payload.details_tab.method_type, CC_METHOD_TYPE);
  assert.equal(CC_METHOD_TYPE, "Cape Cod");
  assert.equal(payload.method_metadata.source_kind, CC_SOURCE_KIND);
  assert.equal(CC_SOURCE_KIND, "cape_cod");
  assert.equal(payload.details_tab.statistic_decimal_places, 2);
  assert.equal(payload.method_tab.prior_ultimate_mode, "latest_ultimates");
  assert.equal(payload.method_tab.scaling_type, "percentage");
  assert.equal(payload.method_tab.auto_trend_fit, true);
  assert.deepEqual(payload.method_tab.trend_factor_overrides, new Array(fixture.origin_labels.length).fill(null));
  // v4 persists only the four sections the python contract normalizes: the
  // always-empty placeholder tabs and the audit log are gone from method files.
  assert.deepEqual(
    Object.keys(payload),
    ["json_format", "details_tab", "method_tab", "method_metadata"],
  );
  for (const retired of ["ultimates_tab", "ratios_tab", "audit_log_tab", "audit_log"]) {
    assert.equal(retired in payload, false, `${retired} must not be persisted`);
  }
  assert.equal(payload.method_metadata.owned_revision, "owned-in-method");
  assert.equal(payload.method_metadata.derived_revision, "derived-in-method");
});

test("Cape Cod numbers keep their precision and rates round at eight decimals", () => {
  // A value read from a DFM is carried whole; only a rate box, which offers
  // eight decimals of its own, is canonicalized to that precision.
  assert.equal(roundCapeCodNumber(1.2345675), 1.2345675);
  assert.equal(roundCapeCodNumber(-1.2345675), -1.2345675);
  assert.equal(roundCapeCodNumber(null), null);
  assert.equal(roundCapeCodRate(0.123456785), 0.12345679);
  assert.equal(roundCapeCodRate(-0.123456785), -0.12345679);
  assert.equal(roundCapeCodRate(null), 0);
});

test("Cape Cod trend-rate fit excludes unusable rows and needs two points", () => {
  assert.equal(fitCapeCodTrendRate([1, 2], [0, 0]), 0);
  assert.equal(fitCapeCodTrendRate([1], [1]), 0);
  const fitted = fitCapeCodTrendRate([100, 110, 121], [100, 100, 100]);
  assertClose(fitted, 0.1, 1e-7, "fitted rate");
});

test("Cape Cod dirty refresh rebases local trend factor overrides by origin label", () => {
  assert.deepEqual(rebaseCapeCodTrendFactorOverridesByOriginLabel({
    localOriginLabels: ["2022", "2023"],
    localOverrides: [1.25, null],
    persistedOriginLabels: ["2023", "2024", "2025"],
    persistedOverrides: [1.4, 1.6],
  }), [null, 1.6, null]);

  const restore = functionSlice(
    mainSource,
    "function restoreLocalOwnedState",
    "async function applyPersistedAggregate",
  );
  assert.match(restore, /rebaseCapeCodTrendFactorOverridesByOriginLabel/u);
  assert.match(restore, /localOriginLabels:\s*local\.originLabels/u);
  assert.match(restore, /persistedOriginLabels:\s*nextOriginLabels/u);
});

test("Cape Cod aggregate API sends identity and revision-aware save requests", async () => {
  const requests = [];
  const previousFetch = globalThis.fetch;
  globalThis.fetch = async (path, init) => {
    requests.push({ path, body: JSON.parse(init.body) });
    return {
      ok: true,
      status: 200,
      json: async () => ({ ok: true, method: { json_format: CC_JSON_FORMAT } }),
    };
  };
  try {
    await loadCapeCodMethod({
      project_name: "Project",
      reserving_class: "RC",
      method_name: "CC Selection",
    });
    await saveCapeCodMethod({
      project_name: "Project",
      reserving_class: "RC",
      method: { json_format: CC_JSON_FORMAT },
      notes: "keep",
      expected_owned_revision: "owned",
      expected_derived_revision: "derived",
    });
  } finally {
    globalThis.fetch = previousFetch;
  }

  assert.deepEqual(requests, [{
    path: "/cape-cod/load",
    body: {
      project_name: "Project",
      reserving_class: "RC",
      method_name: "CC Selection",
    },
  }, {
    path: "/cape-cod/save",
    body: {
      project_name: "Project",
      reserving_class: "RC",
      method: { json_format: CC_JSON_FORMAT },
      notes: "keep",
      expected_owned_revision: "owned",
      expected_derived_revision: "derived",
    },
  }]);
});

test("existing Cape Cod open applies only the aggregate v1 method and sidecar", () => {
  const load = functionSlice(
    mainSource,
    "async function tryLoadExistingMethod()",
    "async function reloadPersistedCapeCod",
  );
  const apply = functionSlice(
    mainSource,
    "async function applyPersistedAggregate",
    "async function fetchPersistedCapeCod",
  );
  const init = functionSlice(mainSource, "async function init()", "void init()");

  assert.match(load, /fetchPersistedCapeCod\(\)/u);
  assert.match(load, /applyPersistedAggregate\(result\)/u);
  assert.match(apply, /applyOutputSidecar\(result\?\.sidecar/u);
  assert.match(apply, /applyPayload\(method\)/u);
  assert.match(apply, /ultimates_triangle/u);
  assert.doesNotMatch(
    `${load}\n${apply}`,
    /loadCachedRows|loadConfiguredSourcePayload|refreshCalculations|refreshOriginLabels|readJsonFile|loadSidecar|workspace_paths|dataset\/cache\/load|dataset\/sidecar\/load/u,
  );
  assert.doesNotMatch(init, /loadCachedRows|loadSidecar|refreshCalculations/u);
});

test("Cape Cod persisted apply uses stored vectors without dependency reads", () => {
  const apply = functionSlice(mainSource, "async function applyPayload(payload)", "function snapshotPayload");

  for (const field of [
    "latest_values",
    "exposure_values",
    "prior_ultimate_values",
    "trend_factor_overrides",
    "percentage_developed",
    "developed_exposure_values",
    "expected_ultimate_ratios",
    "detrended_expected_ratios",
    "cape_cod_ultimate",
    "cape_cod_ultimate_ratios",
  ]) {
    assert.match(apply, new RegExp(field, "u"));
  }
  assert.doesNotMatch(apply, /fetch\(|loadCachedRows|loadConfiguredSourcePayload|refreshCalculations|refreshOriginLabels/u);
});

test("Cape Cod dependency clear restores through aggregate load and save uses aggregate persistence", () => {
  const clear = functionSlice(
    mainSource,
    "async function clearDependencySourcePreview",
    "function originLabel",
  );
  const save = functionSlice(
    mainSource,
    "async function saveCapeCod()",
    "function setNotesText",
  );

  assert.match(clear, /reloadPersistedCapeCod/u);
  assert.doesNotMatch(clear, /loadCachedRows|loadConfiguredSourcePayload|dataset\/cache\/load/u);
  assert.match(save, /saveCapeCodMethod/u);
  assert.match(save, /expected_owned_revision:\s*state\.ownedRevision/u);
  assert.match(save, /expected_derived_revision:\s*state\.derivedRevision/u);
  assert.doesNotMatch(save, /readJsonFile|saveJsonFile|saveTextFile|loadCachedRows|loadSidecar|refreshCalculations|dataset\/sidecar\/save/u);
});
