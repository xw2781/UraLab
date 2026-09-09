// Browser-side mirror of python-api/src/arcrho_api/cape_cod_contract.py.
// Every formula here is the ResQ Generalised Cape Cod calculation; keep this
// module in exact lockstep with the canonical python contract.

export const CC_JSON_FORMAT = "arcrho-cape-cod-v4";
export const CC_METHOD_TYPE = "Cape Cod";
export const CC_SOURCE_KIND = "cape_cod";

export const CC_PRIOR_ULTIMATE_MODES = ["latest_ultimates", "pattern"];
export const CC_SCALING_TYPES = ["percentage", "unscaled", "auto_scaled"];

function text(value) {
  return String(value ?? "").trim();
}

function finiteNumber(value) {
  if (value === null || value === undefined || value === "" || typeof value === "boolean") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

// Kept at the precision it was observed with (cape_cod_contract._number mirror).
// The expected-loss chain reads a DFM's factors, which carry full double
// precision, so quantizing the copy here would reintroduce the drift from ResQ.
export function roundCapeCodNumber(value) {
  return finiteNumber(value);
}

// Eight-decimal half-away-from-zero canonicalization for rate/factor parameters
// (cape_cod_contract._rate mirror; rates display as six-decimal percentages).
export function roundCapeCodRate(value) {
  const number = finiteNumber(value);
  if (number === null) return 0;
  const rounded = Math.round(Math.abs(number) * 100_000_000) / 100_000_000;
  return number < 0 ? -rounded : rounded;
}

export function roundCapeCodVector(values) {
  return Array.isArray(values) ? values.map(roundCapeCodNumber) : [];
}

function fitVector(values, rowCount, fill = null) {
  const fitted = roundCapeCodVector(values).slice(0, rowCount);
  while (fitted.length < rowCount) fitted.push(fill);
  return fitted;
}

export function normalizeCapeCodPriorUltimateMode(value) {
  const mode = text(value).toLowerCase().replace(/\s+/g, "_").replace(/\//g, "_");
  return CC_PRIOR_ULTIMATE_MODES.includes(mode) ? mode : CC_PRIOR_ULTIMATE_MODES[0];
}

export function normalizeCapeCodScalingType(value) {
  const scaling = text(value).toLowerCase().replace(/\s+/g, "_").replace(/-/g, "_");
  return CC_SCALING_TYPES.includes(scaling) ? scaling : CC_SCALING_TYPES[0];
}

// ResQ FitTrendRate: weighted log regression of the untrended developed ratio
// against origin position, weighted by developed exposure.
export function fitCapeCodTrendRate(latestValues, developedExposureValues) {
  const latest = Array.isArray(latestValues) ? latestValues : [];
  const developed = Array.isArray(developedExposureValues) ? developedExposureValues : [];
  const points = [];
  const count = Math.min(latest.length, developed.length);
  for (let index = 0; index < count; index += 1) {
    const latestValue = finiteNumber(latest[index]);
    const weight = finiteNumber(developed[index]);
    if (latestValue === null || weight === null || weight <= 0) continue;
    const ratio = latestValue / weight;
    if (ratio <= 0) continue;
    points.push([index, Math.log(ratio), weight]);
  }
  if (points.length < 2) return 0;
  let totalWeight = 0;
  let xw = 0;
  let yw = 0;
  let xxw = 0;
  let xyw = 0;
  for (const [x, y, weight] of points) {
    totalWeight += weight;
    xw += x * weight;
    yw += y * weight;
    xxw += x * x * weight;
    xyw += x * y * weight;
  }
  const sxx = xxw - (xw * xw) / totalWeight;
  const sxy = xyw - (xw * yw) / totalWeight;
  if (sxx === 0) return 0;
  return roundCapeCodRate(Math.exp(sxy / sxx) - 1);
}

function percentageDevelopedColumn(latest, prior, pattern, mode, rowCount) {
  const percentages = [];
  for (let index = 0; index < rowCount; index += 1) {
    const priorValue = finiteNumber(prior[index]);
    if (mode === "pattern") {
      percentages.push(roundCapeCodNumber(priorValue));
      continue;
    }
    // A prior ultimate a DFM published carries its own development pattern, and
    // that pattern is the percentage developed. Only a prior ultimate with no
    // DFM behind it falls back to the ratio against Latest.
    const developed = finiteNumber(pattern[index]);
    if (developed !== null) {
      percentages.push(roundCapeCodNumber(developed));
      continue;
    }
    const latestValue = finiteNumber(latest[index]);
    if (latestValue === null || priorValue === null || priorValue === 0) {
      percentages.push(null);
    } else {
      percentages.push(roundCapeCodNumber(latestValue / priorValue));
    }
  }
  return percentages;
}

function expectedUltimateRatioColumn(developedExposure, trendedDevelopedRatios, decay, rowCount) {
  const usable = [];
  for (let index = 0; index < rowCount; index += 1) {
    const weight = finiteNumber(developedExposure[index]);
    const ratio = finiteNumber(trendedDevelopedRatios[index]);
    if (weight === null || ratio === null) continue;
    usable.push([index, weight, ratio]);
  }
  const expected = [];
  for (let index = 0; index < rowCount; index += 1) {
    let numerator = 0;
    let denominator = 0;
    for (const [other, weight, ratio] of usable) {
      const decayed = weight * decay ** Math.abs(index - other);
      numerator += decayed * ratio;
      denominator += decayed;
    }
    expected.push(denominator !== 0 ? roundCapeCodNumber(numerator / denominator) : null);
  }
  return expected;
}

// Mirror of cape_cod_contract._calculate_columns: returns the effective trend
// rate plus every derived Method-tab column, all canonicalized.
export function calculateCapeCodColumns({
  originLabels,
  latestValues,
  exposureValues,
  priorUltimateValues,
  priorUltimatePercentageDeveloped,
  priorUltimateMode,
  trendRate,
  autoTrendFit,
  decayFactor,
  alternativeUltimateCalculation,
  trendFactorOverrides,
} = {}) {
  const labels = Array.isArray(originLabels) ? originLabels : [];
  const rowCount = labels.length;
  const latest = fitVector(latestValues, rowCount);
  const exposure = fitVector(exposureValues, rowCount);
  const prior = fitVector(priorUltimateValues, rowCount);
  const pattern = fitVector(priorUltimatePercentageDeveloped, rowCount);
  const mode = normalizeCapeCodPriorUltimateMode(priorUltimateMode);
  const decayValue = finiteNumber(decayFactor);
  const decay = decayValue === null ? 0 : decayValue;
  const autoFit = autoTrendFit === true;
  let overrides = fitVector(trendFactorOverrides, rowCount);
  if (autoFit) overrides = new Array(rowCount).fill(null);

  const percentages = percentageDevelopedColumn(latest, prior, pattern, mode, rowCount);
  const developedExposure = [];
  for (let index = 0; index < rowCount; index += 1) {
    const exposureValue = finiteNumber(exposure[index]);
    const percentage = finiteNumber(percentages[index]);
    developedExposure.push(
      exposureValue !== null && percentage !== null
        ? roundCapeCodNumber(exposureValue * percentage)
        : null,
    );
  }

  const effectiveTrendRate = autoFit
    ? fitCapeCodTrendRate(latest, developedExposure)
    : roundCapeCodRate(trendRate);

  const trendFactors = [];
  for (let index = 0; index < rowCount; index += 1) {
    const override = finiteNumber(overrides[index]);
    if (override !== null) {
      trendFactors.push(roundCapeCodNumber(override));
    } else {
      trendFactors.push(roundCapeCodNumber((1 + effectiveTrendRate) ** (rowCount - 1 - index)));
    }
  }

  const trendedLatest = [];
  const developmentFactors = [];
  const futureExposure = [];
  const trendedDevelopedRatios = [];
  const rawTrended = [];
  for (let index = 0; index < rowCount; index += 1) {
    const latestValue = finiteNumber(latest[index]);
    const factor = finiteNumber(trendFactors[index]);
    const trended = latestValue !== null && factor !== null ? latestValue * factor : null;
    rawTrended.push(trended);
    trendedLatest.push(roundCapeCodNumber(trended));
    const percentage = finiteNumber(percentages[index]);
    developmentFactors.push(
      percentage !== null && percentage !== 0 ? roundCapeCodNumber(1 / percentage) : null,
    );
    const exposureValue = finiteNumber(exposure[index]);
    const developed = finiteNumber(developedExposure[index]);
    futureExposure.push(
      exposureValue !== null && developed !== null
        ? roundCapeCodNumber(exposureValue - developed)
        : null,
    );
    trendedDevelopedRatios.push(
      trended !== null && developed !== null && developed !== 0
        ? roundCapeCodNumber(trended / developed)
        : null,
    );
  }

  const expected = expectedUltimateRatioColumn(developedExposure, trendedDevelopedRatios, decay, rowCount);

  const detrended = [];
  const futureLatest = [];
  const ultimates = [];
  const ratios = [];
  const alternative = alternativeUltimateCalculation === true;
  for (let index = 0; index < rowCount; index += 1) {
    const expectedValue = finiteNumber(expected[index]);
    const factor = finiteNumber(trendFactors[index]);
    const detrendedValue = expectedValue !== null && factor !== null && factor !== 0
      ? expectedValue / factor
      : null;
    detrended.push(roundCapeCodNumber(detrendedValue));
    const futureExposureValue = finiteNumber(futureExposure[index]);
    const futureValue = futureExposureValue !== null && detrendedValue !== null
      ? futureExposureValue * detrendedValue
      : null;
    futureLatest.push(roundCapeCodNumber(futureValue));
    const latestValue = finiteNumber(latest[index]);
    const percentage = finiteNumber(percentages[index]);
    const exposureValue = finiteNumber(exposure[index]);
    let ultimate = null;
    if (
      alternative
      && latestValue !== null
      && latestValue !== 0
      && percentage === 0
      && detrendedValue !== null
      && exposureValue !== null
    ) {
      ultimate = detrendedValue * exposureValue;
    } else if (latestValue !== null && futureValue !== null) {
      ultimate = latestValue + futureValue;
    }
    ultimates.push(roundCapeCodNumber(ultimate));
    ratios.push(
      ultimate !== null && exposureValue !== null && exposureValue !== 0
        ? roundCapeCodNumber(ultimate / exposureValue)
        : null,
    );
  }

  return {
    trendRate: effectiveTrendRate,
    trendFactorOverrides: overrides,
    trendFactors,
    trendedLatestValues: trendedLatest,
    percentageDeveloped: percentages,
    developmentFactors,
    developedExposureValues: developedExposure,
    futureExposureValues: futureExposure,
    trendedDevelopedRatios,
    expectedUltimateRatios: expected,
    detrendedExpectedRatios: detrended,
    futureLatestValues: futureLatest,
    capeCodUltimate: ultimates,
    capeCodUltimateRatios: ratios,
  };
}

// Mirror of cape_cod_contract.cape_cod_ultimates_triangle: the as-if diagnostic
// ultimates triangle. Returns null when the rows are not a regular triangle
// (one Latest row per origin with n - origin_index cells).
export function computeCapeCodUltimatesTriangle({
  exposureValues,
  percentageDeveloped,
  decayFactor,
  trendRate,
  alternativeUltimateCalculation,
} = {}, latestTriangleRows) {
  const rows = Array.isArray(latestTriangleRows) ? latestTriangleRows : [];
  const rowCount = rows.length;
  if (
    !rowCount
    || rows.some((row, index) => !Array.isArray(row) || row.length !== rowCount - index)
  ) {
    return null;
  }
  const exposure = fitVector(exposureValues, rowCount);
  const percentages = fitVector(percentageDeveloped, rowCount);
  const decayValue = finiteNumber(decayFactor);
  const decay = decayValue === null ? 0 : decayValue;
  const rate = finiteNumber(trendRate) || 0;
  const alternative = alternativeUltimateCalculation === true;

  const result = rows.map((row) => new Array(row.length).fill(null));
  for (let diagonal = 1; diagonal <= rowCount; diagonal += 1) {
    const cells = [];
    for (let origin = 0; origin < diagonal; origin += 1) {
      const column = diagonal - origin; // 1-based development column
      const latestValue = finiteNumber(roundCapeCodNumber(rows[origin][column - 1]));
      // The pattern for development column k is the current Method-tab
      // percentage developed of the origin whose leading diagonal sits in
      // column k (both share the same development age on a regular grid).
      const percentage = finiteNumber(percentages[rowCount - column]);
      cells.push([origin, latestValue, percentage]);
    }
    const newest = diagonal - 1;
    const usable = [];
    for (const [origin, latestValue, percentage] of cells) {
      const exposureValue = finiteNumber(exposure[origin]);
      if (latestValue === null || percentage === null || exposureValue === null) continue;
      const developed = exposureValue * percentage;
      if (developed === 0) continue;
      const factor = (1 + rate) ** (newest - origin);
      usable.push([origin, developed, (factor * latestValue) / developed]);
    }
    for (const [origin, latestValue, percentage] of cells) {
      const exposureValue = finiteNumber(exposure[origin]);
      if (latestValue === null || percentage === null || exposureValue === null) continue;
      let numerator = 0;
      let denominator = 0;
      for (const [other, developed, ratio] of usable) {
        const weight = developed * decay ** Math.abs(origin - other);
        numerator += weight * ratio;
        denominator += weight;
      }
      if (denominator === 0) continue;
      const factor = (1 + rate) ** (newest - origin);
      if (factor === 0) continue;
      const detrendedValue = numerator / denominator / factor;
      const developedExposure = exposureValue * percentage;
      const ultimate = alternative && latestValue !== 0 && percentage === 0
        ? detrendedValue * exposureValue
        : latestValue + (exposureValue - developedExposure) * detrendedValue;
      result[origin][diagonal - origin - 1] = roundCapeCodNumber(ultimate);
    }
  }
  return result;
}

export function rebaseCapeCodTrendFactorOverridesByOriginLabel({
  localOriginLabels,
  localOverrides,
  persistedOriginLabels,
  persistedOverrides,
} = {}) {
  const localLabels = Array.isArray(localOriginLabels) ? localOriginLabels : [];
  const localValues = Array.isArray(localOverrides) ? localOverrides : [];
  const persistedLabels = Array.isArray(persistedOriginLabels) ? persistedOriginLabels : [];
  const persistedValues = Array.isArray(persistedOverrides) ? persistedOverrides : [];
  const localOverrideByLabel = new Map();
  for (let index = 0; index < localLabels.length; index += 1) {
    const label = text(localLabels[index]);
    if (label) localOverrideByLabel.set(label, roundCapeCodNumber(localValues[index]));
  }
  return persistedLabels.map((rawLabel, index) => {
    const label = text(rawLabel);
    if (label && localOverrideByLabel.has(label)) return localOverrideByLabel.get(label);
    return roundCapeCodNumber(persistedValues[index]);
  });
}

export function isCapeCodV1Method(method) {
  return text(method?.json_format) === CC_JSON_FORMAT;
}

export function buildCapeCodMethodPayload({
  details,
  originLabels,
  latestValues,
  exposureValues,
  priorUltimateValues,
  priorUltimatePercentageDeveloped,
  priorUltimateMode,
  trendRate,
  autoTrendFit = false,
  decayFactor,
  scalingType,
  alternativeUltimateCalculation = false,
  trendFactorOverrides,
  methodMetadata = {},
  lastModified,
} = {}) {
  const safeDetails = details && typeof details === "object" ? details : {};
  const labels = Array.isArray(originLabels)
    ? originLabels.map((label) => String(label ?? ""))
    : [];
  const rowCount = labels.length;
  const metadata = methodMetadata && typeof methodMetadata === "object"
    ? { ...methodMetadata }
    : {};
  const columns = calculateCapeCodColumns({
    originLabels: labels,
    latestValues,
    exposureValues,
    priorUltimateValues,
    priorUltimatePercentageDeveloped,
    priorUltimateMode,
    trendRate,
    autoTrendFit,
    decayFactor,
    alternativeUltimateCalculation,
    trendFactorOverrides,
  });
  const rawDecimals = Number(
    safeDetails.statisticDecimalPlaces ?? safeDetails.statistic_decimal_places,
  );
  return {
    json_format: CC_JSON_FORMAT,
    details_tab: {
      name: text(safeDetails.name),
      method_type: CC_METHOD_TYPE,
      output_type: text(safeDetails.outputType ?? safeDetails.output_type),
      dataset_category: text(safeDetails.datasetCategory ?? safeDetails.dataset_category),
      origin_length: Number(safeDetails.originLength ?? safeDetails.origin_length) || 12,
      statistic_decimal_places: Number.isFinite(rawDecimals)
        ? Math.max(0, Math.min(8, rawDecimals))
        : 2,
    },
    method_tab: {
      latest_dataset: text(safeDetails.latestDataset ?? safeDetails.latest_dataset),
      latest_values: fitVector(latestValues, rowCount),
      exposure_dataset: text(safeDetails.exposureDataset ?? safeDetails.exposure_dataset),
      exposure_values: fitVector(exposureValues, rowCount),
      prior_ultimate_dataset: text(
        safeDetails.priorUltimateDataset ?? safeDetails.prior_ultimate_dataset,
      ),
      prior_ultimate_mode: normalizeCapeCodPriorUltimateMode(priorUltimateMode),
      prior_ultimate_values: fitVector(priorUltimateValues, rowCount),
      prior_ultimate_percentage_developed: fitVector(priorUltimatePercentageDeveloped, rowCount),
      trend_rate: columns.trendRate,
      auto_trend_fit: autoTrendFit === true,
      decay_factor: roundCapeCodRate(decayFactor),
      scaling_type: normalizeCapeCodScalingType(scalingType),
      alternative_ultimate_calculation: alternativeUltimateCalculation === true,
      trend_factor_overrides: columns.trendFactorOverrides,
      origin_labels: labels,
      trend_factors: columns.trendFactors,
      trended_latest_values: columns.trendedLatestValues,
      percentage_developed: columns.percentageDeveloped,
      development_factors: columns.developmentFactors,
      developed_exposure_values: columns.developedExposureValues,
      future_exposure_values: columns.futureExposureValues,
      trended_developed_ratios: columns.trendedDevelopedRatios,
      expected_ultimate_ratios: columns.expectedUltimateRatios,
      detrended_expected_ratios: columns.detrendedExpectedRatios,
      future_latest_values: columns.futureLatestValues,
      cape_cod_ultimate: columns.capeCodUltimate,
      cape_cod_ultimate_ratios: columns.capeCodUltimateRatios,
    },
    method_metadata: {
      ...metadata,
      method_type: CC_METHOD_TYPE,
      source_kind: CC_SOURCE_KIND,
      last_modified: text(lastModified ?? metadata.last_modified) || new Date().toISOString(),
    },
  };
}
