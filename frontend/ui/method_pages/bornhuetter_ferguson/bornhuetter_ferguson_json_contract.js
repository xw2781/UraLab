export const BORN_HUETTER_FERGUSON_JSON_FORMAT = "arcrho-bornhuetter-ferguson-v4";
export const BORN_HUETTER_FERGUSON_METHOD_TYPE = "Bornhuetter Ferguson";
export const BORN_HUETTER_FERGUSON_SOURCE_KIND = "bornhuetter_ferguson";

function text(value) {
  return String(value ?? "").trim();
}

// Kept at the precision it was observed with (bornhuetter_ferguson_contract._number
// mirror). A percentage developed and an ultimate come from a DFM that chains its
// factors in full double precision, so quantizing the copy here would put the drift
// from ResQ back one method further down.
export function roundBornhuetterFergusonNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

export function roundBornhuetterFergusonVector(values) {
  return Array.isArray(values) ? values.map(roundBornhuetterFergusonNumber) : [];
}

function fitVector(values, rowCount, fill = null) {
  const fitted = roundBornhuetterFergusonVector(values).slice(0, rowCount);
  while (fitted.length < rowCount) fitted.push(fill);
  return fitted;
}

function nonNegativeWeight(value, fallback = 1) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.max(0, number) : fallback;
}

export function rebaseBornhuetterFergusonWeightsByOriginLabel({
  localOriginLabels,
  localWeights,
  persistedOriginLabels,
  persistedWeights,
} = {}) {
  const localLabels = Array.isArray(localOriginLabels) ? localOriginLabels : [];
  const localValues = Array.isArray(localWeights) ? localWeights : [];
  const persistedLabels = Array.isArray(persistedOriginLabels) ? persistedOriginLabels : [];
  const persistedValues = Array.isArray(persistedWeights) ? persistedWeights : [];
  const localWeightByLabel = new Map();
  for (let index = 0; index < localLabels.length; index += 1) {
    const label = text(localLabels[index]);
    if (label) localWeightByLabel.set(label, nonNegativeWeight(localValues[index]));
  }
  return persistedLabels.map((rawLabel, index) => {
    const label = text(rawLabel);
    if (label && localWeightByLabel.has(label)) return localWeightByLabel.get(label);
    return nonNegativeWeight(persistedValues[index]);
  });
}

function normalizedPriorSources(priorSources, rowCount) {
  return (Array.isArray(priorSources) ? priorSources : [])
    .map((source) => ({
      name: text(source?.name ?? source?.dataset_name ?? source?.dataset),
      values: fitVector(source?.values, rowCount),
      weights: fitVector(source?.weights, rowCount, 1)
        .map((value) => Math.max(0, value ?? 1)),
    }))
    .filter((source) => source.name);
}

export function isBornhuetterFergusonV3Method(method) {
  return text(method?.json_format) === BORN_HUETTER_FERGUSON_JSON_FORMAT;
}

export function buildBornhuetterFergusonMethodPayload({
  details,
  originLabels,
  latestValues,
  dfmUltimateValues,
  priorSources,
  percentageDeveloped,
  selectedPriorValues,
  newUltimate,
  showWeights = true,
  showEffectiveWeights = false,
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
  return {
    json_format: BORN_HUETTER_FERGUSON_JSON_FORMAT,
    details_tab: {
      name: text(safeDetails.name),
      method_type: BORN_HUETTER_FERGUSON_METHOD_TYPE,
      output_type: text(safeDetails.outputType ?? safeDetails.output_type),
      dataset_category: text(safeDetails.datasetCategory ?? safeDetails.dataset_category),
      origin_length: Number(safeDetails.originLength ?? safeDetails.origin_length) || 12,
      statistic_decimal_places: Number.isFinite(Number(
        safeDetails.statisticDecimalPlaces ?? safeDetails.statistic_decimal_places,
      ))
        ? Math.max(0, Math.min(8, Number(
          safeDetails.statisticDecimalPlaces ?? safeDetails.statistic_decimal_places,
        )))
        : 1,
    },
    method_tab: {
      latest_dataset: text(safeDetails.latestDataset ?? safeDetails.latest_dataset),
      dfm_dataset: text(safeDetails.dfmDataset ?? safeDetails.dfm_dataset),
      show_weights: showWeights !== false,
      show_effective_weights: showEffectiveWeights === true,
      prior_datasets: normalizedPriorSources(priorSources, rowCount),
      origin_labels: labels,
      latest_values: fitVector(latestValues, rowCount),
      dfm_ultimate_values: fitVector(dfmUltimateValues, rowCount),
      percentage_developed: fitVector(percentageDeveloped, rowCount),
      selected_prior_values: fitVector(selectedPriorValues, rowCount),
      new_ultimate: fitVector(newUltimate, rowCount),
    },
    method_metadata: {
      ...metadata,
      method_type: BORN_HUETTER_FERGUSON_METHOD_TYPE,
      source_kind: BORN_HUETTER_FERGUSON_SOURCE_KIND,
      last_modified: text(lastModified ?? metadata.last_modified) || new Date().toISOString(),
    },
  };
}
