export const RESULT_SELECTION_JSON_FORMAT = "arcrho-result-selection-v4";
export const RESULT_SELECTION_VALUE_DECIMAL_PLACES = 6;

function text(value) {
  return String(value ?? "").trim();
}

function key(value) {
  return text(value).replace(/\s+/g, " ").toLowerCase();
}

export function canonicalRatioBasisNames(names) {
  const out = [];
  const seen = new Set();
  for (const rawName of Array.isArray(names) ? names : []) {
    const name = text(rawName);
    const normalizedName = key(name);
    if (!normalizedName || seen.has(normalizedName)) continue;
    seen.add(normalizedName);
    out.push(name);
    if (out.length >= 3) break;
  }
  return out;
}

// Carried at the precision it was observed with (result_selection_service._round_number
// mirror). Every value is another method's ultimate copied in, so quantizing the copy
// made a weighted average of several ultimates disagree with the same average in ResQ.
export function roundResultSelectionNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

export function roundResultSelectionVector(values) {
  return Array.isArray(values) ? values.map(roundResultSelectionNumber) : [];
}

function fitResultSelectionVector(values, rowCount, fill = null) {
  const fitted = roundResultSelectionVector(values).slice(0, rowCount);
  while (fitted.length < rowCount) fitted.push(fill);
  return fitted;
}

export function normalizeRatioBasisValueSets(valueSets, configuredNames = []) {
  const byName = new Map();
  for (const item of Array.isArray(valueSets) ? valueSets : []) {
    if (!item || typeof item !== "object") continue;
    const name = text(item.name);
    const normalizedName = key(name);
    if (!normalizedName || byName.has(normalizedName)) continue;
    byName.set(normalizedName, {
      name,
      values: roundResultSelectionVector(item.values),
    });
  }

  const ordered = [];
  const seen = new Set();
  for (const name of canonicalRatioBasisNames(configuredNames)) {
    const normalizedName = key(name);
    if (!normalizedName || seen.has(normalizedName)) continue;
    seen.add(normalizedName);
    const stored = byName.get(normalizedName);
    ordered.push({
      name,
      values: stored ? stored.values : [],
    });
  }
  return ordered;
}

export function ratioBasisValuesForName(valueSets, name) {
  const normalizedName = key(name);
  if (!normalizedName) return [];
  const match = (Array.isArray(valueSets) ? valueSets : []).find(
    (item) => key(item?.name) === normalizedName,
  );
  return roundResultSelectionVector(match?.values);
}

export function upsertRatioBasisValueSet(valueSets, name, values) {
  const basisName = text(name);
  const normalizedName = key(basisName);
  if (!normalizedName) return Array.isArray(valueSets) ? valueSets.slice() : [];
  const next = [];
  let replaced = false;
  for (const item of Array.isArray(valueSets) ? valueSets : []) {
    if (key(item?.name) === normalizedName) {
      if (!replaced) {
        next.push({ name: basisName, values: roundResultSelectionVector(values) });
        replaced = true;
      }
      continue;
    }
    const itemName = text(item?.name);
    if (itemName) next.push({ name: itemName, values: roundResultSelectionVector(item?.values) });
  }
  if (!replaced) next.push({ name: basisName, values: roundResultSelectionVector(values) });
  return next;
}

export function buildResultSelectionMethodPayload({
  details,
  originLabels,
  showWeights,
  sources,
  ratioBasisValueSets,
  calculatedUltimate,
  selectedUltimate,
  ultimateOverrides,
  lastModified,
}) {
  const safeDetails = details && typeof details === "object" ? details : {};
  const ratioBases = canonicalRatioBasisNames(safeDetails.ratioBases);
  const activeRatioBasis = ratioBases.find((name) => key(name) === key(safeDetails.ratioBasis)) || "";
  const normalizedBasisValues = normalizeRatioBasisValueSets(ratioBasisValueSets, ratioBases);
  const rowCount = Array.isArray(originLabels) ? originLabels.length : 0;
  for (const item of normalizedBasisValues) {
    if (item.values.length !== rowCount) {
      throw new Error(`Ratio Basis '${item.name}' must contain exactly ${rowCount} origin values before save.`);
    }
  }
  return {
    json_format: RESULT_SELECTION_JSON_FORMAT,
    details_tab: {
      name: text(safeDetails.name),
      output_type: text(safeDetails.outputType),
      origin_length: Number(safeDetails.originLength) || 12,
      ratio_basis_datasets: ratioBases,
      active_ratio_basis_dataset: activeRatioBasis,
      show_ratios_as_percentages: safeDetails.showRatiosAsPercentages !== false,
      statistic_decimal_places: Number.isFinite(Number(safeDetails.statisticDecimalPlaces))
        ? Number(safeDetails.statisticDecimalPlaces)
        : 1,
    },
    method_tab: {
      origin_labels: Array.isArray(originLabels) ? originLabels.map((label) => String(label ?? "")) : [],
      show_weights: showWeights !== false,
      loaded_datasets: (Array.isArray(sources) ? sources : []).map((source) => ({
        name: text(source?.name),
        dataset_type: text(source?.datasetType ?? source?.dataset_type),
        data_format: text(source?.dataFormat ?? source?.data_format),
        method_type: text(source?.methodType ?? source?.method_type),
        category: text(source?.category),
        source_kind: text(source?.sourceKind ?? source?.source_kind),
        origin_length: Number(source?.originLength ?? source?.origin_length) > 0
          ? Number(source?.originLength ?? source?.origin_length)
          : null,
        values: fitResultSelectionVector(source?.values, rowCount),
        weights: fitResultSelectionVector(source?.weights, rowCount, 0)
          .map((value) => Math.max(0, value ?? 0)),
      })),
      ratio_basis_values: normalizedBasisValues,
      calculated_ultimate: roundResultSelectionVector(calculatedUltimate),
      selected_ultimate: roundResultSelectionVector(selectedUltimate),
      ultimate_overrides: roundResultSelectionVector(ultimateOverrides),
    },
    method_metadata: {
      last_modified: text(lastModified) || new Date().toISOString(),
    },
  };
}
