// A later value of zero is not a ratio of zero: the origin has nothing to
// develop from, and carrying a 0 into the column would drag every average down
// with it. It reads as "no ratio" instead, exactly as a zero earlier value
// already does, so the cell shows the muted placeholder and no average uses it.
export function calcRatio(a, b) {
  const na = Number(a);
  const nb = Number(b);
  if (!Number.isFinite(na) || !Number.isFinite(nb) || na === 0 || nb === 0) return null;
  const v = nb / na;
  return Number.isFinite(v) ? v : null;
}

// A method saved before that rule holds a stored 0 where it now holds no ratio.
export function persistedRatioOrNull(value) {
  const numeric = ratioNumberOrNull(value);
  return numeric === 0 ? null : numeric;
}

// A cell with no ratio holds null, and Number(null) is 0, so a bare
// Number()/Number.isFinite() pair silently turns "no ratio" into a ratio of zero.
// Every read of a stored or calculated ratio goes through here instead.
export function ratioNumberOrNull(value) {
  if (value === null || value === undefined || value === "") return null;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : null;
}

export function roundRatio(value, decimals = 6) {
  if (!Number.isFinite(value)) return null;
  const f = 10 ** decimals;
  return Math.round(value * f) / f;
}

// `toFixed` and `Math.round` work on the binary double behind the number, so a
// value whose decimal form sits exactly on a half - 2.38625 to four places -
// rounds down, while the engine, the persisted value and ResQ all round it up.
// Shifting the decimal text and rounding there gives the digit a reader expects.
// This is also the ROUND function of a User Entry formula, so the browser and
// `arcrho_api.dfm_contract.round_half_up` agree on every operand.
export function roundHalfUp(value, decimals = 0) {
  if (!Number.isFinite(value)) return null;
  const shifted = Number(`${value}e${decimals}`);
  if (!Number.isFinite(shifted)) return null;
  const restored = Number(`${Math.sign(shifted) * Math.round(Math.abs(shifted))}e-${decimals}`);
  return Number.isFinite(restored) ? restored : null;
}

// The value an average row contributes to a User Entry formula. A row enters
// the formula at the precision the Ratios tab prints it at, the method's own
// Decimal Places, rather than at the six decimals it is stored with, so a
// reviewer can multiply the digits shown on screen and land on the User Entry
// factor exactly. `arcrho_api.dfm_contract.average_row_reference_value` is the
// same rule for the server, and the two must agree on every operand.
export function averageRowReferenceValue(value, decimals) {
  if (!Number.isFinite(value)) return null;
  return roundHalfUp(roundRatio(value), decimals);
}

export function formatRatio(value, decimals = 4) {
  if (!Number.isFinite(value)) return "";
  const restored = roundHalfUp(value, decimals);
  return restored === null ? value.toFixed(decimals) : restored.toFixed(decimals);
}

export function computeVolumeAllForColumn(model, col, excludedSet) {
  return computeAverageForColumn(model, col, excludedSet, { base: "volume", periods: "all" });
}

export function computeVolumeRecentForColumn(model, col, excludedSet, lookback = 8) {
  return computeAverageForColumn(model, col, excludedSet, { base: "volume", periods: lookback });
}

export function computeSimpleRecentForColumn(model, col, excludedSet, lookback = 8) {
  return computeAverageForColumn(model, col, excludedSet, { base: "simple", periods: lookback });
}

// `excludedSet` holds every ratio left out of the average: the ones the user
// struck out plus any high/low pair an "Ex hi/lo" row drops. `windowExcludedSet`
// holds only the user's strikes. A "last N" window is counted over ratios the
// user has not struck, so a high/low pair dropped inside that window still
// occupies its place: "Simple - 5 Ex hi/lo" averages three of the last five
// usable ratios, never five drawn from a wider span. Callers that pass no
// window set get the two treated as one, which is right when no pair is dropped.
export function computeAverageForColumn(model, col, excludedSet, options = {}, windowExcludedSet = excludedSet) {
  const baseRaw = String(options.base || "volume").toLowerCase();
  const base = baseRaw === "volume" ? "volume" : "simple";
  const periodsRaw = options.periods ?? "all";
  const periods = typeof periodsRaw === "string" && periodsRaw.toLowerCase() === "all"
    ? "all"
    : Number(periodsRaw);
  const lookback = Number.isFinite(periods) && periods > 0 ? Math.floor(periods) : null;

  const out = {
    sumA: 0,
    sumB: 0,
    sum: 0,
    totalValid: 0,
    totalIncluded: 0,
    value: null,
  };

  // ResQ benchmark rows do not expose a portable local formula. Their
  // canonical values are therefore owned, frozen values that must be rendered
  // and selected exactly as persisted instead of being treated as a simple
  // average.
  if (baseRaw === "benchmark") {
    const value = Number(Array.isArray(options.values) ? options.values[col] : null);
    if (Number.isFinite(value) && value > 0) {
      out.sum = value;
      out.totalValid = 1;
      out.totalIncluded = 1;
      out.value = value;
    }
    return out;
  }

  if (!model || !Array.isArray(model.values) || !Array.isArray(model.mask)) return out;
  const vals = model.values;
  const mask = model.mask;
  const rowCount = Array.isArray(model.origin_labels) ? model.origin_labels.length : vals.length;

  const includeRow = (r) => {
    const hasA = !!(mask[r] && mask[r][col]);
    const hasB = !!(mask[r] && mask[r][col + 1]);
    if (!hasA || !hasB) return null;
    const ratio = calcRatio(vals?.[r]?.[col], vals?.[r]?.[col + 1]);
    if (!Number.isFinite(ratio)) return null;
    return ratio;
  };

  if (lookback) {
    let picked = 0;
    for (let r = rowCount - 1; r >= 0; r--) {
      if (picked >= lookback) break;
      const ratio = includeRow(r);
      if (!Number.isFinite(ratio)) continue;
      out.totalValid += 1;
      const key = `${r},${col}`;
      if (windowExcludedSet && windowExcludedSet.has(key)) continue;
      picked += 1;
      if (excludedSet && excludedSet.has(key)) continue;
      out.totalIncluded += 1;
      if (base === "volume") {
        const a = Number(vals?.[r]?.[col]);
        const b = Number(vals?.[r]?.[col + 1]);
        if (!Number.isFinite(a) || !Number.isFinite(b) || a === 0) continue;
        out.sumA += a;
        out.sumB += b;
      } else {
        out.sum += ratio;
      }
    }
  } else {
    for (let r = 0; r < rowCount; r++) {
      const ratio = includeRow(r);
      if (!Number.isFinite(ratio)) continue;
      out.totalValid += 1;
      if (excludedSet && excludedSet.has(`${r},${col}`)) continue;
      out.totalIncluded += 1;
      if (base === "volume") {
        const a = Number(vals?.[r]?.[col]);
        const b = Number(vals?.[r]?.[col + 1]);
        if (!Number.isFinite(a) || !Number.isFinite(b) || a === 0) continue;
        out.sumA += a;
        out.sumB += b;
      } else {
        out.sum += ratio;
      }
    }
  }

  if (base === "volume") {
    if (out.sumA) out.value = out.sumB / out.sumA;
  } else if (out.totalIncluded > 0) {
    out.value = out.sum / out.totalIncluded;
  }

  return out;
}
