import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

// ROUND(x, digits) in a User Entry formula rounds half up on the decimal text,
// the way arcrho_api.dfm_contract.round_half_up does on the server, so the
// browser's evaluation of a formula matches the persisted value it re-creates.

const calcUrl = new URL("../ui/method_pages/dfm/dfm_ratio_calc.js", import.meta.url).href;
const { roundHalfUp, formatRatio, averageRowReferenceValue } = await import(calcUrl);

const modelSource = await readFile(
  new URL("../ui/method_pages/dfm/ratios_summary/summary_model.js", import.meta.url),
  "utf8",
);
// The patched module is loaded from a data: URL, so the helper reaches it
// through a global rather than an import.
globalThis.__roundHalfUpUnderTest = roundHalfUp;
const runtimeStub = `const summaryRuntime = {
  roundHalfUp: globalThis.__roundHalfUpUnderTest,
  state: { model: {} },
  summaryRowConfigs: [],
  summaryRowMap: new Map(),
  summaryFormulaEditState: null,
};
const registerSummaryFunctions = (functions) => Object.assign(summaryRuntime, functions);
export { summaryRuntime };`;
const patchedModel = modelSource.replace(
  /import \{[\s\S]*?\} from "\/ui\/method_pages\/dfm\/ratios_summary\/summary_runtime\.js\?v=[^"]*";/u,
  runtimeStub,
);
globalThis.CSS = { escape: (value) => String(value) };
globalThis.document = { body: { contains: () => false } };
globalThis.window = { requestAnimationFrame: () => 0 };
const model = await import(`data:text/javascript;base64,${Buffer.from(patchedModel).toString("base64")}`);
const evaluate = model.summaryRuntime.evaluateSimpleMathExpression;

test("roundHalfUp rounds the decimal text half away from zero", () => {
  assert.equal(roundHalfUp(2.38625, 4), 2.3863);
  assert.equal(roundHalfUp(-2.38625, 4), -2.3863);
  assert.equal(roundHalfUp(1.35735, 4), 1.3574);
  assert.equal(roundHalfUp(1.5), 2);
  assert.equal(roundHalfUp(Number.NaN, 4), null);
  assert.equal(formatRatio(2.38625, 4), "2.3863");
});

test("ROUND in a User Entry formula fixes the operand before it multiplies", () => {
  const rows = new Map([["Simple - 2", 1.35735]]);
  const rounded = evaluate('= ROUND("Simple - 2", 4) * 0.9949', rows);
  assert.ok(Math.abs(rounded - 1.3574 * 0.9949) < 1e-12, `got ${rounded}`);
  const unrounded = evaluate('= "Simple - 2" * 0.9949', rows);
  assert.ok(Math.abs(unrounded - 1.35735 * 0.9949) < 1e-12, `got ${unrounded}`);
  assert.equal(evaluate("= round(2.38625, 4)"), 2.3863);
  assert.equal(evaluate("= ROUND(1.5)"), 2);
  assert.equal(evaluate("= ROUND(1.5) + ROUND(2.5, 0)"), 5);
});

test("an average row enters a formula at the method's decimal places", () => {
  // 400/300 stored at six decimals, read at the four the Ratios tab prints.
  assert.equal(averageRowReferenceValue(1.3333333333, 4), 1.3333);
  assert.equal(averageRowReferenceValue(1.3333333333, 6), 1.333333);
  assert.equal(averageRowReferenceValue(2.386254, 4), 2.3863);
  assert.equal(averageRowReferenceValue(Number.NaN, 4), null);

  const rows = new Map([["Volume - all", averageRowReferenceValue(1.3333333333, 4)]]);
  assert.ok(Math.abs(evaluate('= "Volume - all" * 1.1', rows) - 1.3333 * 1.1) < 1e-12);
});

test("no other word reaches the evaluator", () => {
  assert.equal(evaluate("= FLOOR(1.5)"), null);
  assert.equal(evaluate("= Math.round(1.5)"), null);
  assert.equal(evaluate("= R(1.5)"), null);
  assert.equal(evaluate("= ROUND(1.5) * alert(1)"), null);
});
