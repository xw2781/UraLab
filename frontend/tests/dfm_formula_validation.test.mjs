import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const helperUrl = new URL("../ui/method_pages/dfm/dfm_formula_validation.js", import.meta.url);
const helperSource = await readFile(helperUrl, "utf8");
// A template literal normalises its own line endings to LF, so the checked-in
// CRLF source has to be normalised too or a multi-line stub never matches.
const summarySource = (await readFile(
  new URL("../ui/method_pages/dfm/ratios_summary/summary_formula_bar.js", import.meta.url),
  "utf8",
)).replaceAll("\r\n", "\n");
const ratiosTabSource = await readFile(
  new URL("../ui/method_pages/dfm/dfm_ratios_tab.js", import.meta.url),
  "utf8",
);
const formulaTextUrl = new URL(
  "../ui/shared/components/formula_bar/formula_text.js",
  import.meta.url,
).href;
const validation = await import(`data:text/javascript;base64,${Buffer.from(helperSource).toString("base64")}`);
const summaryFormatterSource = summarySource
  .replace(
    '"/ui/shared/components/formula_bar/formula_text.js?v=20260908a"',
    JSON.stringify(formulaTextUrl),
  )
  .replace(
    'import { attachArcrhoTooltip } from "/ui/shared/components/tooltip/tooltip.js?v=20260812a";',
    "const attachArcrhoTooltip = (target, text) => { target.tooltipText = text; };",
  )
  .replace(
    'import { installDfmDatasetAutocomplete } from "/ui/method_pages/dfm/dfm_dataset_autocomplete.js?v=20260814b";',
    "const installDfmDatasetAutocomplete = () => {};",
  )
  .replace(
    `import {
  getCachedDfmDatasetReferenceValues,
  resolveDfmDatasetReferencesInFormulaDetailed,
} from "/ui/method_pages/dfm/dfm_dataset_formula.js?v=20260820a";`,
    'const getCachedDfmDatasetReferenceValues = (formula) => (globalThis.__dfmCachedReferenceValues || (() => []))(formula); const resolveDfmDatasetReferencesInFormulaDetailed = async (formula) => { globalThis.__dfmTooltipResolutionCalls = (globalThis.__dfmTooltipResolutionCalls || 0) + 1; globalThis.__dfmTooltipReferenceFormula = formula; if (globalThis.__dfmResolveReferenceValues) globalThis.__dfmCachedReferenceValues = globalThis.__dfmResolveReferenceValues; return { resolvedFormula: "=1.00264" }; };',
  )
  .replace(
    `import {
  registerSummaryFunctions,
  summaryRuntime,
} from "/ui/method_pages/dfm/ratios_summary/summary_runtime.js?v=20260819a";`,
    'const summaryRuntime = { formatUserEntryFormulaEvaluationValue: (value) => Number(value).toFixed(4) }; const registerSummaryFunctions = (functions) => Object.assign(summaryRuntime, functions);',
  )
  .concat("\nexport { tokenizeFormula, formatFormulaText, openDfmFormulaDataset, renderFormulaBarDisplay, updateFormulaBarDisplayMode };\n");
const summaryFormatter = await import(
  `data:text/javascript;base64,${Buffer.from(summaryFormatterSource).toString("base64")}`,
);

class FakeClassList {
  constructor() {
    this.values = new Set();
  }

  add(value) {
    this.values.add(value);
  }

  remove(value) {
    this.values.delete(value);
  }

  contains(value) {
    return this.values.has(value);
  }

  toggle(value, force) {
    const next = force === undefined ? !this.values.has(value) : !!force;
    if (next) this.values.add(value);
    else this.values.delete(value);
    return next;
  }
}

class FakeElement {
  constructor(tagName) {
    this.tagName = String(tagName || "").toUpperCase();
    this.children = [];
    this.dataset = {};
    this.attributes = new Map();
    this.listeners = new Map();
    this.className = "";
    this.classList = new FakeClassList();
    this.textContent = "";
  }

  set innerHTML(_value) {
    this.children = [];
  }

  appendChild(child) {
    this.children.push(child);
    return child;
  }

  setAttribute(name, value) {
    this.attributes.set(name, String(value));
  }

  addEventListener(type, listener) {
    this.listeners.set(type, listener);
  }
}

class FakeInput {
  constructor(value = "= bad") {
    this.value = value;
    this.dataset = {};
    this.style = { display: "none" };
    this.readOnly = false;
    this.isConnected = true;
    this.attributes = new Map();
    this.focusCount = 0;
    this.selection = null;
  }

  setAttribute(name, value) {
    this.attributes.set(name, String(value));
  }

  getAttribute(name) {
    return this.attributes.get(name) ?? null;
  }

  removeAttribute(name) {
    this.attributes.delete(name);
  }

  focus() {
    this.focusCount += 1;
    this.displayWhenFocused = this.style.display;
  }

  setSelectionRange(start, end) {
    this.selection = [start, end];
  }
}

test("formula errors use an in-page error state and preserve the draft", () => {
  const bar = { classList: new FakeClassList() };
  const input = new FakeInput("= 0");
  const error = { id: "formula-error", hidden: true, textContent: "" };

  validation.showFormulaValidationError({
    barEl: bar,
    inputEl: input,
    errorEl: error,
    message: "Enter a number greater than 0.",
  });

  assert.equal(input.value, "= 0");
  assert.equal(input.getAttribute("aria-invalid"), "true");
  assert.equal(input.getAttribute("aria-describedby"), "formula-error");
  assert.equal(error.hidden, false);
  assert.equal(error.textContent, "Enter a number greater than 0.");
  assert.equal(bar.classList.contains("hasValidationError"), true);

  validation.clearFormulaValidationError({ barEl: bar, inputEl: input, errorEl: error });
  assert.equal(input.getAttribute("aria-invalid"), null);
  assert.equal(error.hidden, true);
  assert.equal(bar.classList.contains("hasValidationError"), false);
});

test("formula errors preserve unrelated accessible descriptions", () => {
  const input = new FakeInput();
  const error = { id: "formula-error", hidden: true, textContent: "" };
  input.setAttribute("aria-describedby", "formula-help");

  validation.showFormulaValidationError({ inputEl: input, errorEl: error, message: "Invalid." });
  assert.equal(input.getAttribute("aria-describedby"), "formula-help formula-error");

  validation.clearFormulaValidationError({ inputEl: input, errorEl: error });
  assert.equal(input.getAttribute("aria-describedby"), "formula-help");
});

test("formula error tooltip is positioned above the bar inside the visible Ratios width", () => {
  const layout = validation.computeFormulaValidationTooltipLayout({
    barRect: { left: 100, top: 300, right: 700, bottom: 326, width: 600, height: 26 },
    anchorRect: { left: 150, width: 300 },
    hostRect: { left: 80, top: 100, right: 800, bottom: 600 },
    tooltipRect: { width: 240, height: 32 },
    viewportWidth: 1000,
    viewportHeight: 700,
  });

  assert.deepEqual(layout, {
    left: 154,
    top: 262,
    maxWidth: 520,
    arrowX: 24,
    placement: "above",
    visible: true,
  });
});

test("formula error tooltip clamps horizontally and flips only at the viewport top edge", () => {
  const layout = validation.computeFormulaValidationTooltipLayout({
    barRect: { left: 0, top: 10, right: 320, bottom: 36, width: 320, height: 26 },
    anchorRect: { left: 280, width: 100 },
    hostRect: { left: 0, top: 0, right: 320, bottom: 200 },
    tooltipRect: { width: 200, height: 40 },
    viewportWidth: 320,
    viewportHeight: 200,
  });

  assert.equal(layout.left, 112);
  assert.equal(layout.top, 42);
  assert.equal(layout.maxWidth, 304);
  assert.equal(layout.arrowX, 190);
  assert.equal(layout.placement, "below");
  assert.equal(layout.visible, true);
});

test("formula error tooltip stays inside the host after horizontal table scrolling", () => {
  const layout = validation.computeFormulaValidationTooltipLayout({
    barRect: { left: -600, top: 300, right: 900, bottom: 326, width: 1500, height: 26 },
    anchorRect: { left: 100, width: 200 },
    hostRect: { left: 0, top: 100, right: 400, bottom: 600 },
    tooltipRect: { width: 300, height: 32 },
    viewportWidth: 800,
    viewportHeight: 700,
  });

  assert.equal(layout.visible, true);
  assert.equal(layout.placement, "above");
  assert.ok(layout.left >= 8);
  assert.ok(layout.left + 300 <= 392);
  assert.ok(layout.arrowX >= 10 && layout.arrowX <= 290);
});

test("formula error tooltip hides for collapsed or offscreen geometry", () => {
  const base = {
    barRect: { left: 100, top: 300, right: 700, bottom: 326, width: 600, height: 26 },
    anchorRect: { left: 150, width: 300 },
    hostRect: { left: 80, top: 100, right: 800, bottom: 600 },
    tooltipRect: { width: 240, height: 32 },
    viewportWidth: 1000,
    viewportHeight: 700,
  };
  const cases = [
    { hostRect: { ...base.hostRect, right: base.hostRect.left } },
    { hostRect: { ...base.hostRect, bottom: base.hostRect.top } },
    { barRect: { ...base.barRect, right: base.barRect.left } },
    { barRect: { ...base.barRect, bottom: base.barRect.top } },
    { tooltipRect: { width: 0, height: 32 } },
    { tooltipRect: { width: 240, height: 0 } },
    { hostRect: { left: 1200, top: 100, right: 1500, bottom: 600 } },
  ];

  cases.forEach((override) => {
    const layout = validation.computeFormulaValidationTooltipLayout({ ...base, ...override });
    assert.equal(layout.visible, false);
  });
});

test("formula recovery reveals the real input before focusing and restores selection", () => {
  const input = new FakeInput("= invalid");
  const display = { style: { display: "" } };

  const restored = validation.revealAndFocusFormulaInput({
    inputEl: input,
    displayEl: display,
    selectionStart: 3,
    selectionEnd: 7,
  });

  assert.equal(restored, true);
  assert.equal(input.displayWhenFocused, "");
  assert.equal(display.style.display, "none");
  assert.deepEqual(input.selection, [3, 7]);
});

test("validation leases always release pending and read-only state", () => {
  const input = new FakeInput();
  const lease = validation.beginFormulaValidationLease(input, { timeoutMs: 1000 });

  assert.equal(input.dataset.formulaCommitPending, "1");
  assert.equal(input.readOnly, true);
  assert.equal(input.getAttribute("aria-busy"), "true");

  lease.finish();
  assert.equal(input.dataset.formulaCommitPending, undefined);
  assert.equal(input.readOnly, false);
  assert.equal(input.getAttribute("aria-busy"), null);
});

test("validation timeout aborts the request and remains releasable", async () => {
  const input = new FakeInput();
  const lease = validation.beginFormulaValidationLease(input, { timeoutMs: 5 });
  const abortAwareRequest = new Promise((resolve) => {
    lease.signal.addEventListener("abort", resolve, { once: true });
  });

  await abortAwareRequest;
  assert.equal(lease.timedOut, true);
  assert.equal(lease.signal.aborted, true);

  lease.finish();
  assert.equal(input.dataset.formulaCommitPending, undefined);
  assert.equal(input.readOnly, false);
});

test("cancel aborts validation and restores a pre-existing read-only state", () => {
  const input = new FakeInput();
  input.readOnly = true;
  const lease = validation.beginFormulaValidationLease(input, { timeoutMs: 1000 });

  lease.cancel();
  assert.equal(lease.signal.aborted, true);
  assert.equal(input.dataset.formulaCommitPending, undefined);
  assert.equal(input.readOnly, true);
  assert.equal(input.getAttribute("aria-busy"), null);
});

test("a newer validation lease cannot be cleared by stale cleanup", () => {
  const input = new FakeInput();
  const first = validation.beginFormulaValidationLease(input, { timeoutMs: 1000 });
  const second = validation.beginFormulaValidationLease(input, { timeoutMs: 1000 });

  assert.equal(first.signal.aborted, true);
  first.finish();
  assert.equal(input.dataset.formulaCommitPending, "1");
  assert.equal(input.readOnly, true);

  second.finish();
  assert.equal(input.dataset.formulaCommitPending, undefined);
  assert.equal(input.readOnly, false);
});

test("DFM ratio validation does not use native alerts", () => {
  assert.doesNotMatch(summarySource, /\balert\s*\(/);
  assert.doesNotMatch(ratiosTabSource, /\balert\s*\(/);
  assert.match(summarySource, /setAttribute\("role", "alert"\)/);
});

test("formula display formatting preserves all bracket contents verbatim", () => {
  assert.equal(
    summaryFormatter.formatFormulaText(
      '= "Simple - 2" * [Accounting Cutoff][-1] * [C 01 - Growth Adjustment][row / 2]',
    ),
    '= "Simple - 2" * [Accounting Cutoff][-1] * [C 01 - Growth Adjustment][row / 2]',
  );

  assert.deepEqual(
    summaryFormatter.tokenizeFormula("=[Dataset + Name][-1]")
      .filter((token) => token.type === "bracket")
      .map((token) => token.text),
    ["[Dataset + Name]", "[-1]"],
  );
});

test("non-editing formula display renders average formulas and datasets as reference pills", async () => {
  const posted = [];
  const priorDocument = globalThis.document;
  const priorWindow = globalThis.window;
  // Every reference already carries a value, so rendering reads the session
  // cache instead of asking the API for one.
  globalThis.__dfmCachedReferenceValues = (formula) => summaryFormatter
    .tokenizeFormula(formula)
    .filter((token) => token.datasetName)
    .map(() => 1.0118);
  globalThis.document = {
    createElement: (tagName) => new FakeElement(tagName),
    createTextNode: (textContent) => ({ nodeType: 3, textContent }),
  };
  globalThis.window = {
    parent: {
      postMessage: (message, targetOrigin) => posted.push({ message, targetOrigin }),
    },
  };

  try {
    const display = new FakeElement("div");
    summaryFormatter.renderFormulaBarDisplay(
      display,
      '= "Simple - 2" * [Accounting Cutoff][2025 Q4] * [C 01 - Growth Adjustment][2024, 12 months]',
    );
    const averageFormulaPills = display.children.filter((child) => child.className === "fmtRowRef");
    assert.deepEqual(averageFormulaPills.map((pill) => pill.textContent), ["Simple - 2"]);
    const pills = display.children.filter((child) => child.className === "fmtDatasetRef");
    assert.deepEqual(pills.map((pill) => pill.textContent), [
      "Accounting Cutoff @ 2025 Q4",
      "C 01 - Growth Adjustment @ 2024, 12 months",
    ]);
    assert.deepEqual(pills.map((pill) => pill.dataset.datasetName), [
      "Accounting Cutoff",
      "C 01 - Growth Adjustment",
    ]);
    assert.deepEqual(pills.map((pill) => pill.dataset.coordinateLabel), [
      "2025 Q4",
      "2024, 12 months",
    ]);
    assert.equal(
      pills[0].attributes.get("aria-label"),
      "Open dataset Accounting Cutoff at 2025 Q4 in Dataset Viewer",
    );
    const renderedText = display.children.map((child) => child.textContent).join("");
    assert.match(renderedText, /^= Simple - 2/u);
    assert.match(renderedText, /Accounting Cutoff @ 2025 Q4/u);
    assert.doesNotMatch(renderedText, /\[2025 Q4\]/u);
    assert.doesNotMatch(renderedText, /\[-1\]/u);

    const modeDisplay = new FakeElement("div");
    modeDisplay.style = {};
    const modeInput = {
      value: "=[Accounting Cutoff][-1]",
      dataset: { displayFormula: "=[Accounting Cutoff][2025 Q4]" },
      style: {},
    };
    const modeBar = {
      querySelector(selector) {
        return selector === "#dfmSummaryFormulaBarInput" ? modeInput : modeDisplay;
      },
    };
    summaryFormatter.updateFormulaBarDisplayMode(modeBar, false);
    assert.match(
      modeDisplay.children.map((child) => child.textContent).join(""),
      /Accounting Cutoff @ 2025 Q4/u,
    );
    assert.doesNotMatch(modeDisplay.children.map((child) => child.textContent).join(""), /\[2025 Q4\]/u);
    assert.doesNotMatch(modeDisplay.children.map((child) => child.textContent).join(""), /\[-1\]/u);
    const modePill = modeDisplay.children.find((child) => child.className === "fmtDatasetRef");
    assert.equal(modePill.classList.contains("isNeutral"), false);
    globalThis.__dfmTooltipResolutionCalls = 0;
    assert.equal(typeof modePill.tooltipText, "function");
    assert.equal(await modePill.tooltipText(), "1.0026");
    assert.equal(await modePill.tooltipText(), "1.0026");
    assert.equal(globalThis.__dfmTooltipResolutionCalls, 1);
    assert.equal(globalThis.__dfmTooltipReferenceFormula, "=[Accounting Cutoff][-1]");

    const excelDisplay = new FakeElement("div");
    summaryFormatter.renderFormulaBarDisplay(
      excelDisplay,
      "='C:\\Data\\[Book.xlsx]Sheet 1'!A1",
    );
    assert.equal(excelDisplay.children.some((child) => child.className === "fmtDatasetRef"), false);

    let propagationStopped = false;
    pills[0].listeners.get("click")({
      preventDefault() {},
      stopPropagation() { propagationStopped = true; },
    });
    assert.equal(propagationStopped, true);
    assert.deepEqual(posted, [{
      message: {
        type: "arcrho:project-instance-open-dependent-dataset",
        datasetName: "Accounting Cutoff",
        openMethod: false,
      },
      targetOrigin: "*",
    }]);
  } finally {
    delete globalThis.__dfmCachedReferenceValues;
    if (priorDocument === undefined) delete globalThis.document;
    else globalThis.document = priorDocument;
    if (priorWindow === undefined) delete globalThis.window;
    else globalThis.window = priorWindow;
    delete globalThis.__dfmTooltipResolutionCalls;
    delete globalThis.__dfmTooltipReferenceFormula;
  }
});

test("a dataset reference worth exactly 1 renders as a quiet pill", async () => {
  const priorDocument = globalThis.document;
  const priorWindow = globalThis.window;
  globalThis.document = {
    createElement: (tagName) => new FakeElement(tagName),
    createTextNode: (textContent) => ({ nodeType: 3, textContent }),
  };
  globalThis.window = { parent: { postMessage: () => {} } };
  const formula = '= "Simple - 3" * [Accounting Cutoff][-1] * [Growth Adjustment - Counts][-1]';
  const neutralFlags = (display) => display.children
    .filter((child) => child.className === "fmtDatasetRef")
    .map((pill) => pill.classList.contains("isNeutral"));

  try {
    globalThis.__dfmTooltipResolutionCalls = 0;
    globalThis.__dfmCachedReferenceValues = () => [1, 1.0012];
    const display = new FakeElement("div");
    summaryFormatter.renderFormulaBarDisplay(display, formula);
    assert.deepEqual(neutralFlags(display), [true, false]);
    assert.equal(globalThis.__dfmTooltipResolutionCalls, 0);

    // Binary rounding around 1 still reads as neutral; a real adjustment, and a
    // reference that failed to read, do not.
    globalThis.__dfmCachedReferenceValues = () => [1 + 1e-12, null];
    const rounded = new FakeElement("div");
    summaryFormatter.renderFormulaBarDisplay(rounded, formula);
    assert.deepEqual(neutralFlags(rounded), [true, false]);

    // A cache reading of the formula that found a different number of
    // references cannot be lined up with the pills, so none of them is greyed.
    globalThis.__dfmCachedReferenceValues = () => [1];
    const mismatched = new FakeElement("div");
    summaryFormatter.renderFormulaBarDisplay(mismatched, formula);
    assert.deepEqual(neutralFlags(mismatched), [false, false]);

    // Nothing resolved yet: the pills start blue and settle after one batched
    // read fills the cache for the whole formula.
    globalThis.__dfmTooltipResolutionCalls = 0;
    globalThis.__dfmCachedReferenceValues = () => [];
    globalThis.__dfmResolveReferenceValues = () => [1, 1.0012];
    const cold = new FakeElement("div");
    summaryFormatter.renderFormulaBarDisplay(cold, formula);
    assert.deepEqual(neutralFlags(cold), [false, false]);
    assert.equal(globalThis.__dfmTooltipResolutionCalls, 1);
    assert.equal(globalThis.__dfmTooltipReferenceFormula, formula);
    await new Promise((resolve) => setImmediate(resolve));
    assert.deepEqual(neutralFlags(cold), [true, false]);
  } finally {
    delete globalThis.__dfmCachedReferenceValues;
    delete globalThis.__dfmResolveReferenceValues;
    delete globalThis.__dfmTooltipResolutionCalls;
    delete globalThis.__dfmTooltipReferenceFormula;
    if (priorDocument === undefined) delete globalThis.document;
    else globalThis.document = priorDocument;
    if (priorWindow === undefined) delete globalThis.window;
    else globalThis.window = priorWindow;
  }
});
