// A length control the open dataset has no use for is locked rather than
// removed. The Data tab locks Development Length on a Vector, which has one
// column of values and no development dimension.
import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const requestControllerSource = await readFile(
  new URL("../ui/shared/tabs/data/data_tab_request_controller.js", import.meta.url),
  "utf8",
);
const persistenceControllerSource = await readFile(
  new URL("../ui/shared/tabs/data/data_tab_persistence_controller.js", import.meta.url),
  "utf8",
);
const preferencesControllerSource = await readFile(
  new URL("../ui/shared/tabs/data/data_tab_preferences_controller.js", import.meta.url),
  "utf8",
);
const gridInteractionsSource = await readFile(
  new URL("../ui/shared/tabs/data/dataset_grid_interactions.js", import.meta.url),
  "utf8",
);
const runControllerSource = await readFile(
  new URL("../ui/shared/dataset/dataset_run_controller.js", import.meta.url),
  "utf8",
);
const datasetViewerViewSource = await readFile(
  new URL("../ui/dataset_viewer/dataset_viewer_view.js", import.meta.url),
  "utf8",
);
const datasetViewerCss = await readFile(
  new URL("../ui/dataset_viewer/dataset_viewer.css", import.meta.url),
  "utf8",
);
const dataControlsSource = await readFile(
  new URL("../ui/shared/tabs/data/data_tab_controls.js", import.meta.url),
  "utf8",
);
const inputsControllerSource = await readFile(
  new URL("../ui/shared/tabs/data/data_tab_inputs_controller.js", import.meta.url),
  "utf8",
);
const hostControllerSource = await readFile(
  new URL("../ui/shared/tabs/data/data_tab_host_controller.js", import.meta.url),
  "utf8",
);
const dfmPageSource = await readFile(
  new URL("../ui/method_pages/dfm/dfm.html", import.meta.url),
  "utf8",
);

// The controller imports its siblings by their server-absolute `/ui/...` paths,
// which Node cannot resolve. Only the tooltip binding is reached on the paths
// under test, so the imports are swapped for no-op stubs and the module is
// loaded from source rather than pulling in the whole page graph.
function importRequestController() {
  const stubbed = requestControllerSource.replace(
    /^import\s*\{([\s\S]*?)\}\s*from\s*"\/ui\/[^"]*";$/gmu,
    (_match, names) => `const {${names}} = __moduleStubs;`,
  );
  const source = `const __moduleStubs = new Proxy({}, { get: () => () => {} });\n${stubbed}`;
  return import(`data:text/javascript;base64,${Buffer.from(source).toString("base64")}`);
}

class FakeElement {
  constructor(tag = "div") {
    this.tagName = String(tag).toUpperCase();
    this.attributes = new Map();
    this.children = [];
    this.dataset = {};
    this.handlers = new Map();
    this.className = "";
    this.textContent = "";
    this.tabIndex = 0;
    this.classList = {
      contains: (name) => this.className.split(/\s+/u).includes(name),
      toggle: (name, force) => {
        const classes = new Set(this.className.split(/\s+/u).filter(Boolean));
        if (force === undefined ? classes.has(name) : !force) classes.delete(name);
        else classes.add(name);
        this.className = [...classes].join(" ");
      },
    };
  }

  set innerHTML(_value) { this.children = []; }

  get innerHTML() { return ""; }

  setAttribute(name, value) { this.attributes.set(name, String(value)); }

  getAttribute(name) { return this.attributes.has(name) ? this.attributes.get(name) : null; }

  hasAttribute(name) { return this.attributes.has(name); }

  removeAttribute(name) { this.attributes.delete(name); }

  appendChild(child) { this.children.push(child); return child; }

  addEventListener(type, handler) { this.handlers.set(type, handler); }

  focus() {}

  dispatchEvent(event) { this.handlers.get(event?.type)?.(event); return true; }

  querySelector(selector) {
    return selector === ".lenSelectValue" ? this.valueLabel : null;
  }

  fire(type, event = {}) {
    this.handlers.get(type)?.({ preventDefault() {}, stopPropagation() {}, ...event });
  }
}

class FakeSelect extends FakeElement {
  constructor(values) {
    super("select");
    this.options = values.map((value) => ({ value: String(value), textContent: String(value) }));
    this.selectedIndex = 0;
  }

  set innerHTML(_value) {
    this.children = [];
    this.options = [];
    this.selectedIndex = -1;
  }

  get innerHTML() { return ""; }

  appendChild(child) {
    this.children.push(child);
    this.options.push(child);
    if (this.selectedIndex < 0) this.selectedIndex = 0;
    return child;
  }

  get value() { return this.options[this.selectedIndex]?.value ?? ""; }

  set value(next) {
    const index = this.options.findIndex((option) => option.value === String(next));
    if (index >= 0) this.selectedIndex = index;
  }
}

function buildLengthControlDom() {
  const elements = {};
  for (const name of ["origin", "dev", "originStored", "devStored"]) {
    const wrap = new FakeElement();
    wrap.className = "lenSelectWrap";
    const button = new FakeElement("button");
    button.valueLabel = new FakeElement("span");
    button.valueLabel.textContent = "12";
    const dropdown = new FakeElement();
    const select = new FakeSelect([12, 6, 3, 1]);
    elements[`${name}LenWrap`] = wrap;
    elements[`${name}LenDisplay`] = button;
    elements[`${name}LenDropdown`] = dropdown;
    elements[`${name}LenSelect`] = select;
  }
  return elements;
}

async function createLengthControlRuntime() {
  const elements = buildLengthControlDom();
  const originalDocument = globalThis.document;
  globalThis.document = {
    body: new FakeElement("body"),
    getElementById: (id) => elements[id] || null,
    createElement: (tag) => new FakeElement(tag),
    querySelectorAll: () => [],
    addEventListener() {},
  };
  const runtime = {
    state: { dirty: new Map(), model: null },
    config: {},
    isTemporaryDatasetView: false,
    qs: new URLSearchParams(""),
    temporaryDatasetSessionId: "",
    LEN_DROPDOWN_CONFIG: {
      originLenSelect: {
        wrapId: "originLenWrap",
        buttonId: "originLenDisplay",
        dropdownId: "originLenDropdown",
      },
      devLenSelect: {
        wrapId: "devLenWrap",
        buttonId: "devLenDisplay",
        dropdownId: "devLenDropdown",
      },
      originStoredLenSelect: {
        wrapId: "originStoredLenWrap",
        buttonId: "originStoredLenDisplay",
        dropdownId: "originStoredLenDropdown",
      },
      devStoredLenSelect: {
        wrapId: "devStoredLenWrap",
        buttonId: "devStoredLenDisplay",
        dropdownId: "devStoredLenDropdown",
      },
    },
    createDatasetDependencyGuard: () => ({}),
    showProjectDropdown() {},
    showDatasetDropdown() {},
  };
  const { registerDataTabRequestController } = await importRequestController();
  registerDataTabRequestController(runtime);
  runtime.wireLenDropdowns();
  return { runtime, elements, restore: () => { globalThis.document = originalDocument; } };
}

test("a locked length control shows its fixed value and cannot be opened", async () => {
  const { runtime, elements, restore } = await createLengthControlRuntime();
  try {
    const wrap = elements.devLenWrap;
    const button = elements.devLenDisplay;
    const select = elements.devLenSelect;

    // Unlocked, the trigger reads the select and opens its list on click.
    assert.equal(button.valueLabel.textContent, "12");
    button.fire("click");
    assert.equal(wrap.classList.contains("open"), true);

    runtime.setLenSelectLock("devLenSelect", {
      locked: true,
      displayValue: "0",
      reason: "A vector has no development periods.",
    });

    // Locking closes the open list and repaints the trigger as a fixed 0.
    assert.equal(wrap.classList.contains("open"), false);
    assert.equal(button.valueLabel.textContent, "0");
    assert.equal(wrap.classList.contains("is-locked"), true);
    assert.equal(button.getAttribute("aria-disabled"), "true");
    assert.equal(button.tabIndex, -1);
    assert.equal(wrap.getAttribute("data-locked-reason"), "A vector has no development periods.");

    // The value underneath is untouched, so nothing that reads the stored
    // development length sees a 0 the user never chose.
    assert.equal(select.value, "12");

    // Click, keyboard, and wheel all refuse while locked.
    button.fire("click");
    assert.equal(wrap.classList.contains("open"), false);
    button.fire("keydown", { key: "ArrowDown" });
    assert.equal(wrap.classList.contains("open"), false);
    button.fire("wheel", { deltaY: 1 });
    assert.equal(select.value, "12");
    assert.equal(button.valueLabel.textContent, "0");

    // A later repaint of the list cannot restore the real value to the trigger.
    runtime.renderLenDropdownOptions("devLenSelect");
    assert.equal(button.valueLabel.textContent, "0");
    runtime.setLenSelectValue("devLenSelect", "6");
    assert.equal(button.valueLabel.textContent, "0");
    assert.equal(select.value, "6");

    // The neighbouring control is unaffected by the lock.
    assert.equal(elements.originLenWrap.classList.contains("is-locked"), false);
    elements.originLenDisplay.fire("click");
    assert.equal(elements.originLenWrap.classList.contains("open"), true);

    // Unlocking hands the trigger back to the select.
    runtime.setLenSelectLock("devLenSelect", { locked: false });
    assert.equal(wrap.classList.contains("is-locked"), false);
    assert.equal(button.getAttribute("aria-disabled"), "false");
    assert.equal(button.tabIndex, 0);
    assert.equal(button.valueLabel.textContent, "6");
    button.fire("click");
    assert.equal(wrap.classList.contains("open"), true);
  } finally {
    restore();
  }
});

test("the Data tab locks Development Length to 0 for a vector dataset", () => {
  // The lock follows the resolved data format of the open dataset.
  assert.match(
    persistenceControllerSource,
    /normalizeDatasetModeText\(getDatasetRunDataFormat\(\)\) === "vector"/u,
  );
  assert.match(persistenceControllerSource, /setLenSelectLock\("devLenSelect", \{/u);
  assert.match(persistenceControllerSource, /displayValue: "0"/u);
  // It is reapplied wherever the length controls are repainted, so switching
  // between a Triangle and a Vector Dataset Type cannot strand the old state.
  assert.match(persistenceControllerSource, /updateVectorDevelopmentLengthControl\(\);/u);
  // The stored development length is never rewritten to 0.
  assert.doesNotMatch(persistenceControllerSource, /setLenSelectValue\("devLenSelect", "0"\)/u);
  // The lock is generic to the top-bar length controls: the shared dropdown
  // helpers decide neither which control is locked nor what it then reads.
  assert.doesNotMatch(requestControllerSource, /isLenSelectLocked\("devLenSelect"\)/u);
  assert.doesNotMatch(requestControllerSource, /displayValue: "0"/u);
  assert.match(datasetViewerCss, /\.lenSelectWrap\.is-locked \.lenSelectDisplay \{/u);
});


// ---------------------------------------------------------------------------
// A display length is only ever a whole multiple of the period the dataset's
// own file is stored at, the window opens at the display shape the sidecar
// saved, and a coarser view is read-only because its cells are a roll-up.

test("a length control mutes the lengths the stored period rules out", async () => {
  const { runtime, elements, restore } = await createLengthControlRuntime();
  try {
    assert.deepEqual(runtime.lenChoicesForStoredLength(1), [12, 6, 3, 1]);
    assert.deepEqual(runtime.lenChoicesForStoredLength(3), [12, 6, 3]);
    assert.deepEqual(runtime.lenChoicesForStoredLength(6), [12, 6]);
    assert.deepEqual(runtime.lenChoicesForStoredLength(12), [12]);
    // Not stated yet: before a sidecar has loaded nothing is known, so the
    // whole ladder stays open rather than being narrowed to a guess.
    assert.deepEqual(runtime.lenChoicesForStoredLength(0), [12, 6, 3, 1]);

    const select = elements.originLenSelect;
    const values = () => select.options.map((option) => option.value);
    const muted = () => select.options.filter((option) => runtime.lenOptionIsUnavailable(option)).map((option) => option.value);

    runtime.setLenSelectValue("originLenSelect", "6");
    runtime.setLenSelectStoredLength("originLenSelect", 3);
    // The ladder itself never shrinks: only what can be chosen changes.
    assert.deepEqual(values(), ["12", "6", "3", "1"]);
    assert.deepEqual(muted(), ["1"]);
    // A length still offered is kept.
    assert.equal(select.value, "6");
    assert.equal(elements.originLenDisplay.valueLabel.textContent, "6");

    // One that has just been muted lands on the stored period itself, which is
    // where the values live.
    runtime.setLenSelectStoredLength("originLenSelect", 12);
    assert.deepEqual(values(), ["12", "6", "3", "1"]);
    assert.deepEqual(muted(), ["6", "3", "1"]);
    assert.equal(select.value, "12");
    assert.equal(elements.originLenDisplay.valueLabel.textContent, "12");

    // A muted length cannot be set, by the list or by anything reading the
    // saved display shape back into the control.
    assert.equal(runtime.setLenSelectValue("originLenSelect", "6"), false);
    assert.equal(select.value, "12");

    // Opening it again clears the muting without disturbing the value.
    runtime.setLenSelectStoredLength("originLenSelect", 1);
    assert.deepEqual(muted(), []);
    assert.equal(select.value, "12");

    // The list the user actually sees carries every length, with the ones the
    // stored period rules out marked rather than dropped, and each of those
    // says why on hover.
    runtime.setLenSelectStoredLength("originLenSelect", 6);
    const rows = elements.originLenDropdown.children;
    assert.deepEqual(rows.map((row) => row.textContent), ["12", "6", "3", "1"]);
    assert.deepEqual(
      rows.map((row) => row.classList.contains("is-unavailable")),
      [false, false, true, true],
    );
    assert.deepEqual(rows.map((row) => row.getAttribute("aria-disabled")), [null, null, "true", "true"]);
    assert.match(runtime.lenUnavailableReason(6), /stored at 6/u);
    assert.equal(runtime.lenUnavailableReason(0), "");

    // A muted row neither takes the highlight nor accepts a click.
    rows[2].fire("mouseenter");
    assert.equal(rows[2].classList.contains("active"), false);
    rows[2].fire("mousedown");
    assert.equal(select.value, "12");
    rows[1].fire("mousedown");
    assert.equal(select.value, "6");

    // The keyboard and the wheel step over the muted rows rather than resting
    // on one, so neither can reach a length the dataset cannot be shown at.
    const activeIndex = () => elements.originLenDropdown.children.findIndex((row) => row.classList.contains("active"));
    for (let step = 0; step < 4; step += 1) {
      elements.originLenDisplay.fire("keydown", { key: "ArrowDown" });
      assert.ok(activeIndex() === 0 || activeIndex() === 1, `the highlight landed on a muted row: ${activeIndex()}`);
    }
    for (let step = 0; step < 4; step += 1) {
      elements.originLenDisplay.fire("keydown", { key: "ArrowUp" });
      assert.ok(activeIndex() === 0 || activeIndex() === 1, `the highlight landed on a muted row: ${activeIndex()}`);
    }
    runtime.setLenSelectValue("originLenSelect", "6");
    elements.originLenDisplay.fire("wheel", { deltaY: 1 });
    assert.equal(select.value, "6");
  } finally {
    restore();
  }
});

test("the offered lengths follow the open dataset's stored period", () => {
  // The stored pair comes from the sidecar, on both the load and the save, and
  // is cleared when there is no sidecar to read it from.
  assert.match(persistenceControllerSource, /function applyStoredLengthsFromResponse\(payload\)/u);
  assert.match(persistenceControllerSource, /runtime\.currentDatasetStoredOriginLength = Number\(source\.stored_origin_length\) \|\| 0;/u);
  assert.match(persistenceControllerSource, /runtime\.currentDatasetStoredDevelopmentLength = Number\(source\.stored_development_length\) \|\| 0;/u);
  assert.match(persistenceControllerSource, /applyStoredLengthsFromResponse\(data\.exists \? data : null\);/u);
  assert.match(persistenceControllerSource, /applyStoredLengthsFromResponse\(resp\.data\);/u);
  // A hand-entered dataset that still holds nothing has no stored period to
  // protect, so the whole ladder stays open until its first real save.
  assert.match(
    persistenceControllerSource,
    /function storedLengthIsPending\(\) \{\s*return currentDatasetIsManualTriangleOrVector\(\) && datasetValuesAreAllZero\(\);/u,
  );
  assert.match(persistenceControllerSource, /applyStoredLengthChoices\(\);/u);
  // Narrowing runs before the saved display shape is written into the control,
  // so the length the window reopens at is never dropped for want of an option.
  assert.match(
    persistenceControllerSource,
    /applyStoredLengthChoices\(\);\s*setLenSelectValue\("originLenSelect", String\(normalized\.origin_length\)\);/u,
  );
});

test("the stored period reads off the list and the Stored at control, not a caption", () => {
  // Nothing in the strip repeats the stored period as a caption: the list and
  // the `Stored at` control beside each length carry it.
  assert.doesNotMatch(datasetViewerViewSource, /lenStoredNote/u);
  assert.doesNotMatch(datasetViewerCss, /lenStoredNote/u);
  assert.match(datasetViewerCss, /#datasetTopBar \.lenDropdown \.lenOption\.is-unavailable/u);
  // Every length stays in the list; the stored period only decides which of
  // them are muted.
  assert.match(persistenceControllerSource, /setLenSelectStoredLength\("originLenSelect", stored\.origin_length\);/u);
  assert.match(persistenceControllerSource, /setLenSelectStoredLength\("devLenSelect", stored\.development_length\);/u);
  assert.match(requestControllerSource, /for \(const value of LEN_CHOICES\) \{/u);
  assert.match(requestControllerSource, /option\.dataset\.unavailable = "1";/u);
  // Hovering a length control still names the shape the file is held at, once
  // that shape is settled.
  assert.match(persistenceControllerSource, /`This dataset is stored at \$\{value\}\.`/u);
  assert.match(requestControllerSource, /wrap\.getAttribute\("data-locked-reason"\) \|\| wrap\.getAttribute\("data-hint"\)/u);
  // While the dataset is still empty the `Stored at` control is the live
  // answer, so the hint that used to say so has been retired.
  assert.doesNotMatch(persistenceControllerSource, /its first save stores it at/u);
  assert.match(persistenceControllerSource, /setLenSelectHint\("originLenSelect", pending \? "" : storedLengthHintText\(recorded\.origin_length\)\);/u);
  // A vector has no development dimension, so it shows no development hint.
  assert.match(persistenceControllerSource, /pending \|\| vector \? "" : storedLengthHintText\(recorded\.development_length\)/u);
});

// ---------------------------------------------------------------------------
// ResQ puts a `Stored at` spinner beside each length in its Edit Triangle
// dialog. ArcRho now does the same: the origin one is a dimmed readout and the
// development one can lower the store while the dataset is still empty.

test("a Stored at control offers the periods that divide the length beside it", async () => {
  const { runtime, elements, restore } = await createLengthControlRuntime();
  try {
    // The store may be finer than the display but never coarser, and it has to
    // divide it evenly, which is the opposite of the display control's rule.
    assert.deepEqual(runtime.lenChoicesForDisplayLength(12), [12, 6, 3, 1]);
    assert.deepEqual(runtime.lenChoicesForDisplayLength(6), [6, 3, 1]);
    assert.deepEqual(runtime.lenChoicesForDisplayLength(3), [3, 1]);
    assert.deepEqual(runtime.lenChoicesForDisplayLength(1), [1]);
    // Nothing known yet leaves the whole ladder open.
    assert.deepEqual(runtime.lenChoicesForDisplayLength(0), [12, 6, 3, 1]);

    const select = elements.devStoredLenSelect;
    const muted = () => select.options.filter((option) => runtime.lenOptionIsUnavailable(option)).map((option) => option.value);

    runtime.setLenSelectDisplayLength("devStoredLenSelect", 12);
    assert.deepEqual(muted(), []);
    assert.equal(runtime.setLenSelectValue("devStoredLenSelect", "1"), true);
    assert.equal(select.value, "1");

    // Showing the dataset at 6 rules out a store of 12, and the ladder itself
    // keeps every rung.
    runtime.setLenSelectDisplayLength("devStoredLenSelect", 6);
    assert.deepEqual(select.options.map((option) => option.value), ["12", "6", "3", "1"]);
    assert.deepEqual(muted(), ["12"]);
    // A store of 1 still divides 6, so it is kept.
    assert.equal(select.value, "1");

    // A store that has just been ruled out lands on the displayed period, which
    // is where an unlowered store sits.
    runtime.setLenSelectValue("devStoredLenSelect", "3");
    runtime.setLenSelectDisplayLength("devStoredLenSelect", 12);
    runtime.setLenSelectValue("devStoredLenSelect", "12");
    runtime.setLenSelectDisplayLength("devStoredLenSelect", 3);
    assert.equal(select.value, "3");
    assert.match(runtime.storedLenUnavailableReason(6), /shown at 6/u);
    assert.equal(runtime.storedLenUnavailableReason(0), "");
  } finally {
    restore();
  }
});

test("both hosts show a Stored at value beside each length", () => {
  // The Dataset window draws the pair as two of the same length control.
  assert.match(datasetViewerViewSource, /id="originStoredLenWrap"/u);
  assert.match(datasetViewerViewSource, /id="originStoredLenSelect"/u);
  assert.match(datasetViewerViewSource, /id="devStoredLenWrap"/u);
  assert.match(datasetViewerViewSource, /id="devStoredLenSelect"/u);
  assert.equal((datasetViewerViewSource.match(/class="lenStoredLabel">Stored at:/gu) || []).length, 2);
  assert.match(datasetViewerCss, /#datasetTopBar \.lenStoredLabel \{/u);
  // The DFM Data tab hosts the same two controls in its own spinner shape.
  assert.match(dfmPageSource, /data-target="originStoredLenSelect"/u);
  assert.match(dfmPageSource, /data-target="devStoredLenSelect"/u);
  assert.equal((dfmPageSource.match(/class="lenStoredLabel">Stored at:/gu) || []).length, 2);
  // Both are the shared length control, so they take the same list, lock and
  // tooltip treatment rather than a second implementation.
  assert.match(inputsControllerSource, /originStoredLenSelect: \{\s*wrapId: "originStoredLenWrap",/u);
  assert.match(inputsControllerSource, /devStoredLenSelect: \{\s*wrapId: "devStoredLenWrap",/u);
});

test("the origin store is read-only and the development store is live only while empty", () => {
  // As in ResQ, the Origin Length control fixes the origin store while the
  // dataset is empty, so its readout is never editable.
  assert.match(
    persistenceControllerSource,
    /applyStoredLenControl\("originStoredLenSelect", \{[\s\S]*?enabled: false,/u,
  );
  assert.match(persistenceControllerSource, /The origin period is fixed by Origin Length while the dataset is empty\./u);
  // The development store moves only while the dataset holds no value, which is
  // the one time ResQ allows it.
  assert.match(
    persistenceControllerSource,
    /applyStoredLenControl\("devStoredLenSelect", \{[\s\S]*?enabled: pending && !vector && !isDfmDataTabHost\(\),/u,
  );
  assert.match(persistenceControllerSource, /Stored at can be changed only while the dataset is empty\./u);
  // A control that cannot be changed is dimmed in place rather than removed, so
  // the period the file is held at is always on screen.
  assert.match(persistenceControllerSource, /setLenSelectLock\(selectId, \{ locked: !enabled, displayValue: displayValue \|\| shown, reason \}\);/u);
  // A vector has no development dimension, so its store reads 0 beside the 0 on
  // its Development Length.
  assert.match(persistenceControllerSource, /displayValue: vector \? "0" : "",/u);
});

test("an empty dataset's store follows the display until the user lowers it", () => {
  // ResQ moves an empty triangle's store with its display, so a choice is
  // remembered only against the display length it was made at.
  assert.match(persistenceControllerSource, /function chooseStoredDevelopmentLength\(value\)/u);
  assert.match(
    persistenceControllerSource,
    /storedDevelopmentChoiceDisplay = getCurrentLengthControlValues\(\)\.development_length;/u,
  );
  assert.match(persistenceControllerSource, /storedDevelopmentChoiceDisplay === display/u);
  // A fresh sidecar answer spends any store the user had asked for.
  assert.match(persistenceControllerSource, /storedDevelopmentChoice = 0;\s*storedDevelopmentChoiceDisplay = 0;/u);
  // Lowering the store leaves the display alone: no length is rewritten and no
  // reload is scheduled from that control.
  assert.match(
    dataControlsSource,
    /devStoredSel\.addEventListener\("change", \(\) => \{\s*chooseStoredDevelopmentLength\(devStoredSel\.value\);\s*if \(typeof refreshDatasetSettingsDirty === "function"\) refreshDatasetSettingsDirty\(\);\s*devStoredSel\.blur\(\);\s*\}\);/u,
  );
});

test("the save states the period the dataset's file is written at", () => {
  // Only a still-empty hand-entered dataset states a store; everywhere else the
  // sidecar keeps the period it already records.
  assert.match(persistenceControllerSource, /function storedDevelopmentLengthForSave\(\)/u);
  assert.match(
    persistenceControllerSource,
    /if \(isDfmDataTabHost\(\) \|\| !storedLengthIsPending\(\) \|\| currentDatasetIsVector\(\)\) return 0;/u,
  );
  assert.match(persistenceControllerSource, /stored_development_length: storedDevelopmentLengthForSave\(\) \|\| null,/u);
  // Asking for a finer store is a change Save has to carry on its own.
  assert.match(persistenceControllerSource, /\|\| storedDevelopmentLengthIsDirty\(\)/u);
});

// ResQ refuses a write at a coarser origin display and accepts one at a coarser
// development display, so ArcRho tests the two axes apart: the origin one locks
// the grid, the development one only warns what a save there will do.

test("only a coarser origin view is read-only, and it names that axis", () => {
  // The read-only chain reads the origin test alone; the combined test that
  // locked both axes is gone.
  assert.match(preferencesControllerSource, /\|\| datasetOriginDisplayIsCoarserThanStored\(\);/u);
  assert.match(preferencesControllerSource, /&& datasetOriginDisplayIsCoarserThanStored\(\);/u);
  assert.doesNotMatch(preferencesControllerSource, /datasetDisplayIsCoarserThanStored/u);
  assert.match(
    persistenceControllerSource,
    /function datasetOriginDisplayIsCoarserThanStored\(\) \{\s*if \(storedLengthIsPending\(\)\) return false;/u,
  );
  assert.match(
    persistenceControllerSource,
    /Values can be entered only at the stored origin period \(Origin \$\{stored\.origin_length\}\)\. Set the origin length back to edit\./u,
  );
  // One place decides the wording, so every refusal names the rule that
  // stopped it rather than blaming a generated dataset.
  assert.match(preferencesControllerSource, /function getDatasetReadOnlyMessage\(\)/u);
  assert.doesNotMatch(gridInteractionsSource, /setStatus\("Generated datasets are read-only\."\)/u);
  assert.match(gridInteractionsSource, /readOnlyMessage = \(\) => "Generated datasets are read-only\.",/u);
  assert.doesNotMatch(runControllerSource, /setStatus\("Generated datasets are read-only\."\)/u);
  // The Links tab already inherits the same rule.
  assert.match(persistenceControllerSource, /isDatasetReadOnly\(\) \|\| isDfmDataTabHost\(\)/u);
});

test("a coarser development view is editable and says what a save there does", () => {
  // The development test is its own function and nothing in the read-only
  // chain reads it, so typing, paste and links stay live at that view.
  assert.match(
    persistenceControllerSource,
    /function datasetDevelopmentDisplayIsCoarserThanStored\(\) \{\s*if \(storedLengthIsPending\(\) \|\| currentDatasetIsVector\(\)\) return false;/u,
  );
  assert.doesNotMatch(preferencesControllerSource, /datasetDevelopmentDisplayIsCoarserThanStored/u);
  // A save there rewrites the whole stored triangle, so one sentence stays on
  // the status line for as long as the view is up rather than waiting for a
  // refusal that never comes.
  assert.match(
    persistenceControllerSource,
    /Saving here writes each value into the stored period \(Development \$\{stored\.development_length\}\) at its own column age and clears the stored periods between\./u,
  );
  assert.match(runControllerSource, /datasetCoarseDevelopmentNote = \(\) => "",/u);
  assert.match(runControllerSource, /setStatus\(datasetCoarseDevelopmentNote\(\) \|\| meta \|\| "Ready"\);/u);
  assert.match(hostControllerSource, /datasetCoarseDevelopmentNote,/u);
});

test("a vector keeps its one Period wording and has no development view to relax", () => {
  assert.match(
    persistenceControllerSource,
    /Values can be entered only at the stored period \(Period \$\{stored\.origin_length\}\)\. Set the length back to edit\./u,
  );
  // A vector has no development dimension, so the development test is false
  // for one and the note it feeds never appears.
  assert.match(persistenceControllerSource, /storedLengthIsPending\(\) \|\| currentDatasetIsVector\(\)/u);
});

test("a length change is a display setting the save keeps, not a value edit", () => {
  // Changing a length dirties the settings and Save stays available: the
  // display shape is persisted even though no value may be written at it.
  assert.match(persistenceControllerSource, /left\.origin_length === right\.origin_length/u);
  assert.match(
    persistenceControllerSource,
    /saveBlocked: isTemporaryDatasetView \|\| runtime\.datasetInstanceNameConflict \|\| !hasContext \|\| isDraftGridUnavailable\(\)/u,
  );
  // Going back down to the stored period is always allowed, so an edit is
  // never one save away from being locked out: the floor is the stored pair.
  assert.match(persistenceControllerSource, /function getManualDatasetLengthBaseline\(\) \{\s*const stored = getStoredLengthPair\(\);/u);
});
