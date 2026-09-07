import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const source = await readFile(
  new URL("../ui/shared/components/formula_hover/formula_hover.js", import.meta.url),
  "utf8",
);
const styles = await readFile(
  new URL("../ui/shared/components/formula_hover/formula_hover.css", import.meta.url),
  "utf8",
);
const sharedStyles = await readFile(
  new URL("../ui/shared/components/formula_bar/formula_bar.css", import.meta.url),
  "utf8",
);
const sharedModuleUrl = (name) => new URL(
  `../ui/shared/components/formula_bar/${name}`,
  import.meta.url,
).href;
const sharedSpecifier = (name) => new RegExp(
  `"/ui/shared/components/formula_bar/${name.replace(".", "\\.")}\\?v=[^"]*"`,
  "u",
);
const replaceSharedImport = (text, name, replacement, origin) => {
  const specifier = sharedSpecifier(name);
  if (!specifier.test(text)) throw new Error(`${name} import not found in ${origin}`);
  return text.replace(specifier, JSON.stringify(replacement));
};
const asModule = (text) => `data:text/javascript;base64,${Buffer.from(text).toString("base64")}`;
// The drag controller reaches for the layout module by the same runtime path,
// so it is inlined with that import resolved rather than loaded from disk.
const dragSource = await readFile(
  new URL("../ui/shared/components/formula_bar/formula_bar_drag.js", import.meta.url),
  "utf8",
);
const dragModuleUrl = asModule(replaceSharedImport(
  dragSource,
  "formula_bar_layout.js",
  sharedModuleUrl("formula_bar_layout.js"),
  "formula_bar_drag.js",
));
// Resolved to real modules so the test exercises the shared layout, tokenizer,
// and drag controller rather than stand-ins.
const patchedSource = [
  ["formula_bar_layout.js", sharedModuleUrl("formula_bar_layout.js")],
  ["formula_text.js", sharedModuleUrl("formula_text.js")],
  ["formula_bar_drag.js", dragModuleUrl],
].reduce(
  (text, [name, replacement]) => replaceSharedImport(text, name, replacement, "formula_hover.js"),
  source,
);
const formulaHover = await import(
  `data:text/javascript;base64,${Buffer.from(patchedSource).toString("base64")}`
);

class FakeClassList {
  constructor(owner) {
    this.owner = owner;
  }

  values() {
    return new Set(String(this.owner.className || "").split(/\s+/u).filter(Boolean));
  }

  write(values) {
    this.owner.className = Array.from(values).join(" ");
  }

  add(...names) {
    const values = this.values();
    names.forEach((name) => values.add(name));
    this.write(values);
  }

  remove(...names) {
    const values = this.values();
    names.forEach((name) => values.delete(name));
    this.write(values);
  }

  contains(name) {
    return this.values().has(name);
  }

  toggle(name, force) {
    const values = this.values();
    const enabled = force === undefined ? !values.has(name) : !!force;
    if (enabled) values.add(name);
    else values.delete(name);
    this.write(values);
    return enabled;
  }
}

class FakeElement {
  constructor(tagName, ownerDocument) {
    this.tagName = String(tagName).toUpperCase();
    this.ownerDocument = ownerDocument;
    this.children = [];
    this.parentElement = null;
    this.attributes = new Map();
    this.listeners = new Map();
    this.className = "";
    this.classList = new FakeClassList(this);
    this.dataset = {};
    this.style = {};
    this.hidden = false;
    this.readOnly = false;
    this.textContent = "";
    this.value = "";
    this._rect = { left: 0, top: 0, right: 0, bottom: 0, width: 0, height: 0 };
  }

  get isConnected() {
    let current = this;
    while (current) {
      if (current === this.ownerDocument.documentElement) return true;
      current = current.parentElement;
    }
    return false;
  }

  appendChild(child) {
    child.parentElement = this;
    this.children.push(child);
    return child;
  }

  setAttribute(name, value) {
    this.attributes.set(name, String(value));
  }

  getAttribute(name) {
    if (this.attributes.has(name)) return this.attributes.get(name);
    if (name === "href" && this.href !== undefined) return String(this.href);
    return null;
  }

  removeAttribute(name) {
    this.attributes.delete(name);
  }

  addEventListener(type, listener) {
    if (!this.listeners.has(type)) this.listeners.set(type, []);
    this.listeners.get(type).push(listener);
  }

  dispatch(type, properties = {}) {
    const event = {
      type,
      target: this,
      currentTarget: this,
      preventDefault() {},
      stopPropagation() {},
      ...properties,
    };
    for (const listener of this.listeners.get(type) || []) listener(event);
    return event;
  }

  focus() {
    this.ownerDocument.activeElement = this;
    this.dispatch("focus");
  }

  select() {
    this.selected = true;
  }

  contains(candidate) {
    if (candidate === this) return true;
    return this.children.some((child) => child.contains(candidate));
  }

  getBoundingClientRect() {
    return this._rect;
  }

  remove() {
    if (!this.parentElement) return;
    const index = this.parentElement.children.indexOf(this);
    if (index >= 0) this.parentElement.children.splice(index, 1);
    this.parentElement = null;
  }
}

class FakeDocument {
  constructor(windowRef) {
    this.defaultView = windowRef;
    this.listeners = new Map();
    this.activeElement = null;
    this.documentElement = new FakeElement("html", this);
    this.documentElement.clientWidth = windowRef.innerWidth;
    this.documentElement.clientHeight = windowRef.innerHeight;
    this.head = new FakeElement("head", this);
    this.body = new FakeElement("body", this);
    this.documentElement.appendChild(this.head);
    this.documentElement.appendChild(this.body);
  }

  createElement(tagName) {
    return new FakeElement(tagName, this);
  }

  createTextNode(data) {
    const node = new FakeElement("#text", this);
    node.textContent = String(data ?? "");
    return node;
  }

  allElements() {
    const elements = [];
    const visit = (element) => {
      elements.push(element);
      element.children.forEach(visit);
    };
    visit(this.documentElement);
    return elements;
  }

  getElementById(id) {
    return this.allElements().find((element) => element.id === id) || null;
  }

  querySelectorAll(selector) {
    if (selector !== 'link[rel="stylesheet"]') return [];
    return this.allElements().filter(
      (element) => element.tagName === "LINK" && element.rel === "stylesheet",
    );
  }

  addEventListener(type, listener) {
    if (!this.listeners.has(type)) this.listeners.set(type, []);
    this.listeners.get(type).push(listener);
  }
}

function byClass(documentRef, className) {
  return documentRef.allElements().find((element) => element.classList.contains(className));
}

function setup(options = {}) {
  const windowListeners = new Map();
  const windowRef = {
    innerWidth: 800,
    innerHeight: 600,
    setTimeout,
    clearTimeout,
    requestAnimationFrame: (callback) => callback(),
    addEventListener(type, listener) {
      if (!windowListeners.has(type)) windowListeners.set(type, new Set());
      windowListeners.get(type).add(listener);
    },
    removeEventListener(type, listener) {
      windowListeners.get(type)?.delete(listener);
    },
  };
  const documentRef = new FakeDocument(windowRef);
  const anchor = documentRef.createElement("td");
  anchor._rect = { left: 40, top: 100, right: 140, bottom: 128, width: 100, height: 28 };
  documentRef.body.appendChild(anchor);
  const commits = [];
  const dismisses = [];
  const editStarts = [];
  const statuses = [];
  const drafts = [];
  const closes = [];
  const controller = formulaHover.createFormulaHoverEditor({
    documentRef,
    windowRef,
    hideDelayMs: options.hideDelayMs ?? 5,
    onDismiss: (context) => dismisses.push(context),
    onEditStart: (context) => editStarts.push(context),
    onStatus: (message) => statuses.push(message),
    onDraftChange: (value) => drafts.push(value),
    onClosed: (context) => closes.push(context),
    shouldStayOpenUnfocused: () => !!options.stayOpenUnfocused,
    onCommit: async (request) => {
      commits.push(request);
      return options.commitResult || { ok: true };
    },
  });
  return { anchor, closes, commits, controller, dismisses, documentRef, drafts, editStarts, statuses, windowRef };
}

/** Open the editor and give the bar a real size, which the fake DOM will not. */
function openSizedBar(context, linkContext = LINK_CONTEXT, openOptions = {}) {
  context.controller.open(context.anchor, linkContext, openOptions);
  const root = byClass(context.documentRef, "arFormulaHover");
  root._rect = { left: 0, top: 0, right: 300, bottom: 30, width: 300, height: 30 };
  return root;
}

const LINK_CONTEXT = {
  reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!A1:B2",
  anchorDisplayRow: 1,
  anchorDisplayColumn: 2,
};

test("the linked-cell editor wears the shared formula bar and adds only its own placement", () => {
  assert.ok(source.includes("/ui/shared/components/formula_bar/formula_bar.css?v="));
  assert.ok(source.includes("/ui/shared/components/formula_hover/formula_hover.css?v="));
  assert.ok(source.includes('root.className = "arFormulaBar arFormulaHover"'));
  // Nothing visual is redefined here: the shared bar owns border, shadow, and type.
  assert.doesNotMatch(styles, /box-shadow|border-radius|font:\s*12px/u);
  assert.match(styles, /position:\s*fixed/u);
});

test("the formula holds one line and keeps a stretch of bar to click past its end", () => {
  // The bar wraps only for its error message. A starting size the bar can
  // always fit is what stops the formula from taking the error's line when the
  // measured width falls short of the text: it is shrunk and clipped instead.
  assert.match(sharedStyles, /\.arFormulaBarInput\s*\{[^}]*flex:\s*1 1 40px/su);
  assert.match(sharedStyles, /\.arFormulaBarDisplay\s*\{[^}]*flex:\s*1 1 40px/su);
  // Padding past the formula is what a click landing to the right of it hits.
  assert.match(sharedStyles, /\.arFormulaBarInput\s*\{[^}]*padding:\s*0 20px 0 0/su);
  assert.match(sharedStyles, /\.arFormulaBarDisplay\s*\{[^}]*padding:\s*0 20px 0 0/su);
});

const NOTE = "This dataset's cells are linked. Set Origin Length to 1 to view or edit the formula.";

test("a notice takes the place of the formula and can never be typed into", () => {
  // What a linked cell says while the window shows the dataset at a coarser
  // period than the file is stored at: there is no formula to show, because no
  // cell on screen is one a link names.
  const context = setup();
  context.controller.open(context.anchor, { note: NOTE, readOnly: false });
  const root = byClass(context.documentRef, "arFormulaHover");
  const input = byClass(context.documentRef, "arFormulaBarInput");
  const display = byClass(context.documentRef, "arFormulaBarDisplay");

  assert.ok(root.classList.contains("isOpen"));
  assert.ok(root.classList.contains("isNote"));
  assert.equal(display.textContent, NOTE);
  assert.equal(display.style.display, "");
  // The notice is read, never edited: the input stays hidden and read-only
  // however the caller asked for the bar to open.
  assert.equal(input.style.display, "none");
  assert.equal(input.value, "");
  assert.equal(input.readOnly, true);
  assert.equal(context.controller.isEditing(), false);
  assert.equal(context.controller.getDraft(), "");
});

test("a notice never leaves the bar dressed as a formula for the next cell", () => {
  const context = setup();
  context.controller.open(context.anchor, { note: NOTE });
  const root = byClass(context.documentRef, "arFormulaHover");
  assert.ok(root.classList.contains("isNote"));

  context.controller.open(context.anchor, LINK_CONTEXT);
  assert.equal(root.classList.contains("isNote"), false);
  const display = byClass(context.documentRef, "arFormulaBarDisplay");
  // Back to the colorized formula rendering, not the notice's plain sentence.
  assert.ok(display.children.some((child) => child.classList.contains("fmtExcelRef")));
  const input = byClass(context.documentRef, "arFormulaBarInput");
  assert.equal(input.value, LINK_CONTEXT.reference);
});

test("every cell can carry the same notice and still be told apart", () => {
  // The notice is one sentence for the whole grid, so the caller keys the bar
  // by cell; without that a click on one cell would close the bar opened on
  // its neighbour.
  const context = setup();
  const secondCell = context.documentRef.createElement("td");
  secondCell._rect = { left: 140, top: 100, right: 240, bottom: 128, width: 100, height: 28 };
  context.documentRef.body.appendChild(secondCell);
  context.controller.attach(context.anchor, { note: NOTE }, { key: "note:0,0" });
  context.controller.attach(secondCell, { note: NOTE }, { key: "note:0,1" });

  context.anchor.dispatch("click", { button: 0 });
  const root = byClass(context.documentRef, "arFormulaHover");
  assert.ok(root.classList.contains("isOpen"));
  secondCell.dispatch("click", { button: 0 });
  assert.ok(root.classList.contains("isOpen"));
  assert.equal(context.anchor.getAttribute("aria-description"), NOTE);
});

test("the bar is measured again once the stylesheets it is measured in arrive", () => {
  const context = setup();
  const root = openSizedBar(context);
  context.controller.reposition();
  assert.equal(root.style.width, "300px");

  // The first bar of a session opens before its styles do, so the width it took
  // was measured in the page's font rather than the bar's.
  root._rect = { ...root._rect, right: 460, width: 460 };
  const link = context.documentRef.allElements().find((element) => element.tagName === "LINK");
  assert.ok(link, "the bar injects its stylesheets");
  link.dispatch("load");

  assert.equal(root.style.width, "460px");
});

test("array formula positioning uses the complete range and never overlaps it", () => {
  const context = setup();
  const rangeRect = { left: 40, top: 10, right: 240, bottom: 120, width: 200, height: 110 };
  context.controller.open(context.anchor, LINK_CONTEXT, { positionRect: rangeRect });
  const root = byClass(context.documentRef, "arFormulaHover");
  root._rect = { left: 0, top: 0, right: 300, bottom: 30, width: 300, height: 30 };
  context.controller.reposition();

  assert.equal(root.dataset.placement, "below");
  assert.equal(root.style.top, "124px");
  assert.ok(Number.parseInt(root.style.top, 10) > rangeRect.bottom);
  // Sized to its content rather than a fixed panel width.
  assert.equal(root.style.width, "300px");
});

test("moving across one array keeps the canonical anchor and formula-bar position", () => {
  const context = setup();
  const secondCell = context.documentRef.createElement("td");
  secondCell._rect = { left: 140, top: 128, right: 240, bottom: 156, width: 100, height: 28 };
  context.documentRef.body.appendChild(secondCell);
  const rangeRect = { left: 40, top: 100, right: 240, bottom: 156, width: 200, height: 56 };
  const options = {
    resolveAnchor: () => context.anchor,
    positionRect: () => rangeRect,
  };
  context.controller.attach(context.anchor, LINK_CONTEXT, options);
  context.controller.attach(secondCell, LINK_CONTEXT, options);

  context.anchor.dispatch("mouseenter");
  const root = byClass(context.documentRef, "arFormulaHover");
  root._rect = { left: 0, top: 0, right: 300, bottom: 30, width: 300, height: 30 };
  context.controller.reposition();
  const firstPosition = { left: root.style.left, top: root.style.top, placement: root.dataset.placement };

  secondCell.dispatch("mouseenter");
  assert.deepEqual(
    { left: root.style.left, top: root.style.top, placement: root.dataset.placement },
    firstPosition,
  );
});

test("formula hover positioning prefers above the cell and flips below near the viewport top", () => {
  assert.deepEqual(
    formulaHover.calculateFormulaHoverPosition(
      { left: 40, top: 100, bottom: 125 },
      { width: 300, height: 30 },
      { width: 800, height: 600 },
    ),
    { left: 40, top: 66, width: 300, placement: "above" },
  );
  assert.deepEqual(
    formulaHover.calculateFormulaHoverPosition(
      { left: 40, top: 10, bottom: 35 },
      { width: 300, height: 30 },
      { width: 800, height: 600 },
    ),
    { left: 40, top: 39, width: 300, placement: "below" },
  );
});

test("linked-cell hover shows the full formula and delays hiding while entering the bar", async () => {
  const context = setup();
  context.controller.attach(context.anchor, LINK_CONTEXT);
  context.anchor.dispatch("mouseenter");

  const root = byClass(context.documentRef, "arFormulaHover");
  const input = byClass(context.documentRef, "arFormulaBarInput");
  assert.equal(root.classList.contains("isOpen"), true);
  assert.equal(root.dataset.placement, "above");
  assert.equal(input.value, LINK_CONTEXT.reference);
  assert.match(context.anchor.getAttribute("aria-description"), /press F2/u);

  context.anchor.dispatch("mouseleave");
  root.dispatch("mouseenter");
  await new Promise((resolve) => setTimeout(resolve, 10));
  assert.equal(root.classList.contains("isOpen"), true);

  root.dispatch("mouseleave");
  await new Promise((resolve) => setTimeout(resolve, 10));
  assert.equal(root.classList.contains("isOpen"), false);
});

test("Enter commits the edited formula and Escape cancels without a commit", async () => {
  const context = setup();
  context.controller.open(context.anchor, LINK_CONTEXT, { focus: true });
  const root = byClass(context.documentRef, "arFormulaHover");
  const input = byClass(context.documentRef, "arFormulaBarInput");
  assert.equal(context.documentRef.activeElement, input);
  assert.equal(input.selected, true);
  assert.deepEqual(context.editStarts, [{ ...LINK_CONTEXT, formula: LINK_CONTEXT.reference }]);

  const nextFormula = "='C:\\Data\\[Book.xlsx]Sheet 1'!C3:D4";
  input.value = nextFormula;
  let prevented = false;
  input.dispatch("keydown", {
    key: "Enter",
    preventDefault() { prevented = true; },
    stopPropagation() {},
  });
  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(prevented, true);
  assert.deepEqual(context.commits, [{ formula: nextFormula, context: { ...LINK_CONTEXT, formula: LINK_CONTEXT.reference } }]);
  assert.equal(root.getAttribute("aria-busy"), "false");
  assert.equal(root.classList.contains("isOpen"), false);

  context.controller.open(context.anchor, LINK_CONTEXT, { focus: true });
  input.value = "changed but canceled";
  input.dispatch("keydown", { key: "Escape" });
  assert.equal(input.value, LINK_CONTEXT.reference);
  assert.equal(root.classList.contains("isOpen"), false);
  assert.equal(context.commits.length, 1);
  assert.equal(context.dismisses.length, 2);
});

test("failed formula commits retain the editor with an accessible error", async () => {
  const context = setup({ commitResult: { ok: false, error: "Workbook unavailable." } });
  context.controller.open(context.anchor, LINK_CONTEXT, { focus: true });
  const root = byClass(context.documentRef, "arFormulaHover");
  const input = byClass(context.documentRef, "arFormulaBarInput");
  input.value = "='C:\\Missing\\[Book.xlsx]Sheet 1'!A1";

  assert.equal(await context.controller.commit(), false);
  assert.equal(root.classList.contains("isOpen"), true);
  assert.equal(root.classList.contains("has-error"), true);
  assert.equal(input.getAttribute("aria-invalid"), "true");
  assert.equal(byClass(context.documentRef, "arFormulaHoverError").textContent, "Workbook unavailable.");
  assert.equal(context.documentRef.activeElement, input);
  assert.equal(context.statuses.at(-1), "Workbook unavailable.");
});

test("the editor renders the formula read-only and swaps to the input on focus", () => {
  const context = setup();
  context.controller.open(context.anchor, LINK_CONTEXT);
  const input = byClass(context.documentRef, "arFormulaBarInput");
  const display = byClass(context.documentRef, "arFormulaBarDisplay");

  assert.equal(input.style.display, "none", "the raw input stays out of the way until edited");
  assert.equal(display.style.display, "");
  const excelRef = display.children.find((child) => child.classList.contains("fmtExcelRef"));
  assert.ok(excelRef, "the external reference is colorized");
  assert.equal(excelRef.textContent, LINK_CONTEXT.reference.replace(/^=/u, ""));

  input.focus();
  assert.equal(input.style.display, "");
  assert.equal(display.style.display, "none");
});

test("clicking a linked cell pins the editor open and clicking it again releases it", async () => {
  const context = setup();
  context.controller.attach(context.anchor, LINK_CONTEXT);
  const findRoot = () => byClass(context.documentRef, "arFormulaHover");

  // Hovering alone stays transient: leaving still hides it.
  context.anchor.dispatch("mouseenter");
  assert.equal(findRoot().classList.contains("isOpen"), true);
  context.anchor.dispatch("mouseleave");
  await new Promise((resolve) => setTimeout(resolve, 10));
  assert.equal(findRoot().classList.contains("isOpen"), false);

  // A click pins it, and leaving no longer hides it.
  context.anchor.dispatch("click", { button: 0 });
  assert.equal(findRoot().classList.contains("isOpen"), true);
  context.anchor.dispatch("mouseleave");
  await new Promise((resolve) => setTimeout(resolve, 10));
  assert.equal(findRoot().classList.contains("isOpen"), true, "a pinned editor stays put");

  // Clicking the same cell again releases it.
  context.anchor.dispatch("click", { button: 0 });
  assert.equal(findRoot().classList.contains("isOpen"), false);
  assert.match(context.anchor.getAttribute("aria-description"), /Hover, click, or press F2/u);
});

test("a link in the grid's first row still gets its bar above the range", () => {
  const context = setup();
  // The grid starts at y=25; the linked range fills its first rows, so there is
  // no room above inside the grid itself.
  const gridEl = context.documentRef.createElement("div");
  gridEl.id = "tableWrap";
  gridEl._rect = { left: 20, top: 25, right: 700, bottom: 400, width: 680, height: 375 };
  gridEl.clientWidth = 680;
  gridEl.scrollTop = 0;
  context.documentRef.body.appendChild(gridEl);

  const rangeRect = { left: 150, top: 57, right: 280, bottom: 315, width: 130, height: 258 };
  context.controller.open(context.anchor, LINK_CONTEXT, { positionRect: rangeRect });
  const root = byClass(context.documentRef, "arFormulaHover");
  root._rect = { left: 0, top: 0, right: 300, bottom: 30, width: 300, height: 30 };
  context.controller.reposition();

  assert.equal(root.dataset.placement, "above");
  // 57 - 30 - 4, which is above the grid's own top edge and still on screen.
  assert.equal(root.style.top, "23px");
  assert.equal(root.style.left, "150px");
});

test("the fx badge carries the bar off its cell, and a press alone moves nothing", () => {
  const context = setup();
  const root = openSizedBar(context);
  const badge = byClass(context.documentRef, "arFormulaBarFxIcon");

  badge.dispatch("pointerdown", { pointerId: 1, button: 0, clientX: 100, clientY: 200 });
  badge.dispatch("pointermove", { pointerId: 1, clientX: 102, clientY: 201 });
  assert.equal(root.classList.contains("isDragPlaced"), false, "under the slop of an ordinary press");

  badge.dispatch("pointermove", { pointerId: 1, clientX: 160, clientY: 260 });
  assert.equal(root.classList.contains("isDragPlaced"), true);
  assert.equal(root.classList.contains("isDragging"), true);
  assert.equal(root.style.left, "60px");
  assert.equal(root.style.top, "60px");

  badge.dispatch("pointerup", { pointerId: 1 });
  assert.equal(root.classList.contains("isDragging"), false);
  assert.equal(root.classList.contains("isDragPlaced"), true, "where it was dropped is where it stays");

  // Re-measuring must not pull the bar back over the cell it was moved off.
  context.controller.reposition();
  assert.equal(root.style.left, "60px");
  assert.equal(root.style.top, "60px");
});

test("a dropped bar is kept on screen", () => {
  const context = setup();
  const root = openSizedBar(context);
  const badge = byClass(context.documentRef, "arFormulaBarFxIcon");

  badge.dispatch("pointerdown", { pointerId: 2, button: 0, clientX: 0, clientY: 0 });
  badge.dispatch("pointermove", { pointerId: 2, clientX: 4000, clientY: 4000 });

  // 800 - 8 - 300 wide, and 600 - 8 - 30 tall.
  assert.equal(root.style.left, "492px");
  assert.equal(root.style.top, "562px");
});

test("a hand-placed bar belongs to one formula and is handed back for any other", () => {
  const context = setup();
  const root = openSizedBar(context);
  const badge = byClass(context.documentRef, "arFormulaBarFxIcon");
  badge.dispatch("pointerdown", { pointerId: 3, button: 0, clientX: 100, clientY: 100 });
  badge.dispatch("pointermove", { pointerId: 3, clientX: 200, clientY: 180 });
  assert.equal(root.classList.contains("isDragPlaced"), true);

  context.controller.open(context.anchor, {
    ...LINK_CONTEXT,
    reference: "='C:\\Data\\[Book.xlsx]Sheet 1'!D4",
  });
  assert.equal(root.classList.contains("isDragPlaced"), false);
  assert.equal(root.dataset.placement, "above");
});

test("a dataset reference is colorized apart from an Excel one", () => {
  const context = setup();
  openSizedBar(context, { ...LINK_CONTEXT, reference: "=[C 82 - Prior Qtr Selected][1:7]" });
  const display = byClass(context.documentRef, "arFormulaBarDisplay");

  const internal = display.children.filter((child) => child.classList.contains("fmtInternalRef"));
  assert.deepEqual(
    internal.map((child) => child.textContent),
    ["[C 82 - Prior Qtr Selected]", "[1:7]"],
  );
  assert.equal(display.children.some((child) => child.classList.contains("fmtExcelRef")), false);
});

test("the editor survives losing focus while its formula is being pointed at another window", () => {
  const context = setup({ stayOpenUnfocused: true });
  const root = openSizedBar(context, LINK_CONTEXT, { focus: true });
  const input = byClass(context.documentRef, "arFormulaBarInput");

  input.value = "=[";
  input.dispatch("input");
  assert.deepEqual(context.drafts, ["=["]);

  // Clicking a cell in the other window takes focus out of this one.
  input.dispatch("blur");
  assert.equal(root.classList.contains("isOpen"), true);
  assert.equal(input.style.display, "", "the draft is still the thing on show");
  assert.equal(context.controller.isEditing(), true);

  assert.equal(context.controller.setDraft("=[Paid Claims][1:3]", { focus: true }), true);
  assert.equal(input.value, "=[Paid Claims][1:3]");
  assert.equal(context.documentRef.activeElement, input);
});

test("closing the editor is reported once, whether or not it held focus", () => {
  const context = setup();
  openSizedBar(context);
  context.controller.hide();
  assert.equal(context.closes.length, 1);
  assert.equal(context.dismisses.length, 0, "nothing had focus to give back");
  assert.deepEqual(context.closes[0], { ...LINK_CONTEXT, formula: LINK_CONTEXT.reference });
});
