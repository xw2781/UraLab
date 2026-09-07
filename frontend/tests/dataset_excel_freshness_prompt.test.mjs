import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const persistenceSource = await readFile(
  new URL("../ui/shared/tabs/data/data_tab_persistence_controller.js", import.meta.url),
  "utf8",
);
const messageBoxSource = await readFile(
  new URL("../ui/shared/components/message_box/message_box.js", import.meta.url),
  "utf8",
);
const messageBoxStyles = await readFile(
  new URL("../ui/shared/components/message_box/message_box.css", import.meta.url),
  "utf8",
);
const linkAlertSource = await readFile(
  new URL("../ui/shared/integrations/excel_link_alert.js", import.meta.url),
  "utf8",
);
const linkAlert = await import(
  `data:text/javascript;base64,${Buffer.from(
    linkAlertSource
      .replace(
        /import \{[\s\S]*?\} from "\/ui\/shared\/integrations\/excel_api\.js\?v=[^"]+";/u,
        "const openExcelWorkbook = () => {};",
      )
      .replace(
        /import \{[\s\S]*?\} from "\/ui\/shared\/components\/message_box\/message_box\.js\?v=[^"]+";/u,
        "const showPageMessageBox = () => Promise.resolve();",
      ),
  ).toString("base64")}`
);

test("DSV validates its Excel links once after linked sidecar data loads", () => {
  assert.match(persistenceSource, /datasetExcelLinkCheckedKeys = new Set\(\)/);
  assert.match(persistenceSource, /window\.setTimeout\(async \(\) => \{/);
  assert.match(persistenceSource, /validateLinks\(\s*state\.fileMtime/);
  assert.match(persistenceSource, /if \(data\.exists\) scheduleDatasetExcelLinkCheck/);
});

test("a reload reads Excel again only where a workbook has been saved since the dataset", () => {
  // Changing the view a dataset is shown at reloads it, and the reload used to
  // read every linked workbook from the top. The dataset's own CSV already
  // holds those figures, so the workbooks are stated first and read only when
  // one is newer than that file.
  assert.match(
    persistenceSource,
    /if \(options\?\.forceReload === true\) \{\s*await refreshDatasetExternalLinksIfWorkbooksChanged\(\{ isCurrent \}\);/u,
  );
  assert.match(
    persistenceSource,
    /const changed = await runtime\.datasetExternalLinks\.findNewerWorkbooks\(state\.fileMtime\);/u,
  );
  // An unreachable workbook is not a changed one: the stored figures stand.
  assert.match(
    persistenceSource,
    /if \(!changed\?\.ok \|\| !changed\.newerWorkbooks\?\.length\) \{\s*return \{ linkedCellCount: 0, changedCount: 0, failedCount: 0 \};\s*\}\s*return refreshDatasetExternalLinks\(options\);/u,
  );
});

test("a broken reference replaces the newer-workbook prompt", () => {
  // The alert comes first and returns, so the "Linked Excel File Updated"
  // prompt never queues behind a reference that cannot be refreshed anyway.
  assert.match(
    persistenceSource,
    /if \(result\.failures\.length\) \{[\s\S]*?await reportDatasetExcelLinkFailures\(result\.failures, \{ isCurrent \}\);\s*return;\s*\}\s*if \(!result\.newerWorkbookCount\) return;/u,
  );
  // The grid repaints before the alert so the red cells are already visible
  // behind it, and stay visible once it is dismissed.
  assert.match(
    persistenceSource,
    /async function reportDatasetExcelLinkFailures\(failures, options = \{\}\) \{[\s\S]*?renderTable\(\);/u,
  );
  assert.match(
    persistenceSource,
    /await showExcelLinkFailureAlert\(\{[\s\S]*?valueNoun: "linked dataset cell",\s*\}\);/u,
  );
});

test("every failed refresh reaches the alert, named or not", () => {
  // A refresh the user asked for that did not do what they asked must say so in
  // the window; the status bar is only for a refresh that succeeded.
  assert.match(
    persistenceSource,
    /const unnamedCount = failures\.length \? 0 : Number\(result\.failedCount\) \|\| 0;\s*if \(failures\.length \|\| unnamedCount\) \{/u,
  );
  assert.match(
    persistenceSource,
    /reportDatasetExcelLinkFailures\(failures, \{\s*isCurrent,\s*unnamedCount,\s*reason: result\.error,\s*\}\);/u,
  );
  assert.doesNotMatch(persistenceSource, /linked dataset cell\$\{result\.failedCount === 1 \? "" : "s"\} failed/u);
});

test("a refresh that never came back still names what it tried and what it kept", () => {
  const message = linkAlert.describeExcelLinkFailures([], {
    valueNoun: "linked dataset cell",
    unnamedCount: 10,
    reason: "Excel refresh failed.",
  });

  assert.match(message, /^10 linked dataset cells could not be refreshed\./u);
  assert.match(message, /saved values are kept, so nothing in this window changed/u);
  assert.match(message, /\n\nExcel refresh failed\.$/u);
  // Nothing to name means nothing is claimed about a reference.
  assert.equal(message.includes("→"), false);
  assert.equal(linkAlert.describeExcelLinkFailures([], { unnamedCount: 0 }), "");
});

test("freshness prompt defaults to keeping values and refreshes only on request", () => {
  assert.match(persistenceSource, /okLabel: "Keep Current Values"/);
  assert.match(persistenceSource, /actions: \[\{ id: "refresh", label: "Refresh from Excel" \}\]/);
  assert.match(persistenceSource, /balancedActions: true/);
  assert.match(persistenceSource, /if \(choice === "refresh" && isCurrent\(\)\) \{\s*await refreshDatasetExternalLinks\(\{\s*isCurrent,\s*markRefreshedCellsDirty: true,/);
  assert.match(persistenceSource, /runtime\.datasetExternalLinks\.refreshAll\([\s\S]*?markRefreshedCellsDirty: options\?\.markRefreshedCellsDirty === true/);
  assert.match(persistenceSource, /getDataTabLinksController\(\)\?\.refresh\?\.\(\);\s*updateDatasetSaveUi\(\);/);
  assert.match(messageBoxSource, /okButton\.textContent = String\(okLabel \|\| "OK"\)/);
  assert.match(messageBoxSource, /pageMessageBoxActionsBalanced/);
  assert.match(messageBoxStyles, /\.pageMessageBoxActionsBalanced \{\s*gap: 8px;/);
  assert.match(messageBoxStyles, /\.pageMessageBoxActionsBalanced \.pageMessageBoxButton \{\s*flex: 1 1 0;/);
});

test("the alert names each broken reference, its destination, and the reason", () => {
  const message = linkAlert.describeExcelLinkFailures([
    {
      workbookPath: "C:\\Data\\Inputs.xlsx",
      worksheet: "Sheet1",
      sourceCell: "B5",
      destination: "2019 Q1 / Ultimate",
      error: "Not numeric: '#REF!'",
    },
    {
      workbookPath: "C:\\Data\\Inputs.xlsx",
      worksheet: "Sheet1",
      sourceCell: "B6",
      destination: "2019 Q2 / Ultimate",
      error: "Not numeric: '#REF!'",
    },
  ], { valueNoun: "linked dataset cell" });

  assert.match(message, /^2 linked dataset cells could not be read from Inputs\.xlsx\./u);
  assert.match(message, /saved values are kept and shown in red until the reference is fixed/u);
  assert.match(message, /Sheet1!B5 → 2019 Q1 \/ Ultimate: Not numeric: '#REF!'/u);
  assert.match(message, /Sheet1!B6 → 2019 Q2 \/ Ultimate: Not numeric: '#REF!'/u);
});

test("the alert lists only the first few failures and says how many are left", () => {
  const message = linkAlert.describeExcelLinkFailures(
    Array.from({ length: 40 }, (_value, index) => ({
      workbookPath: "C:\\Data\\Inputs.xlsx",
      worksheet: "Sheet1",
      sourceCell: `B${index + 1}`,
      destination: `Row ${index + 1}`,
      error: "Not numeric: '#REF!'",
    })),
    { valueNoun: "linked dataset cell" },
  );

  assert.match(message, /^40 linked dataset cells could not be read/u);
  assert.match(message, /\nand 34 more\.$/u);
  assert.equal(message.includes("B7 →"), false);
});
