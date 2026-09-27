import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import {
  IMPLEMENTED_ACTIONS,
  INTERNAL_ACTIONS,
  executeOfficeAction,
  getImplementedAction,
  negotiateRequirementSets
} from "../src/actions.mjs";

test("source contract metadata matches every Office.js action registration", async () => {
  const contractFiles = [
    "../../src/ExcelMcp.Core/Commands/Table/ITableCommands.cs",
    "../../src/ExcelMcp.Core/Commands/Table/ITableColumnCommands.cs",
    "../../src/ExcelMcp.Core/Commands/ConditionalFormat/IConditionalFormattingCommands.cs",
    "../../src/ExcelMcp.Core/Commands/Sheet/ISheetCommands.cs"
  ];
  const contracts = new Map();
  for (const relativePath of contractFiles) {
    const source = await readFile(new URL(relativePath, import.meta.url), "utf8");
    const category = source.match(/\[ServiceCategory\("([^"]+)"/)?.[1];
    assert.ok(category, `${relativePath} must declare ServiceCategory`);
    for (const match of source.matchAll(
      /\[ServiceAction\("([^"]+)"\), OfficeAddInAction\("([^"]+)", mutation: (true|false)\)\]/g
    )) {
      contracts.set(`${category}.${match[1]}`, {
        requirementSet: match[2],
        mutation: match[3] === "true"
      });
    }
  }

  assert.deepEqual([...IMPLEMENTED_ACTIONS].sort(), [...contracts.keys()].sort());
  for (const [name, expected] of contracts) {
    const actual = getImplementedAction(name);
    assert.equal(actual.requirementSet, expected.requirementSet, name);
    assert.equal(actual.mutation, expected.mutation, name);
  }
});

test("negotiates numbered ExcelApi and ExcelApiDesktop independently", () => {
  const supported = new Set(["ExcelApi:1.1", "ExcelApi:1.16", "ExcelApiDesktop:1.1"]);
  const requirements = {
    isSetSupported(name, version) {
      return supported.has(`${name}:${version}`);
    }
  };

  assert.deepEqual(negotiateRequirementSets(requirements), {
    excelApi: ["1.1", "1.16"],
    excelApiDesktop: ["1.1"]
  });
});

test("registers disabled screenshot geometry actions as internal Desktop 1.1 mutations", () => {
  assert.deepEqual(INTERNAL_ACTIONS, [
    "screenshot.prepare-range-geometry",
    "screenshot.prepare-sheet-geometry",
    "screenshot.restore-view"
  ]);
  for (const name of INTERNAL_ACTIONS) {
    const action = getImplementedAction(name);
    assert.equal(action.requirementFamily, "ExcelApiDesktop", name);
    assert.equal(action.requirementSet, "1.1", name);
    assert.equal(action.mutation, true, name);
  }
  assert.equal(IMPLEMENTED_ACTIONS.includes("screenshot.prepare-range-geometry"), false);
});

test("prepares range geometry using converted edges and restores view with an owned token", async () => {
  const fixture = createGeometryContext();
  const runtime = {
    requirementSets: { excelApi: ["1.21"], excelApiDesktop: ["1.1"] },
    async run(callback) {
      return callback(fixture.context);
    }
  };
  const identity = {
    sessionId: "session-1",
    workbookUrl: "file:///tmp/owned.xlsx"
  };

  const prepared = await executeOfficeAction({
    ...identity,
    requestId: "request-prepare",
    action: "screenshot.prepare-range-geometry",
    payload: { sheetName: "Data", rangeAddress: "B2:D8" }
  }, runtime);

  assert.match(prepared.captureToken, /^[0-9a-f-]{36}$/);
  assert.deepEqual(prepared.worksheet, { id: "sheet-data", name: "Data" });
  assert.deepEqual(prepared.range, { address: "Data!B2:D8" });
  assert.deepEqual(prepared.excelWindow, {
    index: 0,
    windowNumber: 42,
    type: "workbook",
    state: "normal"
  });
  assert.deepEqual(prepared.screenRect, {
    left: 220,
    top: 320,
    right: 820,
    bottom: 1120,
    width: 600,
    height: 800,
    unit: "physicalPixel",
    origin: "topLeftGlobalScreen"
  });
  assert.deepEqual(prepared.excelWindowScreenRect, {
    left: 40,
    top: 60,
    right: 2440,
    bottom: 1860,
    width: 2400,
    height: 1800,
    unit: "physicalPixel",
    origin: "topLeftGlobalScreen"
  });
  assert.deepEqual(prepared.viewState, {
    activeWorksheetId: "sheet-summary",
    activeWorksheetName: "Summary",
    selectionAddress: "Summary!C3:E4",
    activeCellAddress: "Summary!D3",
    visibleRangeAddress: "Summary!A1:M30",
    scrollRow: 7,
    scrollColumn: 3,
    zoom: 125,
    view: "normalView",
    windowState: "normal",
    freezePanes: true,
    split: false,
    splitRow: 1,
    splitColumn: 0,
    splitHorizontal: 18,
    splitVertical: 0
  });
  assert.deepEqual(fixture.scrollCalls, [
    { left: 100, top: 150, width: 300, height: 400, start: true }
  ]);
  assert.equal(fixture.selectedRange, "Data!B2:D8");

  fixture.window.scrollRow = 50;
  fixture.window.scrollColumn = 20;
  fixture.window.zoom = 80;
  await executeOfficeAction({
    ...identity,
    requestId: "request-restore",
    action: "screenshot.restore-view",
    payload: { captureToken: prepared.captureToken }
  }, runtime);

  assert.equal(fixture.activeSheet, "Summary");
  assert.equal(fixture.selectedRange, "Summary!C3:E4");
  assert.equal(fixture.window.scrollRow, 7);
  assert.equal(fixture.window.scrollColumn, 3);
  assert.equal(fixture.window.zoom, 125);
  await assert.rejects(
    executeOfficeAction({
      ...identity,
      requestId: "request-restore-again",
      action: "screenshot.restore-view",
      payload: { captureToken: prepared.captureToken }
    }, runtime),
    /capture token is unknown or has already been restored/
  );
});

test("capture token is bound to the exact workbook session", async () => {
  const fixture = createGeometryContext();
  const runtime = {
    requirementSets: { excelApi: ["1.21"], excelApiDesktop: ["1.1"] },
    async run(callback) {
      return callback(fixture.context);
    }
  };
  const prepared = await executeOfficeAction({
    sessionId: "session-owner",
    workbookUrl: "file:///tmp/owner.xlsx",
    requestId: "request-prepare",
    action: "screenshot.prepare-sheet-geometry",
    payload: { sheetName: "Data" }
  }, runtime);

  await assert.rejects(
    executeOfficeAction({
      sessionId: "session-2",
      workbookUrl: "file:///tmp/other.xlsx",
      requestId: "request-restore-wrong",
      action: "screenshot.restore-view",
      payload: { captureToken: prepared.captureToken }
    }, runtime),
    /capture token is not bound to this exact workbook session/
  );
  await executeOfficeAction({
    sessionId: "session-owner",
    workbookUrl: "file:///tmp/owner.xlsx",
    requestId: "request-restore-correct",
    action: "screenshot.restore-view",
    payload: { captureToken: prepared.captureToken }
  }, runtime);
});

test("rejects geometry dispatch without Desktop 1.1 before Excel.run", async () => {
  await assert.rejects(
    executeOfficeAction({
      sessionId: "session-1",
      workbookUrl: "file:///tmp/owned.xlsx",
      requestId: "request-prepare",
      action: "screenshot.prepare-range-geometry",
      payload: { sheetName: "Data", rangeAddress: "B2:D8" }
    }, {
      requirementSets: { excelApi: ["1.21"], excelApiDesktop: [] },
      run: async () => assert.fail("Excel.run must not start without Desktop 1.1")
    }),
    /requires ExcelApiDesktop 1\.1/
  );
});

test("rejects geometry when the application active window belongs to another workbook", async () => {
  const fixture = createGeometryContext();
  fixture.window.activeWorksheet = {
    id: "foreign-sheet",
    name: "Foreign",
    load() {}
  };

  await assert.rejects(
    executeOfficeAction({
      sessionId: "session-foreign",
      workbookUrl: "file:///tmp/owned.xlsx",
      requestId: "request-foreign",
      action: "screenshot.prepare-range-geometry",
      payload: { sheetName: "Data", rangeAddress: "B2:D8" }
    }, {
      requirementSets: { excelApi: ["1.21"], excelApiDesktop: ["1.1"] },
      async run(callback) {
        return callback(fixture.context);
      }
    }),
    /active Excel window does not belong to this exact workbook/
  );
  assert.deepEqual(fixture.scrollCalls, []);
});

test("rejects geometry that cannot satisfy the native integral pixel schema", async () => {
  const fixture = createGeometryContext();
  fixture.window.pointsToScreenPixelsX = (points) => ({ value: points * 2 + 0.5 });

  await assert.rejects(
    executeOfficeAction({
      sessionId: "session-fractional",
      workbookUrl: "file:///tmp/owned.xlsx",
      requestId: "request-fractional",
      action: "screenshot.prepare-range-geometry",
      payload: { sheetName: "Data", rangeAddress: "B2:D8" }
    }, {
      requirementSets: { excelApi: ["1.21"], excelApiDesktop: ["1.1"] },
      async run(callback) {
        return callback(fixture.context);
      }
    }),
    /screen rectangle edges must be 32-bit integers/
  );
});

test("rejects a non-positive Desktop window number", async () => {
  const fixture = createGeometryContext();
  fixture.window.windowNumber = 0;

  await assert.rejects(
    executeOfficeAction({
      sessionId: "session-window-number",
      workbookUrl: "file:///tmp/owned.xlsx",
      requestId: "request-window-number",
      action: "screenshot.prepare-range-geometry",
      payload: { sheetName: "Data", rangeAddress: "B2:D8" }
    }, {
      requirementSets: { excelApi: ["1.21"], excelApiDesktop: ["1.1"] },
      async run(callback) {
        return callback(fixture.context);
      }
    }),
    /window number must be a positive 32-bit integer/
  );
});

test("routes a table mutation through Excel.run and returns contract-shaped result", async () => {
  let syncCount = 0;
  const table = {};
  const context = {
    workbook: {
      tables: {
        add(address, hasHeaders) {
          assert.equal(address, "'Data'!A1:B3");
          assert.equal(hasHeaders, true);
          return table;
        }
      }
    },
    async sync() {
      syncCount++;
    }
  };
  const runtime = {
    async run(callback) {
      return callback(context);
    }
  };

  const result = await executeOfficeAction({
    action: "table.create",
    payload: {
      sheetName: "Data",
      tableName: "Sales",
      rangeAddress: "A1:B3",
      hasHeaders: true,
      tableStyle: "TableStyleMedium2"
    }
  }, runtime);

  assert.equal(table.name, "Sales");
  assert.equal(table.style, "TableStyleMedium2");
  assert.deepEqual(result, {
    success: true,
    errorMessage: null,
    action: "create",
    message: "Table 'Sales' created."
  });
  assert.equal(syncCount, 1);
});

test("rejects an action when the active workbook lacks its required API", async () => {
  const action = getImplementedAction("conditionalformat.add-rule");
  assert.equal(action.requirementSet, "1.6");

  await assert.rejects(
    executeOfficeAction({
      action: "conditionalformat.add-rule",
      payload: {
        sheetName: "Data",
        rangeAddress: "A1:A3",
        ruleType: "cellValue",
        operatorType: "greater",
        formula1: "10"
      }
    }, {
      requirementSets: { excelApi: ["1.5"], excelApiDesktop: [] },
      run: async () => assert.fail("Excel.run must not start without the required API")
    }),
    /requires ExcelApi 1\.6/
  );
});

test("serializes dispatched Excel.run actions", async () => {
  let active = 0;
  let maximum = 0;
  let releaseFirst;
  const firstBlocked = new Promise((resolve) => {
    releaseFirst = resolve;
  });
  let call = 0;
  const runtime = {
    async run(callback) {
      call++;
      active++;
      maximum = Math.max(maximum, active);
      if (call === 1) {
        await firstBlocked;
      }
      try {
        return await callback({
          workbook: { worksheets: createWorksheetCollection() },
          sync: async () => {}
        });
      } finally {
        active--;
      }
    }
  };

  const first = executeOfficeAction({
    action: "sheet.copy",
    payload: { sourceName: "Data", targetName: "Data Copy" }
  }, runtime);
  const second = executeOfficeAction({
    action: "sheet.move",
    payload: { sheetName: "Data", afterSheet: "Summary" }
  }, runtime);
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(maximum, 1);
  releaseFirst();
  await Promise.all([first, second]);
  assert.equal(maximum, 1);
});

function createWorksheetCollection() {
  const sheets = new Map([
    ["Data", { name: "Data", position: 0 }],
    ["Summary", { name: "Summary", position: 1 }]
  ]);
  return {
    getItem(name) {
      const sheet = sheets.get(name);
      if (!sheet) {
        throw new Error(`Missing sheet ${name}`);
      }
      sheet.load = () => {};
      sheet.copy = () => {
        const copy = { name: "", position: sheets.size, load() {} };
        return copy;
      };
      return sheet;
    },
    load() {},
    items: [...sheets.values()]
  };
}

function createGeometryContext() {
  const scrollCalls = [];
  const ranges = new Map();
  let activeSheet = "Summary";
  let selectedRange = "Summary!C3:E4";
  const window = {
    index: 0,
    windowNumber: 42,
    type: "workbook",
    windowState: "normal",
    left: 10,
    top: 20,
    width: 1200,
    height: 900,
    usableWidth: 1100,
    usableHeight: 760,
    scrollRow: 7,
    scrollColumn: 3,
    zoom: 125,
    view: "normalView",
    freezePanes: true,
    split: false,
    splitRow: 1,
    splitColumn: 0,
    splitHorizontal: 18,
    splitVertical: 0,
    load() {},
    scrollIntoView(left, top, width, height, start) {
      scrollCalls.push({ left, top, width, height, start });
    },
    pointsToScreenPixelsX(points) {
      return { value: points * 2 + 20 };
    },
    pointsToScreenPixelsY(points) {
      return { value: points * 2 + 20 };
    }
  };

  function range(address, sheet) {
    const key = `${sheet.name}!${address}`;
    if (!ranges.has(key)) {
      ranges.set(key, {
        address: key,
        left: 100,
        top: 150,
        width: 300,
        height: 400,
        worksheet: sheet,
        load() {},
        select() {
          selectedRange = key;
        }
      });
    }
    return ranges.get(key);
  }

  const summary = {
    id: "sheet-summary",
    name: "Summary",
    load() {},
    activate() {
      activeSheet = this.name;
    },
    getRange(address) {
      return range(address.replace(/^Summary!/, ""), this);
    },
    getUsedRange() {
      return range("A1:M30", this);
    }
  };
  const data = {
    id: "sheet-data",
    name: "Data",
    load() {},
    activate() {
      activeSheet = this.name;
    },
    getRange(address) {
      return range(address.replace(/^Data!/, ""), this);
    },
    getUsedRange() {
      return range("A1:H40", this);
    }
  };
  const sheets = new Map([
    ["Summary", summary],
    ["sheet-summary", summary],
    ["Data", data],
    ["sheet-data", data]
  ]);
  const selected = range("C3:E4", summary);
  const activeCell = range("D3", summary);
  const visible = range("A1:M30", summary);
  window.activeWorksheet = summary;
  window.activeCell = activeCell;
  window.visibleRange = visible;

  const context = {
    workbook: {
      application: { activeWindow: window },
      worksheets: {
        getActiveWorksheet() {
          return sheets.get(activeSheet);
        },
        getItem(nameOrId) {
          const sheet = sheets.get(nameOrId);
          if (!sheet) throw new Error(`Missing sheet ${nameOrId}`);
          return sheet;
        }
      },
      getSelectedRange() {
        return selectedRange === selected.address
          ? selected
          : ranges.get(selectedRange);
      }
    },
    async sync() {}
  };

  return {
    context,
    window,
    scrollCalls,
    get activeSheet() {
      return activeSheet;
    },
    get selectedRange() {
      return selectedRange;
    }
  };
}
