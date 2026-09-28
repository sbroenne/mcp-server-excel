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
    "../../src/ExcelMcp.Core/Commands/Sheet/ISheetCommands.cs",
    "../../src/ExcelMcp.Core/Commands/Chart/IChartCommands.cs",
    "../../src/ExcelMcp.Core/Commands/Chart/IChartConfigCommands.cs",
    "../../src/ExcelMcp.Core/Commands/PivotTable/IPivotTableCommands.cs",
    "../../src/ExcelMcp.Core/Commands/PivotTable/IPivotTableFieldCommands.cs",
    "../../src/ExcelMcp.Core/Commands/PivotTable/IPivotTableCalcCommands.cs",
    "../../src/ExcelMcp.Core/Commands/Slicer/ISlicerCommands.cs"
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

test("rejects a single capture rectangle that is not contained in the Excel window", async () => {
  const fixture = createGeometryContext();
  fixture.window.width = 100;

  await assert.rejects(
    executeOfficeAction({
      sessionId: "session-outside",
      workbookUrl: "file:///tmp/owned.xlsx",
      requestId: "request-outside",
      action: "screenshot.prepare-range-geometry",
      payload: { sheetName: "Data", rangeAddress: "B2:D8" }
    }, {
      requirementSets: { excelApi: ["1.21"], excelApiDesktop: ["1.1"] },
      async run(callback) {
        return callback(fixture.context);
      }
    }),
    /does not fit within the Excel window/
  );
});

test("caps sheet geometry to the top-left 500 rows and 50 columns", async () => {
  const fixture = createGeometryContext();
  fixture.usedRange.rowIndex = 0;
  fixture.usedRange.columnIndex = 0;
  fixture.usedRange.rowCount = 800;
  fixture.usedRange.columnCount = 80;

  const prepared = await executeOfficeAction({
    sessionId: "session-sheet-cap",
    workbookUrl: "file:///tmp/owned.xlsx",
    requestId: "request-sheet-cap",
    action: "screenshot.prepare-sheet-geometry",
    payload: { sheetName: "Data" }
  }, {
    requirementSets: { excelApi: ["1.21"], excelApiDesktop: ["1.1"] },
    async run(callback) {
      return callback(fixture.context);
    }
  });

  assert.deepEqual(fixture.indexedRangeCalls, [{
    rowIndex: 0,
    columnIndex: 0,
    rowCount: 500,
    columnCount: 50
  }]);
  assert.deepEqual(prepared.range, { address: "Data!A1:AX500" });
  assert.equal(prepared.truncated, true);
  assert.match(prepared.message, /top-left 500 rows and 50 columns/);
  await executeOfficeAction({
    sessionId: "session-sheet-cap",
    workbookUrl: "file:///tmp/owned.xlsx",
    requestId: "request-sheet-cap-restore",
    action: "screenshot.restore-view",
    payload: { captureToken: prepared.captureToken }
  }, {
    requirementSets: { excelApi: ["1.21"], excelApiDesktop: ["1.1"] },
    async run(callback) {
      return callback(fixture.context);
    }
  });
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

test("keeps inexact PivotChart, OLAP, grouping, and slicer discovery routes disabled", () => {
  for (const action of [
    "chart.create-from-pivottable",
    "chart.list",
    "pivottable.create-from-datamodel",
    "pivottable.get-cache-options",
    "pivottablefield.group-by-date",
    "pivottablefield.group-items",
    "slicer.list-slicers",
    "slicer.set-slicer-selection"
  ]) {
    assert.equal(getImplementedAction(action), null, action);
  }
});

test("creates a regular chart with target-range precedence and exact result shape", async () => {
  const fixture = createChartContext();
  const result = await executeOfficeAction({
    action: "chart.create-from-range",
    payload: {
      sheetName: "Data",
      sourceRangeAddress: "A1:B4",
      chartType: "Column3DClustered",
      left: 900,
      top: 700,
      width: 500,
      height: 350,
      chartName: "Sales Chart",
      targetRange: "D2:H14"
    }
  }, {
    requirementSets: { excelApi: ["1.21"], excelApiDesktop: [] },
    async run(callback) {
      return callback(fixture.context);
    }
  });

  assert.deepEqual(fixture.addCalls, [{
    type: "3DColumnClustered",
    sourceAddress: "Data!A1:B4",
    seriesBy: "Auto"
  }]);
  assert.deepEqual(fixture.positionCalls, [{
    start: "Data!D2:H14:first",
    end: "Data!D2:H14:last"
  }]);
  assert.deepEqual(result, {
    success: true,
    errorMessage: null,
    action: "create",
    message: "IMPORTANT: You MUST take a screenshot(capture-sheet) to verify the chart does not overlap the data.",
    chartName: "Sales Chart",
    sheetName: "Data",
    chartType: "Column3DClustered",
    isPivotChart: false,
    linkedPivotTable: null,
    left: 40,
    top: 60,
    width: 320,
    height: 240
  });
});

test("uses Core automatic chart padding and reports data collisions", async () => {
  const automatic = createChartContext();
  const runtime = (fixture) => ({
    requirementSets: { excelApi: ["1.21"], excelApiDesktop: [] },
    async run(callback) {
      return callback(fixture.context);
    }
  });

  const placed = await executeOfficeAction({
    action: "chart.create-from-range",
    payload: {
      sheetName: "Data",
      sourceRangeAddress: "A1:B4",
      chartType: "ColumnClustered",
      left: 0,
      top: 0,
      width: 400,
      height: 300,
      chartName: "Automatic"
    }
  }, runtime(automatic));

  assert.equal(placed.left, 10);
  assert.equal(placed.top, 50);

  const colliding = createChartContext();
  colliding.usedRange.height = 100;
  const result = await executeOfficeAction({
    action: "chart.create-from-range",
    payload: {
      sheetName: "Data",
      sourceRangeAddress: "A1:B4",
      chartType: "ColumnClustered",
      chartName: "Collision",
      targetRange: "D2:H14"
    }
  }, runtime(colliding));

  assert.equal(
    result.message,
    "OVERLAP WARNING: Chart overlaps data area Data!A1:B4. Use chart move or fit-to-range to reposition, then screenshot(capture-sheet) to verify layout."
  );
});

test("routes chart configuration with Core enum mappings and 1-based series indexes", async () => {
  const fixture = createChartContext();
  const runtime = {
    requirementSets: { excelApi: ["1.21"], excelApiDesktop: [] },
    async run(callback) {
      return callback(fixture.context);
    }
  };

  await executeOfficeAction({
    action: "chartconfig.set-axis-scale",
    payload: {
      chartName: "Existing",
      axis: "ValueSecondary",
      minimumScale: null,
      maximumScale: 100,
      majorUnit: 10,
      minorUnit: null
    }
  }, runtime);
  await executeOfficeAction({
    action: "chartconfig.set-series-chart-type",
    payload: { chartName: "Existing", seriesIndex: 2, chartType: "LineMarkers" }
  }, runtime);
  const plot = await executeOfficeAction({
    action: "chartconfig.get-plot-options",
    payload: { chartName: "Existing" }
  }, runtime);

  assert.deepEqual(fixture.axisCalls.at(-1), ["Value", "Secondary"]);
  assert.equal(fixture.chart.axes.lastAxis.minimum, "");
  assert.equal(fixture.chart.axes.lastAxis.maximum, 100);
  assert.equal(fixture.chart.axes.lastAxis.majorUnit, 10);
  assert.equal(fixture.chart.axes.lastAxis.minorUnit, "");
  assert.deepEqual(fixture.seriesCalls, [1]);
  assert.equal(fixture.chart.seriesItems[1].chartType, "LineMarkers");
  assert.deepEqual(plot, {
    success: true,
    errorMessage: null,
    chartName: "Existing",
    plotBy: "Columns",
    displayBlanksAs: "Gaps",
    plotVisibleOnly: true
  });
});

test("adds trendlines with 1-based indexes and exact shared result shape", async () => {
  const fixture = createChartContext();
  const result = await executeOfficeAction({
    action: "chartconfig.add-trendline",
    payload: {
      chartName: "Existing",
      seriesIndex: 1,
      trendlineType: "Polynomial",
      order: 3,
      forward: 2,
      displayEquation: true,
      displayRSquared: false,
      name: "Forecast"
    }
  }, {
    requirementSets: { excelApi: ["1.21"], excelApiDesktop: [] },
    async run(callback) {
      return callback(fixture.context);
    }
  });

  assert.deepEqual(result, {
    success: true,
    errorMessage: null,
    action: "add-trendline",
    message: "Trendline added to series 1 of 'Existing'.",
    chartName: "Existing",
    seriesIndex: 1,
    trendlineIndex: 1,
    type: "Polynomial",
    name: "Forecast"
  });
  assert.equal(fixture.trendline.polynomialOrder, 3);
  assert.equal(fixture.trendline.forwardPeriod, 2);
  assert.equal(fixture.trendline.showEquation, true);
  assert.equal(fixture.trendline.showRSquared, false);
});

test("creates ordinary range PivotTables with exact shared DTO fields", async () => {
  const fixture = createPivotContext();
  const result = await executeOfficeAction({
    action: "pivottable.create-from-range",
    payload: {
      sourceSheet: "Source",
      sourceRange: "A1:C4",
      destinationSheet: "Report",
      destinationCell: "E3",
      pivotTableName: "SalesPivot"
    }
  }, {
    requirementSets: { excelApi: ["1.21"], excelApiDesktop: [] },
    async run(callback) {
      return callback(fixture.context);
    }
  });

  assert.deepEqual(fixture.addCalls, [{
    name: "SalesPivot",
    source: "Source!A1:C4",
    destination: "Report!E3"
  }]);
  assert.deepEqual(result, {
    success: true,
    errorMessage: null,
    pivotTableName: "SalesPivot",
    sheetName: "Report",
    range: "Report!E3:G6",
    sourceData: "Source!A1:C4",
    sourceRowCount: 3,
    availableFields: ["Region", "Sales", "Units"]
  });
});

test("rejects OLAP PivotTable mutations instead of treating metadata availability as parity", async () => {
  const fixture = createPivotContext("Unknown");
  await assert.rejects(
    executeOfficeAction({
      action: "pivottable.delete",
      payload: { pivotTableName: "ModelPivot" }
    }, {
      requirementSets: { excelApi: ["1.21"], excelApiDesktop: [] },
      async run(callback) {
        return callback(fixture.context);
      }
    }),
    /OLAP and Power Pivot require the trusted VBA capability/
  );
  assert.equal(fixture.pivot.deleted, false);
});

test("routes faithful ordinary PivotTable field and layout actions", async () => {
  const fixture = createOrdinaryPivotFieldContext();
  const runtime = {
    requirementSets: { excelApi: ["1.21"], excelApiDesktop: [] },
    async run(callback) {
      return callback(fixture.context);
    }
  };

  const filtered = await executeOfficeAction({
    action: "pivottablefield.set-field-filter",
    payload: {
      pivotTableName: "SalesPivot",
      fieldName: "Region",
      selectedValues: ["North"]
    }
  }, runtime);
  const sorted = await executeOfficeAction({
    action: "pivottablefield.sort-field",
    payload: {
      pivotTableName: "SalesPivot",
      fieldName: "Region",
      direction: "Descending"
    }
  }, runtime);
  const formatted = await executeOfficeAction({
    action: "pivottablefield.set-field-format",
    payload: {
      pivotTableName: "SalesPivot",
      fieldName: "Sales",
      numberFormat: "$#,##0.00"
    }
  }, runtime);
  const renamed = await executeOfficeAction({
    action: "pivottablefield.set-field-name",
    payload: {
      pivotTableName: "SalesPivot",
      fieldName: "Region",
      customName: "Sales Region"
    }
  }, runtime);
  const subtotals = await executeOfficeAction({
    action: "pivottablecalc.set-subtotals",
    payload: {
      pivotTableName: "SalesPivot",
      fieldName: "Region",
      showSubtotals: false
    }
  }, runtime);
  await executeOfficeAction({
    action: "pivottablecalc.set-layout",
    payload: { pivotTableName: "SalesPivot", rowLayout: 1 }
  }, runtime);
  await executeOfficeAction({
    action: "pivottablecalc.set-grand-totals",
    payload: {
      pivotTableName: "SalesPivot",
      showRowGrandTotals: false,
      showColumnGrandTotals: true
    }
  }, runtime);

  assert.deepEqual(fixture.regionField.appliedFilter, {
    manualFilter: { selectedItems: ["North"] }
  });
  assert.equal(fixture.regionField.sortedBy, "Descending");
  assert.equal(fixture.salesHierarchy.numberFormat, "$#,##0.00");
  assert.equal(fixture.regionHierarchy.name, "Sales Region");
  assert.deepEqual(fixture.regionField.subtotals, { automatic: false });
  assert.equal(fixture.pivot.layout.layoutType, "Tabular");
  assert.equal(fixture.pivot.layout.showRowGrandTotals, false);
  assert.equal(fixture.pivot.layout.showColumnGrandTotals, true);
  assert.deepEqual(filtered, {
    success: true,
    errorMessage: null,
    fieldName: "Region",
    selectedItems: ["North"],
    availableItems: ["North", "South"],
    visibleRowCount: 3,
    totalRowCount: 5,
    showAll: false
  });
  assert.deepEqual(sorted, {
    success: true,
    errorMessage: null,
    fieldName: "Region",
    customName: "Region",
    area: "Row",
    position: 1,
    availableValues: [],
    dataType: ""
  });
  assert.equal(formatted.numberFormat, "$#,##0.00");
  assert.equal(renamed.customName, "Sales Region");
  assert.deepEqual(subtotals, {
    success: true,
    errorMessage: null,
    fieldName: "Region",
    customName: "",
    area: "Hidden",
    position: 0,
    availableValues: [],
    dataType: "",
    workflowHint: "Subtotals disabled for field. Only detail rows visible."
  });
});

test("reads ordinary PivotTable data and removes placed hierarchies", async () => {
  const fixture = createOrdinaryPivotFieldContext();
  const runtime = {
    requirementSets: { excelApi: ["1.21"], excelApiDesktop: [] },
    async run(callback) {
      return callback(fixture.context);
    }
  };
  const data = await executeOfficeAction({
    action: "pivottablecalc.get-data",
    payload: { pivotTableName: "SalesPivot" }
  }, runtime);
  const removed = await executeOfficeAction({
    action: "pivottablefield.remove-field",
    payload: { pivotTableName: "SalesPivot", fieldName: "Region" }
  }, runtime);

  assert.deepEqual(data, {
    success: true,
    errorMessage: null,
    pivotTableName: "SalesPivot",
    values: [
      ["Region", "Sum of Sales"],
      ["North", 40],
      ["South", 20]
    ],
    columnHeaders: [],
    rowHeaders: [],
    dataRowCount: 3,
    dataColumnCount: 2,
    grandTotals: {}
  });
  assert.deepEqual(fixture.pivot.rowHierarchies.removed, [fixture.regionHierarchy]);
  assert.deepEqual(removed, {
    success: true,
    errorMessage: null,
    fieldName: "Region",
    customName: "",
    area: "Hidden",
    position: 0,
    availableValues: [],
    dataType: ""
  });
});

test("ordinary PivotTable field routes fail closed on source, item, and layout mismatches", async () => {
  const olap = createOrdinaryPivotFieldContext("Unknown");
  await assert.rejects(
    executeOfficeAction({
      action: "pivottablefield.sort-field",
      payload: {
        pivotTableName: "SalesPivot",
        fieldName: "Region",
        direction: "Ascending"
      }
    }, {
      requirementSets: { excelApi: ["1.21"], excelApiDesktop: [] },
      async run(callback) {
        return callback(olap.context);
      }
    }),
    /OLAP and Power Pivot require the trusted VBA capability/
  );

  const fixture = createOrdinaryPivotFieldContext();
  const runtime = {
    requirementSets: { excelApi: ["1.21"], excelApiDesktop: [] },
    async run(callback) {
      return callback(fixture.context);
    }
  };
  await assert.rejects(
    executeOfficeAction({
      action: "pivottablefield.set-field-filter",
      payload: {
        pivotTableName: "SalesPivot",
        fieldName: "Region",
        selectedValues: ["Missing"]
      }
    }, runtime),
    /does not contain selected item/
  );
  await assert.rejects(
    executeOfficeAction({
      action: "pivottablecalc.set-layout",
      payload: { pivotTableName: "SalesPivot", rowLayout: 3 }
    }, runtime),
    /rowLayout must be 0/
  );
});

test("keeps unrepresentable PivotTable field placement and aggregation actions disabled", () => {
  for (const action of [
    "pivottablefield.list-fields",
    "pivottablefield.add-row-field",
    "pivottablefield.add-column-field",
    "pivottablefield.add-filter-field",
    "pivottablefield.add-value-field",
    "pivottablefield.set-field-function"
  ]) {
    assert.equal(getImplementedAction(action), null, action);
  }
});

test("creates a table slicer with exact known source identity and item captions", async () => {
  const fixture = createPivotContext();
  const result = await executeOfficeAction({
    action: "slicer.create-table-slicer",
    payload: {
      tableName: "Sales",
      columnName: "Region",
      slicerName: "RegionSlicer",
      destinationSheet: "Report",
      position: "J2"
    }
  }, {
    requirementSets: { excelApi: ["1.21"], excelApiDesktop: [] },
    async run(callback) {
      return callback(fixture.context);
    }
  });

  assert.deepEqual(result, {
    success: true,
    errorMessage: null,
    name: "RegionSlicer",
    caption: "RegionSlicer",
    fieldName: "Region",
    sheetName: "Report",
    position: "J2",
    selectedItems: ["North"],
    availableItems: ["North", "South"],
    connectedPivotTables: [],
    connectedTable: "Sales",
    sourceType: "Table",
    workflowHint: "Slicer 'RegionSlicer' created for column 'Region' in table 'Sales'. Use SetTableSlicerSelection to filter data."
  });
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

function createChartContext() {
  const addCalls = [];
  const positionCalls = [];
  const axisCalls = [];
  const seriesCalls = [];
  const trendline = {
    name: "",
    type: "",
    load() {}
  };
  const trendlines = {
    items: [],
    add(type) {
      trendline.type = type;
      this.items.push(trendline);
      return trendline;
    },
    getCount() {
      return { value: this.items.length };
    },
    getItem(index) {
      return this.items[index];
    }
  };
  const seriesItems = [
    { chartType: "ColumnClustered", trendlines },
    { chartType: "ColumnClustered", trendlines }
  ];
  const axes = {
    lastAxis: null,
    getItem(type, group) {
      axisCalls.push([type, group]);
      const axis = {
        minimum: 0,
        maximum: 0,
        majorUnit: 0,
        minorUnit: 0,
        title: {},
        majorGridlines: { visible: true, load() {} },
        minorGridlines: { visible: false, load() {} },
        load() {}
      };
      this.lastAxis = axis;
      return axis;
    }
  };
  const chart = {
    name: "Existing",
    chartType: "ColumnClustered",
    left: 10,
    top: 20,
    width: 400,
    height: 300,
    plotBy: "Columns",
    displayBlanksAs: "NotPlotted",
    plotVisibleOnly: true,
    isNullObject: false,
    title: {},
    legend: {},
    dataLabels: {},
    axes,
    seriesItems,
    series: {
      getItemAt(index) {
        seriesCalls.push(index);
        return seriesItems[index];
      }
    },
    load() {},
    setPosition(start, end) {
      positionCalls.push({ start: start.address, end: end.address });
      this.left = 40;
      this.top = 60;
      this.width = 320;
      this.height = 240;
    },
    delete() {}
  };
  const nullChart = {
    isNullObject: true,
    load() {}
  };
  const ranges = new Map();
  const range = (address) => {
    if (!ranges.has(address)) {
      ranges.set(address, {
        address,
        getCell() {
          return { address: `${address}:first` };
        },
        getLastCell() {
          return { address: `${address}:last` };
        }
      });
    }
    return ranges.get(address);
  };
  const charts = {
    items: [chart],
    add(type, source, seriesBy) {
      addCalls.push({ type, sourceAddress: source.address, seriesBy });
      chart.chartType = type;
      return chart;
    },
    getItemOrNullObject(name) {
      return name === chart.name ? chart : nullChart;
    },
    load() {}
  };
  const usedRange = {
    address: "Data!A1:B4",
    left: 0,
    top: 0,
    width: 100,
    height: 40,
    load() {}
  };
  const sheet = {
    name: "Data",
    charts,
    load() {},
    getRange(address) {
      return range(`Data!${address}`);
    },
    getUsedRange() {
      return usedRange;
    }
  };
  const context = {
    workbook: {
      worksheets: {
        items: [sheet],
        load() {},
        getItem() {
          return sheet;
        }
      },
      tables: {
        getItem() {
          return { getRange: () => range("Data!A1:B4") };
        }
      }
    },
    async sync() {}
  };
  return {
    context,
    chart,
    trendline,
    usedRange,
    addCalls,
    positionCalls,
    axisCalls,
    seriesCalls
  };
}

function createPivotContext(sourceType = "LocalRange") {
  const addCalls = [];
  const sourceRange = {
    address: "Source!A1:C4",
    rowCount: 4,
    values: [
      ["Region", "Sales", "Units"],
      ["North", 10, 1],
      ["South", 20, 2],
      ["North", 30, 3]
    ],
    load() {}
  };
  const destination = { address: "Report!E3" };
  const position = { address: "Report!J2", left: 500, top: 30, load() {} };
  const pivotRange = { address: "Report!E3:G6", load() {} };
  const pivot = {
    deleted: false,
    layout: { getRange: () => pivotRange },
    getDataSourceType() {
      return { value: sourceType };
    },
    delete() {
      this.deleted = true;
    }
  };
  const slicer = {
    name: "",
    caption: "",
    left: 0,
    top: 0,
    load() {},
    getSelectedItems() {
      return { value: ["north-key"] };
    },
    slicerItems: {
      items: [
        { key: "north-key", name: "North" },
        { key: "south-key", name: "South" }
      ],
      load() {}
    }
  };
  const sourceSheet = {
    name: "Source",
    load() {},
    getRange() {
      return sourceRange;
    }
  };
  const reportSheet = {
    name: "Report",
    load() {},
    getRange(address) {
      return address === "J2" ? position : destination;
    }
  };
  const tables = {
    Sales: {
      name: "Sales",
      getRange() {
        return sourceRange;
      }
    }
  };
  const context = {
    workbook: {
      worksheets: {
        getItem(name) {
          return name === "Source" ? sourceSheet : reportSheet;
        }
      },
      tables: {
        getItem(name) {
          return tables[name];
        }
      },
      pivotTables: {
        add(name, source, target) {
          addCalls.push({ name, source: source.address, destination: target.address });
          return pivot;
        },
        getItem() {
          return pivot;
        }
      },
      slicers: {
        add(source, field, sheet) {
          assert.equal(source, tables.Sales);
          assert.equal(field, "Region");
          assert.equal(sheet, reportSheet);
          return slicer;
        }
      }
    },
    async sync() {}
  };
  return { context, pivot, addCalls };
}

function createOrdinaryPivotFieldContext(sourceType = "LocalRange") {
  const regionField = {
    name: "Region",
    subtotals: { automatic: true },
    items: {
      items: [{ name: "North" }, { name: "South" }],
      load() {}
    },
    load() {},
    applyFilter(filter) {
      this.appliedFilter = filter;
    },
    sortByLabels(direction) {
      this.sortedBy = direction;
    }
  };
  const salesField = {
    name: "Sales",
    items: { items: [], load() {} },
    load() {}
  };
  const regionHierarchy = {
    name: "Region",
    position: 0,
    fields: {
      items: [regionField],
      getItem() {
        return regionField;
      },
      load() {}
    },
    load() {}
  };
  const salesHierarchy = {
    name: "Sum of Sales",
    position: 0,
    numberFormat: "General",
    summarizeBy: "Sum",
    field: salesField,
    load() {}
  };
  const emptyCollection = () => ({
    items: [],
    removed: [],
    load() {},
    remove(item) {
      this.removed.push(item);
      this.items = this.items.filter((candidate) => candidate !== item);
    }
  });
  const rowHierarchies = emptyCollection();
  rowHierarchies.items = [regionHierarchy];
  const columnHierarchies = emptyCollection();
  const filterHierarchies = emptyCollection();
  const dataHierarchies = emptyCollection();
  dataHierarchies.items = [salesHierarchy];
  const fullRange = {
    values: [
      ["Region", "Sum of Sales"],
      ["North", 40],
      ["South", 20]
    ],
    rowCount: 5,
    columnCount: 2,
    load() {}
  };
  const pivot = {
    rowHierarchies,
    columnHierarchies,
    filterHierarchies,
    dataHierarchies,
    layout: {
      layoutType: "Compact",
      showRowGrandTotals: true,
      showColumnGrandTotals: true,
      getRange() {
        return fullRange;
      },
      load() {}
    },
    getDataSourceType() {
      return { value: sourceType };
    }
  };
  const context = {
    workbook: {
      pivotTables: {
        getItem(name) {
          assert.equal(name, "SalesPivot");
          return pivot;
        }
      }
    },
    async sync() {
      if (regionField.appliedFilter) {
        fullRange.rowCount = 3;
      }
    }
  };
  return {
    context,
    pivot,
    regionField,
    salesField,
    regionHierarchy,
    salesHierarchy
  };
}

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
  const indexedRangeCalls = [];
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
    },
    getRangeByIndexes(rowIndex, columnIndex, rowCount, columnCount) {
      indexedRangeCalls.push({ rowIndex, columnIndex, rowCount, columnCount });
      return range("A1:AX500", this);
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
  const usedRange = data.getUsedRange();
  usedRange.rowIndex = 0;
  usedRange.columnIndex = 0;
  usedRange.rowCount = 40;
  usedRange.columnCount = 8;
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
    indexedRangeCalls,
    usedRange,
    get activeSheet() {
      return activeSheet;
    },
    get selectedRange() {
      return selectedRange;
    }
  };
}
