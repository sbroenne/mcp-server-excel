import { EXCEL_API_REQUIREMENT_SETS, EXCEL_API_DESKTOP_REQUIREMENT_SETS } from "./constants.mjs";

const actions = new Map();
const publicActions = [];
const internalActions = [];
const captureStates = new Map();
const captureTokenBySession = new Map();
let dispatchTail = Promise.resolve();

registerPublic("table.list", "1.1", false, listTables);
registerPublic("table.create", "1.1", true, createTable);
registerPublic("table.rename", "1.1", true, renameTable);
registerPublic("table.delete", "1.2", true, deleteTable);
registerPublic("table.read", "1.1", false, readTable);
registerPublic("table.resize", "1.13", true, resizeTable);
registerPublic("table.toggle-totals", "1.1", true, toggleTotals);
registerPublic("table.set-column-total", "1.1", true, setColumnTotal);
registerPublic("table.append", "1.4", true, appendRows);
registerPublic("table.get-data", "1.3", false, getTableData);
registerPublic("table.set-style", "1.1", true, setTableStyle);
registerPublic("tablecolumn.apply-filter", "1.2", true, applyFilter);
registerPublic("tablecolumn.apply-filter-values", "1.2", true, applyFilterValues);
registerPublic("tablecolumn.clear-filters", "1.2", true, clearFilters);
registerPublic("tablecolumn.get-filters", "1.2", false, getFilters);
registerPublic("tablecolumn.add-column", "1.4", true, addColumn);
registerPublic("tablecolumn.remove-column", "1.1", true, removeColumn);
registerPublic("tablecolumn.rename-column", "1.4", true, renameColumn);
registerPublic("tablecolumn.get-structured-reference", "1.1", false, getStructuredReference);
registerPublic("tablecolumn.sort", "1.2", true, sortTable);
registerPublic("tablecolumn.sort-multi", "1.2", true, sortTableMulti);
registerPublic("tablecolumn.get-column-number-format", "1.1", false, getColumnNumberFormat);
registerPublic("tablecolumn.set-column-number-format", "1.1", true, setColumnNumberFormat);
registerPublic("conditionalformat.add-rule", "1.6", true, addConditionalRule);
registerPublic("conditionalformat.clear-rules", "1.6", true, clearConditionalRules);
registerPublic("conditionalformat.list-rules", "1.6", false, listConditionalRules);
registerPublic("conditionalformat.list-worksheet-rules", "1.6", false, listWorksheetConditionalRules);
registerPublic("sheet.copy", "1.7", true, copyWorksheet);
registerPublic("sheet.move", "1.1", true, moveWorksheet);
registerPublic("chart.create-from-range", "1.1", true, createChartFromRange);
registerPublic("chart.create-from-table", "1.1", true, createChartFromTable);
registerPublic("chart.delete", "1.1", true, deleteChart);
registerPublic("chart.move", "1.1", true, moveChart);
registerPublic("chart.fit-to-range", "1.1", true, fitChartToRange);
registerPublic("chartconfig.set-chart-type", "1.7", true, setChartType);
registerPublic("chartconfig.set-title", "1.1", true, setChartTitle);
registerPublic("chartconfig.set-axis-title", "1.7", true, setChartAxisTitle);
registerPublic("chartconfig.get-axis-number-format", "1.8", false, getChartAxisNumberFormat);
registerPublic("chartconfig.set-axis-number-format", "1.8", true, setChartAxisNumberFormat);
registerPublic("chartconfig.show-legend", "1.1", true, showChartLegend);
registerPublic("chartconfig.set-style", "1.8", true, setChartStyle);
registerPublic("chartconfig.set-data-labels", "1.8", true, setChartDataLabels);
registerPublic("chartconfig.set-axis-scale", "1.7", true, setChartAxisScale);
registerPublic("chartconfig.get-gridlines", "1.7", false, getChartGridlines);
registerPublic("chartconfig.set-gridlines", "1.7", true, setChartGridlines);
registerPublic("chartconfig.set-series-chart-type", "1.7", true, setSeriesChartType);
registerPublic("chartconfig.get-plot-options", "1.8", false, getChartPlotOptions);
registerPublic("chartconfig.set-plot-options", "1.8", true, setChartPlotOptions);
registerPublic("chartconfig.add-trendline", "1.8", true, addChartTrendline);
registerPublic("chartconfig.delete-trendline", "1.7", true, deleteChartTrendline);
registerPublic("chartconfig.set-trendline", "1.8", true, setChartTrendline);
registerPublic("pivottable.create-from-range", "1.8", true, createPivotTableFromRange);
registerPublic("pivottable.create-from-table", "1.8", true, createPivotTableFromTable);
registerPublic("pivottable.delete", "1.15", true, deletePivotTable);
registerPublic("pivottablefield.remove-field", "1.15", true, removePivotField);
registerPublic("pivottablefield.set-field-name", "1.15", true, setPivotFieldName);
registerPublic("pivottablefield.set-field-format", "1.15", true, setPivotFieldFormat);
registerPublic("pivottablefield.set-field-filter", "1.15", true, setPivotFieldFilter);
registerPublic("pivottablefield.sort-field", "1.15", true, sortPivotField);
registerPublic("pivottablecalc.get-data", "1.15", false, getPivotTableData);
registerPublic("pivottablecalc.set-layout", "1.15", true, setPivotTableLayout);
registerPublic("pivottablecalc.set-subtotals", "1.15", true, setPivotFieldSubtotals);
registerPublic("pivottablecalc.set-grand-totals", "1.15", true, setPivotTableGrandTotals);
registerPublic("slicer.create-slicer", "1.15", true, createPivotSlicer);
registerPublic("slicer.create-table-slicer", "1.10", true, createTableSlicer);
registerInternal("screenshot.prepare-range-geometry", "1.1", true, prepareRangeGeometry);
registerInternal("screenshot.prepare-sheet-geometry", "1.1", true, prepareSheetGeometry);
registerInternal("screenshot.restore-view", "1.1", true, restoreCaptureView);

export const IMPLEMENTED_ACTIONS = Object.freeze(publicActions);
export const INTERNAL_ACTIONS = Object.freeze(internalActions);

export function getImplementedAction(name) {
  return actions.get(name) ?? null;
}

export function negotiateRequirementSets(requirements) {
  return {
    excelApi: EXCEL_API_REQUIREMENT_SETS.filter((version) =>
      requirements.isSetSupported("ExcelApi", version)),
    excelApiDesktop: EXCEL_API_DESKTOP_REQUIREMENT_SETS.filter((version) =>
      requirements.isSetSupported("ExcelApiDesktop", version))
  };
}

export function executeOfficeAction(request, runtime = {}) {
  const action = actions.get(request.action);
  if (!action) {
    return Promise.reject(new Error(`Office.js action '${request.action}' is not implemented.`));
  }
  const negotiated = runtime.requirementSets;
  const negotiatedVersions = action.requirementFamily === "ExcelApiDesktop"
    ? negotiated?.excelApiDesktop
    : negotiated?.excelApi;
  if (negotiated && !supports(negotiatedVersions ?? [], action.requirementSet)) {
    return Promise.reject(new Error(
      `Office.js action '${request.action}' requires ${action.requirementFamily} ${action.requirementSet}; ` +
      `the active workbook reports ${negotiatedVersions?.at(-1) ?? `no ${action.requirementFamily} set`}.`
    ));
  }
  const run = runtime.run ?? globalThis.Excel?.run;
  if (typeof run !== "function") {
    return Promise.reject(new Error("Excel.run is unavailable in the active task pane."));
  }

  const execute = () => run((context) => action.handler(context, request.payload ?? {}, request));
  const pending = dispatchTail.then(execute, execute);
  dispatchTail = pending.catch(() => {});
  return pending;
}

function registerPublic(name, requirementSet, mutation, handler) {
  publicActions.push(name);
  register(name, "ExcelApi", requirementSet, mutation, handler);
}

function registerInternal(name, requirementSet, mutation, handler) {
  internalActions.push(name);
  register(name, "ExcelApiDesktop", requirementSet, mutation, handler);
}

function register(name, requirementFamily, requirementSet, mutation, handler) {
  actions.set(name, Object.freeze({
    name,
    requirementFamily,
    requirementSet,
    mutation,
    handler
  }));
}

function supports(versions, minimum) {
  return versions.some((version) => compareVersions(version, minimum) >= 0);
}

function compareVersions(left, right) {
  const [leftMajor, leftMinor] = left.split(".").map(Number);
  const [rightMajor, rightMinor] = right.split(".").map(Number);
  return leftMajor - rightMajor || leftMinor - rightMinor;
}

async function prepareRangeGeometry(context, payload, request) {
  const sheet = getWorksheet(context, payload.sheetName);
  const range = sheet.getRange(requiredString(payload.rangeAddress, "rangeAddress"));
  return prepareGeometry(context, sheet, range, request);
}

async function prepareSheetGeometry(context, payload, request) {
  const sheet = getWorksheet(context, payload.sheetName);
  const usedRange = sheet.getUsedRange();
  usedRange.load("rowIndex,columnIndex,rowCount,columnCount");
  await context.sync();
  const rowCount = Math.min(usedRange.rowCount, 500);
  const columnCount = Math.min(usedRange.columnCount, 50);
  const truncated = rowCount !== usedRange.rowCount || columnCount !== usedRange.columnCount;
  const range = truncated
    ? sheet.getRangeByIndexes(
        usedRange.rowIndex,
        usedRange.columnIndex,
        rowCount,
        columnCount
      )
    : usedRange;
  return prepareGeometry(context, sheet, range, request, {
    truncated,
    ...(truncated
      ? { message: "The worksheet capture was limited to its top-left 500 rows and 50 columns." }
      : {})
  });
}

async function prepareGeometry(context, sheet, range, request, resultMetadata = {}) {
  const owner = captureOwner(request);
  if (captureTokenBySession.has(owner.key)) {
    throw new Error("This exact workbook session already has a capture awaiting view restoration.");
  }

  const window = context.workbook.application.activeWindow;
  const activeSheet = window.activeWorksheet;
  const workbookActiveSheet = context.workbook.worksheets.getActiveWorksheet();
  const activeCell = window.activeCell;
  const visibleRange = window.visibleRange;
  const selection = context.workbook.getSelectedRange();
  window.load([
    "index", "windowNumber", "type", "windowState",
    "left", "top", "width", "height", "usableWidth", "usableHeight",
    "scrollRow", "scrollColumn", "zoom", "view",
    "freezePanes", "split", "splitRow", "splitColumn",
    "splitHorizontal", "splitVertical"
  ].join(","));
  activeSheet.load("id,name");
  workbookActiveSheet.load("id,name");
  activeCell.load("address");
  visibleRange.load("address");
  selection.load("address");
  sheet.load("id,name");
  range.load("address,left,top,width,height");
  await context.sync();
  if (activeSheet.id !== workbookActiveSheet.id) {
    throw new Error("The active Excel window does not belong to this exact workbook.");
  }
  if (lowerFirst(window.type) !== "workbook") {
    throw new Error("The active Excel window is not a workbook window.");
  }

  const viewState = {
    activeWorksheetId: activeSheet.id,
    activeWorksheetName: activeSheet.name,
    selectionAddress: selection.address,
    activeCellAddress: activeCell.address,
    visibleRangeAddress: visibleRange.address,
    scrollRow: window.scrollRow,
    scrollColumn: window.scrollColumn,
    zoom: window.zoom,
    view: window.view,
    windowState: window.windowState,
    freezePanes: window.freezePanes,
    split: window.split,
    splitRow: window.splitRow,
    splitColumn: window.splitColumn,
    splitHorizontal: window.splitHorizontal,
    splitVertical: window.splitVertical
  };
  const rangePoints = rectangle(range.left, range.top, range.width, range.height);
  const windowPoints = rectangle(window.left, window.top, window.width, window.height);

  sheet.activate();
  range.select();
  window.scrollIntoView(
    rangePoints.left,
    rangePoints.top,
    rangePoints.width,
    rangePoints.height,
    true
  );
  await context.sync();

  const rangePixels = convertRectangle(window, rangePoints);
  const windowPixels = convertRectangle(window, windowPoints);
  await context.sync();
  const screenRect = resolvedRectangle(rangePixels, "The screen");
  const excelWindowScreenRect = resolvedRectangle(windowPixels, "The Excel window screen");
  if (!containsRectangle(excelWindowScreenRect, screenRect)) {
    throw new Error(
      "The requested range does not fit within the Excel window as a single capture rectangle. " +
      "Tiled Office.js capture is not implemented."
    );
  }

  const captureToken = globalThis.crypto.randomUUID();
  captureStates.set(captureToken, { owner, viewState });
  captureTokenBySession.set(owner.key, captureToken);
  return success({
    captureToken,
    workbookUrl: owner.workbookUrl,
    worksheet: {
      id: requiredString(sheet.id, "worksheet.id"),
      name: requiredString(sheet.name, "worksheet.name")
    },
    range: { address: requiredString(range.address, "range.address") },
    excelWindow: {
      index: window.index,
      windowNumber: positiveInt32(window.windowNumber, "The Excel window number"),
      type: lowerFirst(window.type),
      state: lowerFirst(window.windowState)
    },
    screenRect,
    excelWindowScreenRect,
    viewState,
    ...resultMetadata
  });
}

async function restoreCaptureView(context, payload, request) {
  const captureToken = requiredString(payload.captureToken, "captureToken");
  const state = captureStates.get(captureToken);
  if (!state) {
    throw new Error("The capture token is unknown or has already been restored.");
  }
  const owner = captureOwner(request);
  if (state.owner.key !== owner.key) {
    throw new Error("The capture token is not bound to this exact workbook session.");
  }

  const window = context.workbook.application.activeWindow;
  const sheet = context.workbook.worksheets.getItem(state.viewState.activeWorksheetId);
  const selection = sheet.getRange(localAddress(state.viewState.selectionAddress));
  sheet.activate();
  window.windowState = state.viewState.windowState;
  window.view = state.viewState.view;
  window.zoom = state.viewState.zoom;
  window.freezePanes = state.viewState.freezePanes;
  window.split = state.viewState.split;
  window.splitRow = state.viewState.splitRow;
  window.splitColumn = state.viewState.splitColumn;
  window.splitHorizontal = state.viewState.splitHorizontal;
  window.splitVertical = state.viewState.splitVertical;
  window.scrollRow = state.viewState.scrollRow;
  window.scrollColumn = state.viewState.scrollColumn;
  selection.select();
  await context.sync();

  captureStates.delete(captureToken);
  captureTokenBySession.delete(owner.key);
  return success({ restored: true, captureToken });
}

function captureOwner(request) {
  const sessionId = requiredString(request.sessionId, "sessionId");
  const workbookUrl = requiredString(request.workbookUrl, "workbookUrl");
  return {
    sessionId,
    workbookUrl,
    key: `${sessionId}\n${workbookUrl}`
  };
}

function rectangle(left, top, width, height) {
  return {
    left,
    top,
    right: left + width,
    bottom: top + height,
    width,
    height
  };
}

function convertRectangle(window, points) {
  return {
    left: window.pointsToScreenPixelsX(points.left),
    top: window.pointsToScreenPixelsY(points.top),
    right: window.pointsToScreenPixelsX(points.right),
    bottom: window.pointsToScreenPixelsY(points.bottom)
  };
}

function resolvedRectangle(pixels, name) {
  const left = pixels.left.value;
  const top = pixels.top.value;
  const right = pixels.right.value;
  const bottom = pixels.bottom.value;
  if (![left, top, right, bottom].every(isInt32)) {
    throw new Error(`${name} rectangle edges must be 32-bit integers.`);
  }
  const width = right - left;
  const height = bottom - top;
  if (!isInt32(width) || !isInt32(height) || width <= 0 || height <= 0) {
    throw new Error(`${name} rectangle dimensions must be positive 32-bit integers.`);
  }
  return {
    left,
    top,
    right,
    bottom,
    width,
    height,
    unit: "physicalPixel",
    origin: "topLeftGlobalScreen"
  };
}

function positiveInt32(value, name) {
  if (!isInt32(value) || value <= 0) {
    throw new Error(`${name} must be a positive 32-bit integer.`);
  }
  return value;
}

function containsRectangle(outer, inner) {
  return inner.left >= outer.left
    && inner.top >= outer.top
    && inner.right <= outer.right
    && inner.bottom <= outer.bottom;
}

function isInt32(value) {
  return Number.isInteger(value) && value >= -2147483648 && value <= 2147483647;
}

function localAddress(address) {
  return address.slice(address.lastIndexOf("!") + 1);
}

async function listTables(context) {
  const tables = context.workbook.tables;
  tables.load("items/name,items/style,items/showHeaders,items/showTotals");
  await context.sync();
  const infos = [];
  for (const table of tables.items) {
    infos.push(await readTableInfo(context, table));
  }
  return success({ tables: infos });
}

async function createTable(context, payload) {
  const sheetName = requiredString(payload.sheetName, "sheetName");
  const tableName = requiredString(payload.tableName, "tableName");
  const rangeAddress = requiredString(payload.rangeAddress, "rangeAddress");
  const table = context.workbook.tables.add(
    `${quoteSheetName(sheetName)}!${rangeAddress}`,
    payload.hasHeaders !== false
  );
  table.name = tableName;
  if (payload.tableStyle) {
    table.style = requiredString(payload.tableStyle, "tableStyle");
  }
  await context.sync();
  return operation("create", `Table '${tableName}' created.`);
}

async function renameTable(context, payload) {
  const table = getTable(context, payload.tableName);
  const newName = requiredString(payload.newName, "newName");
  table.name = newName;
  await context.sync();
  return operation("rename", `Table '${payload.tableName}' renamed to '${newName}'.`);
}

async function deleteTable(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  const range = getTable(context, tableName).convertToRange();
  range.load("address");
  await context.sync();
  return operation("delete", `Table '${tableName}' converted to a range.`);
}

async function readTable(context, payload) {
  return success({ table: await readTableInfo(context, getTable(context, payload.tableName)) });
}

async function resizeTable(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  const table = getTable(context, tableName);
  const sheet = table.getRange().worksheet;
  sheet.load("name");
  await context.sync();
  table.resize(sheet.getRange(requiredString(payload.newRange, "newRange")));
  await context.sync();
  return operation("resize", `Table '${tableName}' resized to '${payload.newRange}'.`);
}

async function toggleTotals(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  getTable(context, tableName).showTotals = requiredBoolean(payload.showTotals, "showTotals");
  await context.sync();
  return operation("toggle-totals", `Totals row ${payload.showTotals ? "shown" : "hidden"} for '${tableName}'.`);
}

async function setColumnTotal(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  const columnName = requiredString(payload.columnName, "columnName");
  const column = getTable(context, tableName).columns.getItem(columnName);
  column.totalRowFunction = normalizeTotalFunction(payload.totalFunction);
  await context.sync();
  return operation("set-column-total", `Totals function set for '${tableName}[${columnName}]'.`);
}

async function appendRows(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  if (!Array.isArray(payload.rows) || payload.rows.length === 0
      || payload.rows.some((row) => !Array.isArray(row))) {
    throw new TypeError("rows must be a non-empty two-dimensional array.");
  }
  getTable(context, tableName).rows.add(null, payload.rows);
  await context.sync();
  return operation("append", `Appended ${payload.rows.length} row(s) to '${tableName}'.`);
}

async function getTableData(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  const table = getTable(context, tableName);
  const headers = table.getHeaderRowRange();
  const rows = table.rows;
  headers.load("values");
  rows.load("count");
  await context.sync();
  let data = [];
  let rowCount = 0;
  const columnCount = headers.values[0]?.length ?? 0;
  if (rows.count > 0) {
    const body = table.getDataBodyRange();
    const view = payload.visibleOnly === true ? body.getVisibleView() : body;
    view.load("values,rowCount,columnCount");
    await context.sync();
    data = view.values;
    rowCount = view.rowCount;
  }
  return success({
    tableName,
    headers: headers.values[0] ?? [],
    data,
    rowCount,
    columnCount
  });
}

async function setTableStyle(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  getTable(context, tableName).style = requiredString(payload.tableStyle, "tableStyle");
  await context.sync();
  return operation("set-style", `Style set for table '${tableName}'.`);
}

async function applyFilter(context, payload) {
  const { tableName, columnName, column } = getColumn(context, payload);
  column.filter.applyCustomFilter(requiredString(payload.criteria, "criteria"));
  await context.sync();
  return operation("apply-filter", `Filter applied to '${tableName}[${columnName}]'.`);
}

async function applyFilterValues(context, payload) {
  const { tableName, columnName, column } = getColumn(context, payload);
  if (!Array.isArray(payload.values) || payload.values.some((value) => typeof value !== "string")) {
    throw new TypeError("values must be an array of strings.");
  }
  column.filter.applyValuesFilter(payload.values);
  await context.sync();
  return operation("apply-filter-values", `Value filter applied to '${tableName}[${columnName}]'.`);
}

async function clearFilters(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  getTable(context, tableName).clearFilters();
  await context.sync();
  return operation("clear-filters", `Filters cleared from '${tableName}'.`);
}

async function getFilters(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  const columns = getTable(context, tableName).columns;
  columns.load("items/name,items/index,items/filter/criteria");
  await context.sync();
  const columnFilters = columns.items.map((column) => {
    const criteria = column.filter.criteria ?? {};
    const values = Array.isArray(criteria.values) ? criteria.values.map(String) : null;
    const criterion = criteria.criterion1 ?? null;
    return {
      columnName: column.name,
      columnIndex: column.index + 1,
      isFiltered: Boolean(criterion || values?.length),
      ...(criterion ? { criteria: criterion } : {}),
      ...(values ? { filterValues: values } : {})
    };
  });
  return success({
    tableName,
    columnFilters,
    hasActiveFilters: columnFilters.some((filter) => filter.isFiltered)
  });
}

async function addColumn(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  const columnName = requiredString(payload.columnName, "columnName");
  const position = payload.position == null ? null : requiredPositiveInteger(payload.position, "position") - 1;
  getTable(context, tableName).columns.add(position, null, columnName);
  await context.sync();
  return operation("add-column", `Column '${columnName}' added to '${tableName}'.`);
}

async function removeColumn(context, payload) {
  const { tableName, columnName, column } = getColumn(context, payload);
  column.delete();
  await context.sync();
  return operation("remove-column", `Column '${columnName}' removed from '${tableName}'.`);
}

async function renameColumn(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  const oldName = requiredString(payload.oldName, "oldName");
  const newName = requiredString(payload.newName, "newName");
  getTable(context, tableName).columns.getItem(oldName).name = newName;
  await context.sync();
  return operation("rename-column", `Column '${oldName}' renamed to '${newName}'.`);
}

async function getStructuredReference(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  const region = requiredString(payload.region, "region");
  const table = getTable(context, tableName);
  const range = tableRegionRange(table, region, payload.columnName);
  range.load("address,rowCount,columnCount");
  range.worksheet.load("name");
  await context.sync();
  const suffix = structuredReferenceSuffix(region, payload.columnName);
  return success({
    tableName,
    region,
    rangeAddress: range.address,
    structuredReference: `${tableName}${suffix}`,
    sheetName: range.worksheet.name,
    ...(payload.columnName ? { columnName: payload.columnName } : {}),
    rowCount: range.rowCount,
    columnCount: range.columnCount
  });
}

async function sortTable(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  const column = getTable(context, tableName).columns.getItem(requiredString(payload.columnName, "columnName"));
  column.load("index");
  await context.sync();
  getTable(context, tableName).sort.apply([
    { key: column.index, ascending: payload.ascending !== false }
  ], false, false);
  await context.sync();
  return operation("sort", `Table '${tableName}' sorted.`);
}

async function sortTableMulti(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  if (!Array.isArray(payload.sortColumns) || payload.sortColumns.length === 0) {
    throw new TypeError("sortColumns must be a non-empty array.");
  }
  const table = getTable(context, tableName);
  const columns = payload.sortColumns.map((item) => {
    const column = table.columns.getItem(requiredString(item.columnName, "columnName"));
    column.load("index");
    return { column, ascending: item.ascending !== false };
  });
  await context.sync();
  table.sort.apply(columns.map((item) => ({
    key: item.column.index,
    ascending: item.ascending
  })), false, false);
  await context.sync();
  return operation("sort-multi", `Table '${tableName}' sorted by ${columns.length} columns.`);
}

async function getColumnNumberFormat(context, payload) {
  const { column } = getColumn(context, payload);
  const range = column.getDataBodyRange();
  range.load("address,numberFormat,rowCount,columnCount");
  range.worksheet.load("name");
  await context.sync();
  return success({
    sheetName: range.worksheet.name,
    rangeAddress: range.address,
    formats: range.numberFormat,
    rowCount: range.rowCount,
    columnCount: range.columnCount
  });
}

async function setColumnNumberFormat(context, payload) {
  const { tableName, columnName, column } = getColumn(context, payload);
  const range = column.getDataBodyRange();
  range.load("rowCount,columnCount");
  await context.sync();
  const formatCode = requiredString(payload.formatCode, "formatCode");
  range.numberFormat = Array.from(
    { length: range.rowCount },
    () => Array(range.columnCount).fill(formatCode)
  );
  await context.sync();
  return operation("set-column-number-format", `Number format set for '${tableName}[${columnName}]'.`);
}

async function addConditionalRule(context, payload) {
  const range = getWorksheet(context, payload.sheetName).getRange(
    requiredString(payload.rangeAddress, "rangeAddress")
  );
  const ruleType = requiredString(payload.ruleType, "ruleType").replaceAll("-", "").toLowerCase();
  const officeType = conditionalOfficeType(ruleType);
  const format = range.conditionalFormats.add(officeType);
  switch (ruleType) {
    case "cellvalue":
      format.cellValue.rule = {
        formula1: requiredString(payload.formula1, "formula1"),
        operator: conditionalOperator(payload.operatorType),
        ...(payload.formula2 ? { formula2: payload.formula2 } : {})
      };
      applyConditionalFormat(format.cellValue.format, payload);
      break;
    case "expression":
      format.custom.rule.formula = requiredString(payload.formula1, "formula1");
      applyConditionalFormat(format.custom.format, payload);
      break;
    case "colorscale":
      format.colorScale.criteria = colorScaleCriteria(payload);
      break;
    case "databar":
      configureDataBar(format.dataBar, payload);
      break;
    case "iconset":
      configureIconSet(format.iconSet, payload);
      break;
    case "top10":
      format.topBottom.rule = topBottomRule(payload);
      applyConditionalFormat(format.topBottom.format, payload);
      break;
    case "aboveaverage":
    case "timeperiod":
    case "uniquevalues":
    case "blankscondition":
      format.preset.rule = { criterion: presetCriterion(ruleType, payload) };
      applyConditionalFormat(format.preset.format, payload);
      break;
  }
  await context.sync();
  return operation("add-rule", `Conditional formatting rule added to '${payload.rangeAddress}'.`);
}

async function clearConditionalRules(context, payload) {
  getWorksheet(context, payload.sheetName)
    .getRange(requiredString(payload.rangeAddress, "rangeAddress"))
    .conditionalFormats.clearAll();
  await context.sync();
  return operation("clear-rules", `Conditional formatting cleared from '${payload.rangeAddress}'.`);
}

async function listConditionalRules(context, payload) {
  const sheet = getWorksheet(context, payload.sheetName);
  const rangeAddress = requiredString(payload.rangeAddress, "rangeAddress");
  sheet.load("name");
  return success({
    sheetName: await loadSheetName(context, sheet),
    rangeAddress,
    rules: await loadConditionalRules(context, sheet.getRange(rangeAddress).conditionalFormats)
  });
}

async function listWorksheetConditionalRules(context, payload) {
  const sheet = getWorksheet(context, payload.sheetName);
  sheet.load("name");
  return success({
    sheetName: await loadSheetName(context, sheet),
    rules: await loadConditionalRules(context, sheet.getUsedRange().conditionalFormats)
  });
}

async function copyWorksheet(context, payload) {
  const sourceName = requiredString(payload.sourceName, "sourceName");
  const targetName = requiredString(payload.targetName, "targetName");
  const source = context.workbook.worksheets.getItem(sourceName);
  const copy = source.copy("After", source);
  copy.name = targetName;
  await context.sync();
  return operation("copy", `Worksheet '${sourceName}' copied to '${targetName}'.`);
}

async function moveWorksheet(context, payload) {
  const sheetName = requiredString(payload.sheetName, "sheetName");
  if (payload.beforeSheet && payload.afterSheet) {
    throw new TypeError("Specify beforeSheet or afterSheet, not both.");
  }
  const worksheets = context.workbook.worksheets;
  worksheets.load("items/name,items/position");
  await context.sync();
  const sheet = worksheets.getItem(sheetName);
  let targetPosition = worksheets.items.length - 1;
  if (payload.beforeSheet) {
    targetPosition = worksheets.getItem(payload.beforeSheet).position;
  } else if (payload.afterSheet) {
    targetPosition = worksheets.getItem(payload.afterSheet).position + 1;
  }
  if (sheet.position < targetPosition) {
    targetPosition--;
  }
  sheet.position = targetPosition;
  await context.sync();
  return operation("move", `Worksheet '${sheetName}' moved.`);
}

async function readTableInfo(context, table) {
  const range = table.getRange();
  const columns = table.columns;
  const rows = table.rows;
  range.worksheet.load("name");
  range.load("address,columnCount");
  columns.load("items/name");
  rows.load("count");
  table.load("name,style,showHeaders,showTotals");
  await context.sync();
  return {
    name: table.name,
    sheetName: range.worksheet.name,
    range: range.address,
    hasHeaders: table.showHeaders,
    tableStyle: table.style || null,
    rowCount: rows.count,
    columnCount: range.columnCount,
    columns: columns.items.map((column) => column.name),
    showTotals: table.showTotals
  };
}

async function loadConditionalRules(context, collection) {
  collection.load("items/type,items/priority,items/stopIfTrue");
  await context.sync();
  const loaded = collection.items.map((item) => {
    const range = item.getRange();
    range.load("address");
    const detail = loadConditionalDetail(item);
    return { item, range, detail };
  });
  await context.sync();
  return loaded
    .sort((left, right) => left.item.priority - right.item.priority)
    .map(({ item, range, detail }) => ({
      type: resolvedConditionalType(detail),
      appliesTo: range.address,
      priority: item.priority + 1,
      ...(item.stopIfTrue == null ? {} : { stopIfTrue: item.stopIfTrue }),
      ...readConditionalDetail(detail)
    }));
}

async function loadSheetName(context, sheet) {
  await context.sync();
  return sheet.name;
}

function loadConditionalDetail(item) {
  const type = conditionalTypeName(item.type);
  const source = {
    cellValue: item.cellValue,
    expression: item.custom,
    colorScale: item.colorScale,
    dataBar: item.dataBar,
    iconSet: item.iconSet,
    top10: item.topBottom,
    presetCriteria: item.preset
  }[type];
  if (!source) {
    return { type };
  }
  if (type === "cellValue" || type === "expression" || type === "top10"
      || type === "presetCriteria") {
    source.load("rule");
    source.format.fill.load("color");
    source.format.font.load("color,bold,italic");
  } else if (type === "colorScale") {
    source.load("criteria");
  } else if (type === "dataBar") {
    source.load("barDirection,lowerBoundRule,showDataBarOnly,upperBoundRule");
    source.positiveFormat.load("fillColor");
    source.negativeFormat.load("fillColor,matchPositiveFillColor");
  } else if (type === "iconSet") {
    source.load("criteria,reverseIconOrder,showIconOnly,style");
  }
  return { type, source };
}

function readConditionalDetail(detail) {
  const { type, source } = detail;
  if (!source) return {};
  if (type === "cellValue") {
    return {
      operator: conditionalOperatorName(source.rule.operator),
      formula1: source.rule.formula1,
      ...(source.rule.formula2 ? { formula2: source.rule.formula2 } : {}),
      ...readConditionalFormat(source.format)
    };
  }
  if (type === "expression") {
    return { formula1: source.rule.formula, ...readConditionalFormat(source.format) };
  }
  if (type === "colorScale") {
    const criteria = source.criteria;
    return {
      colorScaleCriteria: [criteria.minimum, criteria.midpoint, criteria.maximum]
        .filter(Boolean)
        .map((item) => ({
          type: conditionalThresholdName(item.type),
          ...(item.formula == null ? {} : { value: String(item.formula) }),
          ...(item.color ? { color: item.color } : {})
        }))
    };
  }
  if (type === "dataBar") {
    return {
      dataBar: {
        fillColor: source.positiveFormat.fillColor,
        ...(source.negativeFormat.matchPositiveFillColor
          ? {}
          : { barColorNegative: source.negativeFormat.fillColor }),
        direction: lowerFirst(source.barDirection),
        showValue: !source.showDataBarOnly,
        minType: conditionalThresholdName(source.lowerBoundRule.type),
        ...(source.lowerBoundRule.formula == null
          ? {}
          : { minValue: String(source.lowerBoundRule.formula) }),
        maxType: conditionalThresholdName(source.upperBoundRule.type),
        ...(source.upperBoundRule.formula == null
          ? {}
          : { maxValue: String(source.upperBoundRule.formula) })
      }
    };
  }
  if (type === "iconSet") {
    return {
      iconSet: {
        id: iconSetName(source.style),
        reverse: source.reverseIconOrder,
        showIconOnly: source.showIconOnly,
        criteria: source.criteria.map((item) => ({
          operator: conditionalIconOperatorName(item.operator),
          value: item.formula,
          type: lowerFirst(item.type)
        }))
      }
    };
  }
  if (type === "top10") {
    return {
      top10: {
        rank: source.rule.rank,
        percent: String(source.rule.type).endsWith("Percent"),
        topBottom: String(source.rule.type).startsWith("Bottom") ? "bottom" : "top"
      },
      ...readConditionalFormat(source.format)
    };
  }
  const resolvedType = resolvedConditionalType(detail);
  const criterion = String(source.rule.criterion);
  if (resolvedType === "aboveAverage") {
    return { aboveBelow: aboveBelowName(criterion), ...readConditionalFormat(source.format) };
  }
  if (resolvedType === "timePeriod") {
    return { datePeriod: datePeriodName(criterion), ...readConditionalFormat(source.format) };
  }
  return readConditionalFormat(source.format);
}

function applyConditionalFormat(format, payload) {
  if (payload.interiorPattern != null) {
    throw new TypeError("interiorPattern is not supported by the Office.js conditional format API.");
  }
  if (payload.interiorColor) {
    format.fill.color = payload.interiorColor;
  }
  if (payload.fontColor) {
    format.font.color = payload.fontColor;
  }
  if (payload.fontBold != null) {
    format.font.bold = payload.fontBold;
  }
  if (payload.fontItalic != null) {
    format.font.italic = payload.fontItalic;
  }
  if (payload.borderStyle || payload.borderColor) {
    for (const side of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"]) {
      const border = format.borders.getItem(side);
      if (payload.borderStyle) border.style = conditionalBorderStyle(payload.borderStyle);
      if (payload.borderColor) border.color = payload.borderColor;
    }
  }
}

async function createChartFromRange(context, payload) {
  const sheet = getWorksheet(context, payload.sheetName);
  const source = sheet.getRange(requiredString(payload.sourceRangeAddress, "sourceRangeAddress"));
  return createChart(context, sheet, source, payload);
}

async function createChartFromTable(context, payload) {
  const sheet = getWorksheet(context, payload.sheetName);
  const source = getTable(context, payload.tableName).getRange();
  return createChart(context, sheet, source, payload);
}

async function createChart(context, sheet, source, payload) {
  const chartType = officeChartType(payload.chartType);
  const charts = sheet.charts;
  const chart = charts.add(chartType, source, "Auto");
  if (payload.chartName != null) {
    chart.name = requiredString(payload.chartName, "chartName");
  }
  chart.load("name");
  const usedRange = sheet.getUsedRange();
  usedRange.load("address,left,top,width,height");
  charts.load("items/name,items/left,items/top,items/width,items/height");
  await context.sync();

  if (payload.targetRange) {
    const target = sheet.getRange(requiredString(payload.targetRange, "targetRange"));
    chart.setPosition(target.getCell(0, 0), target.getLastCell());
  } else if (Number(payload.left) !== 0 || Number(payload.top) !== 0) {
    chart.left = optionalFiniteNumber(payload.left, 0, "left");
    chart.top = optionalFiniteNumber(payload.top, 0, "top");
    chart.width = optionalPositiveNumber(payload.width, 400, "width");
    chart.height = optionalPositiveNumber(payload.height, 300, "height");
  } else {
    const otherChartBottom = charts.items
      .filter((candidate) => candidate.name !== chart.name)
      .reduce((bottom, candidate) => Math.max(bottom, candidate.top + candidate.height), 0);
    chart.left = 10;
    chart.top = Math.max(usedRange.top + usedRange.height, otherChartBottom) + 10;
    chart.width = optionalPositiveNumber(payload.width, 400, "width");
    chart.height = optionalPositiveNumber(payload.height, 300, "height");
  }

  chart.load("name,chartType,left,top,width,height");
  sheet.load("name");
  await context.sync();
  const warnings = chartCollisionWarnings(chart, usedRange, charts.items);
  return success({
    action: "create",
    message: chartPositionMessage(warnings, charts.items.length),
    chartName: chart.name,
    sheetName: sheet.name,
    chartType: chartTypeName(chart.chartType),
    isPivotChart: false,
    linkedPivotTable: null,
    left: chart.left,
    top: chart.top,
    width: chart.width,
    height: chart.height
  });
}

async function deleteChart(context, payload) {
  const chart = await findChart(context, payload.chartName);
  const name = requiredString(payload.chartName, "chartName");
  chart.delete();
  await context.sync();
  return operation("delete", `Chart '${name}' deleted.`);
}

async function moveChart(context, payload) {
  const chart = await findChart(context, payload.chartName);
  if (payload.left != null) chart.left = finiteNumber(payload.left, "left");
  if (payload.top != null) chart.top = finiteNumber(payload.top, "top");
  if (payload.width != null) chart.width = positiveNumber(payload.width, "width");
  if (payload.height != null) chart.height = positiveNumber(payload.height, "height");
  await context.sync();
  return operation("move", `Chart '${payload.chartName}' moved.`);
}

async function fitChartToRange(context, payload) {
  const chart = await findChart(context, payload.chartName);
  const sheet = getWorksheet(context, payload.sheetName);
  const target = sheet.getRange(requiredString(payload.rangeAddress, "rangeAddress"));
  chart.setPosition(target.getCell(0, 0), target.getLastCell());
  await context.sync();
  return operation("fit-to-range", `Chart '${payload.chartName}' fitted to '${payload.rangeAddress}'.`);
}

async function setChartType(context, payload) {
  const chart = await findChart(context, payload.chartName);
  chart.chartType = officeChartType(payload.chartType);
  await context.sync();
  return operation("set-chart-type", `Chart type set for '${payload.chartName}'.`);
}

async function setChartTitle(context, payload) {
  const chart = await findChart(context, payload.chartName);
  const title = typeof payload.title === "string"
    ? payload.title
    : throwType("title must be a string.");
  chart.title.text = title;
  chart.title.visible = title.length > 0;
  await context.sync();
  return operation("set-title", `Chart title updated for '${payload.chartName}'.`);
}

async function setChartAxisTitle(context, payload) {
  const chart = await findChart(context, payload.chartName);
  const axis = chartAxis(chart, payload.axis);
  const title = typeof payload.title === "string"
    ? payload.title
    : throwType("title must be a string.");
  axis.title.text = title;
  axis.title.visible = title.length > 0;
  await context.sync();
  return operation("set-axis-title", `Axis title updated for '${payload.chartName}'.`);
}

async function getChartAxisNumberFormat(context, payload) {
  const axis = chartAxis(await findChart(context, payload.chartName), payload.axis);
  axis.load("numberFormat");
  await context.sync();
  return axis.numberFormat;
}

async function setChartAxisNumberFormat(context, payload) {
  const axis = chartAxis(await findChart(context, payload.chartName), payload.axis);
  axis.numberFormat = requiredString(payload.numberFormat, "numberFormat");
  await context.sync();
  return operation("set-axis-number-format", `Axis number format updated for '${payload.chartName}'.`);
}

async function showChartLegend(context, payload) {
  const chart = await findChart(context, payload.chartName);
  chart.legend.visible = requiredBoolean(payload.visible, "visible");
  if (payload.legendPosition != null) {
    chart.legend.position = enumValue(payload.legendPosition, {
      bottom: "Bottom",
      corner: "Corner",
      custom: "Custom",
      left: "Left",
      right: "Right",
      top: "Top"
    }, "legendPosition");
  }
  await context.sync();
  return operation("show-legend", `Chart legend updated for '${payload.chartName}'.`);
}

async function setChartStyle(context, payload) {
  const styleId = requiredPositiveInteger(payload.styleId, "styleId");
  if (styleId > 48) {
    throw new TypeError("styleId must be between 1 and 48.");
  }
  const chart = await findChart(context, payload.chartName);
  chart.style = styleId;
  await context.sync();
  return operation("set-style", `Chart style updated for '${payload.chartName}'.`);
}

async function setChartDataLabels(context, payload) {
  const chart = await findChart(context, payload.chartName);
  const index = payload.seriesIndex == null ? 0 : nonNegativeInteger(payload.seriesIndex, "seriesIndex");
  const labels = index === 0
    ? chart.dataLabels
    : chart.series.getItemAt(index - 1).dataLabels;
  for (const property of [
    "showValue", "showPercentage", "showSeriesName", "showCategoryName", "showBubbleSize"
  ]) {
    if (payload[property] != null) {
      labels[property] = requiredBoolean(payload[property], property);
    }
  }
  if (payload.separator != null) {
    labels.separator = typeof payload.separator === "string"
      ? payload.separator
      : throwType("separator must be a string.");
  }
  if (payload.labelPosition != null) {
    labels.position = enumValue(payload.labelPosition, {
      bestfit: "BestFit",
      center: "Center",
      above: "Top",
      below: "Bottom",
      left: "Left",
      right: "Right",
      insidebase: "InsideBase",
      insideend: "InsideEnd",
      outsideend: "OutsideEnd"
    }, "labelPosition");
  }
  await context.sync();
  return operation("set-data-labels", `Data labels updated for '${payload.chartName}'.`);
}

async function setChartAxisScale(context, payload) {
  const axis = chartAxis(await findChart(context, payload.chartName), payload.axis);
  axis.minimum = payload.minimumScale == null ? "" : finiteNumber(payload.minimumScale, "minimumScale");
  axis.maximum = payload.maximumScale == null ? "" : finiteNumber(payload.maximumScale, "maximumScale");
  axis.majorUnit = payload.majorUnit == null ? "" : positiveNumber(payload.majorUnit, "majorUnit");
  axis.minorUnit = payload.minorUnit == null ? "" : positiveNumber(payload.minorUnit, "minorUnit");
  await context.sync();
  return operation("set-axis-scale", `Axis scale updated for '${payload.chartName}'.`);
}

async function getChartGridlines(context, payload) {
  const chart = await findChart(context, payload.chartName);
  const valueAxis = chart.axes.getItem("Value", "Primary");
  const categoryAxis = chart.axes.getItem("Category", "Primary");
  valueAxis.majorGridlines.load("visible");
  valueAxis.minorGridlines.load("visible");
  categoryAxis.majorGridlines.load("visible");
  categoryAxis.minorGridlines.load("visible");
  await context.sync();
  return success({
    action: "get-gridlines",
    message: `Gridlines read for '${payload.chartName}'.`,
    chartName: requiredString(payload.chartName, "chartName"),
    gridlines: {
      hasValueMajorGridlines: valueAxis.majorGridlines.visible,
      hasValueMinorGridlines: valueAxis.minorGridlines.visible,
      hasCategoryMajorGridlines: categoryAxis.majorGridlines.visible,
      hasCategoryMinorGridlines: categoryAxis.minorGridlines.visible
    }
  });
}

async function setChartGridlines(context, payload) {
  const axis = chartAxis(await findChart(context, payload.chartName), payload.axis);
  if (payload.showMajor != null) {
    axis.majorGridlines.visible = requiredBoolean(payload.showMajor, "showMajor");
  }
  if (payload.showMinor != null) {
    axis.minorGridlines.visible = requiredBoolean(payload.showMinor, "showMinor");
  }
  await context.sync();
  return operation("set-gridlines", `Gridlines updated for '${payload.chartName}'.`);
}

async function setSeriesChartType(context, payload) {
  const chart = await findChart(context, payload.chartName);
  const index = requiredPositiveInteger(payload.seriesIndex, "seriesIndex");
  chart.series.getItemAt(index - 1).chartType = officeChartType(payload.chartType);
  await context.sync();
  return operation("set-series-chart-type", `Series ${index} type updated for '${payload.chartName}'.`);
}

async function getChartPlotOptions(context, payload) {
  const chart = await findChart(context, payload.chartName);
  chart.load("plotBy,displayBlanksAs,plotVisibleOnly");
  await context.sync();
  const displayBlanksAs = {
    NotPlotted: "Gaps",
    Zero: "Zero",
    Interplotted: "Interpolated"
  }[chart.displayBlanksAs];
  if (!displayBlanksAs) {
    throw new Error(`Office.js returned unsupported blank-cell plotting mode '${chart.displayBlanksAs}'.`);
  }
  return success({
    chartName: requiredString(payload.chartName, "chartName"),
    plotBy: chart.plotBy === "Rows" ? "Rows" : "Columns",
    displayBlanksAs,
    plotVisibleOnly: chart.plotVisibleOnly
  });
}

async function setChartPlotOptions(context, payload) {
  const chart = await findChart(context, payload.chartName);
  if (payload.plotBy != null) {
    chart.plotBy = enumValue(payload.plotBy, { rows: "Rows", columns: "Columns" }, "plotBy");
  }
  if (payload.displayBlanksAs != null) {
    chart.displayBlanksAs = enumValue(payload.displayBlanksAs, {
      gaps: "NotPlotted",
      zero: "Zero",
      interpolated: "Interplotted"
    }, "displayBlanksAs");
  }
  if (payload.plotVisibleOnly != null) {
    chart.plotVisibleOnly = requiredBoolean(payload.plotVisibleOnly, "plotVisibleOnly");
  }
  await context.sync();
  return operation("set-plot-options", `Plot options updated for '${payload.chartName}'.`);
}

async function addChartTrendline(context, payload) {
  const chart = await findChart(context, payload.chartName);
  const seriesIndex = requiredPositiveInteger(payload.seriesIndex, "seriesIndex");
  const series = chart.series.getItemAt(seriesIndex - 1);
  const trendlines = series.trendlines;
  const type = trendlineType(payload.trendlineType);
  const trendline = trendlines.add(type);
  applyTrendlineProperties(trendline, payload, true);
  trendline.load("name,type");
  const count = trendlines.getCount();
  await context.sync();
  return success({
    action: "add-trendline",
    message: `Trendline added to series ${seriesIndex} of '${payload.chartName}'.`,
    chartName: requiredString(payload.chartName, "chartName"),
    seriesIndex,
    trendlineIndex: count.value,
    type: chartTrendlineTypeName(trendline.type),
    name: trendline.name || null
  });
}

async function deleteChartTrendline(context, payload) {
  const seriesIndex = requiredPositiveInteger(payload.seriesIndex, "seriesIndex");
  const trendlineIndex = requiredPositiveInteger(payload.trendlineIndex, "trendlineIndex");
  const chart = await findChart(context, payload.chartName);
  chart.series.getItemAt(seriesIndex - 1).trendlines.getItem(trendlineIndex - 1).delete();
  await context.sync();
  return operation("delete-trendline", `Trendline ${trendlineIndex} deleted from '${payload.chartName}'.`);
}

async function setChartTrendline(context, payload) {
  const seriesIndex = requiredPositiveInteger(payload.seriesIndex, "seriesIndex");
  const trendlineIndex = requiredPositiveInteger(payload.trendlineIndex, "trendlineIndex");
  const chart = await findChart(context, payload.chartName);
  const trendline = chart.series.getItemAt(seriesIndex - 1).trendlines.getItem(trendlineIndex - 1);
  applyTrendlineProperties(trendline, payload, false);
  await context.sync();
  return operation("set-trendline", `Trendline ${trendlineIndex} updated for '${payload.chartName}'.`);
}

function applyTrendlineProperties(trendline, payload, creating) {
  if (payload.order != null) {
    const order = requiredPositiveInteger(payload.order, "order");
    if (order < 2 || order > 6) throw new TypeError("order must be between 2 and 6.");
    trendline.polynomialOrder = order;
  }
  if (payload.period != null) trendline.movingAveragePeriod = requiredPositiveInteger(payload.period, "period");
  if (payload.forward != null) trendline.forwardPeriod = nonNegativeNumber(payload.forward, "forward");
  if (payload.backward != null) trendline.backwardPeriod = nonNegativeNumber(payload.backward, "backward");
  if (payload.intercept != null) trendline.intercept = finiteNumber(payload.intercept, "intercept");
  if (creating || payload.displayEquation != null) {
    trendline.showEquation = requiredBoolean(payload.displayEquation ?? false, "displayEquation");
  }
  if (creating || payload.displayRSquared != null) {
    trendline.showRSquared = requiredBoolean(payload.displayRSquared ?? false, "displayRSquared");
  }
  if (payload.name != null) trendline.name = requiredString(payload.name, "name");
}

async function createPivotTableFromRange(context, payload) {
  const sourceSheet = getWorksheet(context, payload.sourceSheet);
  const source = sourceSheet.getRange(requiredString(payload.sourceRange, "sourceRange"));
  return createPivotTable(context, source, payload);
}

async function createPivotTableFromTable(context, payload) {
  const source = getTable(context, payload.tableName);
  return createPivotTable(context, source, payload);
}

async function createPivotTable(context, source, payload) {
  const destinationSheet = getWorksheet(context, payload.destinationSheet);
  const destination = destinationSheet.getRange(requiredString(payload.destinationCell, "destinationCell"));
  const sourceRange = source.getRange ? source.getRange() : source;
  sourceRange.load("address,rowCount,values");
  const pivotTableName = requiredString(payload.pivotTableName, "pivotTableName");
  const pivot = context.workbook.pivotTables.add(pivotTableName, source, destination);
  const pivotRange = pivot.layout.getRange();
  pivotRange.load("address");
  destinationSheet.load("name");
  await context.sync();
  return success({
    pivotTableName,
    sheetName: destinationSheet.name,
    range: pivotRange.address,
    sourceData: sourceRange.address,
    sourceRowCount: Math.max(0, sourceRange.rowCount - 1),
    availableFields: (sourceRange.values[0] ?? []).map(String)
  });
}

async function deletePivotTable(context, payload) {
  const pivotTableName = requiredString(payload.pivotTableName, "pivotTableName");
  const pivot = context.workbook.pivotTables.getItem(pivotTableName);
  await requireOrdinaryPivot(context, pivot);
  pivot.delete();
  await context.sync();
  return operation("delete", `PivotTable '${pivotTableName}' deleted.`);
}

async function removePivotField(context, payload) {
  const state = await ordinaryPivotFieldState(context, payload);
  state.collection.remove(state.hierarchy);
  await context.sync();
  return pivotFieldResult(payload.fieldName, {
    area: "Hidden",
    position: 0
  });
}

async function setPivotFieldName(context, payload) {
  const state = await ordinaryPivotFieldState(context, payload);
  const customName = requiredString(payload.customName, "customName");
  state.hierarchy.name = customName;
  await context.sync();
  return pivotFieldResult(payload.fieldName, {
    customName,
    area: state.area,
    position: state.position
  });
}

async function setPivotFieldFormat(context, payload) {
  const state = await ordinaryPivotFieldState(context, payload);
  if (state.area !== "Value") {
    throw new Error(
      `Field '${payload.fieldName}' is not in the Values area. Only value fields can have number formats.`
    );
  }
  const numberFormat = requiredString(payload.numberFormat, "numberFormat");
  state.hierarchy.numberFormat = numberFormat;
  state.hierarchy.load("numberFormat");
  await context.sync();
  return pivotFieldResult(payload.fieldName, {
    customName: state.hierarchy.name,
    area: "Value",
    position: state.position,
    numberFormat: state.hierarchy.numberFormat
  });
}

async function setPivotFieldFilter(context, payload) {
  const state = await ordinaryPivotFieldState(context, payload);
  if (state.area === "Value") {
    throw new Error(`Field '${payload.fieldName}' is in the Values area and cannot accept item filters.`);
  }
  if (!Array.isArray(payload.selectedValues)
      || payload.selectedValues.some((item) => typeof item !== "string")) {
    throw new TypeError("selectedValues must be an array of strings.");
  }
  if (new Set(payload.selectedValues).size !== payload.selectedValues.length) {
    throw new TypeError("selectedValues must not contain duplicates.");
  }

  state.field.items.load("items/name");
  const beforeRange = state.pivot.layout.getRange();
  beforeRange.load("rowCount");
  await context.sync();
  const totalRowCount = beforeRange.rowCount;
  const availableItems = state.field.items.items.map((item) => item.name);
  const unknown = payload.selectedValues.filter((item) => !availableItems.includes(item));
  if (unknown.length > 0) {
    throw new Error(
      `Field '${payload.fieldName}' does not contain selected item(s): ${unknown.join(", ")}.`
    );
  }

  state.field.applyFilter({
    manualFilter: { selectedItems: payload.selectedValues }
  });
  const afterRange = state.pivot.layout.getRange();
  afterRange.load("rowCount");
  await context.sync();
  return success({
    fieldName: requiredString(payload.fieldName, "fieldName"),
    selectedItems: [...payload.selectedValues],
    availableItems,
    visibleRowCount: afterRange.rowCount,
    totalRowCount,
    showAll: payload.selectedValues.length === availableItems.length
  });
}

async function sortPivotField(context, payload) {
  const state = await ordinaryPivotFieldState(context, payload);
  if (state.area === "Value") {
    throw new Error(`Field '${payload.fieldName}' is in the Values area and cannot be sorted by labels.`);
  }
  const direction = enumValue(
    payload.direction ?? "Ascending",
    { ascending: "Ascending", descending: "Descending" },
    "direction"
  );
  state.field.sortByLabels(direction);
  await context.sync();
  return pivotFieldResult(payload.fieldName, {
    customName: state.hierarchy.name,
    area: state.area,
    position: state.position
  });
}

async function getPivotTableData(context, payload) {
  const pivotTableName = requiredString(payload.pivotTableName, "pivotTableName");
  const pivot = context.workbook.pivotTables.getItem(pivotTableName);
  await requireOrdinaryPivot(context, pivot);
  const range = pivot.layout.getRange();
  range.load("values,rowCount,columnCount");
  await context.sync();
  return success({
    pivotTableName,
    values: range.values.map((row) => [...row]),
    columnHeaders: [],
    rowHeaders: [],
    dataRowCount: range.values.length,
    dataColumnCount: range.values[0]?.length ?? 0,
    grandTotals: {}
  });
}

async function setPivotTableLayout(context, payload) {
  const pivot = await getOrdinaryPivot(context, payload.pivotTableName);
  const rowLayout = nonNegativeInteger(payload.rowLayout, "rowLayout");
  const layoutType = ["Compact", "Tabular", "Outline"][rowLayout];
  if (!layoutType) {
    throw new TypeError("rowLayout must be 0 (Compact), 1 (Tabular), or 2 (Outline).");
  }
  pivot.layout.layoutType = layoutType;
  await context.sync();
  return success();
}

async function setPivotFieldSubtotals(context, payload) {
  const state = await ordinaryPivotFieldState(context, payload);
  if (state.area !== "Row") {
    throw new Error(`Field '${payload.fieldName}' is not in the Row area.`);
  }
  const showSubtotals = requiredBoolean(payload.showSubtotals, "showSubtotals");
  state.field.subtotals = { automatic: showSubtotals };
  await context.sync();
  return pivotFieldResult(payload.fieldName, {
    workflowHint: showSubtotals
      ? "Subtotals enabled for field. Automatic function selected based on data type."
      : "Subtotals disabled for field. Only detail rows visible."
  });
}

async function setPivotTableGrandTotals(context, payload) {
  const pivot = await getOrdinaryPivot(context, payload.pivotTableName);
  pivot.layout.showRowGrandTotals = requiredBoolean(
    payload.showRowGrandTotals,
    "showRowGrandTotals"
  );
  pivot.layout.showColumnGrandTotals = requiredBoolean(
    payload.showColumnGrandTotals,
    "showColumnGrandTotals"
  );
  await context.sync();
  return success();
}

async function createPivotSlicer(context, payload) {
  const pivotTableName = requiredString(payload.pivotTableName, "pivotTableName");
  const pivot = context.workbook.pivotTables.getItem(pivotTableName);
  await requireOrdinaryPivot(context, pivot);
  return createSlicer(context, pivot, payload.fieldName, payload, {
    connectedPivotTables: [pivotTableName],
    sourceType: "PivotTable"
  });
}

async function createTableSlicer(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  return createSlicer(context, getTable(context, tableName), payload.columnName, payload, {
    connectedPivotTables: [],
    connectedTable: tableName,
    sourceType: "Table"
  });
}

async function createSlicer(context, source, fieldNameValue, payload, sourceResult) {
  const fieldName = requiredString(fieldNameValue, "fieldName");
  const sheet = getWorksheet(context, payload.destinationSheet);
  const position = sheet.getRange(requiredString(payload.position, "position"));
  position.load("left,top,address");
  const slicer = context.workbook.slicers.add(source, fieldName, sheet);
  slicer.name = requiredString(payload.slicerName, "slicerName");
  slicer.caption = slicer.name;
  slicer.load("name,caption");
  await context.sync();
  slicer.left = position.left;
  slicer.top = position.top;
  const selected = slicer.getSelectedItems();
  slicer.slicerItems.load("items/key,items/name");
  sheet.load("name");
  await context.sync();
  const itemByKey = new Map(slicer.slicerItems.items.map((item) => [item.key, item.name]));
  return success({
    name: slicer.name,
    caption: slicer.caption,
    fieldName,
    sheetName: sheet.name,
    position: localAddress(position.address),
    selectedItems: selected.value.map((key) => itemByKey.get(key) ?? key),
    availableItems: slicer.slicerItems.items.map((item) => item.name),
    ...sourceResult,
    workflowHint: sourceResult.sourceType === "Table"
      ? `Slicer '${slicer.name}' created for column '${fieldName}' in table '${sourceResult.connectedTable}'. Use SetTableSlicerSelection to filter data.`
      : `Slicer '${slicer.name}' created for field '${fieldName}'. Use SetSlicerSelection to filter data, or connect additional PivotTables to this slicer.`
  });
}

async function requireOrdinaryPivot(context, pivot) {
  const sourceType = pivot.getDataSourceType();
  await context.sync();
  if (sourceType.value !== "LocalRange" && sourceType.value !== "LocalTable") {
    throw new Error(
      "This Office.js route supports only ordinary PivotTables backed by a local range or table. " +
      "OLAP and Power Pivot require the trusted VBA capability."
    );
  }
}

async function getOrdinaryPivot(context, value) {
  const pivot = context.workbook.pivotTables.getItem(
    requiredString(value, "pivotTableName")
  );
  await requireOrdinaryPivot(context, pivot);
  return pivot;
}

async function ordinaryPivotFieldState(context, payload) {
  const pivot = await getOrdinaryPivot(context, payload.pivotTableName);
  const fieldName = requiredString(payload.fieldName, "fieldName");
  const collections = [
    ["Row", pivot.rowHierarchies],
    ["Column", pivot.columnHierarchies],
    ["Filter", pivot.filterHierarchies],
    ["Value", pivot.dataHierarchies]
  ];
  for (const [, collection] of collections) {
    collection.load("items/name,items/position");
  }
  await context.sync();

  for (const [area, collection] of collections) {
    for (const hierarchy of collection.items) {
      if (area === "Value") {
        hierarchy.field.load("name");
      } else {
        hierarchy.fields.load("items/name");
      }
    }
  }
  await context.sync();

  for (const [area, collection] of collections) {
    const hierarchy = collection.items.find((candidate) =>
      candidate.name === fieldName
      || (area === "Value"
        ? candidate.field.name === fieldName
        : candidate.fields.items.some((field) => field.name === fieldName)));
    if (hierarchy) {
      return {
        pivot,
        collection,
        hierarchy,
        field: area === "Value"
          ? hierarchy.field
          : hierarchy.fields.items.find((field) => field.name === fieldName)
            ?? hierarchy.fields.items[0],
        area,
        position: hierarchy.position + 1
      };
    }
  }
  throw new Error(`Field '${fieldName}' is not currently placed in any area.`);
}

function pivotFieldResult(fieldNameValue, values = {}) {
  return success({
    fieldName: requiredString(fieldNameValue, "fieldName"),
    customName: "",
    area: "Hidden",
    position: 0,
    availableValues: [],
    dataType: "",
    ...values
  });
}

async function findChart(context, value) {
  const chartName = requiredString(value, "chartName");
  const worksheets = context.workbook.worksheets;
  worksheets.load("items/name");
  await context.sync();
  const candidates = worksheets.items.map((sheet) => {
    const chart = sheet.charts.getItemOrNullObject(chartName);
    chart.load("isNullObject");
    return chart;
  });
  await context.sync();
  const chart = candidates.find((candidate) => !candidate.isNullObject);
  if (!chart) throw new Error(`Chart '${chartName}' was not found.`);
  return chart;
}

function chartAxis(chart, value) {
  const axis = requiredString(value, "axis").toLowerCase();
  const mapping = {
    primary: ["Category", "Primary"],
    secondary: ["Value", "Primary"],
    category: ["Category", "Primary"],
    value: ["Value", "Primary"],
    categorysecondary: ["Category", "Secondary"],
    valuesecondary: ["Value", "Secondary"]
  };
  if (!mapping[axis]) throw new TypeError(`Unsupported axis '${value}'.`);
  return chart.axes.getItem(...mapping[axis]);
}

function officeChartType(value) {
  const name = requiredString(value, "chartType");
  const mapping = {
    Column3DClustered: "3DColumnClustered",
    Column3DStacked: "3DColumnStacked",
    Column3DStacked100: "3DColumnStacked100",
    Column3D: "3DColumn",
    Bar3DClustered: "3DBarClustered",
    Bar3DStacked: "3DBarStacked",
    Bar3DStacked100: "3DBarStacked100",
    Line3D: "3DLine",
    Pie3D: "3DPie",
    PieExploded3D: "3DPieExploded",
    Area3D: "3DArea",
    Area3DStacked: "3DAreaStacked",
    Area3DStacked100: "3DAreaStacked100",
    BoxWhisker: "Boxwhisker"
  };
  if (name === "ColumnLineCombo") {
    throw new TypeError(
      "ColumnLineCombo is not a native Office.js chart type; create a regular chart and set individual series types."
    );
  }
  return mapping[name] ?? name;
}

function chartTypeName(value) {
  const mapping = {
    "3DColumnClustered": "Column3DClustered",
    "3DColumnStacked": "Column3DStacked",
    "3DColumnStacked100": "Column3DStacked100",
    "3DColumn": "Column3D",
    "3DBarClustered": "Bar3DClustered",
    "3DBarStacked": "Bar3DStacked",
    "3DBarStacked100": "Bar3DStacked100",
    "3DLine": "Line3D",
    "3DPie": "Pie3D",
    "3DPieExploded": "PieExploded3D",
    "3DArea": "Area3D",
    "3DAreaStacked": "Area3DStacked",
    "3DAreaStacked100": "Area3DStacked100",
    Boxwhisker: "BoxWhisker"
  };
  return mapping[value] ?? value;
}

function trendlineType(value) {
  return enumValue(value, {
    linear: "Linear",
    exponential: "Exponential",
    logarithmic: "Logarithmic",
    polynomial: "Polynomial",
    power: "Power",
    movingaverage: "MovingAverage"
  }, "trendlineType");
}

function chartTrendlineTypeName(value) {
  return String(value);
}

function chartCollisionWarnings(chart, usedRange, charts) {
  const warnings = [];
  if (rectanglesOverlap(chart, usedRange)) {
    warnings.push(`Chart overlaps data area ${usedRange.address}`);
  }
  for (const existing of charts) {
    if (existing.name !== chart.name && rectanglesOverlap(chart, existing)) {
      warnings.push(`Chart overlaps existing chart '${existing.name}'`);
    }
  }
  return warnings;
}

function rectanglesOverlap(left, right) {
  return left.left < right.left + right.width
    && left.left + left.width > right.left
    && left.top < right.top + right.height
    && left.top + left.height > right.top;
}

function chartPositionMessage(warnings, chartCount) {
  if (warnings.length > 0) {
    return `OVERLAP WARNING: ${warnings.join("; ")}. Use chart move or fit-to-range to reposition, then screenshot(capture-sheet) to verify layout.`;
  }
  if (chartCount >= 2) {
    return `IMPORTANT: ${chartCount} charts now on this sheet. You MUST take a screenshot(capture-sheet) to verify no charts overlap each other or the data.`;
  }
  return "IMPORTANT: You MUST take a screenshot(capture-sheet) to verify the chart does not overlap the data.";
}

function getTable(context, name) {
  return context.workbook.tables.getItem(requiredString(name, "tableName"));
}

function getColumn(context, payload) {
  const tableName = requiredString(payload.tableName, "tableName");
  const columnName = requiredString(payload.columnName, "columnName");
  return {
    tableName,
    columnName,
    column: getTable(context, tableName).columns.getItem(columnName)
  };
}

function getWorksheet(context, name) {
  return name === "" || name == null
    ? context.workbook.worksheets.getActiveWorksheet()
    : context.workbook.worksheets.getItem(requiredString(name, "sheetName"));
}

function tableRegionRange(table, region, columnName) {
  const normalized = region.toLowerCase();
  const column = columnName ? table.columns.getItem(columnName) : null;
  if (normalized === "all") return column ? column.getRange() : table.getRange();
  if (normalized === "data") return column ? column.getDataBodyRange() : table.getDataBodyRange();
  if (normalized === "headers") return column ? column.getHeaderRowRange() : table.getHeaderRowRange();
  if (normalized === "totals") return column ? column.getTotalRowRange() : table.getTotalRowRange();
  if (normalized === "thisrow") {
    throw new Error("ThisRow structured references require formula context and have no fixed range.");
  }
  throw new TypeError("region must be All, Data, Headers, Totals, or ThisRow.");
}

function structuredReferenceSuffix(region, columnName) {
  const token = {
    all: "#All",
    data: "#Data",
    headers: "#Headers",
    totals: "#Totals",
    thisrow: "@"
  }[region.toLowerCase()];
  const column = columnName ? `[${escapeStructuredName(columnName)}]` : "";
  return token === "@" ? `[@${escapeStructuredName(columnName ?? "")}]` : `[[${token}]${column ? `,${column}` : ""}]`;
}

function escapeStructuredName(value) {
  return value.replaceAll("'", "''").replaceAll("[", "'[").replaceAll("]", "']");
}

function conditionalOfficeType(ruleType) {
  const types = {
    cellvalue: "CellValue",
    expression: "Custom",
    colorscale: "ColorScale",
    databar: "DataBar",
    iconset: "IconSet",
    top10: "TopBottom",
    aboveaverage: "PresetCriteria",
    timeperiod: "PresetCriteria",
    uniquevalues: "PresetCriteria",
    blankscondition: "PresetCriteria"
  };
  if (!types[ruleType]) {
    throw new TypeError(`Unsupported conditional formatting type '${ruleType}'.`);
  }
  return types[ruleType];
}

function colorScaleCriteria(payload) {
  const result = {
    minimum: colorScaleCriterion(
      payload.colorScaleMinType ?? "minimum",
      payload.colorScaleMinValue,
      payload.colorScaleMinColor ?? "#F8696B"
    ),
    maximum: colorScaleCriterion(
      payload.colorScaleMaxType ?? "maximum",
      payload.colorScaleMaxValue,
      payload.colorScaleMaxColor ?? "#63BE7B"
    )
  };
  if (payload.colorScaleMidType || payload.colorScaleMidValue || payload.colorScaleMidColor) {
    result.midpoint = colorScaleCriterion(
      payload.colorScaleMidType ?? "percentile",
      payload.colorScaleMidValue ?? "50",
      payload.colorScaleMidColor ?? "#FFEB84"
    );
  }
  return result;
}

function colorScaleCriterion(type, value, color) {
  return {
    type: conditionalThresholdType(type, true),
    color,
    ...(value == null ? {} : { formula: String(value) })
  };
}

function configureDataBar(dataBar, payload) {
  if (payload.dataBarColor) {
    dataBar.positiveFormat.fillColor = payload.dataBarColor;
  }
  if (payload.dataBarNegativeColor) {
    dataBar.negativeFormat.fillColor = payload.dataBarNegativeColor;
    dataBar.negativeFormat.matchPositiveFillColor = false;
  }
  if (payload.dataBarDirection) {
    dataBar.barDirection = enumValue(payload.dataBarDirection, {
      context: "Context",
      lefttoright: "LeftToRight",
      righttoleft: "RightToLeft"
    }, "dataBarDirection");
  }
  if (payload.dataBarShowValue != null) {
    dataBar.showDataBarOnly = !payload.dataBarShowValue;
  }
  if (payload.dataBarMinType) {
    dataBar.lowerBoundRule = thresholdRule(payload.dataBarMinType, payload.dataBarMinValue);
  }
  if (payload.dataBarMaxType) {
    dataBar.upperBoundRule = thresholdRule(payload.dataBarMaxType, payload.dataBarMaxValue);
  }
}

function thresholdRule(type, value) {
  return {
    type: conditionalThresholdType(type, false),
    ...(value == null ? {} : { formula: String(value) })
  };
}

function configureIconSet(iconSet, payload) {
  if (payload.iconSetId) {
    iconSet.style = iconSetStyle(payload.iconSetId);
  }
  if (payload.iconSetReverse != null) {
    iconSet.reverseIconOrder = payload.iconSetReverse;
  }
  if (payload.iconSetShowIconOnly != null) {
    iconSet.showIconOnly = payload.iconSetShowIconOnly;
  }
  const criteria = [1, 2, 3, 4]
    .map((index) => {
      const type = payload[`iconThreshold${index}Type`];
      const value = payload[`iconThreshold${index}Value`];
      if (type == null && value == null) return null;
      return {
        type: conditionalIconThresholdType(type ?? "percent"),
        formula: String(value ?? "0"),
        operator: "GreaterThanOrEqual"
      };
    })
    .filter(Boolean);
  if (criteria.length > 0) {
    iconSet.criteria = criteria;
  }
}

function topBottomRule(payload) {
  const rank = requiredPositiveInteger(payload.rank ?? 10, "rank");
  const top = (payload.topBottom ?? "top").toLowerCase();
  if (top !== "top" && top !== "bottom") {
    throw new TypeError("topBottom must be top or bottom.");
  }
  return {
    rank,
    type: `${top === "top" ? "Top" : "Bottom"}${payload.top10Percent ? "Percent" : "Items"}`
  };
}

function presetCriterion(ruleType, payload) {
  if (ruleType === "uniquevalues") return "UniqueValues";
  if (ruleType === "blankscondition") return "Blanks";
  if (ruleType === "aboveaverage") {
    return enumValue(payload.aboveBelow ?? "aboveAverage", {
      aboveaverage: "AboveAverage",
      belowaverage: "BelowAverage",
      abovestddev: "OneStdDevAboveAverage",
      belowstddev: "OneStdDevBelowAverage",
      equalaboveaverage: "EqualOrAboveAverage",
      equalbelowaverage: "EqualOrBelowAverage"
    }, "aboveBelow");
  }
  return enumValue(requiredString(payload.datePeriod, "datePeriod"), {
    today: "Today",
    yesterday: "Yesterday",
    tomorrow: "Tomorrow",
    last7days: "LastSevenDays",
    thisweek: "ThisWeek",
    lastweek: "LastWeek",
    nextweek: "NextWeek",
    thismonth: "ThisMonth",
    lastmonth: "LastMonth",
    nextmonth: "NextMonth"
  }, "datePeriod");
}

function conditionalThresholdType(value, colorScale) {
  return enumValue(value, {
    automaticminimum: "Automatic",
    automaticmaximum: "Automatic",
    minimum: "LowestValue",
    maximum: "HighestValue",
    lowestvalue: "LowestValue",
    highestvalue: "HighestValue",
    number: "Number",
    percent: "Percent",
    percentile: "Percentile",
    formula: "Formula"
  }, colorScale ? "color scale criterion type" : "data bar threshold type");
}

function conditionalIconThresholdType(value) {
  return enumValue(value, {
    number: "Number",
    percent: "Percent",
    percentile: "Percentile",
    formula: "Formula"
  }, "icon threshold type");
}

function iconSetStyle(value) {
  const normalized = value.replaceAll("-", "").toLowerCase();
  const aliases = {
    "3arrows": "ThreeArrows",
    "3trafficlights1": "ThreeTrafficLights1",
    "4ratings": "FourRating",
    "4rating": "FourRating",
    "5quarters": "FiveQuarters"
  };
  return aliases[normalized] ?? value;
}

function enumValue(value, values, name) {
  const normalized = requiredString(value, name).replaceAll("-", "").toLowerCase();
  if (!values[normalized]) {
    throw new TypeError(`Unsupported ${name} '${value}'.`);
  }
  return values[normalized];
}

function conditionalOperator(value) {
  const normalized = requiredString(value, "operatorType").replaceAll("-", "").toLowerCase();
  const operators = {
    equal: "EqualTo",
    notequal: "NotEqualTo",
    greater: "GreaterThan",
    less: "LessThan",
    greaterequal: "GreaterThanOrEqual",
    lessequal: "LessThanOrEqual",
    between: "Between",
    notbetween: "NotBetween"
  };
  if (!operators[normalized]) {
    throw new TypeError(`Unsupported conditional formatting operator '${value}'.`);
  }
  return operators[normalized];
}

function conditionalTypeName(value) {
  return {
    cellvalue: "cellValue",
    custom: "expression",
    colorscale: "colorScale",
    databar: "dataBar",
    iconset: "iconSet",
    topbottom: "top10",
    presetcriteria: "presetCriteria"
  }[String(value).replace(/^ConditionalFormatType\./, "").toLowerCase()] ?? String(value);
}

function resolvedConditionalType(detail) {
  if (detail.type !== "presetCriteria") return detail.type;
  const criterion = String(detail.source.rule.criterion);
  if (["UniqueValues", "DuplicateValues"].includes(criterion)) return "uniqueValues";
  if (["Blanks", "NonBlanks"].includes(criterion)) return "blanksCondition";
  if (criterion.includes("Average")) return "aboveAverage";
  return "timePeriod";
}

function readConditionalFormat(format) {
  return {
    ...(format.fill.color ? { interiorColor: format.fill.color } : {}),
    ...(format.font.color ? { fontColor: format.font.color } : {}),
    ...(format.font.bold == null ? {} : { fontBold: format.font.bold }),
    ...(format.font.italic == null ? {} : { fontItalic: format.font.italic })
  };
}

function conditionalOperatorName(value) {
  return {
    EqualTo: "equal",
    NotEqualTo: "notEqual",
    GreaterThan: "greater",
    LessThan: "less",
    GreaterThanOrEqual: "greaterEqual",
    LessThanOrEqual: "lessEqual",
    Between: "between",
    NotBetween: "notBetween"
  }[String(value)] ?? lowerFirst(String(value));
}

function conditionalThresholdName(value) {
  return {
    Automatic: "automatic",
    LowestValue: "minimum",
    HighestValue: "maximum"
  }[String(value)] ?? lowerFirst(String(value));
}

function conditionalIconOperatorName(value) {
  return String(value) === "GreaterThan" ? "greater" : "greaterEqual";
}

function aboveBelowName(value) {
  return {
    OneStdDevAboveAverage: "aboveStdDev",
    OneStdDevBelowAverage: "belowStdDev",
    EqualOrAboveAverage: "equalAboveAverage",
    EqualOrBelowAverage: "equalBelowAverage"
  }[value] ?? lowerFirst(value);
}

function datePeriodName(value) {
  return value === "LastSevenDays" ? "last7Days" : lowerFirst(value);
}

function iconSetName(value) {
  return String(value)
    .replace(/^Three/, "3")
    .replace(/^Four/, "4")
    .replace(/^Five/, "5")
    .replace("Rating", "Ratings");
}

function conditionalBorderStyle(value) {
  return enumValue(value, {
    none: "None",
    continuous: "Continuous",
    dash: "Dash",
    dot: "Dot",
    dashdot: "DashDot",
    dashdotdot: "DashDotDot",
    double: "Double",
    slantdashdot: "SlantDashDot"
  }, "borderStyle");
}

function lowerFirst(value) {
  return value.length === 0 ? value : value[0].toLowerCase() + value.slice(1);
}

function normalizeTotalFunction(value) {
  const normalized = requiredString(value, "totalFunction").replaceAll("-", "").toLowerCase();
  const values = {
    sum: "Sum",
    count: "Count",
    average: "Average",
    min: "Min",
    max: "Max",
    countnums: "CountNumbers",
    stddev: "StandardDeviation",
    var: "Variance",
    none: "None"
  };
  if (!values[normalized]) {
    throw new TypeError(`Unsupported totals function '${value}'.`);
  }
  return values[normalized];
}

function quoteSheetName(name) {
  return `'${name.replaceAll("'", "''")}'`;
}

function requiredString(value, name) {
  if (typeof value !== "string" || value.length === 0) {
    throw new TypeError(`${name} must be a non-empty string.`);
  }
  return value;
}

function requiredBoolean(value, name) {
  if (typeof value !== "boolean") {
    throw new TypeError(`${name} must be a boolean.`);
  }
  return value;
}

function requiredPositiveInteger(value, name) {
  if (!Number.isInteger(value) || value < 1) {
    throw new TypeError(`${name} must be a positive integer.`);
  }
  return value;
}

function nonNegativeInteger(value, name) {
  if (!Number.isInteger(value) || value < 0) {
    throw new TypeError(`${name} must be a non-negative integer.`);
  }
  return value;
}

function finiteNumber(value, name) {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new TypeError(`${name} must be a finite number.`);
  }
  return value;
}

function positiveNumber(value, name) {
  const result = finiteNumber(value, name);
  if (result <= 0) {
    throw new TypeError(`${name} must be greater than zero.`);
  }
  return result;
}

function nonNegativeNumber(value, name) {
  const result = finiteNumber(value, name);
  if (result < 0) {
    throw new TypeError(`${name} must be non-negative.`);
  }
  return result;
}

function optionalFiniteNumber(value, fallback, name) {
  return value == null ? fallback : finiteNumber(value, name);
}

function optionalPositiveNumber(value, fallback, name) {
  return value == null ? fallback : positiveNumber(value, name);
}

function throwType(message) {
  throw new TypeError(message);
}

function success(value = {}) {
  return { success: true, errorMessage: null, ...value };
}

function operation(action, message) {
  return success({ action, message });
}
