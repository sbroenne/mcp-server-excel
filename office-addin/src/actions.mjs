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

function success(value = {}) {
  return { success: true, errorMessage: null, ...value };
}

function operation(action, message) {
  return success({ action, message });
}
