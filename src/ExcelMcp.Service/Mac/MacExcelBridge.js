ObjC.import("Foundation");
ObjC.import("AppKit");

function json(value) {
    return JSON.stringify(value);
}

function findWorkbookByPath(excel, filePath) {
    const target = $.NSString.stringWithString(filePath).stringByStandardizingPath.js;
    const workbooks = excel.workbooks;
    for (let index = 0; index < workbooks.length; index++) {
        const candidate = workbooks[index];
        const fullName = candidate.fullName();
        if (fullName && $.NSString.stringWithString(fullName).stringByStandardizingPath.js === target) {
            return candidate;
        }
    }
    return null;
}

function findWorkbookByName(excel, filePath) {
    const targetName = $.NSString.stringWithString(filePath).lastPathComponent.js.toLocaleLowerCase();
    const workbooks = excel.workbooks;
    for (let index = 0; index < workbooks.length; index++) {
        const candidate = workbooks[index];
        if (candidate.name().toLocaleLowerCase() === targetName) {
            return candidate;
        }
    }
    return null;
}

function workbookByPath(excel, filePath) {
    const workbook = findWorkbookByPath(excel, filePath);
    if (workbook) return workbook;
    throw new Error("Workbook is not open in this ExcelMcp session.");
}

function worksheetByName(workbook, sheetName) {
    const sheets = workbook.worksheets;
    for (let index = 0; index < sheets.length; index++) {
        if (sheets[index].name() === sheetName) {
            return sheets[index];
        }
    }
    throw new Error(`Worksheet '${sheetName}' does not exist.`);
}

function excelProcessId() {
    const applications = $.NSRunningApplication.runningApplicationsWithBundleIdentifier(
        "com.microsoft.Excel");
    if (Number(applications.count) !== 1) {
        throw new Error("Expected exactly one running Microsoft Excel application process.");
    }
    return Number(applications.objectAtIndex(0).processIdentifier);
}

function workbookWindowIdentity(workbook, windowNumber) {
    if (!Number.isInteger(windowNumber) || windowNumber <= 0) {
        throw new Error("windowNumber must be a positive integer.");
    }
    const windows = workbook.windows;
    for (let index = 0; index < windows.length; index++) {
        if (Number(windows[index].windowNumber()) === windowNumber) {
            const windowId = Number(windows[index].id());
            if (!Number.isInteger(windowId) || windowId <= 0) {
                throw new Error("Excel returned an invalid native window identifier.");
            }
            return { processId: excelProcessId(), windowId, windowNumber };
        }
    }
    throw new Error(`Workbook window number ${windowNumber} does not exist.`);
}

function scenarioByName(sheet, scenarioName) {
    const scenarios = sheet.scenarios;
    for (let index = 0; index < scenarios.length; index++) {
        if (scenarios[index].name() === scenarioName) {
            return scenarios[index];
        }
    }
    throw new Error(`Scenario '${scenarioName}' does not exist on worksheet '${sheet.name()}'.`);
}

function worksheetNames(workbook) {
    const names = [];
    const sheets = workbook.worksheets;
    for (let index = 0; index < sheets.length; index++) {
        names.push(sheets[index].name());
    }
    return names;
}

function findAddedWorksheetName(workbook, previousNames) {
    const currentNames = worksheetNames(workbook);
    let addedName = null;
    for (let index = 0; index < currentNames.length; index++) {
        if (!previousNames.includes(currentNames[index])) {
            if (addedName !== null) {
                throw new Error("Excel created more than one scenario summary worksheet.");
            }
            addedName = currentNames[index];
        }
    }
    if (addedName === null) {
        throw new Error("Excel did not create a scenario summary worksheet.");
    }
    return addedName;
}

function validateScenarioValues(changingRange, values) {
    if (!Array.isArray(values) || values.length === 0) {
        throw new Error("At least one scenario value is required.");
    }
    const changingCellCount = Number(changingRange.countLarge());
    if (changingCellCount > 32) {
        throw new Error("A scenario cannot contain more than 32 changing cells.");
    }
    if (changingCellCount !== values.length) {
        throw new Error(
            `Scenario values count (${values.length}) must match changing cells count (${changingCellCount}).`);
    }
}

const pythonUnavailableMessage =
    "Python in Excel is not available in this Excel session. It requires a licensed Microsoft 365 account " +
    "with the Python in Excel feature enabled and internet access; it is not available with perpetual-license " +
    "Excel (2016/2019/2021/2024) or offline.";
const pythonTransientMarkers = ["#BUSY!", "#CONNECT!", "#BLOCKED!"];
const pythonErrorMessages = {
    "-2146826288": "#NULL! - Invalid intersection of ranges",
    "-2146826281": "#DIV/0! - Division by zero",
    "-2146826273": "#VALUE! - Wrong type of argument",
    "-2146826265": "#REF! - Invalid cell reference",
    "-2146826259": "#NAME? - Unrecognized formula name",
    "-2146826252": "#NUM! - Invalid numeric value",
    "-2146826246": "#N/A - Value not available",
    "-2146826243": "#SPILL! - Dynamic array result cannot spill",
    "-2146826240": "#UNKNOWN! - Excel cannot identify the data type",
    "-2146826239": "#FIELD! - Referenced data field is unavailable",
    "-2146826238": "#CALC! - Excel cannot complete the calculation"
};

function firstScalar(value) {
    let current = value;
    while (Array.isArray(current)) {
        current = current.length ? current[0] : null;
    }
    return current;
}

function pythonFormulaUnavailable(formula, value, text) {
    return /^\s*=(?:_xlfn\.)?PY\(/i.test(formula)
        && (value === -2146826259 || text === "#NAME?");
}

function normalizedPythonFormula(formula) {
    return formula.trim().replace(/^=_xlfn\.PY/i, "=PY");
}

function pythonCellState(range) {
    const value = firstScalar(range.value());
    const rawText = firstScalar(range.text());
    const text = String(rawText == null ? "" : rawText);
    return { value, text };
}

function pythonCalculationDone(excel) {
    const state = excel.calculationState();
    return state === 0 || /done/i.test(String(state));
}

function pythonFormulaReturnType(formula) {
    const match = formula.match(/,\s*(\d+)\s*\)\s*$/);
    return match ? Number.parseInt(match[1], 10) : 0;
}

function pythonRangeAddress(excel, range) {
    return excel.getAddress(range);
}

function pythonErrorResult(result, value, returnType, text) {
    const knownError = pythonErrorMessages[String(value)];
    if (returnType === 1 && !(knownError && text.trim().toLocaleUpperCase() === knownError.split(" - ")[0])) {
        result.success = true;
        result.isPythonObject = true;
        if (/[A-Za-z]/.test(text)) result.typeName = text;
        result.message =
            "Cell holds a Python Object (rich data type such as a DataFrame). " +
            "Value2 cannot expose rich Python object data via COM automation - set returnType=0 " +
            "(Excel Value) instead if you need to read the underlying data.";
        return result;
    }
    if (knownError) {
        result.success = false;
        result.errorMessage = knownError;
        return result;
    }

    result.success = false;
    result.isPythonError = true;
    result.errorMessage = "#PYTHON! - Python code raised an error (syntax or runtime exception)";
    return result;
}

function normalizeMatrix(value) {
    return Array.isArray(value) ? value : [[value]];
}

function flattenValues(value) {
    const matrix = normalizeMatrix(value);
    const result = [];
    for (let row = 0; row < matrix.length; row++) {
        const values = Array.isArray(matrix[row]) ? matrix[row] : [matrix[row]];
        for (let column = 0; column < values.length; column++) {
            result.push(values[column]);
        }
    }
    return result;
}

function repeatMatrix(value, rowCount, columnCount) {
    const result = [];
    for (let row = 0; row < rowCount; row++) {
        const values = [];
        for (let column = 0; column < columnCount; column++) {
            values.push(value);
        }
        result.push(values);
    }
    return result;
}

function absoluteRangeAddress(excel, range) {
    return String(excel.getAddress(range));
}

function rangeTopLeft(excel, range) {
    const address = absoluteRangeAddress(excel, range).split(",")[0];
    const match = address.match(/(?:^|!)\$?([A-Z]+)\$?(\d+)/i);
    if (!match) {
        throw new Error(`Excel returned an unsupported range address '${address}'.`);
    }
    let column = 0;
    for (const character of match[1].toLocaleUpperCase()) {
        column = (column * 26) + character.charCodeAt(0) - 64;
    }
    return { row: Number.parseInt(match[2], 10), column };
}

function columnName(column) {
    let remaining = column;
    let name = "";
    while (remaining > 0) {
        remaining--;
        name = String.fromCharCode(65 + (remaining % 26)) + name;
        remaining = Math.floor(remaining / 26);
    }
    return name;
}

function matrixRange(excel, sheet, anchor, rowCount, columnCount) {
    const start = rangeTopLeft(excel, anchor);
    const endRow = start.row + rowCount - 1;
    const endColumn = start.column + columnCount - 1;
    const address =
        `$${columnName(start.column)}$${start.row}:$${columnName(endColumn)}$${endRow}`;
    return sheet.ranges.byName(address);
}

function rangeValueResult(excel, filePath, sheetName, range, emptyAsNoCells) {
    const rawValue = range.value();
    const values = emptyAsNoCells && rawValue == null ? [] : normalizeMatrix(rawValue);
    return {
        success: true,
        filePath,
        sheetName,
        rangeAddress: absoluteRangeAddress(excel, range),
        values,
        rowCount: values.length,
        columnCount: values.length ? values[0].length : 0,
        cellErrors: []
    };
}

function usedRangeResult(excel, filePath, sheetName, sheet) {
    try {
        return rangeValueResult(excel, filePath, sheetName, sheet.usedRange(), true);
    } catch (error) {
        const message = String(error && error.message ? error.message : error);
        if (!/object you are trying to access does not exist/i.test(message)) {
            throw error;
        }
        return {
            success: true,
            filePath,
            sheetName,
            rangeAddress: "$A$1",
            values: [],
            rowCount: 0,
            columnCount: 0,
            cellErrors: []
        };
    }
}

function requireMatrixShape(matrix, rowCount, columnCount, parameterName) {
    if (!Array.isArray(matrix) || matrix.length !== rowCount) {
        const actualRows = Array.isArray(matrix) ? matrix.length : 0;
        throw new Error(
            `${parameterName} array row count (${actualRows}) doesn't match range row count (${rowCount})`);
    }
    for (let row = 0; row < matrix.length; row++) {
        if (!Array.isArray(matrix[row]) || matrix[row].length !== columnCount) {
            const actualColumns = Array.isArray(matrix[row]) ? matrix[row].length : 0;
            throw new Error(
                `${parameterName} array row ${row + 1} column count (${actualColumns}) ` +
                `doesn't match range column count (${columnCount})`);
        }
    }
}

function requireSupported(command, supported) {
    if (!supported.includes(command)) {
        const error = new Error(
            `Action '${command}' is not yet supported by the macOS Excel backend. ` +
            "The Windows backend remains unchanged; this capability is explicitly gated.");
        error.category = "PlatformNotSupported";
        throw error;
    }
}

function findNamedItem(workbook, name) {
    const items = workbook.namedItems;
    for (let index = 0; index < items.length; index++) {
        if (items[index].name().toLowerCase() === name.toLowerCase()) {
            return items[index];
        }
    }
    return null;
}

function namedRangeValue(range) {
    const raw = range.value();
    const count = Number(range.rows.length) * Number(range.columns.length);
    const value = count === 1 ? firstScalar(raw) : normalizeMatrix(raw);
    let valueType = "null";
    if (Array.isArray(value)) valueType = "Array";
    else if (typeof value === "number") valueType = "Double";
    else if (typeof value === "boolean") valueType = "Boolean";
    else if (typeof value === "string") valueType = "String";
    else if (value != null) {
        throw new Error("Excel returned an unsupported named range value type.");
    }
    return value == null ? { valueType } : { value, valueType };
}

function namedRangeList(workbook, filePath) {
    const result = [];
    const items = workbook.namedItems;
    for (let index = 0; index < items.length; index++) {
        const item = items[index];
        if (!item.visible()) continue;
        const name = item.name();
        const localName = name.substring(name.lastIndexOf("!") + 1).replace(/^'|'$/g, "").toLowerCase();
        if (localName.startsWith("_xlnm.") || localName === "_filterdatabase") continue;
        const info = { name, refersTo: item.references(), valueType: "Unavailable" };
        try {
            const range = item.referenceRange();
            if (range.areas().length > 1) {
                info.valueType = "MultiAreaRange";
                info.valueOmittedReason = "Named range resolves to multiple areas; list omits multi-area value previews.";
            } else {
                const count = Number(range.rows.length) * Number(range.columns.length);
                if (!Number.isSafeInteger(count) || count < 1) {
                    throw new Error("Excel returned an invalid named range cell count.");
                }
                info.cellCount = count;
                if (count > 10000) {
                    info.valueType = "RangeTooLarge";
                    info.valueOmittedReason =
                        `Named range contains ${count} cells, which exceeds the list preview limit of 10000.`;
                } else {
                    Object.assign(info, namedRangeValue(range));
                }
            }
        } catch (error) {
            if (![-1728, -1700, -1004].includes(Number(error.errorNumber))) throw error;
            info.valueOmittedReason = error.message || String(error);
        }
        result.push(info);
    }
    return { success: true, filePath, namedRanges: result };
}

function dispatchNamedRange(excel, workbook, command, args) {
    if (command === "namedrange.list") {
        return namedRangeList(workbook, args.filePath);
    }
    const item = findNamedItem(workbook, args.name);
    if (command === "namedrange.create") {
        if (item) throw new Error(`Named range '${args.name}' already exists.`);
        workbook.namedItems.push(excel.NamedItem({ name: args.name, references: args.reference }));
    } else {
        if (!item) throw new Error(`Named range '${args.name}' not found.`);
        if (command === "namedrange.update") item.references = args.reference;
        else if (command === "namedrange.delete") item.delete();
        else if (command === "namedrange.write") item.referenceRange().value = args.parsedValue;
        else if (command === "namedrange.read") {
            return Object.assign(
                { name: args.name, refersTo: item.references() },
                namedRangeValue(item.referenceRange()));
        } else {
            requireSupported(command, []);
        }
    }
    return { success: true, filePath: args.filePath };
}

function run(argv) {
    const command = argv[0];
    const args = JSON.parse(argv[1] || "{}");
    const excel = Application("Microsoft Excel");
    excel.includeStandardAdditions = false;

    try {
        if (command === "session.prepare-open") {
            if (findWorkbookByPath(excel, args.filePath)) {
                throw new Error("Workbook is already open in shared Excel. Reuse its owning session or close it before opening a new session.");
            }
            if (findWorkbookByName(excel, args.filePath)) {
                throw new Error("A workbook with the same name is already open in shared Excel. Close it before opening this workbook.");
            }
            return json({ success: true, errorMessage: "" });
        }
        if (command === "session.open") {
            let workbook = null;
            for (let attempt = 0; attempt < 100 && !workbook; attempt++) {
                workbook = findWorkbookByPath(excel, args.filePath);
                if (!workbook) delay(0.1);
            }
            if (!workbook) throw new Error("LaunchServices did not open the requested workbook within ten seconds.");
            workbook.windows[0].visible = !!args.show;
            return json({ success: true, errorMessage: "" });
        }
        if (command === "session.close") {
            const workbook = workbookByPath(excel, args.filePath);
            if (command.startsWith("namedrange.")) {
                return json(dispatchNamedRange(excel, workbook, command, args));
            }
            workbook.close({ saving: args.save ? "yes" : "no" });
            return json({ success: true, errorMessage: "" });
        }
        if (command === "session.close-if-saved") {
            const workbook = workbookByPath(excel, args.filePath);
            if (!workbook.saved()) {
                const error = new Error(
                    "Power Query package updates on macOS require a saved workbook. " +
                    "Save or discard the current workbook changes, then retry.");
                error.category = "InvalidOperation";
                throw error;
            }
            workbook.close({ saving: "no" });
            return json({ success: true, errorMessage: "" });
        }
        if (command === "session.is-open") {
            return json({
                success: true,
                errorMessage: "",
                open: !!findWorkbookByPath(excel, args.filePath)
            });
        }

        const workbook = workbookByPath(excel, args.filePath);
        if (command === "workbook.state") {
            return json({
                success: true,
                filePath: args.filePath,
                saved: workbook.saved()
            });
        }
        if (command === "screenshot.window-identity") {
            const identity = workbookWindowIdentity(workbook, args.windowNumber);
            return json({
                success: true,
                filePath: args.filePath,
                processId: identity.processId,
                windowId: identity.windowId,
                windowNumber: identity.windowNumber
            });
        }
        if (command === "sheet.list") {
            const sheets = workbook.worksheets;
            const result = [];
            for (let index = 0; index < sheets.length; index++) {
                result.push({
                    name: sheets[index].name(),
                    index: index + 1,
                    visible: sheets[index].visible() !== false
                });
            }
            return json({ success: true, filePath: args.filePath, worksheets: result });
        }
        if (command === "sheet.rename") {
            const sheet = worksheetByName(workbook, args.oldName);
            sheet.name = args.newName;
            return json({ success: true, filePath: args.filePath, oldName: args.oldName, newName: args.newName });
        }
        if (command === "sheet.set-visibility"
            || command === "sheet.get-visibility"
            || command === "sheet.show"
            || command === "sheet.hide"
            || command === "sheet.very-hide"
            || command === "sheet.set-tab-color"
            || command === "sheet.get-tab-color"
            || command === "sheet.clear-tab-color") {
            const sheet = worksheetByName(workbook, args.sheetName);
            if (command === "sheet.set-visibility") {
                sheet.visible = sheetVisibility(args.visibility).appleEvent;
                return json({ success: true, filePath: args.filePath });
            }
            if (command === "sheet.show") {
                sheet.visible = "sheet visible";
                return json({ success: true, filePath: args.filePath });
            }
            if (command === "sheet.hide") {
                sheet.visible = "sheet hidden";
                return json({ success: true, filePath: args.filePath });
            }
            if (command === "sheet.very-hide") {
                sheet.visible = "sheet very hidden";
                return json({ success: true, filePath: args.filePath });
            }
            if (command === "sheet.get-visibility") {
                const visibility = currentSheetVisibility(sheet);
                return json({
                    success: true,
                    filePath: args.filePath,
                    visibility: visibility.value,
                    visibilityName: visibility.name
                });
            }
            if (command === "sheet.set-tab-color") {
                const red = requireRgb(args.red);
                const green = requireRgb(args.green);
                const blue = requireRgb(args.blue);
                sheet.sheetTab.color = [red, green, blue];
                return json({ success: true, filePath: args.filePath });
            }
            if (command === "sheet.clear-tab-color") {
                sheet.sheetTab.colorIndex = "color index none";
                return json({ success: true, filePath: args.filePath });
            }
            if (command === "sheet.get-tab-color") {
                const tab = sheet.sheetTab;
                const colorIndex = tab.colorIndex();
                const noColor = typeof colorIndex === "number"
                    ? colorIndex < 0
                    : /none|automatic/i.test(String(colorIndex));
                if (noColor) {
                    return json({ success: true, filePath: args.filePath, hasColor: false });
                }

                const color = tab.color();
                const red = color[0];
                const green = color[1];
                const blue = color[2];
                const hex = [red, green, blue]
                    .map(component => component.toString(16).padStart(2, "0"))
                    .join("")
                    .toLocaleUpperCase();
                return json({
                    success: true,
                    filePath: args.filePath,
                    hasColor: true,
                    red,
                    green,
                    blue,
                    hexColor: `#${hex}`
                });
            }
        }
        if (command.startsWith("range.")) {
            if (command === "range.copy"
                || command === "range.copy-values"
                || command === "range.copy-formulas") {
                const sourceSheet = worksheetByName(workbook, args.sourceSheet);
                const targetSheet = worksheetByName(workbook, args.targetSheet);
                const source = sourceSheet.ranges.byName(args.sourceRange);
                const targetAnchor = targetSheet.ranges.byName(args.targetRange);
                const target = matrixRange(
                    excel, targetSheet, targetAnchor, source.rows.length, source.columns.length);
                if (command === "range.copy") {
                    source.copyRange({ destination: target });
                } else if (command === "range.copy-values") {
                    target.value = normalizeMatrix(source.value());
                } else {
                    target.formulaR1c1 = normalizeMatrix(source.formulaR1c1());
                }
                return json({
                    success: true,
                    filePath: args.filePath,
                    action: command.substring("range.".length)
                });
            }

            const sheet = worksheetByName(workbook, args.sheetName);
            if (command === "range.get-used-range") {
                return json(usedRangeResult(excel, args.filePath, args.sheetName, sheet));
            }
            const requestedAddress = command === "range.get-current-region"
                ? args.cellAddress
                : args.rangeAddress;
            const range = sheet.ranges.byName(requestedAddress);
            if (command === "range.get-current-region") {
                return json(rangeValueResult(
                    excel, args.filePath, args.sheetName, range.currentRegion(), true));
            }
            if (command === "range.get-values") {
                const values = normalizeMatrix(range.value());
                return json({
                    success: true,
                    filePath: args.filePath,
                    sheetName: args.sheetName,
                    rangeAddress: range.address(),
                    values,
                    rowCount: values.length,
                    columnCount: values.length ? values[0].length : 0,
                    cellErrors: []
                });
            }
            if (command === "range.set-values") {
                range.value = args.values;
                return json({ success: true, filePath: args.filePath, action: "set-values" });
            }
            if (command === "range.get-formulas") {
                const formulas = normalizeMatrix(range.formula());
                const values = normalizeMatrix(range.value());
                return json({
                    success: true,
                    filePath: args.filePath,
                    sheetName: args.sheetName,
                    rangeAddress: range.address(),
                    formulas,
                    values,
                    rowCount: formulas.length,
                    columnCount: formulas.length ? formulas[0].length : 0,
                    cellErrors: []
                });
            }
            if (command === "range.set-formulas") {
                range.formula = args.formulas;
                return json({ success: true, filePath: args.filePath, action: "set-formulas" });
            }
            if (command === "range.clear-all") {
                range.clearRange();
                return json({ success: true, filePath: args.filePath, action: "clear-all" });
            }
            if (command === "range.clear-contents") {
                range.clearContents();
                return json({ success: true, filePath: args.filePath, action: "clear-contents" });
            }
            if (command === "range.clear-formats") {
                range.clearFormats();
                return json({ success: true, filePath: args.filePath, action: "clear-formats" });
            }
            if (command === "range.get-number-formats") {
                const rowCount = range.rows.length;
                const columnCount = range.columns.length;
                const rawFormats = range.numberFormat();
                let formats;
                if (Array.isArray(rawFormats)) {
                    formats = normalizeMatrix(rawFormats);
                } else if (rawFormats == null) {
                    formats = [];
                    const start = rangeTopLeft(excel, range);
                    for (let row = 0; row < rowCount; row++) {
                        const rowFormats = [];
                        for (let column = 0; column < columnCount; column++) {
                            const address =
                                `$${columnName(start.column + column)}$${start.row + row}`;
                            const cell = sheet.ranges.byName(address);
                            rowFormats.push(String(cell.numberFormat() || "General"));
                        }
                        formats.push(rowFormats);
                    }
                } else {
                    formats = repeatMatrix(rawFormats || "General", rowCount, columnCount);
                }
                return json({
                    success: true,
                    filePath: args.filePath,
                    sheetName: args.sheetName,
                    rangeAddress: range.address(),
                    formats,
                    rowCount,
                    columnCount
                });
            }
            if (command === "range.set-number-format") {
                range.numberFormat = args.formatCode;
                return json({ success: true, filePath: args.filePath, action: "set-number-format" });
            }
            if (command === "range.set-number-formats") {
                const rowCount = range.rows.length;
                const columnCount = range.columns.length;
                requireMatrixShape(args.formats, rowCount, columnCount, "Format");
                const start = rangeTopLeft(excel, range);
                for (let row = 0; row < rowCount; row++) {
                    for (let column = 0; column < columnCount; column++) {
                        const address =
                            `$${columnName(start.column + column)}$${start.row + row}`;
                        const cell = sheet.ranges.byName(address);
                        const requestedFormat = args.formats[row][column];
                        cell.numberFormat = requestedFormat;
                        const storedFormat = sheet.ranges.byName(address).numberFormat();
                        if (String(storedFormat) !== requestedFormat) {
                            throw new Error(
                                `Excel did not preserve number format '${requestedFormat}' for ${address}.`);
                        }
                    }
                }
                return json({ success: true, filePath: args.filePath, action: "set-number-formats" });
            }
            if (command === "range.get-info") {
                const rawNumberFormat = range.numberFormat();
                const numberFormat = firstScalar(rawNumberFormat);
                return json({
                    success: true,
                    filePath: args.filePath,
                    sheetName: args.sheetName,
                    address: absoluteRangeAddress(excel, range),
                    rowCount: range.rows.length,
                    columnCount: range.columns.length,
                    numberFormat: numberFormat == null ? null : String(numberFormat),
                    left: range.leftPosition(),
                    top: range.top(),
                    width: range.width(),
                    height: range.height()
                });
            }
        }

        if (command.startsWith("rangeformat.")) {
            const sheet = worksheetByName(workbook, args.sheetName);
            const range = sheet.ranges.byName(args.rangeAddress);
            if (command === "rangeformat.set-column-width") {
                range.columnWidth = args.columnWidth;
                return json({ success: true, filePath: args.filePath, action: "set-column-width" });
            }
            if (command === "rangeformat.set-row-height") {
                range.rowHeight = args.rowHeight;
                return json({ success: true, filePath: args.filePath, action: "set-row-height" });
            }
            if (command === "rangeformat.auto-fit-columns") {
                range.columns.autofit();
                return json({ success: true, filePath: args.filePath, action: "auto-fit-columns" });
            }
            if (command === "rangeformat.auto-fit-rows") {
                range.rows.autofit();
                return json({ success: true, filePath: args.filePath, action: "auto-fit-rows" });
            }
            if (command === "rangeformat.merge-cells") {
                range.merge();
                return json({ success: true, filePath: args.filePath, action: "merge-cells" });
            }
            if (command === "rangeformat.unmerge-cells") {
                range.unmerge();
                return json({ success: true, filePath: args.filePath, action: "unmerge-cells" });
            }
            if (command === "rangeformat.get-merge-info") {
                const addresses = [];
                const seen = new Set();
                if (range.mergeCells() !== false) {
                    const cells = range.cells;
                    for (let index = 0; index < cells.length; index++) {
                        const cell = cells[index];
                        if (cell.mergeCells() === true) {
                            const address = absoluteRangeAddress(excel, cell.mergeArea());
                            if (!seen.has(address)) {
                                seen.add(address);
                                addresses.push(address);
                            }
                        }
                    }
                }
                return json({
                    success: true,
                    filePath: args.filePath,
                    sheetName: args.sheetName,
                    rangeAddress: args.rangeAddress,
                    isMerged: addresses.length > 0,
                    mergedRanges: addresses
                });
            }
        }

        if (command.startsWith("rangelink.")) {
            const range = worksheetByName(workbook, args.sheetName).ranges.byName(args.rangeAddress);
            if (command === "rangelink.set-cell-lock") {
                range.locked = args.locked;
                return json({ success: true, filePath: args.filePath, action: "set-cell-lock" });
            }
            if (command === "rangelink.get-cell-lock") {
                return json({
                    success: true,
                    filePath: args.filePath,
                    sheetName: args.sheetName,
                    rangeAddress: args.rangeAddress,
                    isLocked: Boolean(range.cells[0].locked())
                });
            }
        }

        if (command === "pythoninexcel.set-formula" || command === "pythoninexcel.get-result") {
            const range = worksheetByName(workbook, args.sheetName).ranges.byName(args.rangeAddress);
            if (command === "pythoninexcel.set-formula") {
                const escapedCode = String(args.code).replace(/"/g, "\"\"");
                const formula = `=PY("${escapedCode}",${args.returnType})`;
                range.formula2 = [[formula]];
                try {
                    range.calculate();
                } catch (_) {
                    // Excel may already be calculating asynchronously.
                }

                const verificationRange =
                    worksheetByName(workbook, args.sheetName).ranges.byName(args.rangeAddress);
                const state = pythonCellState(verificationRange);
                const storedFormula = String(firstScalar(verificationRange.formula2()) || "");
                if (normalizedPythonFormula(storedFormula) !== formula) {
                    return json({
                        success: false,
                        filePath: args.filePath,
                        action: "set-formula",
                        errorMessage:
                            "Excel did not preserve the requested PY() formula through Range.Formula2. " +
                            "The formula was not accepted as a reliable Python in Excel serialization."
                    });
                }
                if (pythonFormulaUnavailable(storedFormula, state.value, state.text)) {
                    return json({
                        success: false,
                        filePath: args.filePath,
                        action: "set-formula",
                        errorMessage: pythonUnavailableMessage
                    });
                }
                return json({
                    success: true,
                    filePath: args.filePath,
                    action: "set-formula",
                    message:
                        `Set Python in Excel formula on '${pythonRangeAddress(excel, verificationRange)}'. ` +
                        "Use get-result to read the computed value once the cloud Python backend finishes."
                });
            }

            const formula = String(firstScalar(range.formula2()) || "");
            const result = {
                success: false,
                filePath: args.filePath,
                sheetName: args.sheetName,
                rangeAddress: pythonRangeAddress(excel, range),
                formula,
                isPythonObject: false,
                isPythonError: false
            };
            if (!/PY\(/i.test(formula)) {
                result.errorMessage =
                    `Cell '${result.rangeAddress}' does not contain a Python in Excel (PY()) formula.`;
                return json(result);
            }

            try {
                excel.calculate(workbook);
            } catch (_) {
                // Calculation may already be running; polling below remains authoritative.
            }

            const deadline = Date.now() + (args.maxWaitSeconds * 1000);
            let state = { value: null, text: "" };
            let nonBusyReads = 0;
            let calculationDone = false;
            let converged = false;
            let lastMarker = "";
            do {
                state = pythonCellState(range);
                calculationDone = pythonCalculationDone(excel);
                if (pythonFormulaUnavailable(formula, state.value, state.text)) {
                    result.errorMessage = pythonUnavailableMessage;
                    return json(result);
                }

                const textMarker = pythonTransientMarkers.includes(state.text) ? state.text : "";
                const cellBusy = state.value === -2146826237 || textMarker.length > 0;
                lastMarker = state.value === -2146826237 ? "#BUSY!" : textMarker;
                nonBusyReads = cellBusy ? 0 : nonBusyReads + 1;
                if (!cellBusy && (calculationDone || nonBusyReads >= 3)) {
                    converged = true;
                    break;
                }

                const remainingSeconds = (deadline - Date.now()) / 1000;
                if (remainingSeconds > 0) delay(Math.min(0.5, remainingSeconds));
            } while (Date.now() < deadline);

            if (!converged) {
                const observed = lastMarker.length
                    ? `the cell still reads as ${lastMarker}`
                    : !calculationDone
                        ? "the workbook is still calculating"
                        : "the result did not settle";
                const guidance = lastMarker === "#CONNECT!"
                    ? "Excel could not connect to the Microsoft-hosted Python service. Check internet access, " +
                        "the signed-in Microsoft 365 account, and connected experiences, then retry."
                    : lastMarker === "#BLOCKED!"
                        ? "Excel blocked a required cloud resource. Check the account's Python in Excel license " +
                            "and organization-managed privacy, security, and connected-service policies."
                        : "The Microsoft-hosted Python backend may be under cold-start load - call get-result " +
                            "again, or increase maxWaitSeconds.";
                result.errorMessage =
                    `Python in Excel result did not finish within ${args.maxWaitSeconds}s (${observed}). ${guidance}`;
                return json(result);
            }

            const returnType = pythonFormulaReturnType(formula);
            if (Number.isInteger(state.value) && state.value < 0) {
                return json(pythonErrorResult(result, state.value, returnType, state.text));
            }
            result.success = true;
            result.value = state.value;
            return json(result);
        }

        if (command === "calculation.calculate") {
            if (args.scope === "range") {
                worksheetByName(workbook, args.sheetName).ranges.byName(args.rangeAddress).calculate();
            } else {
                excel.calculate(workbook);
            }
            return json({ success: true, filePath: args.filePath, action: "calculate" });
        }

        if (command === "analysis.goal-seek") {
            if (!args.sheetName) throw new Error("sheetName is required.");
            if (!args.formulaCell) throw new Error("formulaCell is required.");
            if (!Number.isFinite(args.goal)) throw new Error("goal must be a finite number.");
            if (!args.changingCell) throw new Error("changingCell is required.");

            const sheet = worksheetByName(workbook, args.sheetName);
            const formulaRange = sheet.ranges.byName(args.formulaCell);
            const changingRange = sheet.ranges.byName(args.changingCell);
            const converged = formulaRange.goalSeek({
                goal: args.goal,
                changingCell: changingRange
            });
            return json({
                success: true,
                converged: !!converged,
                formulaValue: Number(formulaRange.value()),
                changingValue: Number(changingRange.value()),
                message: converged
                    ? `Goal Seek reached ${args.goal} in '${args.formulaCell}'.`
                    : `Goal Seek completed without converging on ${args.goal} in '${args.formulaCell}'.`
            });
        }

        if (command === "analysis.list-scenarios") {
            if (!args.sheetName) throw new Error("sheetName is required.");

            const sheet = worksheetByName(workbook, args.sheetName);
            const scenarios = sheet.scenarios;
            const result = [];
            for (let index = 0; index < scenarios.length; index++) {
                const scenario = scenarios[index];
                result.push({
                    name: scenario.name(),
                    changingCells: scenario.changingCells().address(),
                    values: flattenValues(scenario.getValues()),
                    comment: scenario.excelComment() || "",
                    locked: !!scenario.locked(),
                    hidden: !!scenario.hidden()
                });
            }
            return json({
                success: true,
                sheetName: args.sheetName,
                scenarios: result,
                message: `Found ${result.length} scenario(s) on '${args.sheetName}'.`
            });
        }

        if (command === "analysis.update-scenario") {
            if (!args.sheetName) throw new Error("sheetName is required.");
            if (!args.scenarioName) throw new Error("scenarioName is required.");
            if (!args.changingCells) throw new Error("changingCells is required.");

            const sheet = worksheetByName(workbook, args.sheetName);
            const changingRange = sheet.ranges.byName(args.changingCells);
            validateScenarioValues(changingRange, args.values);
            scenarioByName(sheet, args.scenarioName).changeScenario({
                changingCells: changingRange,
                values: args.values
            });
            return json({
                success: true,
                message: `Scenario '${args.scenarioName}' updated on '${args.sheetName}'.`
            });
        }

        if (command === "analysis.delete-scenario") {
            if (!args.sheetName) throw new Error("sheetName is required.");
            if (!args.scenarioName) throw new Error("scenarioName is required.");

            const sheet = worksheetByName(workbook, args.sheetName);
            scenarioByName(sheet, args.scenarioName).delete();
            return json({
                success: true,
                message: `Scenario '${args.scenarioName}' deleted on '${args.sheetName}'.`
            });
        }

        if (command === "analysis.create-scenario-summary") {
            if (!args.sheetName) throw new Error("sheetName is required.");
            const reportType = args.reportType || "summary";
            if (reportType !== "summary" && reportType !== "pivot-table") {
                throw new Error("reportType must be 'summary' or 'pivot-table'.");
            }

            const sheet = worksheetByName(workbook, args.sheetName);
            const existingSheetNames = worksheetNames(workbook);
            const parameters = {
                reportType: reportType === "summary" ? "standard summary" : "summary pivot table"
            };
            if (args.resultCells) {
                parameters.resultCells = sheet.ranges.byName(args.resultCells);
            }
            sheet.createSummaryForScenarios(parameters);
            const reportSheetName = findAddedWorksheetName(workbook, existingSheetNames);
            return json({
                success: true,
                reportSheetName,
                reportType,
                message: `Scenario ${reportType} report created.`
            });
        }

        if (command === "analysis.create-data-table") {
            if (!args.sheetName) throw new Error("sheetName is required.");
            if (!args.tableRange) throw new Error("tableRange is required.");
            if (!args.rowInputCell && !args.columnInputCell) {
                throw new Error("Provide rowInputCell, columnInputCell, or both for a data table.");
            }

            const sheet = worksheetByName(workbook, args.sheetName);
            const parameters = {};
            if (args.rowInputCell) {
                parameters.rowInput = sheet.ranges.byName(args.rowInputCell);
            }
            if (args.columnInputCell) {
                parameters.columnInput = sheet.ranges.byName(args.columnInputCell);
            }
            sheet.ranges.byName(args.tableRange).dataTable(parameters);
            return json({
                success: true,
                filePath: args.filePath,
                action: "create-data-table",
                message: `Data table created in '${args.sheetName}'!${args.tableRange}.`
            });
        }

        requireSupported(command, []);
    } catch (error) {
        return json({
            success: false,
            errorMessage: error.message || String(error),
            errorCategory: error.category || "ComInterop",
            errorNumber: error.errorNumber || null
        });
    }
}
