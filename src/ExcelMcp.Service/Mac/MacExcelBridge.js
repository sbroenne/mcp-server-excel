ObjC.import("Foundation");

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
    if (applications.count !== 1) {
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

function requireSupported(command, supported) {
    if (!supported.includes(command)) {
        const error = new Error(
            `Action '${command}' is not yet supported by the macOS Excel backend. ` +
            "The Windows backend remains unchanged; this capability is explicitly gated.");
        error.category = "PlatformNotSupported";
        throw error;
    }
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
        if (command.startsWith("range.")) {
            const sheet = worksheetByName(workbook, args.sheetName);
            const range = sheet.ranges.byName(args.rangeAddress);
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
                const formats = Array.isArray(rawFormats)
                    ? normalizeMatrix(rawFormats)
                    : repeatMatrix(rawFormats || "General", rowCount, columnCount);
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
