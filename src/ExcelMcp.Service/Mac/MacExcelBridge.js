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

function normalizeMatrix(value) {
    return Array.isArray(value) ? value : [[value]];
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
        if (command === "session.create") {
            const workbook = excel.Workbook().make();
            const format = args.macroEnabled ? "macro enabled XML file format" : "Excel XML file format";
            excel.saveWorkbookAs(workbook, { filename: args.filePath, fileFormat: format });
            workbook.windows[0].visible = !!args.show;
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

        const workbook = workbookByPath(excel, args.filePath);
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
        if (command === "sheet.create") {
            requireSupported(command, []);
        }
        if (command === "sheet.rename") {
            const sheet = worksheetByName(workbook, args.oldName);
            sheet.name = args.newName;
            return json({ success: true, filePath: args.filePath, oldName: args.oldName, newName: args.newName });
        }
        if (command === "sheet.delete") {
            worksheetByName(workbook, args.sheetName).delete();
            return json({ success: true, filePath: args.filePath, sheetName: args.sheetName });
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
                range.clear();
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
        }

        if (command === "calculation.calculate") {
            if (args.scope === "range") {
                worksheetByName(workbook, args.sheetName).ranges.byName(args.rangeAddress).calculate();
            } else {
                excel.calculate(workbook);
            }
            return json({ success: true, filePath: args.filePath, action: "calculate" });
        }

        if (command === "vba.run") {
            const parameters = args.parameters || [];
            if (parameters.length > 30) {
                throw new Error("Excel supports at most 30 VBA procedure parameters.");
            }
            const options = {};
            for (let index = 0; index < parameters.length; index++) {
                options[`arg${index + 1}`] = parameters[index];
            }
            const workbookName = workbook.name().replace(/'/g, "''");
            const qualifiedName = `'${workbookName}'!${args.procedureName}`;
            const result = parameters.length === 0
                ? excel.runVBMacro(qualifiedName)
                : excel.runVBMacro(qualifiedName, options);
            return json({
                success: true,
                filePath: args.filePath,
                action: "run",
                message: result === undefined ? "VBA procedure completed." : String(result)
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
