using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel range operations - values, formulas, clearing, copying, inserting, deleting, finding, sorting, and hyperlinks.
/// </summary>
[McpServerToolType]
public static partial class ExcelRangeTool
{
    /// <summary>
    /// Unified Excel range operations - ALL data manipulation.
    /// DATA FORMAT: Values/formulas are JSON 2D arrays [[row1col1, row1col2], [row2col1, row2col2]]. Example single cell: [[100]] or [['=SUM(A:A)']]. Example range: [[1,2,3], [4,5,6], [7,8,9]].
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="excelPath">Excel file path (.xlsx or .xlsm)</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action (required for all range operations)</param>
    /// <param name="sheetName">Worksheet name (empty for named ranges, required for most operations)</param>
    /// <param name="rangeAddress">Range address (e.g., 'A1:D10') or named range (e.g., 'SalesData'). For named ranges, leave sheetName empty.</param>
    /// <param name="values">2D array of values for set-values (JSON array of arrays, e.g., [[1,2],[3,4]])</param>
    /// <param name="formulas">2D array of formulas for set-formulas (JSON array of arrays, e.g., [["=A1+B1","=SUM(A:A)"]])</param>
    /// <param name="sourceSheet">Source sheet name (for copy operations)</param>
    /// <param name="sourceRange">Source range address (for copy operations)</param>
    /// <param name="targetSheet">Target sheet name (for copy operations)</param>
    /// <param name="targetRange">Target range address (for copy operations)</param>
    /// <param name="shift">Shift direction for insert-cells/delete-cells: Down, Right, Up, Left</param>
    /// <param name="searchValue">Search value (for find/replace operations)</param>
    /// <param name="replaceValue">Replace value (for replace operation)</param>
    /// <param name="matchCase">Match case (for find/replace, default: false)</param>
    /// <param name="matchEntireCell">Match entire cell (for find/replace, default: false)</param>
    /// <param name="searchFormulas">Search formulas (for find/replace, default: true)</param>
    /// <param name="searchValues">Search values (for find/replace, default: true)</param>
    /// <param name="replaceAll">Replace all occurrences (for replace, default: true)</param>
    /// <param name="sortColumns">Sort columns (JSON array, e.g., [{"columnIndex":1,"ascending":true}])</param>
    /// <param name="hasHeaders">Has header row (for sort, default: true)</param>
    /// <param name="cellAddress">Cell address for single-cell operations (hyperlinks, current-region)</param>
    /// <param name="url">Hyperlink URL (for add-hyperlink)</param>
    /// <param name="displayText">Hyperlink display text (for add-hyperlink, optional)</param>
    /// <param name="tooltip">Hyperlink tooltip (for add-hyperlink, optional)</param>
    /// <param name="formatCode">Excel format code for set-number-format. ALWAYS use US format codes (e.g., '$#,##0.00', '0.00%', 'm/d/yyyy'). The server auto-translates to the user's locale.</param>
    /// <param name="formats">2D array of format codes for set-number-formats (JSON array of arrays, e.g., [['$#,##0','0.00%'],['m/d/yyyy','General']])</param>
    /// <param name="styleName">Built-in Excel style name (for set-style: 'Heading 1', 'Accent1', 'Good', 'Total', 'Currency', 'Percent', 'Normal', etc. - recommended for consistent formatting)</param>
    /// <param name="fontName">Font name (for format-range, e.g., 'Arial', 'Calibri')</param>
    /// <param name="fontSize">Font size (for format-range, e.g., 11, 12, 14)</param>
    /// <param name="bold">Bold font (for format-range)</param>
    /// <param name="italic">Italic font (for format-range)</param>
    /// <param name="underline">Underline font (for format-range)</param>
    /// <param name="fontColor">Font color (for format-range, #RRGGBB or color index)</param>
    /// <param name="fillColor">Fill color (for format-range, #RRGGBB or color index)</param>
    /// <param name="borderStyle">Border style (for format-range: none, continuous, dash, dot, double, etc.)</param>
    /// <param name="borderColor">Border color (for format-range, #RRGGBB or color index)</param>
    /// <param name="borderWeight">Border weight (for format-range: hairline, thin, medium, thick)</param>
    /// <param name="horizontalAlignment">Horizontal alignment (for format-range: left, center, right, justify, distributed)</param>
    /// <param name="verticalAlignment">Vertical alignment (for format-range: top, center, bottom, justify, distributed)</param>
    /// <param name="wrapText">Wrap text in cells (for format-range)</param>
    /// <param name="orientation">Text orientation in degrees (for format-range, 0-90 or -90)</param>
    /// <param name="validationType">Data validation type (for validate-range: list, whole, decimal, date, time, textLength, custom)</param>
    /// <param name="validationOperator">Data validation operator (for validate-range: between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual)</param>
    /// <param name="validationFormula1">Validation formula1 (for validate-range). For 'list' type: MUST be worksheet range reference like '=$A$1:$A$10' to create dropdown. For other types: value/formula for comparison.</param>
    /// <param name="validationFormula2">Validation formula2 (for validate-range, second value/formula for between/notBetween)</param>
    /// <param name="showInputMessage">Show input message (for validate-range)</param>
    /// <param name="inputTitle">Input message title (for validate-range)</param>
    /// <param name="inputMessage">Input message text (for validate-range)</param>
    /// <param name="showErrorAlert">Show error alert (for validate-range)</param>
    /// <param name="errorStyle">Error alert style (for validate-range: stop, warning, information)</param>
    /// <param name="errorTitle">Error alert title (for validate-range)</param>
    /// <param name="errorMessage">Error alert message (for validate-range)</param>
    /// <param name="ignoreBlank">Ignore blank cells in validation (for validate-range)</param>
    /// <param name="showDropdown">Show dropdown for list validation (for validate-range)</param>
    /// <param name="locked">Lock status for cells (for set-cell-lock: true = locked, false = unlocked)</param>
    [McpServerTool(Name = "excel_range")]
    [McpMeta("category", "data")]
    public static partial string ExcelRange(
        RangeAction action,
        string excelPath,
        string sessionId,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? rangeAddress,
        [DefaultValue(null)] List<List<object?>>? values,
        [DefaultValue(null)] List<List<string>>? formulas,
        [DefaultValue(null)] string? sourceSheet,
        [DefaultValue(null)] string? sourceRange,
        [DefaultValue(null)] string? targetSheet,
        [DefaultValue(null)] string? targetRange,
        [DefaultValue(null)] string? shift,
        [DefaultValue(null)] string? searchValue,
        [DefaultValue(null)] string? replaceValue,
        [DefaultValue(null)] bool? matchCase,
        [DefaultValue(null)] bool? matchEntireCell,
        [DefaultValue(null)] bool? searchFormulas,
        [DefaultValue(null)] bool? searchValues,
        [DefaultValue(null)] bool? replaceAll,
        [DefaultValue(null)] List<SortColumn>? sortColumns,
        [DefaultValue(null)] bool? hasHeaders,
        [DefaultValue(null)] string? cellAddress,
        [DefaultValue(null)] string? url,
        [DefaultValue(null)] string? displayText,
        [DefaultValue(null)] string? tooltip,
        [DefaultValue(null)] string? formatCode,
        [DefaultValue(null)] List<List<string>>? formats,
        [DefaultValue(null)] string? styleName,
        [DefaultValue(null)] string? fontName,
        [DefaultValue(null)] double? fontSize,
        [DefaultValue(null)] bool? bold,
        [DefaultValue(null)] bool? italic,
        [DefaultValue(null)] bool? underline,
        [DefaultValue(null)] string? fontColor,
        [DefaultValue(null)] string? fillColor,
        [DefaultValue(null)] string? borderStyle,
        [DefaultValue(null)] string? borderColor,
        [DefaultValue(null)] string? borderWeight,
        [DefaultValue(null)] string? horizontalAlignment,
        [DefaultValue(null)] string? verticalAlignment,
        [DefaultValue(null)] bool? wrapText,
        [DefaultValue(null)] int? orientation,
        [DefaultValue(null)] string? validationType,
        [DefaultValue(null)] string? validationOperator,
        [DefaultValue(null)] string? validationFormula1,
        [DefaultValue(null)] string? validationFormula2,
        [DefaultValue(null)] bool? showInputMessage,
        [DefaultValue(null)] string? inputTitle,
        [DefaultValue(null)] string? inputMessage,
        [DefaultValue(null)] bool? showErrorAlert,
        [DefaultValue(null)] string? errorStyle,
        [DefaultValue(null)] string? errorTitle,
        [DefaultValue(null)] string? errorMessage,
        [DefaultValue(null)] bool? ignoreBlank,
        [DefaultValue(null)] bool? showDropdown,
        [DefaultValue(null)] bool? locked)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_range",
            action.ToActionString(),
            excelPath,
            () =>
            {
                var rangeCommands = new RangeCommands();

                // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
                return action switch
                {
                    RangeAction.GetValues => GetValuesAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.SetValues => SetValuesAsync(rangeCommands, sessionId, sheetName, rangeAddress, values),
                    RangeAction.GetFormulas => GetFormulasAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.SetFormulas => SetFormulasAsync(rangeCommands, sessionId, sheetName, rangeAddress, formulas),
                    RangeAction.GetNumberFormats => GetNumberFormatsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.SetNumberFormat => SetNumberFormatAsync(rangeCommands, sessionId, sheetName, rangeAddress, formatCode),
                    RangeAction.SetNumberFormats => SetNumberFormatsAsync(rangeCommands, sessionId, sheetName, rangeAddress, formats),
                    RangeAction.ClearAll => ClearAllAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.ClearContents => ClearContentsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.ClearFormats => ClearFormatsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.Copy => CopyAsync(rangeCommands, sessionId, sourceSheet, sourceRange, targetSheet, targetRange),
                    RangeAction.CopyValues => CopyValuesAsync(rangeCommands, sessionId, sourceSheet, sourceRange, targetSheet, targetRange),
                    RangeAction.CopyFormulas => CopyFormulasAsync(rangeCommands, sessionId, sourceSheet, sourceRange, targetSheet, targetRange),
                    RangeAction.InsertCells => InsertCellsAsync(rangeCommands, sessionId, sheetName, rangeAddress, shift),
                    RangeAction.DeleteCells => DeleteCellsAsync(rangeCommands, sessionId, sheetName, rangeAddress, shift),
                    RangeAction.InsertRows => InsertRowsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.DeleteRows => DeleteRowsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.InsertColumns => InsertColumnsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.DeleteColumns => DeleteColumnsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.Find => FindAsync(rangeCommands, sessionId, sheetName, rangeAddress, searchValue, matchCase, matchEntireCell, searchFormulas, searchValues),
                    RangeAction.Replace => ReplaceAsync(rangeCommands, sessionId, sheetName, rangeAddress, searchValue, replaceValue, matchCase, matchEntireCell, searchFormulas, searchValues, replaceAll),
                    RangeAction.Sort => SortAsync(rangeCommands, sessionId, sheetName, rangeAddress, sortColumns, hasHeaders),
                    RangeAction.GetUsedRange => GetUsedRangeAsync(rangeCommands, sessionId, sheetName),
                    RangeAction.GetCurrentRegion => GetCurrentRegionAsync(rangeCommands, sessionId, sheetName, cellAddress),
                    RangeAction.GetInfo => GetRangeInfoAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.AddHyperlink => AddHyperlinkAsync(rangeCommands, sessionId, sheetName, cellAddress, url, displayText, tooltip),
                    RangeAction.RemoveHyperlink => RemoveHyperlinkAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.ListHyperlinks => ListHyperlinksAsync(rangeCommands, sessionId, sheetName),
                    RangeAction.GetHyperlink => GetHyperlinkAsync(rangeCommands, sessionId, sheetName, cellAddress),
                    RangeAction.GetStyle => GetStyleAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.SetStyle => SetStyleAsync(rangeCommands, sessionId, sheetName, rangeAddress, styleName),
                    RangeAction.FormatRange => FormatRangeAsync(rangeCommands, sessionId, sheetName, rangeAddress, fontName, fontSize, bold, italic, underline, fontColor, fillColor, borderStyle, borderColor, borderWeight, horizontalAlignment, verticalAlignment, wrapText, orientation),
                    RangeAction.ValidateRange => ValidateRangeAsync(rangeCommands, sessionId, sheetName, rangeAddress, validationType, validationOperator, validationFormula1, validationFormula2, showInputMessage, inputTitle, inputMessage, showErrorAlert, errorStyle, errorTitle, errorMessage, ignoreBlank, showDropdown),
                    RangeAction.GetValidation => GetValidationAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.RemoveValidation => RemoveValidationAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.AutoFitColumns => AutoFitColumnsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.AutoFitRows => AutoFitRowsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.MergeCells => MergeCellsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.UnmergeCells => UnmergeCellsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.GetMergeInfo => GetMergeInfoAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.SetCellLock => SetCellLockAsync(rangeCommands, sessionId, sheetName, rangeAddress, locked),
                    RangeAction.GetCellLock => GetCellLockAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    // === VALUE OPERATIONS ===

    private static string GetValuesAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-values");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetValues(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.RangeAddress,
            result.Values,
            result.RowCount,
            result.ColumnCount,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetValuesAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, List<List<object?>>? values)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-values");
        if (values == null || values.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("values", "set-values");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.SetValues(batch, sheetName ?? "", rangeAddress!, values!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === FORMULA OPERATIONS ===

    private static string GetFormulasAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-formulas");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetFormulas(batch, sheetName ?? "", rangeAddress!));

        var formulaCount = result.Formulas.SelectMany(row => row).Count(f => !string.IsNullOrEmpty(f));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.RangeAddress,
            result.Formulas,
            result.Values,
            result.RowCount,
            result.ColumnCount,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetFormulasAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, List<List<string>>? formulas)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-formulas");
        if (formulas == null || formulas.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("formulas", "set-formulas");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.SetFormulas(batch, sheetName ?? "", rangeAddress!, formulas!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === NUMBER FORMAT OPERATIONS ===

    private static string GetNumberFormatsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-number-formats");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetNumberFormats(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.RangeAddress,
            result.Formats,
            result.RowCount,
            result.ColumnCount,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetNumberFormatAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress,
        string? formatCode)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-number-format");

        if (string.IsNullOrEmpty(formatCode))
            ExcelToolsBase.ThrowMissingParameter("formatCode", "set-number-format");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.SetNumberFormat(batch, sheetName ?? "", rangeAddress!, formatCode!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetNumberFormatsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, List<List<string>>? formats)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-number-formats");
        if (formats == null || formats.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("formats", "set-number-formats");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.SetNumberFormats(batch, sheetName ?? "", rangeAddress!, formats!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === CLEAR OPERATIONS ===

    private static string ClearAllAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-all");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.ClearAll(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ClearContentsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-contents");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.ClearContents(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ClearFormatsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-formats");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.ClearFormats(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === COPY OPERATIONS ===

    private static string CopyAsync(RangeCommands commands, string sessionId, string? sourceSheet, string? sourceRange, string? targetSheet, string? targetRange)
    {
        if (string.IsNullOrEmpty(sourceRange))
            ExcelToolsBase.ThrowMissingParameter("sourceRange", "copy");
        if (string.IsNullOrEmpty(targetRange))
            ExcelToolsBase.ThrowMissingParameter("targetRange", "copy");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Copy(batch, sourceSheet ?? "", sourceRange!, targetSheet ?? "", targetRange!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CopyValuesAsync(RangeCommands commands, string sessionId, string? sourceSheet, string? sourceRange, string? targetSheet, string? targetRange)
    {
        if (string.IsNullOrEmpty(sourceRange))
            ExcelToolsBase.ThrowMissingParameter("sourceRange", "copy-values");
        if (string.IsNullOrEmpty(targetRange))
            ExcelToolsBase.ThrowMissingParameter("targetRange", "copy-values");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.CopyValues(batch, sourceSheet ?? "", sourceRange!, targetSheet ?? "", targetRange!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CopyFormulasAsync(RangeCommands commands, string sessionId, string? sourceSheet, string? sourceRange, string? targetSheet, string? targetRange)
    {
        if (string.IsNullOrEmpty(sourceRange))
            ExcelToolsBase.ThrowMissingParameter("sourceRange", "copy-formulas");
        if (string.IsNullOrEmpty(targetRange))
            ExcelToolsBase.ThrowMissingParameter("targetRange", "copy-formulas");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.CopyFormulas(batch, sourceSheet ?? "", sourceRange!, targetSheet ?? "", targetRange!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === INSERT/DELETE OPERATIONS ===

    private static string InsertCellsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? shift)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "insert-cells");
        if (string.IsNullOrEmpty(shift))
            ExcelToolsBase.ThrowMissingParameter("shift", "insert-cells");

        if (!Enum.TryParse<InsertShiftDirection>(shift, true, out var shiftDirection))
        {
            throw new ArgumentException($"Invalid shift direction '{shift}'. Must be 'Down' or 'Right'.", nameof(shift));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.InsertCells(batch, sheetName ?? "", rangeAddress!, shiftDirection));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteCellsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? shift)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "delete-cells");
        if (string.IsNullOrEmpty(shift))
            ExcelToolsBase.ThrowMissingParameter("shift", "delete-cells");

        if (!Enum.TryParse<DeleteShiftDirection>(shift, true, out var shiftDirection))
        {
            throw new ArgumentException($"Invalid shift direction '{shift}'. Must be 'Up' or 'Left'.", nameof(shift));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.DeleteCells(batch, sheetName ?? "", rangeAddress!, shiftDirection));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string InsertRowsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "insert-rows");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.InsertRows(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteRowsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "delete-rows");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.DeleteRows(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string InsertColumnsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "insert-columns");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.InsertColumns(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteColumnsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "delete-columns");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.DeleteColumns(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === FIND/REPLACE OPERATIONS ===

    private static string FindAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? searchValue, bool? matchCase, bool? matchEntireCell, bool? searchFormulas, bool? searchValues)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "find");
        if (string.IsNullOrEmpty(searchValue))
            ExcelToolsBase.ThrowMissingParameter("searchValue", "find");

        var options = new FindOptions
        {
            MatchCase = matchCase ?? false,
            MatchEntireCell = matchEntireCell ?? false,
            SearchFormulas = searchFormulas ?? true,
            SearchValues = searchValues ?? true
        };

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Find(batch, sheetName ?? "", rangeAddress!, searchValue!, options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.RangeAddress,
            result.SearchValue,
            MatchingCells = result.MatchingCells.Take(10).ToList(),
            TotalMatches = result.MatchingCells.Count,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ReplaceAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? searchValue, string? replaceValue, bool? matchCase, bool? matchEntireCell, bool? searchFormulas, bool? searchValues, bool? replaceAll)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "replace");
        if (string.IsNullOrEmpty(searchValue))
            ExcelToolsBase.ThrowMissingParameter("searchValue", "replace");
        if (replaceValue == null)
            ExcelToolsBase.ThrowMissingParameter("replaceValue", "replace");

        var options = new ReplaceOptions
        {
            MatchCase = matchCase ?? false,
            MatchEntireCell = matchEntireCell ?? false,
            SearchFormulas = searchFormulas ?? true,
            SearchValues = searchValues ?? true,
            ReplaceAll = replaceAll ?? true
        };

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.Replace(batch, sheetName ?? "", rangeAddress!, searchValue!, replaceValue!, options);
                return null;
            });

        return JsonSerializer.Serialize(new
        {
            Success = true
        }, ExcelToolsBase.JsonOptions);
    }

    // === SORT OPERATIONS ===

    private static string SortAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, List<SortColumn>? sortColumns, bool? hasHeaders)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "sort");
        if (sortColumns == null || sortColumns.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("sortColumns", "sort");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.Sort(batch, sheetName ?? "", rangeAddress!, sortColumns!, hasHeaders ?? true);
                return null;
            });

        return JsonSerializer.Serialize(new
        {
            Success = true
        }, ExcelToolsBase.JsonOptions);
    }

    // === DISCOVERY OPERATIONS ===

    private static string GetUsedRangeAsync(RangeCommands commands, string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "get-used-range");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetUsedRange(batch, sheetName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.RangeAddress,
            result.Values,
            result.RowCount,
            result.ColumnCount,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetCurrentRegionAsync(RangeCommands commands, string sessionId, string? sheetName, string? cellAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "get-current-region");
        if (string.IsNullOrEmpty(cellAddress))
            ExcelToolsBase.ThrowMissingParameter("cellAddress", "get-current-region");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetCurrentRegion(batch, sheetName!, cellAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.RangeAddress,
            result.Values,
            result.RowCount,
            result.ColumnCount,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetRangeInfoAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-range-info");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetInfo(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.Address,
            result.RowCount,
            result.ColumnCount,
            result.NumberFormat,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === HYPERLINK OPERATIONS ===

    private static string AddHyperlinkAsync(RangeCommands commands, string sessionId, string? sheetName, string? cellAddress, string? url, string? displayText, string? tooltip)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "add-hyperlink");
        if (string.IsNullOrEmpty(cellAddress))
            ExcelToolsBase.ThrowMissingParameter("cellAddress", "add-hyperlink");
        if (string.IsNullOrEmpty(url))
            ExcelToolsBase.ThrowMissingParameter("url", "add-hyperlink");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.AddHyperlink(batch, sheetName!, cellAddress!, url!, displayText, tooltip));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RemoveHyperlinkAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "remove-hyperlink");
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "remove-hyperlink");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.RemoveHyperlink(batch, sheetName!, rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ListHyperlinksAsync(RangeCommands commands, string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "list-hyperlinks");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.ListHyperlinks(batch, sheetName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            ((dynamic)result).SheetName,
            ((dynamic)result).Hyperlinks,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetHyperlinkAsync(RangeCommands commands, string sessionId, string? sheetName, string? cellAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "get-hyperlink");
        if (string.IsNullOrEmpty(cellAddress))
            ExcelToolsBase.ThrowMissingParameter("cellAddress", "get-hyperlink");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetHyperlink(batch, sheetName!, cellAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            ((dynamic)result).CellAddress,
            ((dynamic)result).Url,
            ((dynamic)result).DisplayText,
            ((dynamic)result).Tooltip,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === FORMATTING OPERATIONS ===

    private static string SetStyleAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress,
        string? styleName)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-style");
        if (string.IsNullOrEmpty(styleName))
            ExcelToolsBase.ThrowMissingParameter("styleName", "set-style");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.SetStyle(batch, sheetName ?? "", rangeAddress!, styleName!);
                return null;
            });

        return JsonSerializer.Serialize(new
        {
            Success = true
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetStyleAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-style");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetStyle(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            success = true,
            sheetName,
            rangeAddress,
            styleName = result.StyleName,
            isBuiltInStyle = result.IsBuiltInStyle,
            styleDescription = result.StyleDescription
        }, ExcelToolsBase.JsonOptions);
    }

    private static string FormatRangeAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress,
        string? fontName,
        double? fontSize,
        bool? bold,
        bool? italic,
        bool? underline,
        string? fontColor,
        string? fillColor,
        string? borderStyle,
        string? borderColor,
        string? borderWeight,
        string? horizontalAlignment,
        string? verticalAlignment,
        bool? wrapText,
        int? orientation)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "format-range");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.FormatRange(
                    batch,
                    sheetName ?? "",
                    rangeAddress!,
                    fontName,
                    fontSize,
                    bold,
                    italic,
                    underline,
                    fontColor,
                    fillColor,
                    borderStyle,
                    borderColor,
                    borderWeight,
                    horizontalAlignment,
                    verticalAlignment,
                    wrapText,
                    orientation);
                return null;
            });

        return JsonSerializer.Serialize(new
        {
            Success = true
        }, ExcelToolsBase.JsonOptions);
    }

    // === VALIDATION OPERATIONS ===

    private static string ValidateRangeAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress,
        string? validationType,
        string? validationOperator,
        string? validationFormula1,
        string? validationFormula2,
        bool? showInputMessage,
        string? inputTitle,
        string? inputMessage,
        bool? showErrorAlert,
        string? errorStyle,
        string? errorTitle,
        string? errorMessage,
        bool? ignoreBlank,
        bool? showDropdown)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "validate-range");
        if (string.IsNullOrEmpty(validationType))
            ExcelToolsBase.ThrowMissingParameter("validationType", "validate-range");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.ValidateRange(
                    batch,
                    sheetName ?? "",
                    rangeAddress!,
                    validationType!,
                    validationOperator,
                    validationFormula1,
                    validationFormula2,
                    showInputMessage,
                    inputTitle,
                    inputMessage,
                    showErrorAlert,
                    errorStyle,
                    errorTitle,
                    errorMessage,
                    ignoreBlank,
                    showDropdown);
                return null;
            });

        return JsonSerializer.Serialize(new
        {
            Success = true
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetValidationAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-validation");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetValidation(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            ((dynamic)result).ValidationType,
            ((dynamic)result).ValidationOperator,
            ((dynamic)result).Formula1,
            ((dynamic)result).Formula2,
            ((dynamic)result).ShowInputMessage,
            ((dynamic)result).InputTitle,
            ((dynamic)result).InputMessage,
            ((dynamic)result).ShowErrorAlert,
            ((dynamic)result).ErrorStyle,
            ((dynamic)result).ErrorTitle,
            ValidationErrorMessage = ((dynamic)result).ErrorMessage,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RemoveValidationAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "remove-validation");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.RemoveValidation(batch, sheetName ?? "", rangeAddress!);
                return null;
            });

        return JsonSerializer.Serialize(new
        {
            Success = true
        }, ExcelToolsBase.JsonOptions);
    }

    private static string AutoFitColumnsAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "auto-fit-columns");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.AutoFitColumns(batch, sheetName ?? "", rangeAddress!);
                return null;
            });

        return JsonSerializer.Serialize(new
        {
            Success = true
        }, ExcelToolsBase.JsonOptions);
    }

    private static string AutoFitRowsAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "auto-fit-rows");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.AutoFitRows(batch, sheetName ?? "", rangeAddress!);
                return null;
            });

        return JsonSerializer.Serialize(new
        {
            Success = true
        }, ExcelToolsBase.JsonOptions);
    }

    private static string MergeCellsAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "merge-cells");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.MergeCells(batch, sheetName ?? "", rangeAddress!);
                return null;
            });

        return JsonSerializer.Serialize(new
        {
            Success = true
        }, ExcelToolsBase.JsonOptions);
    }

    private static string UnmergeCellsAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "unmerge-cells");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.UnmergeCells(batch, sheetName ?? "", rangeAddress!);
                return null;
            });

        return JsonSerializer.Serialize(new
        {
            Success = true
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetMergeInfoAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-merge-info");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetMergeInfo(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            ((dynamic)result).IsMerged,
            ((dynamic)result).MergeAddress,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetCellLockAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress,
        bool? locked)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-cell-lock");

        if (locked == null)
            ExcelToolsBase.ThrowMissingParameter("locked", "set-cell-lock");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.SetCellLock(batch, sheetName ?? "", rangeAddress!, locked!.Value);
                return null;
            });

        return JsonSerializer.Serialize(new
        {
            Success = true
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetCellLockAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-cell-lock");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetCellLock(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            ((dynamic)result).Locked,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }
}

