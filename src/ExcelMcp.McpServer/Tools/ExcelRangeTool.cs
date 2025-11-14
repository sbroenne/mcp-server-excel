using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.McpServer.Models;

#pragma warning disable CA1861 // Avoid constant arrays as arguments - workflow hints are contextual per-call

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel range operations - values, formulas, clearing, copying, inserting, deleting, finding, sorting, and hyperlinks.
/// </summary>
[McpServerToolType]
public static class ExcelRangeTool
{
    /// <summary>
    /// Unified Excel range operations - comprehensive data manipulation API.
    /// Supports: values, formulas, number formats, clear, copy, insert/delete, find/replace, sort, discovery, hyperlinks.
    /// </summary>
    [McpServerTool(Name = "excel_range")]
    [Description(@"Unified Excel range operations - ALL data manipulation.

DATA FORMAT:
- Values/formulas: JSON 2D arrays [[row1col1, row1col2], [row2col1, row2col2]]
- Example single cell: [[100]] or [['=SUM(A:A)']]
- Example range: [[1,2,3], [4,5,6], [7,8,9]]
")]
    public static async Task<string> ExcelRange(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        RangeAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [Required]
        [Description("Session ID from excel_file 'open' action (required for all range operations)")]
        string sessionId,

        [Description("Worksheet name (empty for named ranges, required for most operations)")]
        string? sheetName = null,

        [Description("Range address (e.g., 'A1:D10') or named range (e.g., 'SalesData'). For named ranges, leave sheetName empty.")]
        string? rangeAddress = null,

        [Description("2D array of values for set-values (JSON array of arrays, e.g., [[1,2],[3,4]])")]
        List<List<object?>>? values = null,

        [Description("2D array of formulas for set-formulas (JSON array of arrays, e.g., [[\"=A1+B1\",\"=SUM(A:A)\"]])")]
        List<List<string>>? formulas = null,

        [Description("Source sheet name (for copy operations)")]
        string? sourceSheet = null,

        [Description("Source range address (for copy operations)")]
        string? sourceRange = null,

        [Description("Target sheet name (for copy operations)")]
        string? targetSheet = null,

        [Description("Target range address (for copy operations)")]
        string? targetRange = null,

        [Description("Shift direction for insert-cells/delete-cells: Down, Right, Up, Left")]
        string? shift = null,

        [Description("Search value (for find/replace operations)")]
        string? searchValue = null,

        [Description("Replace value (for replace operation)")]
        string? replaceValue = null,

        [Description("Match case (for find/replace, default: false)")]
        bool? matchCase = null,

        [Description("Match entire cell (for find/replace, default: false)")]
        bool? matchEntireCell = null,

        [Description("Search formulas (for find/replace, default: true)")]
        bool? searchFormulas = null,

        [Description("Search values (for find/replace, default: true)")]
        bool? searchValues = null,

        [Description("Replace all occurrences (for replace, default: true)")]
        bool? replaceAll = null,

        [Description("Sort columns (JSON array, e.g., [{\"columnIndex\":1,\"ascending\":true}])")]
        List<SortColumn>? sortColumns = null,

        [Description("Has header row (for sort, default: true)")]
        bool? hasHeaders = null,

        [Description("Cell address for single-cell operations (hyperlinks, current-region)")]
        string? cellAddress = null,

        [Description("Hyperlink URL (for add-hyperlink)")]
        string? url = null,

        [Description("Hyperlink display text (for add-hyperlink, optional)")]
        string? displayText = null,

        [Description("Hyperlink tooltip (for add-hyperlink, optional)")]
        string? tooltip = null,

        [Description("Excel format code for set-number-format (e.g., '$#,##0.00', '0.00%', 'm/d/yyyy')")]
        string? formatCode = null,

        [Description("2D array of format codes for set-number-formats (JSON array of arrays, e.g., [['$#,##0','0.00%'],['m/d/yyyy','General']])")]
        List<List<string>>? formats = null,

        // === FORMATTING PARAMETERS ===

        [Description("Built-in Excel style name (for set-style: 'Heading 1', 'Accent1', 'Good', 'Total', 'Currency', 'Percent', 'Normal', etc. - recommended for consistent formatting)")]
        string? styleName = null,

        [Description("Font name (for format-range, e.g., 'Arial', 'Calibri')")]
        string? fontName = null,

        [Description("Font size (for format-range, e.g., 11, 12, 14)")]
        double? fontSize = null,

        [Description("Bold font (for format-range)")]
        bool? bold = null,

        [Description("Italic font (for format-range)")]
        bool? italic = null,

        [Description("Underline font (for format-range)")]
        bool? underline = null,

        [Description("Font color (for format-range, #RRGGBB or color index)")]
        string? fontColor = null,

        [Description("Fill color (for format-range, #RRGGBB or color index)")]
        string? fillColor = null,

        [Description("Border style (for format-range: none, continuous, dash, dot, double, etc.)")]
        string? borderStyle = null,

        [Description("Border color (for format-range, #RRGGBB or color index)")]
        string? borderColor = null,

        [Description("Border weight (for format-range: hairline, thin, medium, thick)")]
        string? borderWeight = null,

        [Description("Horizontal alignment (for format-range: left, center, right, justify, distributed)")]
        string? horizontalAlignment = null,

        [Description("Vertical alignment (for format-range: top, center, bottom, justify, distributed)")]
        string? verticalAlignment = null,

        [Description("Wrap text in cells (for format-range)")]
        bool? wrapText = null,

        [Description("Text orientation in degrees (for format-range, 0-90 or -90)")]
        int? orientation = null,

        // === VALIDATION PARAMETERS ===

        [Description("Data validation type (for validate-range: list, whole, decimal, date, time, textLength, custom)")]
        string? validationType = null,

        [Description("Data validation operator (for validate-range: between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual)")]
        string? validationOperator = null,

        [Description("Validation formula1 (for validate-range). For 'list' type: MUST be worksheet range reference like '=$A$1:$A$10' to create dropdown. For other types: value/formula for comparison.")]
        string? validationFormula1 = null,

        [Description("Validation formula2 (for validate-range, second value/formula for between/notBetween)")]
        string? validationFormula2 = null,

        [Description("Show input message (for validate-range)")]
        bool? showInputMessage = null,

        [Description("Input message title (for validate-range)")]
        string? inputTitle = null,

        [Description("Input message text (for validate-range)")]
        string? inputMessage = null,

        [Description("Show error alert (for validate-range)")]
        bool? showErrorAlert = null,

        [Description("Error alert style (for validate-range: stop, warning, information)")]
        string? errorStyle = null,

        [Description("Error alert title (for validate-range)")]
        string? errorTitle = null,

        [Description("Error alert message (for validate-range)")]
        string? errorMessage = null,

        [Description("Ignore blank cells in validation (for validate-range)")]
        bool? ignoreBlank = null,

        [Description("Show dropdown for list validation (for validate-range)")]
        bool? showDropdown = null,

        [Description("Lock status for cells (for set-cell-lock: true = locked, false = unlocked)")]
        bool? locked = null,

        // === CONDITIONAL FORMATTING PARAMETERS ===

        [Description("Conditional formatting rule type (for add-conditional-formatting: cellValue, expression, colorScale, dataBar, iconSet, top10, uniqueValues, duplicateValues, blanks, noBlanks, errors, noErrors)")]
        string? ruleType = null,

        [Description("First formula for conditional formatting rule (for add-conditional-formatting, required for most rule types)")]
        string? formula1 = null,

        [Description("Second formula for conditional formatting rule (for add-conditional-formatting, optional, used for 'between' rules)")]
        string? formula2 = null,

        [Description("Format style for conditional formatting (for add-conditional-formatting, optional, e.g., 'highlight', 'databar', 'colorscale')")]
        string? formatStyle = null)
    {
        try
        {
            var rangeCommands = new RangeCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                RangeAction.GetValues => await GetValuesAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.SetValues => await SetValuesAsync(rangeCommands, sessionId, sheetName, rangeAddress, values),
                RangeAction.GetFormulas => await GetFormulasAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.SetFormulas => await SetFormulasAsync(rangeCommands, sessionId, sheetName, rangeAddress, formulas),
                RangeAction.GetNumberFormats => await GetNumberFormatsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.SetNumberFormat => await SetNumberFormatAsync(rangeCommands, sessionId, sheetName, rangeAddress, formatCode),
                RangeAction.SetNumberFormats => await SetNumberFormatsAsync(rangeCommands, sessionId, sheetName, rangeAddress, formats),
                RangeAction.ClearAll => await ClearAllAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.ClearContents => await ClearContentsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.ClearFormats => await ClearFormatsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.Copy => await CopyAsync(rangeCommands, sessionId, sourceSheet, sourceRange, targetSheet, targetRange),
                RangeAction.CopyValues => await CopyValuesAsync(rangeCommands, sessionId, sourceSheet, sourceRange, targetSheet, targetRange),
                RangeAction.CopyFormulas => await CopyFormulasAsync(rangeCommands, sessionId, sourceSheet, sourceRange, targetSheet, targetRange),
                RangeAction.InsertCells => await InsertCellsAsync(rangeCommands, sessionId, sheetName, rangeAddress, shift),
                RangeAction.DeleteCells => await DeleteCellsAsync(rangeCommands, sessionId, sheetName, rangeAddress, shift),
                RangeAction.InsertRows => await InsertRowsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.DeleteRows => await DeleteRowsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.InsertColumns => await InsertColumnsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.DeleteColumns => await DeleteColumnsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.Find => await FindAsync(rangeCommands, sessionId, sheetName, rangeAddress, searchValue, matchCase, matchEntireCell, searchFormulas, searchValues),
                RangeAction.Replace => await ReplaceAsync(rangeCommands, sessionId, sheetName, rangeAddress, searchValue, replaceValue, matchCase, matchEntireCell, searchFormulas, searchValues, replaceAll),
                RangeAction.Sort => await SortAsync(rangeCommands, sessionId, sheetName, rangeAddress, sortColumns, hasHeaders),
                RangeAction.GetUsedRange => await GetUsedRangeAsync(rangeCommands, sessionId, sheetName),
                RangeAction.GetCurrentRegion => await GetCurrentRegionAsync(rangeCommands, sessionId, sheetName, cellAddress),
                RangeAction.GetInfo => await GetRangeInfoAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.AddHyperlink => await AddHyperlinkAsync(rangeCommands, sessionId, sheetName, cellAddress, url, displayText, tooltip),
                RangeAction.RemoveHyperlink => await RemoveHyperlinkAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.ListHyperlinks => await ListHyperlinksAsync(rangeCommands, sessionId, sheetName),
                RangeAction.GetHyperlink => await GetHyperlinkAsync(rangeCommands, sessionId, sheetName, cellAddress),
                RangeAction.GetStyle => await GetStyleAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.SetStyle => await SetStyleAsync(rangeCommands, sessionId, sheetName, rangeAddress, styleName),
                RangeAction.FormatRange => await FormatRangeAsync(rangeCommands, sessionId, sheetName, rangeAddress, fontName, fontSize, bold, italic, underline, fontColor, fillColor, borderStyle, borderColor, borderWeight, horizontalAlignment, verticalAlignment, wrapText, orientation),
                RangeAction.ValidateRange => await ValidateRangeAsync(rangeCommands, sessionId, sheetName, rangeAddress, validationType, validationOperator, validationFormula1, validationFormula2, showInputMessage, inputTitle, inputMessage, showErrorAlert, errorStyle, errorTitle, errorMessage, ignoreBlank, showDropdown),
                RangeAction.GetValidation => await GetValidationAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.RemoveValidation => await RemoveValidationAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.AutoFitColumns => await AutoFitColumnsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.AutoFitRows => await AutoFitRowsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.MergeCells => await MergeCellsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.UnmergeCells => await UnmergeCellsAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.GetMergeInfo => await GetMergeInfoAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.AddConditionalFormatting => await AddConditionalFormattingAsync(rangeCommands, sessionId, sheetName, rangeAddress, ruleType, formula1, formula2, formatStyle),
                RangeAction.ClearConditionalFormatting => await ClearConditionalFormattingAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                RangeAction.SetCellLock => await SetCellLockAsync(rangeCommands, sessionId, sheetName, rangeAddress, locked),
                RangeAction.GetCellLock => await GetCellLockAsync(rangeCommands, sessionId, sheetName, rangeAddress),
                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action: {action} ({action.ToActionString()})")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);
            throw; // Unreachable but satisfies compiler
        }
    }

    // === VALUE OPERATIONS ===

    private static async Task<string> GetValuesAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-values");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetValuesAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.RangeAddress,
            result.Values,
            result.RowCount,
            result.ColumnCount,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Read {result.RowCount}x{result.ColumnCount} range from '{result.SheetName}' at {result.RangeAddress}. Data retrieved as 2D array."
                : "Failed to read range values. Verify sheet and range address are valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'set-values' to update cell data", "Use 'get-formulas' to see formulas instead of values", "Use excel_table 'create' to convert range to structured table" }
                : ["Verify sheet name with excel_worksheet 'list'", "Check range address format (e.g., 'A1:D10')", "Use 'get-used-range' to discover data boundaries"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetValuesAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, List<List<object?>>? values)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-values");
        if (values == null || values.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("values", "set-values");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SetValuesAsync(batch, sheetName ?? "", rangeAddress!, values!));

        var rowCount = values!.Count;
        var colCount = values.Count > 0 ? values[0].Count : 0;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Wrote {rowCount}x{colCount} values to {rangeAddress} on '{sheetName}'. Data saved to workbook."
                : "Failed to write values to range. Verify range size matches data dimensions.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-values' to verify written data", "Use 'set-number-format' to apply formatting", "Use excel_worksheet 'save' if using batch mode" }
                : ["Check values array dimensions match range", "Verify sheet exists with excel_worksheet 'list'", "Use 'get-range-info' to check range dimensions"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === FORMULA OPERATIONS ===

    private static async Task<string> GetFormulasAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-formulas");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetFormulasAsync(batch, sheetName ?? "", rangeAddress!));

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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Read {result.RowCount}x{result.ColumnCount} range with {formulaCount} formulas. Empty strings indicate values only."
                : "Failed to read formulas from range. Verify sheet and range address are valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'set-formulas' to update calculations", "Use 'copy-formulas' to duplicate logic to another range", "Inspect formulas for dependencies and references" }
                : ["Verify sheet name with excel_worksheet 'list'", "Check range address format", "Use 'get-values' to verify range exists"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetFormulasAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, List<List<string>>? formulas)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-formulas");
        if (formulas == null || formulas.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("formulas", "set-formulas");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SetFormulasAsync(batch, sheetName ?? "", rangeAddress!, formulas!));

        var rowCount = formulas!.Count;
        var colCount = formulas.Count > 0 ? formulas[0].Count : 0;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Wrote {rowCount}x{colCount} formulas to {rangeAddress}. Formulas will calculate automatically."
                : "Failed to write formulas. Verify formula syntax and range dimensions.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-formulas' to verify written formulas", "Use 'get-values' to see calculated results", "Check for #REF! or #VALUE! errors in results" }
                : ["Check formula syntax (must start with '=')", "Verify range size matches formula array dimensions", "Test formula in Excel UI first"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === NUMBER FORMAT OPERATIONS ===

    private static async Task<string> GetNumberFormatsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-number-formats");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetNumberFormatsAsync(batch, sheetName ?? "", rangeAddress!));

        var uniqueFormats = result.Formats.SelectMany(row => row).Distinct().Take(5).ToList();

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.RangeAddress,
            result.Formats,
            result.RowCount,
            result.ColumnCount,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Retrieved number formats for {result.RowCount}x{result.ColumnCount} range. Found {uniqueFormats.Count} unique formats."
                : "Failed to read number formats. Verify sheet and range are valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'set-number-format' to apply uniform formatting", "Analyze format codes to understand data types", "Use 'set-number-formats' for cell-by-cell formatting" }
                : ["Verify sheet name with excel_worksheet 'list'", "Check range address format", "Use 'get-range-info' to verify range exists"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetNumberFormatAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? formatCode)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-number-format");
        if (string.IsNullOrEmpty(formatCode))
            ExcelToolsBase.ThrowMissingParameter("formatCode", "set-number-format");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SetNumberFormatAsync(batch, sheetName ?? "", rangeAddress!, formatCode!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Applied format '{formatCode}' to range {rangeAddress}. All cells now share this format."
                : "Failed to apply number format. Verify format code syntax.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-values' to see formatted display", "Use 'get-number-formats' to verify format applied", "Common formats: '$#,##0.00' (currency), '0.00%' (percent), 'm/d/yyyy' (date)" }
                : ["Test format code in Excel UI first", "Check format string syntax (e.g., '#,##0.00')", "Use built-in Excel format codes"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetNumberFormatsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, List<List<string>>? formats)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-number-formats");
        if (formats == null || formats.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("formats", "set-number-formats");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SetNumberFormatsAsync(batch, sheetName ?? "", rangeAddress!, formats!));

        var rowCount = formats!.Count;
        var colCount = formats.Count > 0 ? formats[0].Count : 0;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Applied {rowCount}x{colCount} cell-specific formats to range. Each cell has independent formatting."
                : "Failed to apply formats. Verify format array dimensions match range.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-number-formats' to verify applied formats", "Use 'set-number-format' for uniform formatting instead", "Mix currency, percentage, and date formats as needed" }
                : ["Check formats array dimensions match range", "Verify all format codes are valid", "Use 'set-number-format' for simpler uniform formatting"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === CLEAR OPERATIONS ===

    private static async Task<string> ClearAllAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-all");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ClearAllAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Cleared all content and formatting from {rangeAddress}. Range is now completely empty."
                : "Failed to clear range. Verify sheet and range are valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'set-values' to repopulate with new data", "Use 'clear-contents' to preserve formatting next time", "Verify cleared with 'get-values' (should return empty)" }
                : ["Verify sheet exists with excel_worksheet 'list'", "Check range address format", "Ensure sheet is not protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ClearContentsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-contents");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ClearContentsAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Cleared values and formulas from {rangeAddress}. Formatting preserved."
                : "Failed to clear contents. Verify sheet and range are valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'set-values' to write new data with existing formatting", "Use 'get-number-formats' to see preserved formats", "Use 'clear-all' to remove formatting too" }
                : ["Verify sheet name with excel_worksheet 'list'", "Check range address format", "Ensure cells are not locked"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ClearFormatsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-formats");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ClearFormatsAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Cleared formatting from {rangeAddress}. Values and formulas preserved."
                : "Failed to clear formats. Verify sheet and range are valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'set-number-format' to apply new formatting", "Use 'get-values' to verify data still intact", "Use 'clear-contents' to remove data instead" }
                : ["Verify sheet name exists", "Check range address format", "Use 'clear-all' to remove everything"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === COPY OPERATIONS ===

    private static async Task<string> CopyAsync(RangeCommands commands, string sessionId, string? sourceSheet, string? sourceRange, string? targetSheet, string? targetRange)
    {
        if (string.IsNullOrEmpty(sourceRange))
            ExcelToolsBase.ThrowMissingParameter("sourceRange", "copy");
        if (string.IsNullOrEmpty(targetRange))
            ExcelToolsBase.ThrowMissingParameter("targetRange", "copy");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.CopyAsync(batch, sourceSheet ?? "", sourceRange!, targetSheet ?? "", targetRange!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Copied {sourceRange} to {targetRange} (values, formulas, and formatting). Complete duplication."
                : "Failed to copy range. Verify source and target ranges are valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-values' on target to verify copy", "Use 'copy-values' or 'copy-formulas' for selective copying", "Modify target range independently now" }
                : ["Check source range exists with 'get-values'", "Verify target range address format", "Ensure sheets exist with excel_worksheet 'list'"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CopyValuesAsync(RangeCommands commands, string sessionId, string? sourceSheet, string? sourceRange, string? targetSheet, string? targetRange)
    {
        if (string.IsNullOrEmpty(sourceRange))
            ExcelToolsBase.ThrowMissingParameter("sourceRange", "copy-values");
        if (string.IsNullOrEmpty(targetRange))
            ExcelToolsBase.ThrowMissingParameter("targetRange", "copy-values");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.CopyValuesAsync(batch, sourceSheet ?? "", sourceRange!, targetSheet ?? "", targetRange!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Copied values from {sourceRange} to {targetRange}. Formulas and formatting not copied."
                : "Failed to copy values. Verify source and target ranges.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-values' on target to verify", "Original formulas become static values in target", "Use 'set-number-format' to apply formatting to target" }
                : ["Check source range has data with 'get-values'", "Verify target range address", "Use 'copy' for complete duplication"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CopyFormulasAsync(RangeCommands commands, string sessionId, string? sourceSheet, string? sourceRange, string? targetSheet, string? targetRange)
    {
        if (string.IsNullOrEmpty(sourceRange))
            ExcelToolsBase.ThrowMissingParameter("sourceRange", "copy-formulas");
        if (string.IsNullOrEmpty(targetRange))
            ExcelToolsBase.ThrowMissingParameter("targetRange", "copy-formulas");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.CopyFormulasAsync(batch, sourceSheet ?? "", sourceRange!, targetSheet ?? "", targetRange!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Copied formulas from {sourceRange} to {targetRange}. Cell references adjusted automatically."
                : "Failed to copy formulas. Verify source and target ranges.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-formulas' on target to verify adjusted references", "Relative references updated, absolute ($) preserved", "Check for #REF! errors if references broken" }
                : ["Verify source has formulas with 'get-formulas'", "Check target range address", "Use 'copy' to include values and formatting"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === INSERT/DELETE OPERATIONS ===

    private static async Task<string> InsertCellsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? shift)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "insert-cells");
        if (string.IsNullOrEmpty(shift))
            ExcelToolsBase.ThrowMissingParameter("shift", "insert-cells");

        if (!Enum.TryParse<InsertShiftDirection>(shift, true, out var shiftDirection))
        {
            throw new ModelContextProtocol.McpException($"Invalid shift direction '{shift}'. Must be 'Down' or 'Right'.");
        }

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.InsertCellsAsync(batch, sheetName ?? "", rangeAddress!, shiftDirection));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Inserted blank cells at {rangeAddress}, shifted existing cells {shift.ToLowerInvariant()}."
                : $"Failed to insert cells with {shift} shift. Verify range and shift direction.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'set-values' to populate new cells", "Use 'delete-cells' to reverse operation", "Check formulas for updated references" }
                : ["Verify range address format", "Shift must be 'Down' or 'Right'", "Ensure sheet is not protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteCellsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? shift)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "delete-cells");
        if (string.IsNullOrEmpty(shift))
            ExcelToolsBase.ThrowMissingParameter("shift", "delete-cells");

        if (!Enum.TryParse<DeleteShiftDirection>(shift, true, out var shiftDirection))
        {
            throw new ModelContextProtocol.McpException($"Invalid shift direction '{shift}'. Must be 'Up' or 'Left'.");
        }

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.DeleteCellsAsync(batch, sheetName ?? "", rangeAddress!, shiftDirection));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Deleted cells at {rangeAddress}, shifted remaining cells {shift.ToLowerInvariant()}."
                : $"Failed to delete cells with {shift} shift. Verify range and shift direction.",
            suggestedNextActions = result.Success
                ? new[] { "Verify no #REF! errors in formulas", "Use 'insert-cells' to reverse operation", "Check surrounding data integrity" }
                : ["Verify range address format", "Shift must be 'Up' or 'Left'", "Ensure cells are not part of table or locked"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> InsertRowsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "insert-rows");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.InsertRowsAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Inserted entire row(s) above {rangeAddress}. Existing rows shifted down."
                : "Failed to insert rows. Verify range address.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'set-values' to populate new rows", "Check formulas for updated row references", "Use 'delete-rows' to reverse operation" }
                : ["Verify range address (e.g., 'A5' or 'A5:A10')", "Ensure sheet exists", "Check sheet is not protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteRowsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "delete-rows");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.DeleteRowsAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Deleted entire row(s) at {rangeAddress}. Rows below shifted up."
                : "Failed to delete rows. Verify range address.",
            suggestedNextActions = result.Success
                ? new[] { "Verify no #REF! errors in formulas", "Check tables adjusted boundaries correctly", "Use 'insert-rows' to reverse if needed" }
                : ["Verify range address format", "Ensure rows are not part of protected sheet", "Cannot delete if only row in sheet"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> InsertColumnsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "insert-columns");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.InsertColumnsAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Inserted entire column(s) to left of {rangeAddress}. Existing columns shifted right."
                : "Failed to insert columns. Verify range address.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'set-values' to populate new columns", "Check formulas for updated column references", "Use 'delete-columns' to reverse operation" }
                : ["Verify range address (e.g., 'C:C' or 'C1:C10')", "Ensure sheet exists", "Check sheet is not protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteColumnsAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "delete-columns");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.DeleteColumnsAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Deleted entire column(s) at {rangeAddress}. Columns to right shifted left."
                : "Failed to delete columns. Verify range address.",
            suggestedNextActions = result.Success
                ? new[] { "Verify no #REF! errors in formulas", "Check tables adjusted boundaries correctly", "Use 'insert-columns' to reverse if needed" }
                : ["Verify range address format", "Ensure columns are not part of protected sheet", "Cannot delete if only column in sheet"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === FIND/REPLACE OPERATIONS ===

    private static async Task<string> FindAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? searchValue, bool? matchCase, bool? matchEntireCell, bool? searchFormulas, bool? searchValues)
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

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.FindAsync(batch, sheetName ?? "", rangeAddress!, searchValue!, options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.RangeAddress,
            result.SearchValue,
            MatchingCells = result.MatchingCells.Take(10).ToList(),
            TotalMatches = result.MatchingCells.Count,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Found {result.MatchingCells.Count} matches for '{searchValue}' in {rangeAddress}. Showing first 10."
                : "Failed to search range. Verify range address and search options.",
            suggestedNextActions = result.Success && result.MatchingCells.Count > 0
                ? new[] { "Use 'replace' to update matching cells", "Use 'get-values' on specific cell addresses", "Refine search with matchCase or matchEntireCell options" }
                : result.Success
                    ? ["No matches found - try different search value", "Check searchFormulas/searchValues options", "Expand range to search larger area"]
                    : ["Verify range address format", "Ensure sheet exists with excel_worksheet 'list'", "Check search value is not empty"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ReplaceAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? searchValue, string? replaceValue, bool? matchCase, bool? matchEntireCell, bool? searchFormulas, bool? searchValues, bool? replaceAll)
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

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ReplaceAsync(batch, sheetName ?? "", rangeAddress!, searchValue!, replaceValue!, options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Replaced '{searchValue}' with '{replaceValue}' in {rangeAddress}. Changes saved."
                : "Failed to replace values. Verify search value exists.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'find' to verify all occurrences replaced", "Use 'get-values' to check specific cells", "Undo with excel_batch if needed (before commit)" }
                : ["Use 'find' first to verify matches exist", "Check searchValue spelling", "Verify replaceValue is valid for cell type"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === SORT OPERATIONS ===

    private static async Task<string> SortAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, List<SortColumn>? sortColumns, bool? hasHeaders)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "sort");
        if (sortColumns == null || sortColumns.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("sortColumns", "sort");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SortAsync(batch, sheetName ?? "", rangeAddress!, sortColumns!, hasHeaders ?? true));

        var sortDesc = string.Join(", ", sortColumns!.Select(c => $"Column {c.ColumnIndex} {(c.Ascending ? "asc" : "desc")}"));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Sorted {rangeAddress} by {sortDesc}. {(hasHeaders ?? true ? "Headers preserved" : "No headers")}."
                : "Failed to sort range. Verify range and sort columns.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-values' to verify sort order", "Headers treated as data if hasHeaders=false", "Excel supports max 3 sort levels" }
                : ["Verify columnIndex is 1-based within range", "Check range has data to sort", "Ensure sheet is not protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === DISCOVERY OPERATIONS ===

    private static async Task<string> GetUsedRangeAsync(RangeCommands commands, string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "get-used-range");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetUsedRangeAsync(batch, sheetName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.RangeAddress,
            result.Values,
            result.RowCount,
            result.ColumnCount,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Retrieved used range {result.RangeAddress} from '{sheetName}' ({result.RowCount}x{result.ColumnCount} cells with data)."
                : "Failed to get used range. Sheet may be empty or invalid.",
            suggestedNextActions = result.Success
                ? new[] { "Use this range address for other operations", "Clear unused areas with 'clear-all'", "Export to CSV or analyze with excel_table 'create'" }
                : ["Verify sheet exists with excel_worksheet 'list'", "Check if sheet has any data", "Empty sheets return no used range"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetCurrentRegionAsync(RangeCommands commands, string sessionId, string? sheetName, string? cellAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "get-current-region");
        if (string.IsNullOrEmpty(cellAddress))
            ExcelToolsBase.ThrowMissingParameter("cellAddress", "get-current-region");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetCurrentRegionAsync(batch, sheetName!, cellAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.RangeAddress,
            result.Values,
            result.RowCount,
            result.ColumnCount,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Found contiguous region {result.RangeAddress} around {cellAddress} ({result.RowCount}x{result.ColumnCount} cells)."
                : "Failed to get current region. Verify cell address.",
            suggestedNextActions = result.Success
                ? new[] { "Use range address for bulk operations", "Convert to table with excel_table 'create'", "Sort or filter this data block" }
                : ["Verify sheet name with excel_worksheet 'list'", "Check cell address format (e.g., 'A5')", "Cell must be part of data block"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetRangeInfoAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-range-info");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetInfoAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.Address,
            result.RowCount,
            result.ColumnCount,
            result.NumberFormat,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Range {result.Address} is {result.RowCount}x{result.ColumnCount} cells. Format: '{result.NumberFormat ?? "General"}'."
                : "Failed to get range info. Verify range address.",
            suggestedNextActions = result.Success
                ? new[] { "Use dimensions to validate data array sizes", "Check format before setting values", "Address shows absolute Excel reference" }
                : ["Verify range address format", "Ensure sheet exists", "Check range is valid (e.g., 'A1:D10')"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === HYPERLINK OPERATIONS ===

    private static async Task<string> AddHyperlinkAsync(RangeCommands commands, string sessionId, string? sheetName, string? cellAddress, string? url, string? displayText, string? tooltip)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "add-hyperlink");
        if (string.IsNullOrEmpty(cellAddress))
            ExcelToolsBase.ThrowMissingParameter("cellAddress", "add-hyperlink");
        if (string.IsNullOrEmpty(url))
            ExcelToolsBase.ThrowMissingParameter("url", "add-hyperlink");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.AddHyperlinkAsync(batch, sheetName!, cellAddress!, url!, displayText, tooltip));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Added hyperlink to {cellAddress}. Display: '{displayText ?? url}'. URL: {url}"
                : "Failed to add hyperlink. Verify cell address.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-hyperlink' to verify", "Add tooltip for user guidance", "Cell shows clickable link" }
                : ["Verify cell address format (e.g., 'A1')", "Check URL is valid", "Sheet must not be protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RemoveHyperlinkAsync(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "remove-hyperlink");
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "remove-hyperlink");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.RemoveHyperlinkAsync(batch, sheetName!, rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Removed hyperlink(s) from {rangeAddress}. Cell text preserved."
                : "Failed to remove hyperlink. Verify range has hyperlink.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'list-hyperlinks' to verify removal", "Cell formatting preserved", "Works on single cell or range" }
                : ["Check if cell/range has hyperlink", "Verify range address", "Use 'list-hyperlinks' to find hyperlinks"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListHyperlinksAsync(RangeCommands commands, string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "list-hyperlinks");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ListHyperlinksAsync(batch, sheetName!));

        var hyperlinkCount = result.Success ? ((dynamic)result).Hyperlinks?.Count ?? 0 : 0;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            ((dynamic)result).SheetName,
            ((dynamic)result).Hyperlinks,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Found {hyperlinkCount} hyperlink(s) in '{sheetName}'."
                : "Failed to list hyperlinks.",
            suggestedNextActions = result.Success && hyperlinkCount > 0
                ? new[] { "Use 'get-hyperlink' for specific cell details", "Use 'remove-hyperlink' to delete", "Hyperlinks show cell address and URL" }
                : result.Success
                    ? ["No hyperlinks in this sheet", "Add with 'add-hyperlink'", "Check other sheets"]
                    : ["Verify sheet exists with excel_worksheet 'list'", "Check sheet name spelling", "Sheet must not be empty"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetHyperlinkAsync(RangeCommands commands, string sessionId, string? sheetName, string? cellAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "get-hyperlink");
        if (string.IsNullOrEmpty(cellAddress))
            ExcelToolsBase.ThrowMissingParameter("cellAddress", "get-hyperlink");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetHyperlinkAsync(batch, sheetName!, cellAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            ((dynamic)result).CellAddress,
            ((dynamic)result).Url,
            ((dynamic)result).DisplayText,
            ((dynamic)result).Tooltip,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Retrieved hyperlink from {cellAddress}."
                : "No hyperlink at this cell. Verify cell address.",
            suggestedNextActions = result.Success
                ? new[] { "Update with 'add-hyperlink' (overwrites)", "Remove with 'remove-hyperlink'", "DisplayText shows in cell" }
                : ["Use 'list-hyperlinks' to find all hyperlinks", "Verify cell has hyperlink", "Check cell address format"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === FORMATTING OPERATIONS ===

    private static async Task<string> SetStyleAsync(
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

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SetStyleAsync(batch, sheetName ?? "", rangeAddress!, styleName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Applied style '{styleName}' to {rangeAddress}."
                : "Failed to apply style. Verify style name exists.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-style' to verify application", "Apply same style to other ranges", "Style includes font, fill, borders" }
                : ["Check built-in style names (e.g., 'Heading 1', 'Good', 'Bad')", "Verify range address", "Custom styles require Excel UI creation"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetStyleAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-style");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetStyleAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            success = true,
            sheetName,
            rangeAddress,
            styleName = result.StyleName,
            isBuiltInStyle = result.IsBuiltInStyle,
            styleDescription = result.StyleDescription,
            workflowHint = result.IsBuiltInStyle
                ? $"Range uses built-in style '{result.StyleName}'."
                : "Range uses custom or no style.",
            suggestedNextActions = result.IsBuiltInStyle
                ? new[] { "Apply this style to other ranges with 'set-style'", "Document style for reuse", "Built-in styles ensure consistency" }
                : ["Apply a built-in style with 'set-style'", "Use 'format-range' for custom formatting", "Check Excel UI for custom styles"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> FormatRangeAsync(
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

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.FormatRangeAsync(batch, sheetName ?? "", rangeAddress!,
                fontName, fontSize, bold, italic, underline, fontColor,
                fillColor, borderStyle, borderColor, borderWeight,
                horizontalAlignment, verticalAlignment, wrapText, orientation));

        var formatCount = new object?[] { fontName, fontSize, bold, italic, underline, fontColor, fillColor, borderStyle, borderColor, borderWeight, horizontalAlignment, verticalAlignment, wrapText, orientation }.Count(f => f != null);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Applied {formatCount} formatting attribute(s) to {rangeAddress}."
                : "Failed to format range. Verify parameters.",
            suggestedNextActions = result.Success
                ? new[] { "Combine with 'set-number-format' for complete formatting", "Use 'get-style' to document custom formatting", "Apply same format to other ranges" }
                : ["Check color format (hex: '#FF0000')", "Verify alignment values (left/center/right)", "Range must not be protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === VALIDATION OPERATIONS ===

    private static async Task<string> ValidateRangeAsync(
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

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ValidateRangeAsync(batch, sheetName ?? "", rangeAddress!,
                validationType!, validationOperator, validationFormula1, validationFormula2,
                showInputMessage, inputTitle, inputMessage,
                showErrorAlert, errorStyle, errorTitle, errorMessage,
                ignoreBlank, showDropdown));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Applied {validationType} validation to {rangeAddress}. Users will see {(showErrorAlert ?? false ? "error alerts" : "warnings")}."
                : "Failed to apply validation. Verify type and formulas.",
            suggestedNextActions = result.Success
                ? new[] { "Test validation by entering invalid data", "Use 'get-validation' to verify rules", "Dropdown shows for List validation type" }
                : ["Check validationType (List, WholeNumber, Decimal, Date, Time, TextLength, Custom)", "Verify formula1 syntax", "Range must not be protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetValidationAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-validation");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetValidationAsync(batch, sheetName ?? "", rangeAddress!));

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
            result.ErrorMessage,
            workflowHint = result.Success
                ? "Retrieved validation rules. Use for documentation or replication."
                : "No validation rules or failed to retrieve.",
            suggestedNextActions = result.Success
                ? new[] { "Apply same rules to other ranges with 'validate-range'", "Remove with 'remove-validation'", "Document rules for compliance" }
                : ["Range may not have validation rules", "Apply rules with 'validate-range'", "Check range address"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RemoveValidationAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "remove-validation");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.RemoveValidationAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Removed validation from {rangeAddress}. Cells accept any input now."
                : "Failed to remove validation.",
            suggestedNextActions = result.Success
                ? new[] { "Verify with 'get-validation'", "Apply new rules with 'validate-range'", "Data validation removed, not cell data" }
                : ["Range may not have validation", "Check range address", "Verify range is not protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> AutoFitColumnsAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "auto-fit-columns");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.AutoFitColumnsAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Auto-fitted columns in {rangeAddress} to content width."
                : "Failed to auto-fit columns.",
            suggestedNextActions = result.Success
                ? new[] { "Verify column widths in Excel UI", "Combine with 'auto-fit-rows' for full auto-fit", "Works best with wrapped text disabled" }
                : ["Verify range address", "Check if columns contain data", "Range must not be protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> AutoFitRowsAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "auto-fit-rows");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.AutoFitRowsAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Auto-fitted rows in {rangeAddress} to content height."
                : "Failed to auto-fit rows.",
            suggestedNextActions = result.Success
                ? new[] { "Verify row heights in Excel UI", "Enable wrap text for multi-line cells", "Combine with 'auto-fit-columns'" }
                : ["Verify range address", "Check if rows contain data", "Range must not be protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> MergeCellsAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "merge-cells");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.MergeCellsAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Merged cells in {rangeAddress}. Only top-left value preserved."
                : "Failed to merge cells.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-merge-info' to verify merge", "Unmerge with 'unmerge-cells'", "Only top-left value kept, others discarded" }
                : ["Verify range is multi-cell (e.g., 'A1:B2')", "Check if already merged", "Range must not be protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UnmergeCellsAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "unmerge-cells");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.UnmergeCellsAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Unmerged cells in {rangeAddress}. Value remains in original top-left cell."
                : "Failed to unmerge cells.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-merge-info' to verify unmerge", "Remerge with 'merge-cells'", "Value stays in top-left cell only" }
                : ["Check if range contains merged cells", "Verify range address", "Use 'get-merge-info' to find merges"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetMergeInfoAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-merge-info");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetMergeInfoAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            ((dynamic)result).IsMerged,
            ((dynamic)result).MergeAddress,
            result.ErrorMessage,
            workflowHint = result.Success
                ? ((dynamic)result).IsMerged ? $"Cell is part of merge: {((dynamic)result).MergeAddress}" : "Cell is not merged."
                : "Failed to get merge info.",
            suggestedNextActions = result.Success && ((dynamic)result).IsMerged
                ? new[] { "Unmerge with 'unmerge-cells'", "MergeAddress shows full merged range", "Only top-left cell holds value" }
                : result.Success
                    ? ["Merge cells with 'merge-cells'", "Check adjacent cells for merges", "Not merged - cells independent"]
                    : ["Verify cell address format", "Check range exists", "Ensure sheet is not empty"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> AddConditionalFormattingAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress,
        string? ruleType,
        string? formula1,
        string? formula2,
        string? formatStyle)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "add-conditional-formatting");
        if (string.IsNullOrEmpty(ruleType))
            ExcelToolsBase.ThrowMissingParameter("ruleType", "add-conditional-formatting");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.AddConditionalFormattingAsync(batch, sheetName ?? "", rangeAddress!,
                ruleType!, formula1, formula2, formatStyle));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Applied {ruleType} conditional formatting to {rangeAddress}."
                : "Failed to add conditional formatting.",
            suggestedNextActions = result.Success
                ? new[] { "Verify formatting in Excel UI", "Clear with 'clear-conditional-formatting'", "Rules: CellValue, Expression, ColorScale, DataBar, IconSet" }
                : ["Check ruleType (CellValue, Expression, ColorScale, DataBar, IconSet)", "Verify formula syntax", "Range must not be protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ClearConditionalFormattingAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-conditional-formatting");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ClearConditionalFormattingAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Cleared conditional formatting from {rangeAddress}. Static formatting preserved."
                : "Failed to clear conditional formatting.",
            suggestedNextActions = result.Success
                ? new[] { "Add new rules with 'add-conditional-formatting'", "Static cell formatting unaffected", "Rules removed, not cell values" }
                : ["Check if range has conditional formatting", "Verify range address", "Range must not be protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetCellLockAsync(
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

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SetCellLockAsync(batch, sheetName ?? "", rangeAddress!, locked!.Value));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? locked!.Value
                    ? $"Locked cells in {rangeAddress}. Prevents editing when sheet protected."
                    : $"Unlocked cells in {rangeAddress}. Allows editing even when sheet protected."
                : "Failed to set cell lock.",
            suggestedNextActions = result.Success
                ? new[] { "Protect sheet with excel_worksheet 'protect'", "Lock applies only when sheet protected", "Use 'get-cell-lock' to verify" }
                : ["Verify range address", "Check range is not protected", "Lock property independent of protection"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetCellLockAsync(
        RangeCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-cell-lock");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetCellLockAsync(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            ((dynamic)result).Locked,
            result.ErrorMessage,
            workflowHint = result.Success
                ? ((dynamic)result).Locked ? "Cells are locked. Prevents editing when sheet protected." : "Cells are unlocked. Allows editing even when sheet protected."
                : "Failed to get cell lock status.",
            suggestedNextActions = result.Success
                ? new[] { "Toggle with 'set-cell-lock'", "Lock effective only when sheet protected", "By default all cells are locked" }
                : ["Verify range address", "Check range exists", "Ensure sheet is not empty"]
        }, ExcelToolsBase.JsonOptions);
    }
}
