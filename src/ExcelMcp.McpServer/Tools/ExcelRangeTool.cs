using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for core Excel range operations - values, formulas, copy, clear, discovery.
/// Use excel_range_edit for insert/delete/find/sort. Use excel_range_format for styling/validation.
/// Use excel_range_link for hyperlinks and cell protection.
/// Calculation mode and explicit recalculation are handled by excel_calculation_mode, not this tool.
/// </summary>
[McpServerToolType]
public static partial class ExcelRangeTool
{
    /// <summary>
    /// Core range operations: get/set values and formulas, copy ranges, clear content, and discover data regions.
    ///
    /// BEST PRACTICE: Use 'get-values' to check existing data before overwriting.
    /// Use 'clear-contents' (not 'clear-all') to preserve cell formatting when clearing data.
    /// set-values preserves existing formatting; use set-number-format after if format change needed.
    ///
    /// DATA FORMAT: values and formulas are 2D JSON arrays representing rows and columns.
    /// Example: [[row1col1, row1col2], [row2col1, row2col2]]
    /// Single cell returns [[value]] (always 2D).
    ///
    /// REQUIRED PARAMETERS:
    /// - sheetName + rangeAddress for cell operations (e.g., sheetName='Sheet1', rangeAddress='A1:D10')
    /// - For named ranges, use sheetName='' (empty string) and rangeAddress='MyNamedRange'
    ///
    /// COPY OPERATIONS: Use sourceSheetName/sourceRangeAddress for source, targetSheetName/targetRangeAddress for destination.
    ///
    /// NUMBER FORMATS: Use US locale format codes (e.g., '#,##0.00', 'mm/dd/yyyy', '0.00%').
    /// </summary>
    /// <param name="action">The range operation to perform</param>
    /// <param name="excelPath">Full path to the Excel workbook file (e.g., 'C:\Reports\Sales.xlsx')</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action - required for all operations</param>
    /// <param name="sheetName">Name of the worksheet containing the range - REQUIRED for cell addresses, use empty string for named ranges only</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1', 'A1:D10', 'B:D') or named range name (e.g., 'SalesData')</param>
    /// <param name="values">2D array of values to set - rows are outer array, columns are inner array (e.g., [[1,2,3],[4,5,6]] for 2 rows x 3 cols)</param>
    /// <param name="formulas">2D array of formulas to set - include '=' prefix (e.g., [['=A1+B1', '=SUM(A:A)'], ['=C1*2', '=AVERAGE(B:B)']])</param>
    /// <param name="sourceSheetName">Source worksheet name for copy operations</param>
    /// <param name="sourceRangeAddress">Source range address for copy operations</param>
    /// <param name="targetSheetName">Target worksheet name for copy operations (defaults to source sheet if empty)</param>
    /// <param name="targetRangeAddress">Target range address for copy operations - can be single cell for paste destination</param>
    /// <param name="formatCode">Number format code in US locale (e.g., '#,##0.00' for numbers, 'mm/dd/yyyy' for dates, '0.00%' for percentages)</param>
    /// <param name="formatCodes">2D array of format codes to apply cell-by-cell - same dimensions as target range</param>
    /// <param name="cellAddress">Single cell address for get-current-region action (e.g., 'B5') - expands to contiguous data region around this cell</param>
    [McpServerTool(Name = "excel_range", Title = "Excel Range Operations", Destructive = true)]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelRange(
        RangeAction action,
        string excelPath,
        string sessionId,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? rangeAddress,
        [DefaultValue(null)] List<List<object?>>? values,
        [DefaultValue(null)] List<List<string>>? formulas,
        [DefaultValue(null)] string? sourceSheetName,
        [DefaultValue(null)] string? sourceRangeAddress,
        [DefaultValue(null)] string? targetSheetName,
        [DefaultValue(null)] string? targetRangeAddress,
        [DefaultValue(null)] string? formatCode,
        [DefaultValue(null)] List<List<string>>? formatCodes,
        [DefaultValue(null)] string? cellAddress)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_range",
            action.ToActionString(),
            excelPath,
            () =>
            {
                var rangeCommands = new RangeCommands();

                return action switch
                {
                    RangeAction.GetValues => GetValuesAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.SetValues => SetValuesAction(rangeCommands, sessionId, sheetName, rangeAddress, values),
                    RangeAction.GetFormulas => GetFormulasAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.SetFormulas => SetFormulasAction(rangeCommands, sessionId, sheetName, rangeAddress, formulas),
                    RangeAction.GetNumberFormats => GetNumberFormatsAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.SetNumberFormat => SetNumberFormatAction(rangeCommands, sessionId, sheetName, rangeAddress, formatCode),
                    RangeAction.SetNumberFormats => SetNumberFormatsAction(rangeCommands, sessionId, sheetName, rangeAddress, formatCodes),
                    RangeAction.ClearAll => ClearAllAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.ClearContents => ClearContentsAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.ClearFormats => ClearFormatsAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeAction.Copy => CopyAction(rangeCommands, sessionId, sourceSheetName, sourceRangeAddress, targetSheetName, targetRangeAddress),
                    RangeAction.CopyValues => CopyValuesAction(rangeCommands, sessionId, sourceSheetName, sourceRangeAddress, targetSheetName, targetRangeAddress),
                    RangeAction.CopyFormulas => CopyFormulasAction(rangeCommands, sessionId, sourceSheetName, sourceRangeAddress, targetSheetName, targetRangeAddress),
                    RangeAction.GetUsedRange => GetUsedRangeAction(rangeCommands, sessionId, sheetName),
                    RangeAction.GetCurrentRegion => GetCurrentRegionAction(rangeCommands, sessionId, sheetName, cellAddress),
                    RangeAction.GetInfo => GetRangeInfoAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    // === VALUE OPERATIONS ===

    private static string GetValuesAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-values");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetValues(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sheetName = result.SheetName,
            rangeAddress = result.RangeAddress,
            values = result.Values,
            rowCount = result.RowCount,
            columnCount = result.ColumnCount,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetValuesAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, List<List<object?>>? values)
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
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === FORMULA OPERATIONS ===

    private static string GetFormulasAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-formulas");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetFormulas(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sheetName = result.SheetName,
            rangeAddress = result.RangeAddress,
            formulas = result.Formulas,
            values = result.Values,
            rowCount = result.RowCount,
            columnCount = result.ColumnCount,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetFormulasAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, List<List<string>>? formulas)
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
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === NUMBER FORMAT OPERATIONS ===

    private static string GetNumberFormatsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-number-formats");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetNumberFormats(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sheetName = result.SheetName,
            rangeAddress = result.RangeAddress,
            formatCodes = result.Formats,
            rowCount = result.RowCount,
            columnCount = result.ColumnCount,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetNumberFormatAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? formatCode)
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
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetNumberFormatsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, List<List<string>>? formatCodes)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-number-formats");
        if (formatCodes == null || formatCodes.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("formatCodes", "set-number-formats");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.SetNumberFormats(batch, sheetName ?? "", rangeAddress!, formatCodes!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === CLEAR OPERATIONS ===

    private static string ClearAllAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-all");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.ClearAll(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ClearContentsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-contents");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.ClearContents(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ClearFormatsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-formats");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.ClearFormats(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === COPY OPERATIONS ===

    private static string CopyAction(RangeCommands commands, string sessionId, string? sourceSheetName, string? sourceRangeAddress, string? targetSheetName, string? targetRangeAddress)
    {
        if (string.IsNullOrEmpty(sourceRangeAddress))
            ExcelToolsBase.ThrowMissingParameter("sourceRangeAddress", "copy");
        if (string.IsNullOrEmpty(targetRangeAddress))
            ExcelToolsBase.ThrowMissingParameter("targetRangeAddress", "copy");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Copy(batch, sourceSheetName ?? "", sourceRangeAddress!, targetSheetName ?? "", targetRangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CopyValuesAction(RangeCommands commands, string sessionId, string? sourceSheetName, string? sourceRangeAddress, string? targetSheetName, string? targetRangeAddress)
    {
        if (string.IsNullOrEmpty(sourceRangeAddress))
            ExcelToolsBase.ThrowMissingParameter("sourceRangeAddress", "copy-values");
        if (string.IsNullOrEmpty(targetRangeAddress))
            ExcelToolsBase.ThrowMissingParameter("targetRangeAddress", "copy-values");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.CopyValues(batch, sourceSheetName ?? "", sourceRangeAddress!, targetSheetName ?? "", targetRangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CopyFormulasAction(RangeCommands commands, string sessionId, string? sourceSheetName, string? sourceRangeAddress, string? targetSheetName, string? targetRangeAddress)
    {
        if (string.IsNullOrEmpty(sourceRangeAddress))
            ExcelToolsBase.ThrowMissingParameter("sourceRangeAddress", "copy-formulas");
        if (string.IsNullOrEmpty(targetRangeAddress))
            ExcelToolsBase.ThrowMissingParameter("targetRangeAddress", "copy-formulas");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.CopyFormulas(batch, sourceSheetName ?? "", sourceRangeAddress!, targetSheetName ?? "", targetRangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === DISCOVERY OPERATIONS ===

    private static string GetUsedRangeAction(RangeCommands commands, string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "get-used-range");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetUsedRange(batch, sheetName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sheetName = result.SheetName,
            rangeAddress = result.RangeAddress,
            values = result.Values,
            rowCount = result.RowCount,
            columnCount = result.ColumnCount,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetCurrentRegionAction(RangeCommands commands, string sessionId, string? sheetName, string? cellAddress)
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
            sheetName = result.SheetName,
            rangeAddress = result.RangeAddress,
            values = result.Values,
            rowCount = result.RowCount,
            columnCount = result.ColumnCount,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetRangeInfoAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-info");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetInfo(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sheetName = result.SheetName,
            rangeAddress = result.Address,
            rowCount = result.RowCount,
            columnCount = result.ColumnCount,
            formatCode = result.NumberFormat,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }
}

