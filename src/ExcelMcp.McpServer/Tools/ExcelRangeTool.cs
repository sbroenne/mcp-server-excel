using System.ComponentModel;
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
            ServiceRegistry.Range.ToActionString(action),
            () => action switch
            {
                RangeAction.GetValues => ForwardGetValues(sessionId, sheetName, rangeAddress),
                RangeAction.SetValues => ForwardSetValues(sessionId, sheetName, rangeAddress, values),
                RangeAction.GetFormulas => ForwardGetFormulas(sessionId, sheetName, rangeAddress),
                RangeAction.SetFormulas => ForwardSetFormulas(sessionId, sheetName, rangeAddress, formulas),
                RangeAction.GetNumberFormats => ForwardGetNumberFormats(sessionId, sheetName, rangeAddress),
                RangeAction.SetNumberFormat => ForwardSetNumberFormat(sessionId, sheetName, rangeAddress, formatCode),
                RangeAction.SetNumberFormats => ForwardSetNumberFormats(sessionId, sheetName, rangeAddress, formatCodes),
                RangeAction.ClearAll => ForwardClearAll(sessionId, sheetName, rangeAddress),
                RangeAction.ClearContents => ForwardClearContents(sessionId, sheetName, rangeAddress),
                RangeAction.ClearFormats => ForwardClearFormats(sessionId, sheetName, rangeAddress),
                RangeAction.Copy => ForwardCopy(sessionId, sourceSheetName, sourceRangeAddress, targetSheetName, targetRangeAddress),
                RangeAction.CopyValues => ForwardCopyValues(sessionId, sourceSheetName, sourceRangeAddress, targetSheetName, targetRangeAddress),
                RangeAction.CopyFormulas => ForwardCopyFormulas(sessionId, sourceSheetName, sourceRangeAddress, targetSheetName, targetRangeAddress),
                RangeAction.GetUsedRange => ForwardGetUsedRange(sessionId, sheetName),
                RangeAction.GetCurrentRegion => ForwardGetCurrentRegion(sessionId, sheetName, cellAddress),
                RangeAction.GetInfo => ForwardGetInfo(sessionId, sheetName, rangeAddress),
                _ => throw new ArgumentException($"Unknown action: {action} ({ServiceRegistry.Range.ToActionString(action)})", nameof(action))
            });
    }

    // === VALUE OPERATIONS ===

    private static string ForwardGetValues(string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-values");

        return ExcelToolsBase.ForwardToService("range.get-values", sessionId, new { sheetName = sheetName ?? "", range = rangeAddress });
    }

    private static string ForwardSetValues(string sessionId, string? sheetName, string? rangeAddress, List<List<object?>>? values)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-values");
        if (values == null || values.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("values", "set-values");

        return ExcelToolsBase.ForwardToService("range.set-values", sessionId, new { sheetName = sheetName ?? "", range = rangeAddress, values });
    }

    // === FORMULA OPERATIONS ===

    private static string ForwardGetFormulas(string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-formulas");

        return ExcelToolsBase.ForwardToService("range.get-formulas", sessionId, new { sheetName = sheetName ?? "", range = rangeAddress });
    }

    private static string ForwardSetFormulas(string sessionId, string? sheetName, string? rangeAddress, List<List<string>>? formulas)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-formulas");
        if (formulas == null || formulas.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("formulas", "set-formulas");

        return ExcelToolsBase.ForwardToService("range.set-formulas", sessionId, new { sheetName = sheetName ?? "", range = rangeAddress, formulas });
    }

    // === NUMBER FORMAT OPERATIONS ===

    private static string ForwardGetNumberFormats(string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-number-formats");

        return ExcelToolsBase.ForwardToService("range.get-number-formats", sessionId, new { sheetName = sheetName ?? "", range = rangeAddress });
    }

    private static string ForwardSetNumberFormat(string sessionId, string? sheetName, string? rangeAddress, string? formatCode)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-number-format");
        if (string.IsNullOrEmpty(formatCode))
            ExcelToolsBase.ThrowMissingParameter("formatCode", "set-number-format");

        return ExcelToolsBase.ForwardToService("range.set-number-format", sessionId, new { sheetName = sheetName ?? "", range = rangeAddress, formatCode });
    }

    private static string ForwardSetNumberFormats(string sessionId, string? sheetName, string? rangeAddress, List<List<string>>? formatCodes)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-number-formats");
        if (formatCodes == null || formatCodes.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("formatCodes", "set-number-formats");

        return ExcelToolsBase.ForwardToService("range.set-number-formats", sessionId, new { sheetName = sheetName ?? "", range = rangeAddress, formats = formatCodes });
    }

    // === CLEAR OPERATIONS ===

    private static string ForwardClearAll(string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-all");

        return ExcelToolsBase.ForwardToService("range.clear-all", sessionId, new { sheetName = sheetName ?? "", range = rangeAddress });
    }

    private static string ForwardClearContents(string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-contents");

        return ExcelToolsBase.ForwardToService("range.clear-contents", sessionId, new { sheetName = sheetName ?? "", range = rangeAddress });
    }

    private static string ForwardClearFormats(string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-formats");

        return ExcelToolsBase.ForwardToService("range.clear-formats", sessionId, new { sheetName = sheetName ?? "", range = rangeAddress });
    }

    // === COPY OPERATIONS ===

    private static string ForwardCopy(string sessionId, string? sourceSheetName, string? sourceRangeAddress, string? targetSheetName, string? targetRangeAddress)
    {
        if (string.IsNullOrEmpty(sourceRangeAddress))
            ExcelToolsBase.ThrowMissingParameter("sourceRangeAddress", "copy");
        if (string.IsNullOrEmpty(targetRangeAddress))
            ExcelToolsBase.ThrowMissingParameter("targetRangeAddress", "copy");

        return ExcelToolsBase.ForwardToService("range.copy", sessionId, new
        {
            sourceSheet = sourceSheetName ?? "",
            sourceRange = sourceRangeAddress,
            targetSheet = targetSheetName ?? "",
            targetRange = targetRangeAddress
        });
    }

    private static string ForwardCopyValues(string sessionId, string? sourceSheetName, string? sourceRangeAddress, string? targetSheetName, string? targetRangeAddress)
    {
        if (string.IsNullOrEmpty(sourceRangeAddress))
            ExcelToolsBase.ThrowMissingParameter("sourceRangeAddress", "copy-values");
        if (string.IsNullOrEmpty(targetRangeAddress))
            ExcelToolsBase.ThrowMissingParameter("targetRangeAddress", "copy-values");

        return ExcelToolsBase.ForwardToService("range.copy-values", sessionId, new
        {
            sourceSheet = sourceSheetName ?? "",
            sourceRange = sourceRangeAddress,
            targetSheet = targetSheetName ?? "",
            targetRange = targetRangeAddress
        });
    }

    private static string ForwardCopyFormulas(string sessionId, string? sourceSheetName, string? sourceRangeAddress, string? targetSheetName, string? targetRangeAddress)
    {
        if (string.IsNullOrEmpty(sourceRangeAddress))
            ExcelToolsBase.ThrowMissingParameter("sourceRangeAddress", "copy-formulas");
        if (string.IsNullOrEmpty(targetRangeAddress))
            ExcelToolsBase.ThrowMissingParameter("targetRangeAddress", "copy-formulas");

        return ExcelToolsBase.ForwardToService("range.copy-formulas", sessionId, new
        {
            sourceSheet = sourceSheetName ?? "",
            sourceRange = sourceRangeAddress,
            targetSheet = targetSheetName ?? "",
            targetRange = targetRangeAddress
        });
    }

    // === DISCOVERY OPERATIONS ===

    private static string ForwardGetUsedRange(string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "get-used-range");

        return ExcelToolsBase.ForwardToService("range.get-used-range", sessionId, new { sheetName });
    }

    private static string ForwardGetCurrentRegion(string sessionId, string? sheetName, string? cellAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "get-current-region");
        if (string.IsNullOrEmpty(cellAddress))
            ExcelToolsBase.ThrowMissingParameter("cellAddress", "get-current-region");

        return ExcelToolsBase.ForwardToService("range.get-current-region", sessionId, new { sheetName, cellAddress });
    }

    private static string ForwardGetInfo(string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-info");

        return ExcelToolsBase.ForwardToService("range.get-info", sessionId, new { sheetName = sheetName ?? "", range = rangeAddress });
    }
}





