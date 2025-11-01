using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.Range;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel range operations tool for MCP server - unified API for all range data manipulation.
/// Handles values, formulas, clearing, copying, inserting, deleting, finding, sorting, and hyperlinks.
/// Single cell = 1x1 range. Named ranges work transparently via rangeAddress parameter.
///
/// LLM Usage Patterns:
/// - Use "get-values" to read cell/range data (supports named ranges)
/// - Use "set-values" to write data to cells/ranges
/// - Use "get-formulas"/"set-formulas" for formula operations
/// - Use "clear-all"/"clear-contents"/"clear-formats" to empty ranges
/// - Use "copy"/"copy-values"/"copy-formulas" to duplicate data
/// - Use "insert-cells"/"delete-cells" to shift cells
/// - Use "insert-rows"/"delete-rows" for entire row operations
/// - Use "insert-columns"/"delete-columns" for entire column operations
/// - Use "find"/"replace" to search and modify content
/// - Use "sort" to order data
/// - Use "get-used-range" to discover data boundaries
/// - Use "get-current-region" to find contiguous data blocks
/// - Use "get-range-info" to inspect range properties
/// - Use "add-hyperlink"/"remove-hyperlink"/"list-hyperlinks"/"get-hyperlink" for hyperlink management
/// </summary>
[McpServerToolType]
public static class ExcelRangeTool
{
    /// <summary>
    /// Unified Excel range operations - comprehensive data manipulation API.
    /// Supports: values, formulas, number formats, clear, copy, insert/delete, find/replace, sort, discovery, hyperlinks.
    /// Optional batchId for batch sessions.
    /// </summary>
    [McpServerTool(Name = "excel_range")]
    [Description("Excel range operations: get-values, set-values, get-formulas, set-formulas, get-number-formats, set-number-format, set-number-formats, clear-all, clear-contents, clear-formats, copy, copy-values, copy-formulas, insert-cells, delete-cells, insert-rows, delete-rows, insert-columns, delete-columns, find, replace, sort, get-used-range, get-current-region, get-range-info, add-hyperlink, remove-hyperlink, list-hyperlinks, get-hyperlink. Optional batchId for batch sessions.")]
    public static async Task<string> ExcelRange(
        [Required]
        [RegularExpression("^(get-values|set-values|get-formulas|set-formulas|get-number-formats|set-number-format|set-number-formats|clear-all|clear-contents|clear-formats|copy|copy-values|copy-formulas|insert-cells|delete-cells|insert-rows|delete-rows|insert-columns|delete-columns|find|replace|sort|get-used-range|get-current-region|get-range-info|add-hyperlink|remove-hyperlink|list-hyperlinks|get-hyperlink)$")]
        [Description("Action: get-values, set-values, get-formulas, set-formulas, get-number-formats, set-number-format, set-number-formats, clear-all, clear-contents, clear-formats, copy, copy-values, copy-formulas, insert-cells, delete-cells, insert-rows, delete-rows, insert-columns, delete-columns, find, replace, sort, get-used-range, get-current-region, get-range-info, add-hyperlink, remove-hyperlink, list-hyperlinks, get-hyperlink")]
        string action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

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

        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            var rangeCommands = new RangeCommands();

            return action.ToLowerInvariant() switch
            {
                "get-values" => await GetValuesAsync(rangeCommands, excelPath, sheetName, rangeAddress, batchId),
                "set-values" => await SetValuesAsync(rangeCommands, excelPath, sheetName, rangeAddress, values, batchId),
                "get-formulas" => await GetFormulasAsync(rangeCommands, excelPath, sheetName, rangeAddress, batchId),
                "set-formulas" => await SetFormulasAsync(rangeCommands, excelPath, sheetName, rangeAddress, formulas, batchId),
                "get-number-formats" => await GetNumberFormatsAsync(rangeCommands, excelPath, sheetName, rangeAddress, batchId),
                "set-number-format" => await SetNumberFormatAsync(rangeCommands, excelPath, sheetName, rangeAddress, formatCode, batchId),
                "set-number-formats" => await SetNumberFormatsAsync(rangeCommands, excelPath, sheetName, rangeAddress, formats, batchId),
                "clear-all" => await ClearAllAsync(rangeCommands, excelPath, sheetName, rangeAddress, batchId),
                "clear-contents" => await ClearContentsAsync(rangeCommands, excelPath, sheetName, rangeAddress, batchId),
                "clear-formats" => await ClearFormatsAsync(rangeCommands, excelPath, sheetName, rangeAddress, batchId),
                "copy" => await CopyAsync(rangeCommands, excelPath, sourceSheet, sourceRange, targetSheet, targetRange, batchId),
                "copy-values" => await CopyValuesAsync(rangeCommands, excelPath, sourceSheet, sourceRange, targetSheet, targetRange, batchId),
                "copy-formulas" => await CopyFormulasAsync(rangeCommands, excelPath, sourceSheet, sourceRange, targetSheet, targetRange, batchId),
                "insert-cells" => await InsertCellsAsync(rangeCommands, excelPath, sheetName, rangeAddress, shift, batchId),
                "delete-cells" => await DeleteCellsAsync(rangeCommands, excelPath, sheetName, rangeAddress, shift, batchId),
                "insert-rows" => await InsertRowsAsync(rangeCommands, excelPath, sheetName, rangeAddress, batchId),
                "delete-rows" => await DeleteRowsAsync(rangeCommands, excelPath, sheetName, rangeAddress, batchId),
                "insert-columns" => await InsertColumnsAsync(rangeCommands, excelPath, sheetName, rangeAddress, batchId),
                "delete-columns" => await DeleteColumnsAsync(rangeCommands, excelPath, sheetName, rangeAddress, batchId),
                "find" => await FindAsync(rangeCommands, excelPath, sheetName, rangeAddress, searchValue, matchCase, matchEntireCell, searchFormulas, searchValues, batchId),
                "replace" => await ReplaceAsync(rangeCommands, excelPath, sheetName, rangeAddress, searchValue, replaceValue, matchCase, matchEntireCell, searchFormulas, searchValues, replaceAll, batchId),
                "sort" => await SortAsync(rangeCommands, excelPath, sheetName, rangeAddress, sortColumns, hasHeaders, batchId),
                "get-used-range" => await GetUsedRangeAsync(rangeCommands, excelPath, sheetName, batchId),
                "get-current-region" => await GetCurrentRegionAsync(rangeCommands, excelPath, sheetName, cellAddress, batchId),
                "get-range-info" => await GetRangeInfoAsync(rangeCommands, excelPath, sheetName, rangeAddress, batchId),
                "add-hyperlink" => await AddHyperlinkAsync(rangeCommands, excelPath, sheetName, cellAddress, url, displayText, tooltip, batchId),
                "remove-hyperlink" => await RemoveHyperlinkAsync(rangeCommands, excelPath, sheetName, rangeAddress, batchId),
                "list-hyperlinks" => await ListHyperlinksAsync(rangeCommands, excelPath, sheetName, batchId),
                "get-hyperlink" => await GetHyperlinkAsync(rangeCommands, excelPath, sheetName, cellAddress, batchId),
                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action '{action}'. Supported: get-values, set-values, get-formulas, set-formulas, get-number-formats, set-number-format, set-number-formats, clear-all, clear-contents, clear-formats, copy, copy-values, copy-formulas, insert-cells, delete-cells, insert-rows, delete-rows, insert-columns, delete-columns, find, replace, sort, get-used-range, get-current-region, get-range-info, add-hyperlink, remove-hyperlink, list-hyperlinks, get-hyperlink")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action, excelPath);
            throw; // Unreachable but satisfies compiler
        }
    }

    // === VALUE OPERATIONS ===

    private static async Task<string> GetValuesAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-values");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetValuesAsync(batch, sheetName ?? "", rangeAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-values failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetValuesAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, List<List<object?>>? values, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-values");
        if (values == null || values.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("values", "set-values");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.SetValuesAsync(batch, sheetName ?? "", rangeAddress!, values!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-values failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // === FORMULA OPERATIONS ===

    private static async Task<string> GetFormulasAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-formulas");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetFormulasAsync(batch, sheetName ?? "", rangeAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-formulas failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetFormulasAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, List<List<string>>? formulas, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-formulas");
        if (formulas == null || formulas.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("formulas", "set-formulas");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.SetFormulasAsync(batch, sheetName ?? "", rangeAddress!, formulas!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-formulas failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // === NUMBER FORMAT OPERATIONS ===

    private static async Task<string> GetNumberFormatsAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-number-formats");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetNumberFormatsAsync(batch, sheetName ?? "", rangeAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-number-formats failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetNumberFormatAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? formatCode, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-number-format");
        if (string.IsNullOrEmpty(formatCode))
            ExcelToolsBase.ThrowMissingParameter("formatCode", "set-number-format");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.SetNumberFormatAsync(batch, sheetName ?? "", rangeAddress!, formatCode!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-number-format failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetNumberFormatsAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, List<List<string>>? formats, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-number-formats");
        if (formats == null || formats.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("formats", "set-number-formats");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.SetNumberFormatsAsync(batch, sheetName ?? "", rangeAddress!, formats!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-number-formats failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // === CLEAR OPERATIONS ===

    private static async Task<string> ClearAllAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-all");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.ClearAllAsync(batch, sheetName ?? "", rangeAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"clear-all failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ClearContentsAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-contents");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.ClearContentsAsync(batch, sheetName ?? "", rangeAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"clear-contents failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ClearFormatsAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "clear-formats");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.ClearFormatsAsync(batch, sheetName ?? "", rangeAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"clear-formats failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // === COPY OPERATIONS ===

    private static async Task<string> CopyAsync(RangeCommands commands, string filePath, string? sourceSheet, string? sourceRange, string? targetSheet, string? targetRange, string? batchId)
    {
        if (string.IsNullOrEmpty(sourceRange))
            ExcelToolsBase.ThrowMissingParameter("sourceRange", "copy");
        if (string.IsNullOrEmpty(targetRange))
            ExcelToolsBase.ThrowMissingParameter("targetRange", "copy");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.CopyAsync(batch, sourceSheet ?? "", sourceRange!, targetSheet ?? "", targetRange!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"copy failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CopyValuesAsync(RangeCommands commands, string filePath, string? sourceSheet, string? sourceRange, string? targetSheet, string? targetRange, string? batchId)
    {
        if (string.IsNullOrEmpty(sourceRange))
            ExcelToolsBase.ThrowMissingParameter("sourceRange", "copy-values");
        if (string.IsNullOrEmpty(targetRange))
            ExcelToolsBase.ThrowMissingParameter("targetRange", "copy-values");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.CopyValuesAsync(batch, sourceSheet ?? "", sourceRange!, targetSheet ?? "", targetRange!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"copy-values failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CopyFormulasAsync(RangeCommands commands, string filePath, string? sourceSheet, string? sourceRange, string? targetSheet, string? targetRange, string? batchId)
    {
        if (string.IsNullOrEmpty(sourceRange))
            ExcelToolsBase.ThrowMissingParameter("sourceRange", "copy-formulas");
        if (string.IsNullOrEmpty(targetRange))
            ExcelToolsBase.ThrowMissingParameter("targetRange", "copy-formulas");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.CopyFormulasAsync(batch, sourceSheet ?? "", sourceRange!, targetSheet ?? "", targetRange!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"copy-formulas failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // === INSERT/DELETE OPERATIONS ===

    private static async Task<string> InsertCellsAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? shift, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "insert-cells");
        if (string.IsNullOrEmpty(shift))
            ExcelToolsBase.ThrowMissingParameter("shift", "insert-cells");

        if (!Enum.TryParse<InsertShiftDirection>(shift, true, out var shiftDirection))
        {
            throw new ModelContextProtocol.McpException($"Invalid shift direction '{shift}'. Must be 'Down' or 'Right'.");
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.InsertCellsAsync(batch, sheetName ?? "", rangeAddress!, shiftDirection));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"insert-cells failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteCellsAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? shift, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "delete-cells");
        if (string.IsNullOrEmpty(shift))
            ExcelToolsBase.ThrowMissingParameter("shift", "delete-cells");

        if (!Enum.TryParse<DeleteShiftDirection>(shift, true, out var shiftDirection))
        {
            throw new ModelContextProtocol.McpException($"Invalid shift direction '{shift}'. Must be 'Up' or 'Left'.");
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.DeleteCellsAsync(batch, sheetName ?? "", rangeAddress!, shiftDirection));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"delete-cells failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> InsertRowsAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "insert-rows");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.InsertRowsAsync(batch, sheetName ?? "", rangeAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"insert-rows failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteRowsAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "delete-rows");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.DeleteRowsAsync(batch, sheetName ?? "", rangeAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"delete-rows failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> InsertColumnsAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "insert-columns");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.InsertColumnsAsync(batch, sheetName ?? "", rangeAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"insert-columns failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteColumnsAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "delete-columns");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.DeleteColumnsAsync(batch, sheetName ?? "", rangeAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"delete-columns failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // === FIND/REPLACE OPERATIONS ===

    private static async Task<string> FindAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? searchValue, bool? matchCase, bool? matchEntireCell, bool? searchFormulas, bool? searchValues, string? batchId)
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

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.FindAsync(batch, sheetName ?? "", rangeAddress!, searchValue!, options));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"find failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ReplaceAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? searchValue, string? replaceValue, bool? matchCase, bool? matchEntireCell, bool? searchFormulas, bool? searchValues, bool? replaceAll, string? batchId)
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

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.ReplaceAsync(batch, sheetName ?? "", rangeAddress!, searchValue!, replaceValue!, options));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"replace failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // === SORT OPERATIONS ===

    private static async Task<string> SortAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, List<SortColumn>? sortColumns, bool? hasHeaders, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "sort");
        if (sortColumns == null || sortColumns.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("sortColumns", "sort");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.SortAsync(batch, sheetName ?? "", rangeAddress!, sortColumns!, hasHeaders ?? true));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"sort failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // === DISCOVERY OPERATIONS ===

    private static async Task<string> GetUsedRangeAsync(RangeCommands commands, string filePath, string? sheetName, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "get-used-range");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetUsedRangeAsync(batch, sheetName!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-used-range failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetCurrentRegionAsync(RangeCommands commands, string filePath, string? sheetName, string? cellAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "get-current-region");
        if (string.IsNullOrEmpty(cellAddress))
            ExcelToolsBase.ThrowMissingParameter("cellAddress", "get-current-region");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetCurrentRegionAsync(batch, sheetName!, cellAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-current-region failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetRangeInfoAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-range-info");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetRangeInfoAsync(batch, sheetName ?? "", rangeAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-range-info failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // === HYPERLINK OPERATIONS ===

    private static async Task<string> AddHyperlinkAsync(RangeCommands commands, string filePath, string? sheetName, string? cellAddress, string? url, string? displayText, string? tooltip, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "add-hyperlink");
        if (string.IsNullOrEmpty(cellAddress))
            ExcelToolsBase.ThrowMissingParameter("cellAddress", "add-hyperlink");
        if (string.IsNullOrEmpty(url))
            ExcelToolsBase.ThrowMissingParameter("url", "add-hyperlink");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.AddHyperlinkAsync(batch, sheetName!, cellAddress!, url!, displayText, tooltip));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"add-hyperlink failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RemoveHyperlinkAsync(RangeCommands commands, string filePath, string? sheetName, string? rangeAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "remove-hyperlink");
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "remove-hyperlink");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.RemoveHyperlinkAsync(batch, sheetName!, rangeAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"remove-hyperlink failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListHyperlinksAsync(RangeCommands commands, string filePath, string? sheetName, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "list-hyperlinks");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ListHyperlinksAsync(batch, sheetName!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"list-hyperlinks failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetHyperlinkAsync(RangeCommands commands, string filePath, string? sheetName, string? cellAddress, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "get-hyperlink");
        if (string.IsNullOrEmpty(cellAddress))
            ExcelToolsBase.ThrowMissingParameter("cellAddress", "get-hyperlink");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetHyperlinkAsync(batch, sheetName!, cellAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-hyperlink failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
