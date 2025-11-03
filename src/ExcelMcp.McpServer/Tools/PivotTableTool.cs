using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel PivotTable management tool for MCP server.
/// Provides complete PivotTable lifecycle, field management, and analysis capabilities.
///
/// LLM Usage Patterns:
/// - Use "create-from-range" to create PivotTables from data ranges with auto field-type detection
/// - Use "add-row-field" / "add-column-field" / "add-value-field" to build analysis structure
/// - Use "list-fields" to see available fields and their current placement
/// - Use "set-field-filter" to focus analysis on specific data subsets
/// - Use "get-data" to extract PivotTable results as 2D arrays for further analysis
///
/// IMPORTANT:
/// - PivotTables provide dynamic data summarization with drag-and-drop field configuration
/// - Field type detection (Numeric, Text, Date) guides appropriate aggregation functions
/// - Value fields validate aggregation functions (e.g., Sum only for numeric fields)
/// - All operations refresh PivotTable to materialize changes immediately
/// </summary>
[McpServerToolType]
public static class PivotTableTool
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false
    };

    /// <summary>
    /// Manage Excel PivotTables - comprehensive PivotTable creation, field management, and analysis
    /// </summary>
    [McpServerTool(Name = "excel_pivottable")]
    [Description("Manage Excel PivotTables for interactive data summarization. Create PivotTables from ranges, tables, or Data Model tables, add fields to Row/Column/Value/Filter areas, configure aggregations, apply filters, and extract results. Auto-detects field types (numeric, text, date) for LLM guidance. Supports: list, get-info, create-from-range, create-from-table, create-from-datamodel, delete, refresh, list-fields, add-row-field, add-column-field, add-value-field, add-filter-field, remove-field, set-field-function, set-field-name, set-field-format, get-data, set-field-filter, sort-field.")]
    public static async Task<string> PivotTable(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        PivotTableAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [StringLength(255, MinimumLength = 1)]
        [Description("PivotTable name (required for most actions)")]
        string? pivotTableName = null,

        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("Source or destination sheet name")]
        string? sheetName = null,

        [Description("Range address (e.g., 'A1:D100') for create-from-range")]
        string? range = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Field name for field operations")]
        string? fieldName = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Custom name for fields or table name for create-from-table")]
        string? customName = null,

        [Description("Aggregation function (Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP)")]
        string? aggregationFunction = null,

        [Description("Number format string (e.g., '$#,##0.00')")]
        string? numberFormat = null,

        [Description("JSON array of filter values (e.g., '[\"North\",\"South\"]')")]
        string? filterValues = null,

        [Description("Sort direction (Ascending or Descending)")]
        string? sortDirection = null,

        [Description("Destination sheet for create operations")]
        string? destinationSheet = null,

        [Description("Destination cell (e.g., 'A1') for create operations")]
        string? destinationCell = null,

        [Range(1, int.MaxValue)]
        [Description("Position in field area (1-based)")]
        int? position = null,

        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            var pivotCommands = new PivotTableCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                PivotTableAction.List => await ListPivotTables(pivotCommands, excelPath, batchId),
                PivotTableAction.Get => await GetPivotTableInfo(pivotCommands, excelPath, pivotTableName, batchId),
                PivotTableAction.CreateFromRange => await CreateFromRange(pivotCommands, excelPath, sheetName, range, destinationSheet, destinationCell, pivotTableName, batchId),
                PivotTableAction.CreateFromTable => await CreateFromTable(pivotCommands, excelPath, customName, destinationSheet, destinationCell, pivotTableName, batchId),
                PivotTableAction.CreateFromDataModel => await CreateFromDataModel(pivotCommands, excelPath, customName, destinationSheet, destinationCell, pivotTableName, batchId),
                PivotTableAction.Delete => await DeletePivotTable(pivotCommands, excelPath, pivotTableName, batchId),
                PivotTableAction.Refresh => await RefreshPivotTable(pivotCommands, excelPath, pivotTableName, batchId),
                PivotTableAction.ListFields => await ListFields(pivotCommands, excelPath, pivotTableName, batchId),
                PivotTableAction.AddRowField => await AddRowField(pivotCommands, excelPath, pivotTableName, fieldName, position, batchId),
                PivotTableAction.AddColumnField => await AddColumnField(pivotCommands, excelPath, pivotTableName, fieldName, position, batchId),
                PivotTableAction.AddValueField => await AddValueField(pivotCommands, excelPath, pivotTableName, fieldName, aggregationFunction, customName, batchId),
                PivotTableAction.AddFilterField => await AddFilterField(pivotCommands, excelPath, pivotTableName, fieldName, batchId),
                PivotTableAction.RemoveField => await RemoveField(pivotCommands, excelPath, pivotTableName, fieldName, batchId),
                PivotTableAction.SetFieldFunction => await SetFieldFunction(pivotCommands, excelPath, pivotTableName, fieldName, aggregationFunction, batchId),
                PivotTableAction.SetFieldName => await SetFieldName(pivotCommands, excelPath, pivotTableName, fieldName, customName, batchId),
                PivotTableAction.SetFieldFormat => await SetFieldFormat(pivotCommands, excelPath, pivotTableName, fieldName, numberFormat, batchId),
                PivotTableAction.GetData => await GetData(pivotCommands, excelPath, pivotTableName, batchId),
                PivotTableAction.SetFieldFilter => await SetFieldFilter(pivotCommands, excelPath, pivotTableName, fieldName, filterValues, batchId),
                PivotTableAction.SortField => await SortField(pivotCommands, excelPath, pivotTableName, fieldName, sortDirection, batchId),
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

    private static async Task<string> ListPivotTables(PivotTableCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            false,
            async (batch) => await commands.ListAsync(batch)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        if (result.PivotTables == null || !result.PivotTables.Any())
        {

        }

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> GetPivotTableInfo(PivotTableCommands commands, string filePath, string? pivotTableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "get-info");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            false,
            async (batch) => await commands.GetAsync(batch, pivotTableName!)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> CreateFromRange(PivotTableCommands commands, string filePath,
        string? sourceSheet, string? sourceRange, string? destSheet, string? destCell, string? pivotTableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(sourceSheet)) ExcelToolsBase.ThrowMissingParameter(nameof(sourceSheet), "create-from-range");
        if (string.IsNullOrWhiteSpace(sourceRange)) ExcelToolsBase.ThrowMissingParameter(nameof(sourceRange), "create-from-range");
        if (string.IsNullOrWhiteSpace(destSheet)) ExcelToolsBase.ThrowMissingParameter(nameof(destSheet), "create-from-range");
        if (string.IsNullOrWhiteSpace(destCell)) ExcelToolsBase.ThrowMissingParameter(nameof(destCell), "create-from-range");
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "create-from-range");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.CreateFromRangeAsync(batch, sourceSheet!, sourceRange!, destSheet!, destCell!, pivotTableName!)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> CreateFromTable(PivotTableCommands commands, string filePath,
        string? tableName, string? destSheet, string? destCell, string? pivotTableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "create-from-table");
        if (string.IsNullOrWhiteSpace(destSheet)) ExcelToolsBase.ThrowMissingParameter(nameof(destSheet), "create-from-table");
        if (string.IsNullOrWhiteSpace(destCell)) ExcelToolsBase.ThrowMissingParameter(nameof(destCell), "create-from-table");
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "create-from-table");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.CreateFromTableAsync(batch, tableName!, destSheet!, destCell!, pivotTableName!)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> CreateFromDataModel(PivotTableCommands commands, string filePath,
        string? tableName, string? destSheet, string? destCell, string? pivotTableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "create-from-datamodel");
        if (string.IsNullOrWhiteSpace(destSheet)) ExcelToolsBase.ThrowMissingParameter(nameof(destSheet), "create-from-datamodel");
        if (string.IsNullOrWhiteSpace(destCell)) ExcelToolsBase.ThrowMissingParameter(nameof(destCell), "create-from-datamodel");
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "create-from-datamodel");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.CreateFromDataModelAsync(batch, tableName!, destSheet!, destCell!, pivotTableName!)
        );

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"create-from-datamodel failed: {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> DeletePivotTable(PivotTableCommands commands, string filePath, string? pivotTableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "delete");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.DeleteAsync(batch, pivotTableName!)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> RefreshPivotTable(PivotTableCommands commands, string filePath, string? pivotTableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "refresh");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.RefreshAsync(batch, pivotTableName!)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> ListFields(PivotTableCommands commands, string filePath, string? pivotTableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "list-fields");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            false,
            async (batch) => await commands.ListFieldsAsync(batch, pivotTableName!)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> AddRowField(PivotTableCommands commands, string filePath, string? pivotTableName, string? fieldName, int? position, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "add-row-field");
        if (string.IsNullOrWhiteSpace(fieldName)) ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "add-row-field");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.AddRowFieldAsync(batch, pivotTableName!, fieldName!, position)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> AddColumnField(PivotTableCommands commands, string filePath, string? pivotTableName, string? fieldName, int? position, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "add-column-field");
        if (string.IsNullOrWhiteSpace(fieldName)) ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "add-column-field");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.AddColumnFieldAsync(batch, pivotTableName!, fieldName!, position)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> AddValueField(PivotTableCommands commands, string filePath, string? pivotTableName,
        string? fieldName, string? aggregationFunction, string? customName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "add-value-field");
        if (string.IsNullOrWhiteSpace(fieldName)) ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "add-value-field");

        // Parse aggregation function
        AggregationFunction function = AggregationFunction.Sum; // Default
        if (!string.IsNullOrEmpty(aggregationFunction))
        {
            if (!Enum.TryParse<AggregationFunction>(aggregationFunction, true, out function))
            {
                throw new ModelContextProtocol.McpException(
                    $"Invalid aggregation function '{aggregationFunction}'. Valid values: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP");
            }
        }

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.AddValueFieldAsync(batch, pivotTableName!, fieldName!, function, customName)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> AddFilterField(PivotTableCommands commands, string filePath, string? pivotTableName, string? fieldName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "add-filter-field");
        if (string.IsNullOrWhiteSpace(fieldName)) ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "add-filter-field");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.AddFilterFieldAsync(batch, pivotTableName!, fieldName!)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> RemoveField(PivotTableCommands commands, string filePath, string? pivotTableName, string? fieldName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "remove-field");
        if (string.IsNullOrWhiteSpace(fieldName)) ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "remove-field");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.RemoveFieldAsync(batch, pivotTableName!, fieldName!)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> SetFieldFunction(PivotTableCommands commands, string filePath, string? pivotTableName,
        string? fieldName, string? aggregationFunction, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "set-field-function");
        if (string.IsNullOrWhiteSpace(fieldName)) ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "set-field-function");
        if (string.IsNullOrWhiteSpace(aggregationFunction)) ExcelToolsBase.ThrowMissingParameter(nameof(aggregationFunction), "set-field-function");

        if (!Enum.TryParse<AggregationFunction>(aggregationFunction!, true, out var function))
        {
            throw new ModelContextProtocol.McpException(
                $"Invalid aggregation function '{aggregationFunction}'. Valid values: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP");
        }

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.SetFieldFunctionAsync(batch, pivotTableName!, fieldName!, function)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> SetFieldName(PivotTableCommands commands, string filePath, string? pivotTableName,
        string? fieldName, string? customName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "set-field-name");
        if (string.IsNullOrWhiteSpace(fieldName)) ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "set-field-name");
        if (string.IsNullOrWhiteSpace(customName)) ExcelToolsBase.ThrowMissingParameter(nameof(customName), "set-field-name");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.SetFieldNameAsync(batch, pivotTableName!, fieldName!, customName!)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> SetFieldFormat(PivotTableCommands commands, string filePath, string? pivotTableName,
        string? fieldName, string? numberFormat, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "set-field-format");
        if (string.IsNullOrWhiteSpace(fieldName)) ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "set-field-format");
        if (string.IsNullOrWhiteSpace(numberFormat)) ExcelToolsBase.ThrowMissingParameter(nameof(numberFormat), "set-field-format");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.SetFieldFormatAsync(batch, pivotTableName!, fieldName!, numberFormat!)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> GetData(PivotTableCommands commands, string filePath, string? pivotTableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "get-data");

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            false,
            async (batch) => await commands.GetDataAsync(batch, pivotTableName!)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> SetFieldFilter(PivotTableCommands commands, string filePath, string? pivotTableName,
        string? fieldName, string? filterValues, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "set-field-filter");
        if (string.IsNullOrWhiteSpace(fieldName)) ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "set-field-filter");
        if (string.IsNullOrWhiteSpace(filterValues)) ExcelToolsBase.ThrowMissingParameter(nameof(filterValues), "set-field-filter");

        // Parse JSON array of filter values
        List<string> values;
        try
        {
            values = JsonSerializer.Deserialize<List<string>>(filterValues!) ?? [];
        }
        catch (JsonException ex)
        {
            throw new ModelContextProtocol.McpException($"Invalid filterValues JSON: {ex.Message}. Expected format: '[\"value1\",\"value2\"]'");
        }

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.SetFieldFilterAsync(batch, pivotTableName!, fieldName!, values)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static async Task<string> SortField(PivotTableCommands commands, string filePath, string? pivotTableName,
        string? fieldName, string? sortDirection, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName)) ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "sort-field");
        if (string.IsNullOrWhiteSpace(fieldName)) ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "sort-field");

        // Parse sort direction
        SortDirection direction = SortDirection.Ascending; // Default
        if (!string.IsNullOrEmpty(sortDirection))
        {
            if (!Enum.TryParse<SortDirection>(sortDirection, true, out direction))
            {
                throw new ModelContextProtocol.McpException(
                    $"Invalid sort direction '{sortDirection}'. Valid values: Ascending, Descending");
            }
        }

        var result = await ExcelToolsBase.WithBatchAsync(batchId,
            filePath,
            true,
            async (batch) => await commands.SortFieldAsync(batch, pivotTableName!, fieldName!, direction)
        );

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, JsonOptions);
    }
}

