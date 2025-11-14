using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Models;

#pragma warning disable CA1861 // Avoid constant arrays as arguments - workflow hints are contextual per-call

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel PivotTable operations
/// </summary>
[McpServerToolType]
public static class ExcelPivotTableTool
{
    private static readonly JsonSerializerOptions JsonOptions = ExcelToolsBase.JsonOptions;

    [McpServerTool(Name = "excel_pivottable")]
    [Description(@"Excel PivotTable operations - interactive data analysis and summarization.")]
    public static async Task<string> ExcelPivotTable(
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        PivotTableAction action,

        [Description("Path to Excel file (.xlsx or .xlsm)")]
        string excelPath,

        [Required]
        [Description("Session ID from excel_file 'open' action")]
        string sessionId,

        [Description("PivotTable name")]
        string? pivotTableName = null,

        [Description("Source sheet name (for create-from-range)")]
        string? sheetName = null,

        [Description("Source range (for create-from-range)")]
        string? range = null,

        [Description("Excel Table name (for create-from-table)")]
        string? tableName = null,

        [Description("Data Model table name (for create-from-datamodel)")]
        string? dataModelTableName = null,

        [Description("Destination sheet for new PivotTable")]
        string? destinationSheet = null,

        [Description("Destination cell for new PivotTable")]
        string? destinationCell = null,

        [Description("Field name for field operations")]
        string? fieldName = null,

        [Description("Aggregation function: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP")]
        string? aggregationFunction = null,

        [Description("Custom display name for field")]
        string? customName = null,

        [Description("Number format code (e.g., '#,##0.00', '0.00%', 'm/d/yyyy')")]
        string? numberFormat = null,

        [Description("Position for field (1-based, optional)")]
        int? position = null,

        [Description("JSON array of filter values (e.g., '[\"value1\",\"value2\"]')")]
        string? filterValues = null,

        [Description("Sort direction: Ascending, Descending")]
        string? sortDirection = null)
    {
        var commands = new PivotTableCommands();

        try
        {
            return action switch
            {
                PivotTableAction.List => await ListAsync(commands, sessionId),
                PivotTableAction.Get => await GetAsync(commands, sessionId, pivotTableName),
                PivotTableAction.CreateFromRange => await CreateFromRangeAsync(commands, sessionId, sheetName, range, destinationSheet, destinationCell, pivotTableName),
                PivotTableAction.CreateFromTable => await CreateFromTableAsync(commands, sessionId, tableName, destinationSheet, destinationCell, pivotTableName),
                PivotTableAction.CreateFromDataModel => await CreateFromDataModelAsync(commands, sessionId, dataModelTableName, destinationSheet, destinationCell, pivotTableName),
                PivotTableAction.Delete => await DeleteAsync(commands, sessionId, pivotTableName),
                PivotTableAction.Refresh => await RefreshAsync(commands, sessionId, pivotTableName),
                PivotTableAction.ListFields => await ListFieldsAsync(commands, sessionId, pivotTableName),
                PivotTableAction.AddRowField => await AddRowFieldAsync(commands, sessionId, pivotTableName, fieldName, position),
                PivotTableAction.AddColumnField => await AddColumnFieldAsync(commands, sessionId, pivotTableName, fieldName, position),
                PivotTableAction.AddValueField => await AddValueFieldAsync(commands, sessionId, pivotTableName, fieldName, aggregationFunction, customName),
                PivotTableAction.AddFilterField => await AddFilterFieldAsync(commands, sessionId, pivotTableName, fieldName),
                PivotTableAction.RemoveField => await RemoveFieldAsync(commands, sessionId, pivotTableName, fieldName),
                PivotTableAction.SetFieldFunction => await SetFieldFunctionAsync(commands, sessionId, pivotTableName, fieldName, aggregationFunction),
                PivotTableAction.SetFieldName => await SetFieldNameAsync(commands, sessionId, pivotTableName, fieldName, customName),
                PivotTableAction.SetFieldFormat => await SetFieldFormatAsync(commands, sessionId, pivotTableName, fieldName, numberFormat),
                PivotTableAction.GetData => await GetDataAsync(commands, sessionId, pivotTableName),
                PivotTableAction.SetFieldFilter => await SetFieldFilterAsync(commands, sessionId, pivotTableName, fieldName, filterValues),
                PivotTableAction.SortField => await SortFieldAsync(commands, sessionId, pivotTableName, fieldName, sortDirection),
                _ => throw new ModelContextProtocol.McpException($"Unknown action: {action} ({action.ToActionString()})")
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

    private static async Task<string> ListAsync(
        PivotTableCommands commands,
        string sessionId)
    {
        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.ListAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTables,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> GetAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "get-info");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.GetAsync(batch, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTable,
            result.Fields,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> CreateFromRangeAsync(
        PivotTableCommands commands,
        string sessionId,
        string? sheetName,
        string? range,
        string? destinationSheet,
        string? destinationCell,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create-from-range");
        if (string.IsNullOrWhiteSpace(range))
            ExcelToolsBase.ThrowMissingParameter(nameof(range), "create-from-range");
        if (string.IsNullOrWhiteSpace(destinationSheet))
            ExcelToolsBase.ThrowMissingParameter(nameof(destinationSheet), "create-from-range");
        if (string.IsNullOrWhiteSpace(destinationCell))
            ExcelToolsBase.ThrowMissingParameter(nameof(destinationCell), "create-from-range");
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "create-from-range");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.CreateFromRangeAsync(batch, sheetName!, range!,
                destinationSheet!, destinationCell!, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTableName,
            result.SheetName,
            result.Range,
            result.SourceData,
            result.SourceRowCount,
            result.AvailableFields,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> CreateFromTableAsync(
        PivotTableCommands commands,
        string sessionId,
        string? tableName,
        string? destinationSheet,
        string? destinationCell,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "create-from-table");
        if (string.IsNullOrWhiteSpace(destinationSheet))
            ExcelToolsBase.ThrowMissingParameter(nameof(destinationSheet), "create-from-table");
        if (string.IsNullOrWhiteSpace(destinationCell))
            ExcelToolsBase.ThrowMissingParameter(nameof(destinationCell), "create-from-table");
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "create-from-table");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.CreateFromTableAsync(batch, tableName!,
                destinationSheet!, destinationCell!, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTableName,
            result.SheetName,
            result.Range,
            result.SourceData,
            result.SourceRowCount,
            result.AvailableFields,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> CreateFromDataModelAsync(
        PivotTableCommands commands,
        string sessionId,
        string? dataModelTableName,
        string? destinationSheet,
        string? destinationCell,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(dataModelTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(dataModelTableName), "create-from-datamodel");
        if (string.IsNullOrWhiteSpace(destinationSheet))
            ExcelToolsBase.ThrowMissingParameter(nameof(destinationSheet), "create-from-datamodel");
        if (string.IsNullOrWhiteSpace(destinationCell))
            ExcelToolsBase.ThrowMissingParameter(nameof(destinationCell), "create-from-datamodel");
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "create-from-datamodel");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.CreateFromDataModelAsync(batch, dataModelTableName!,
                destinationSheet!, destinationCell!, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTableName,
            result.SheetName,
            result.Range,
            result.SourceData,
            result.SourceRowCount,
            result.AvailableFields,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> DeleteAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "delete");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.DeleteAsync(batch, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> RefreshAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "refresh");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.RefreshAsync(batch, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTableName,
            result.RefreshTime,
            result.SourceRecordCount,
            result.PreviousRecordCount,
            result.StructureChanged,
            result.NewFields,
            result.RemovedFields,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> ListFieldsAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "list-fields");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.ListFieldsAsync(batch, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Fields,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> AddRowFieldAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName,
        int? position)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "add-row-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "add-row-field");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.AddRowFieldAsync(batch, pivotTableName!, fieldName!, position));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> AddColumnFieldAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName,
        int? position)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "add-column-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "add-column-field");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.AddColumnFieldAsync(batch, pivotTableName!, fieldName!, position));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> AddValueFieldAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName,
        string? aggregationFunction,
        string? customName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "add-value-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "add-value-field");

        // Parse aggregation function
        AggregationFunction function = AggregationFunction.Sum; // Default
        if (!string.IsNullOrEmpty(aggregationFunction) &&
            !Enum.TryParse(aggregationFunction, true, out function))
        {
            throw new ModelContextProtocol.McpException(
                $"Invalid aggregation function '{aggregationFunction}'. Valid values: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP");
        }

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.AddValueFieldAsync(batch, pivotTableName!, fieldName!, function, customName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> AddFilterFieldAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "add-filter-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "add-filter-field");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.AddFilterFieldAsync(batch, pivotTableName!, fieldName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> RemoveFieldAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "remove-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "remove-field");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.RemoveFieldAsync(batch, pivotTableName!, fieldName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> SetFieldFunctionAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName,
        string? aggregationFunction)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "set-field-function");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "set-field-function");
        if (string.IsNullOrWhiteSpace(aggregationFunction))
            ExcelToolsBase.ThrowMissingParameter(nameof(aggregationFunction), "set-field-function");

        if (!Enum.TryParse<AggregationFunction>(aggregationFunction!, true, out var function))
        {
            throw new ModelContextProtocol.McpException(
                $"Invalid aggregation function '{aggregationFunction}'. Valid values: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP");
        }

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.SetFieldFunctionAsync(batch, pivotTableName!, fieldName!, function));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> SetFieldNameAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName,
        string? customName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "set-field-name");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "set-field-name");
        if (string.IsNullOrWhiteSpace(customName))
            ExcelToolsBase.ThrowMissingParameter(nameof(customName), "set-field-name");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.SetFieldNameAsync(batch, pivotTableName!, fieldName!, customName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> SetFieldFormatAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName,
        string? numberFormat)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "set-field-format");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "set-field-format");
        if (string.IsNullOrWhiteSpace(numberFormat))
            ExcelToolsBase.ThrowMissingParameter(nameof(numberFormat), "set-field-format");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.SetFieldFormatAsync(batch, pivotTableName!, fieldName!, numberFormat!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> GetDataAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "get-data");

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.GetDataAsync(batch, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTableName,
            result.Values,
            result.ColumnHeaders,
            result.RowHeaders,
            result.DataRowCount,
            result.DataColumnCount,
            result.GrandTotals,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> SetFieldFilterAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName,
        string? filterValues)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "set-field-filter");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "set-field-filter");
        if (string.IsNullOrWhiteSpace(filterValues))
            ExcelToolsBase.ThrowMissingParameter(nameof(filterValues), "set-field-filter");

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

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.SetFieldFilterAsync(batch, pivotTableName!, fieldName!, values));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.SelectedItems,
            result.AvailableItems,
            result.VisibleRowCount,
            result.TotalRowCount,
            result.ShowAll,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static async Task<string> SortFieldAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName,
        string? sortDirection)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "sort-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "sort-field");

        // Parse sort direction
        SortDirection direction = SortDirection.Ascending; // Default
        if (!string.IsNullOrEmpty(sortDirection) &&
            !Enum.TryParse(sortDirection, true, out direction))
        {
            throw new ModelContextProtocol.McpException(
                $"Invalid sort direction '{sortDirection}'. Valid values: Ascending, Descending");
        }

        var result = await ExcelToolsBase.WithSessionAsync(sessionId,
            async batch => await commands.SortFieldAsync(batch, pivotTableName!, fieldName!, direction));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }
}
