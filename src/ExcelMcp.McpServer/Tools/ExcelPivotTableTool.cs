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
    public static string ExcelPivotTable(
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
                PivotTableAction.List => ListAsync(commands, sessionId),
                PivotTableAction.Get => GetAsync(commands, sessionId, pivotTableName),
                PivotTableAction.CreateFromRange => CreateFromRangeAsync(commands, sessionId, sheetName, range, destinationSheet, destinationCell, pivotTableName),
                PivotTableAction.CreateFromTable => CreateFromTableAsync(commands, sessionId, tableName, destinationSheet, destinationCell, pivotTableName),
                PivotTableAction.CreateFromDataModel => CreateFromDataModelAsync(commands, sessionId, dataModelTableName, destinationSheet, destinationCell, pivotTableName),
                PivotTableAction.Delete => DeleteAsync(commands, sessionId, pivotTableName),
                PivotTableAction.Refresh => RefreshAsync(commands, sessionId, pivotTableName),
                PivotTableAction.ListFields => ListFieldsAsync(commands, sessionId, pivotTableName),
                PivotTableAction.AddRowField => AddRowFieldAsync(commands, sessionId, pivotTableName, fieldName, position),
                PivotTableAction.AddColumnField => AddColumnFieldAsync(commands, sessionId, pivotTableName, fieldName, position),
                PivotTableAction.AddValueField => AddValueFieldAsync(commands, sessionId, pivotTableName, fieldName, aggregationFunction, customName),
                PivotTableAction.AddFilterField => AddFilterFieldAsync(commands, sessionId, pivotTableName, fieldName),
                PivotTableAction.RemoveField => RemoveFieldAsync(commands, sessionId, pivotTableName, fieldName),
                PivotTableAction.SetFieldFunction => SetFieldFunctionAsync(commands, sessionId, pivotTableName, fieldName, aggregationFunction),
                PivotTableAction.SetFieldName => SetFieldNameAsync(commands, sessionId, pivotTableName, fieldName, customName),
                PivotTableAction.SetFieldFormat => SetFieldFormatAsync(commands, sessionId, pivotTableName, fieldName, numberFormat),
                PivotTableAction.GetData => GetDataAsync(commands, sessionId, pivotTableName),
                PivotTableAction.SetFieldFilter => SetFieldFilterAsync(commands, sessionId, pivotTableName, fieldName, filterValues),
                PivotTableAction.SortField => SortFieldAsync(commands, sessionId, pivotTableName, fieldName, sortDirection),
                _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"{action.ToActionString()} failed for '{excelPath}': {ex.Message}",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ListAsync(
        PivotTableCommands commands,
        string sessionId)
    {
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.List(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTables,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string GetAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "get-info");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Read(batch, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTable,
            result.Fields,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string CreateFromRangeAsync(
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.CreateFromRange(batch, sheetName!, range!,
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

    private static string CreateFromTableAsync(
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.CreateFromTable(batch, tableName!,
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

    private static string CreateFromDataModelAsync(
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.CreateFromDataModel(batch, dataModelTableName!,
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

    private static string DeleteAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "delete");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Delete(batch, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string RefreshAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "refresh");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Refresh(batch, pivotTableName!));

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

    private static string ListFieldsAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "list-fields");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.ListFields(batch, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Fields,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string AddRowFieldAsync(
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.AddRowField(batch, pivotTableName!, fieldName!, position));

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

    private static string AddColumnFieldAsync(
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.AddColumnField(batch, pivotTableName!, fieldName!, position));

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

    private static string AddValueFieldAsync(
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
            throw new ArgumentException(
                $"Invalid aggregation function '{aggregationFunction}'. Valid values: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP", nameof(aggregationFunction));
        }

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.AddValueField(batch, pivotTableName!, fieldName!, function, customName));

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

    private static string AddFilterFieldAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "add-filter-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "add-filter-field");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.AddFilterField(batch, pivotTableName!, fieldName!));

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

    private static string RemoveFieldAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "remove-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "remove-field");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.RemoveField(batch, pivotTableName!, fieldName!));

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

    private static string SetFieldFunctionAsync(
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
            throw new ArgumentException(
                $"Invalid aggregation function '{aggregationFunction}'. Valid values: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP", nameof(aggregationFunction));
        }

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetFieldFunction(batch, pivotTableName!, fieldName!, function));

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

    private static string SetFieldNameAsync(
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetFieldName(batch, pivotTableName!, fieldName!, customName!));

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

    private static string SetFieldFormatAsync(
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetFieldFormat(batch, pivotTableName!, fieldName!, numberFormat!));

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

    private static string GetDataAsync(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "get-data");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.GetData(batch, pivotTableName!));

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

    private static string SetFieldFilterAsync(
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
            throw new ArgumentException($"Invalid filterValues JSON: {ex.Message}. Expected format: '[\"value1\",\"value2\"]'", nameof(filterValues));
        }

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetFieldFilter(batch, pivotTableName!, fieldName!, values));

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

    private static string SortFieldAsync(
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
            throw new ArgumentException(
                $"Invalid sort direction '{sortDirection}'. Valid values: Ascending, Descending", nameof(sortDirection));
        }

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SortField(batch, pivotTableName!, fieldName!, direction));

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

