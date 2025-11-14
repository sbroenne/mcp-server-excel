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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Found {result.PivotTables.Count} PivotTable(s) - ready for analysis or field management"
                : "Failed to list PivotTables - some PivotTables may have properties that cannot be read",
            suggestedNextActions = result.Success
                ? result.PivotTables.Count == 0
                    ? new[] { "Create PivotTable with create-from-range, create-from-table, or create-from-datamodel", "Load data into workbook first", "Check if PivotTables exist in different sheets" }
                    : [$"Use get to view {result.PivotTables[0].Name} details", "Use list-fields to see available fields", "Use refresh to update data from source"]
                : ["Try opening and saving the workbook in Excel first", "Check if PivotTables are disconnected from data sources", "Some PivotTables may use Data Model or external sources"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"{result.PivotTable.Name}: {result.PivotTable.RowFieldCount} rows, {result.PivotTable.ColumnFieldCount} cols, {result.PivotTable.ValueFieldCount} values, {result.Fields.Count} total fields"
                : $"PivotTable '{pivotTableName}' not found or inaccessible",
            suggestedNextActions = result.Success
                ? new[] { "Use add-row-field, add-column-field, or add-value-field to configure", "Use refresh to update with latest source data", "Use get-data to extract PivotTable results" }
                : ["Use list to see all available PivotTables", "Check PivotTable name spelling", "Verify PivotTable exists in workbook"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"PivotTable '{result.PivotTableName}' created at {result.SheetName}!{result.Range} from {result.SourceRowCount} rows, {result.AvailableFields.Count} fields available"
                : "Failed to create PivotTable from range - check source range contains headers and data",
            suggestedNextActions = result.Success
                ? new[] { $"Use add-row-field with fields: {string.Join(", ", result.AvailableFields.Take(3))}", "Use add-value-field to summarize data", "Use add-filter-field to enable interactive filtering" }
                : ["Verify source range contains header row", "Check source range has at least 2 rows (headers + data)", "Ensure range address is valid (e.g., A1:D100)"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"PivotTable '{result.PivotTableName}' created at {result.SheetName}!{result.Range} from Excel Table '{tableName}' ({result.SourceRowCount} rows)"
                : $"Failed to create PivotTable from table '{tableName}' - verify table exists",
            suggestedNextActions = result.Success
                ? new[] { $"Use add-row-field with fields: {string.Join(", ", result.AvailableFields.Take(3))}", "Use add-value-field to calculate aggregations", "Use refresh when source table data changes" }
                : ["Use table-list to see available Excel Tables", "Check table name spelling", "Ensure table contains data rows"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"PivotTable '{result.PivotTableName}' created from Data Model table '{dataModelTableName}' ({result.SourceRowCount} rows) with DAX measures support"
                : $"Failed to create PivotTable from Data Model table '{dataModelTableName}'",
            suggestedNextActions = result.Success
                ? new[] { $"Use add-row-field with fields: {string.Join(", ", result.AvailableFields.Take(3))}", "Use add-value-field to add DAX measures", "Leverage relationships between Data Model tables" }
                : ["Use dm-list-tables to see available Data Model tables", "Verify Data Model contains data (use dm-refresh)", "Check table name spelling"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"PivotTable '{pivotTableName}' deleted - analysis removed from workbook"
                : $"Failed to delete PivotTable '{pivotTableName}'",
            suggestedNextActions = result.Success
                ? new[] { "Use list to verify deletion", "Create new PivotTable if needed", "Source data remains unchanged" }
                : ["Verify PivotTable name is correct", "Check PivotTable isn't protected", "Use list to see available PivotTables"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? result.StructureChanged
                    ? $"Refreshed '{result.PivotTableName}': {result.SourceRecordCount} rows (was {result.PreviousRecordCount}), structure changed - {result.NewFields.Count} new, {result.RemovedFields.Count} removed"
                    : $"Refreshed '{result.PivotTableName}': {result.SourceRecordCount} rows (was {result.PreviousRecordCount})"
                : $"Failed to refresh PivotTable '{pivotTableName}' - check source data connectivity",
            suggestedNextActions = result.Success
                ? result.StructureChanged
                    ? new[] { $"Review new fields: {string.Join(", ", result.NewFields.Take(3))}", $"Update field configuration if needed", "Verify removed fields didn't break analysis" }
                    : ["Use get-data to extract refreshed results", "Verify data changes as expected", "Continue with field configuration or analysis"]
                : ["Check source data connection is valid", "Verify source table/range still exists", "Ensure source data is accessible"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"{result.Fields.Count} field(s) in '{pivotTableName}' - ready for field configuration"
                : $"Failed to list fields for PivotTable '{pivotTableName}'",
            suggestedNextActions = result.Success
                ? result.Fields.Count == 0
                    ? new[] { "Use add-row-field to add grouping dimensions", "Use add-value-field to add calculations", "Check source data contains fields" }
                    : [$"Use add-row-field with available fields", "Use add-value-field for aggregations", "Use set-field-function to change calculations"]
                : ["Verify PivotTable name is correct", "Use list to see available PivotTables", "Refresh PivotTable if structure changed"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Added '{result.FieldName}' as row field at position {result.Position} - grouping enabled"
                : $"Failed to add row field '{fieldName}' to PivotTable '{pivotTableName}'",
            suggestedNextActions = result.Success
                ? new[] { "Use add-column-field or add-value-field to complete analysis", "Use add-filter-field for interactive filtering", "Use refresh if source data changed" }
                : ["Use list-fields to see available field names", "Check field name spelling", "Verify PivotTable exists (use list)"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Added '{result.FieldName}' as column field at position {result.Position} - cross-tabulation enabled"
                : $"Failed to add column field '{fieldName}' to PivotTable '{pivotTableName}'",
            suggestedNextActions = result.Success
                ? new[] { "Use add-value-field to calculate aggregations across columns", "Use add-row-field for grouping", "Use get-data to extract cross-tab results" }
                : ["Use list-fields to see available field names", "Check field name spelling", "Verify PivotTable exists (use list)"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Added value field '{result.CustomName}' ({result.Function}) at position {result.Position} - aggregation configured"
                : $"Failed to add value field '{fieldName}' to PivotTable '{pivotTableName}'",
            suggestedNextActions = result.Success
                ? new[] { "Use set-field-format to apply number formatting", "Use set-field-name to customize display name", "Use get-data to extract aggregated results" }
                : ["Use list-fields to see available field names", "Check field name spelling", "Verify field is numeric for Sum/Average functions"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Added filter field '{result.FieldName}' - interactive filtering enabled with {result.AvailableValues.Count} possible values"
                : $"Failed to add filter field '{fieldName}' to PivotTable '{pivotTableName}'",
            suggestedNextActions = result.Success
                ? new[] { $"Use set-field-filter with values from: {string.Join(", ", result.AvailableValues.Take(5))}", "Use get-data to view filtered results", "Add more filter fields for multi-dimensional filtering" }
                : ["Use list-fields to see available field names", "Check field name spelling", "Verify PivotTable exists (use list)"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Removed field '{result.FieldName}' from {result.Area} area - PivotTable reconfigured"
                : $"Failed to remove field '{fieldName}' from PivotTable '{pivotTableName}'",
            suggestedNextActions = result.Success
                ? new[] { "Use list-fields to verify field removal", "Use get to view updated PivotTable structure", "Add different fields to reconfigure analysis" }
                : ["Use list-fields to see current field configuration", "Check field name spelling", "Verify field is currently in PivotTable"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Changed '{result.FieldName}' aggregation to {result.Function} - calculation method updated"
                : $"Failed to set aggregation function for field '{fieldName}' in PivotTable '{pivotTableName}'",
            suggestedNextActions = result.Success
                ? new[] { "Use get-data to view updated calculations", "Use set-field-format to apply appropriate number formatting", "Use refresh if source data changed" }
                : ["Verify field is a value field (not row/column/filter)", "Check field name spelling", "Ensure function is valid for data type"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Renamed field '{result.FieldName}' to '{result.CustomName}' - display name customized"
                : $"Failed to rename field '{fieldName}' in PivotTable '{pivotTableName}'",
            suggestedNextActions = result.Success
                ? new[] { "Use get to view updated PivotTable with new field name", "Use set-field-format to apply number formatting", "Continue configuring field properties" }
                : ["Check field name spelling", "Verify field exists in PivotTable (use list-fields)", "Ensure custom name is valid"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Applied number format '{result.NumberFormat}' to field '{result.FieldName}' - values formatted"
                : $"Failed to set number format for field '{fieldName}' in PivotTable '{pivotTableName}'",
            suggestedNextActions = result.Success
                ? new[] { "Use get-data to view formatted values", "Use set-field-name to customize field display name", "Continue configuring analysis" }
                : ["Verify field name spelling", "Check number format code syntax", "Ensure field is a value field"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Extracted {result.DataRowCount}×{result.DataColumnCount} data grid from PivotTable '{result.PivotTableName}' - analysis results ready"
                : $"Failed to extract data from PivotTable '{pivotTableName}'",
            suggestedNextActions = result.Success
                ? new[] { "Process extracted values for further analysis", "Export data to external system", "Use refresh to update with latest source data" }
                : ["Verify PivotTable exists (use list)", "Check PivotTable has data (use get)", "Ensure fields are configured properly"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Filtered field '{result.FieldName}' to {result.SelectedItems.Count} of {result.AvailableItems.Count} values - showing {result.VisibleRowCount}/{result.TotalRowCount} rows"
                : $"Failed to filter field '{fieldName}' in PivotTable '{pivotTableName}'",
            suggestedNextActions = result.Success
                ? new[] { "Use get-data to extract filtered results", $"Selected values: {string.Join(", ", result.SelectedItems.Take(5))}", "Use refresh to update with latest source data" }
                : ["Verify field name spelling", "Check filter values are valid for field", "Ensure field exists in PivotTable (use list-fields)"]
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
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Sorted field '{result.FieldName}' {direction.ToString().ToLowerInvariant()} - data reordered"
                : $"Failed to sort field '{fieldName}' in PivotTable '{pivotTableName}'",
            suggestedNextActions = result.Success
                ? new[] { "Use get-data to extract sorted results", "Use add-value-field if need aggregation", "Continue configuring analysis" }
                : ["Verify field name spelling", "Ensure field exists in PivotTable (use list-fields)", "Check field is in row or column area"]
        }, JsonOptions);
    }
}
