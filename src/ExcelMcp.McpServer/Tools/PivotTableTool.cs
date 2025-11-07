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
        [Description("Excel Table name (ListObject) for create-from-table action")]
        string? tableName = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Data Model table name for create-from-datamodel action")]
        string? dataModelTableName = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Custom display name for value fields (used with add-value-field and set-field-name actions)")]
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
            var commands = new PivotTableCommands();

            // Switch directly on enum - inline all operations for clarity
            switch (action)
            {
                case PivotTableAction.List:
                    {
                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, false,
                            commands.ListAsync);

                        return JsonSerializer.Serialize(new
                        {
                            result.Success,
                            result.PivotTables,
                            result.ErrorMessage,
                            workflowHint = result.Success
                                ? $"Found {result.PivotTables.Count} PivotTable(s) - ready for analysis or field management"
                                : "Failed to list PivotTables - check workbook contains PivotTables",
                            suggestedNextActions = result.Success
                                ? result.PivotTables.Count == 0
                                    ? new[] { "Create PivotTable with create-from-range, create-from-table, or create-from-datamodel", "Load data into workbook first", "Check if PivotTables exist in different sheets" }
                                    : new[] { $"Use get to view {result.PivotTables[0].Name} details", "Use list-fields to see available fields", "Use refresh to update data from source" }
                                : new[] { "Verify file path is correct", "Check workbook opens in Excel", "Ensure workbook permissions allow access" }
                        }, JsonOptions);
                    }

                case PivotTableAction.Get:
                    {
                        if (string.IsNullOrWhiteSpace(pivotTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "get-info");

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, false,
                            async (batch) => await commands.GetAsync(batch, pivotTableName!));

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
                                : new[] { "Use list to see all available PivotTables", "Check PivotTable name spelling", "Verify PivotTable exists in workbook" }
                        }, JsonOptions);
                    }

                case PivotTableAction.CreateFromRange:
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

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.CreateFromRangeAsync(batch, sheetName!, range!,
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
                                : new[] { "Verify source range contains header row", "Check source range has at least 2 rows (headers + data)", "Ensure range address is valid (e.g., A1:D100)" }
                        }, JsonOptions);
                    }

                case PivotTableAction.CreateFromTable:
                    {
                        if (string.IsNullOrWhiteSpace(tableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "create-from-table");
                        if (string.IsNullOrWhiteSpace(destinationSheet))
                            ExcelToolsBase.ThrowMissingParameter(nameof(destinationSheet), "create-from-table");
                        if (string.IsNullOrWhiteSpace(destinationCell))
                            ExcelToolsBase.ThrowMissingParameter(nameof(destinationCell), "create-from-table");
                        if (string.IsNullOrWhiteSpace(pivotTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "create-from-table");

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.CreateFromTableAsync(batch, tableName!,
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
                                : new[] { "Use table-list to see available Excel Tables", "Check table name spelling", "Ensure table contains data rows" }
                        }, JsonOptions);
                    }

                case PivotTableAction.CreateFromDataModel:
                    {
                        if (string.IsNullOrWhiteSpace(dataModelTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(dataModelTableName), "create-from-datamodel");
                        if (string.IsNullOrWhiteSpace(destinationSheet))
                            ExcelToolsBase.ThrowMissingParameter(nameof(destinationSheet), "create-from-datamodel");
                        if (string.IsNullOrWhiteSpace(destinationCell))
                            ExcelToolsBase.ThrowMissingParameter(nameof(destinationCell), "create-from-datamodel");
                        if (string.IsNullOrWhiteSpace(pivotTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "create-from-datamodel");

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.CreateFromDataModelAsync(batch, dataModelTableName!,
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
                                : new[] { "Use dm-list-tables to see available Data Model tables", "Verify Data Model contains data (use dm-refresh)", "Check table name spelling" }
                        }, JsonOptions);
                    }

                case PivotTableAction.Delete:
                    {
                        if (string.IsNullOrWhiteSpace(pivotTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "delete");

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.DeleteAsync(batch, pivotTableName!));

                        return JsonSerializer.Serialize(new
                        {
                            result.Success,
                            result.ErrorMessage,
                            workflowHint = result.Success
                                ? $"PivotTable '{pivotTableName}' deleted - analysis removed from workbook"
                                : $"Failed to delete PivotTable '{pivotTableName}'",
                            suggestedNextActions = result.Success
                                ? new[] { "Use list to verify deletion", "Create new PivotTable if needed", "Source data remains unchanged" }
                                : new[] { "Verify PivotTable name is correct", "Check PivotTable isn't protected", "Use list to see available PivotTables" }
                        }, JsonOptions);
                    }

                case PivotTableAction.Refresh:
                    {
                        if (string.IsNullOrWhiteSpace(pivotTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "refresh");

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.RefreshAsync(batch, pivotTableName!));

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
                                    : new[] { "Use get-data to extract refreshed results", "Verify data changes as expected", "Continue with field configuration or analysis" }
                                : new[] { "Check source data connection is valid", "Verify source table/range still exists", "Ensure source data is accessible" }
                        }, JsonOptions);
                    }

                case PivotTableAction.ListFields:
                    {
                        if (string.IsNullOrWhiteSpace(pivotTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "list-fields");

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, false,
                            async (batch) => await commands.ListFieldsAsync(batch, pivotTableName!));

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
                                    : new[] { $"Use add-row-field with available fields", "Use add-value-field for aggregations", "Use set-field-function to change calculations" }
                                : new[] { "Verify PivotTable name is correct", "Use list to see available PivotTables", "Refresh PivotTable if structure changed" }
                        }, JsonOptions);
                    }

                case PivotTableAction.AddRowField:
                    {
                        if (string.IsNullOrWhiteSpace(pivotTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "add-row-field");
                        if (string.IsNullOrWhiteSpace(fieldName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "add-row-field");

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.AddRowFieldAsync(batch, pivotTableName!, fieldName!, position));

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
                                : new[] { "Use list-fields to see available field names", "Check field name spelling", "Verify PivotTable exists (use list)" }
                        }, JsonOptions);
                    }

                case PivotTableAction.AddColumnField:
                    {
                        if (string.IsNullOrWhiteSpace(pivotTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "add-column-field");
                        if (string.IsNullOrWhiteSpace(fieldName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "add-column-field");

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.AddColumnFieldAsync(batch, pivotTableName!, fieldName!, position));

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
                                : new[] { "Use list-fields to see available field names", "Check field name spelling", "Verify PivotTable exists (use list)" }
                        }, JsonOptions);
                    }

                case PivotTableAction.AddValueField:
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

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.AddValueFieldAsync(batch, pivotTableName!, fieldName!, function, customName));

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
                                : new[] { "Use list-fields to see available field names", "Check field name spelling", "Verify field is numeric for Sum/Average functions" }
                        }, JsonOptions);
                    }

                case PivotTableAction.AddFilterField:
                    {
                        if (string.IsNullOrWhiteSpace(pivotTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "add-filter-field");
                        if (string.IsNullOrWhiteSpace(fieldName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "add-filter-field");

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.AddFilterFieldAsync(batch, pivotTableName!, fieldName!));

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
                                : new[] { "Use list-fields to see available field names", "Check field name spelling", "Verify PivotTable exists (use list)" }
                        }, JsonOptions);
                    }

                case PivotTableAction.RemoveField:
                    {
                        if (string.IsNullOrWhiteSpace(pivotTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "remove-field");
                        if (string.IsNullOrWhiteSpace(fieldName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "remove-field");

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.RemoveFieldAsync(batch, pivotTableName!, fieldName!));

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
                                : new[] { "Use list-fields to see current field configuration", "Check field name spelling", "Verify field is currently in PivotTable" }
                        }, JsonOptions);
                    }

                case PivotTableAction.SetFieldFunction:
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

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.SetFieldFunctionAsync(batch, pivotTableName!, fieldName!, function));

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
                                : new[] { "Verify field is a value field (not row/column/filter)", "Check field name spelling", "Ensure function is valid for data type" }
                        }, JsonOptions);
                    }

                case PivotTableAction.SetFieldName:
                    {
                        if (string.IsNullOrWhiteSpace(pivotTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "set-field-name");
                        if (string.IsNullOrWhiteSpace(fieldName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "set-field-name");
                        if (string.IsNullOrWhiteSpace(customName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(customName), "set-field-name");

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.SetFieldNameAsync(batch, pivotTableName!, fieldName!, customName!));

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
                                : new[] { "Check field name spelling", "Verify field exists in PivotTable (use list-fields)", "Ensure custom name is valid" }
                        }, JsonOptions);
                    }

                case PivotTableAction.SetFieldFormat:
                    {
                        if (string.IsNullOrWhiteSpace(pivotTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "set-field-format");
                        if (string.IsNullOrWhiteSpace(fieldName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(fieldName), "set-field-format");
                        if (string.IsNullOrWhiteSpace(numberFormat))
                            ExcelToolsBase.ThrowMissingParameter(nameof(numberFormat), "set-field-format");

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.SetFieldFormatAsync(batch, pivotTableName!, fieldName!, numberFormat!));

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
                                : new[] { "Verify field name spelling", "Check number format code syntax", "Ensure field is a value field" }
                        }, JsonOptions);
                    }

                case PivotTableAction.GetData:
                    {
                        if (string.IsNullOrWhiteSpace(pivotTableName))
                            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "get-data");

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, false,
                            async (batch) => await commands.GetDataAsync(batch, pivotTableName!));

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
                                : new[] { "Verify PivotTable exists (use list)", "Check PivotTable has data (use get)", "Ensure fields are configured properly" }
                        }, JsonOptions);
                    }

                case PivotTableAction.SetFieldFilter:
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

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.SetFieldFilterAsync(batch, pivotTableName!, fieldName!, values));

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
                                : new[] { "Verify field name spelling", "Check filter values are valid for field", "Ensure field exists in PivotTable (use list-fields)" }
                        }, JsonOptions);
                    }

                case PivotTableAction.SortField:
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

                        var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, true,
                            async (batch) => await commands.SortFieldAsync(batch, pivotTableName!, fieldName!, direction));

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
                                : new[] { "Verify field name spelling", "Ensure field exists in PivotTable (use list-fields)", "Check field is in row or column area" }
                        }, JsonOptions);
                    }

                default:
                    throw new ModelContextProtocol.McpException($"Unknown action: {action} ({action.ToActionString()})");
            }
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
}
