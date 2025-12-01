using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel PivotTable operations
/// </summary>
[McpServerToolType]
public static partial class ExcelPivotTableTool
{
    private static readonly JsonSerializerOptions JsonOptions = ExcelToolsBase.JsonOptions;

    /// <summary>
    /// Excel PivotTable operations - interactive data analysis and summarization.
    /// OLAP add-value-field supports TWO modes: 1) Pre-existing measure: fieldName='Total Sales' or '[Measures].[Total Sales]' (adds existing measure) 2) Auto-create measure: fieldName='Sales' (column name, creates DAX measure with aggregationFunction).
    /// TIMEOUT SAFEGUARD: create-from-datamodel auto-timeouts after 5 minutes to prevent hung OLAP queries.
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="excelPath">Path to Excel file (.xlsx or .xlsm)</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action</param>
    /// <param name="pivotTableName">PivotTable name</param>
    /// <param name="sheetName">Source sheet name (for create-from-range)</param>
    /// <param name="range">Source range (for create-from-range)</param>
    /// <param name="tableName">Excel Table name (for create-from-table)</param>
    /// <param name="dataModelTableName">Data Model table name (for create-from-datamodel)</param>
    /// <param name="destinationSheet">Destination sheet for new PivotTable</param>
    /// <param name="destinationCell">Destination cell for new PivotTable</param>
    /// <param name="fieldName">Field name for field operations (for OLAP add-value-field: use measure name like 'Total Sales' for existing measure OR column name like 'Sales' to create new measure)</param>
    /// <param name="aggregationFunction">Aggregation function: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP</param>
    /// <param name="customName">Custom display name for field</param>
    /// <param name="numberFormat">Number format code (e.g., '#,##0.00', '0.00%', 'm/d/yyyy')</param>
    /// <param name="position">Position for field (1-based, optional)</param>
    /// <param name="filterValues">JSON array of filter values (e.g., '["value1","value2"]')</param>
    /// <param name="sortDirection">Sort direction: Ascending, Descending</param>
    /// <param name="dateGroupingInterval">Date grouping interval: Days, Months, Quarters, Years</param>
    /// <param name="numericGroupingStart">Numeric grouping start value (null = use field minimum)</param>
    /// <param name="numericGroupingEnd">Numeric grouping end value (null = use field maximum)</param>
    /// <param name="numericGroupingInterval">Numeric grouping interval size (e.g., 100 for 0-100, 100-200, ...)</param>
    /// <param name="formula">Formula for calculated field (e.g., '=Revenue-Cost', '=Profit/Revenue')</param>
    /// <param name="layout">Layout form: 0=Compact (all fields in one column), 1=Tabular (separate columns, subtotals bottom), 2=Outline (separate columns, subtotals top)</param>
    /// <param name="subtotalsVisible">Show/hide subtotals for field: true=show automatic subtotals, false=hide</param>
    /// <param name="showRowGrandTotals">Show/hide row grand totals: true=show bottom summary row, false=hide</param>
    /// <param name="showColumnGrandTotals">Show/hide column grand totals: true=show right summary column, false=hide</param>
    [McpServerTool(Name = "excel_pivottable")]
    [McpMeta("category", "analysis")]
    public static partial string ExcelPivotTable(
        PivotTableAction action,
        string excelPath,
        string sessionId,
        string? pivotTableName,
        string? sheetName,
        string? range,
        string? tableName,
        string? dataModelTableName,
        string? destinationSheet,
        string? destinationCell,
        string? fieldName,
        string? aggregationFunction,
        string? customName,
        string? numberFormat,
        int? position,
        string? filterValues,
        string? sortDirection,
        string? dateGroupingInterval,
        double? numericGroupingStart,
        double? numericGroupingEnd,
        double? numericGroupingInterval,
        string? formula,
        int? layout,
        bool? subtotalsVisible,
        bool? showRowGrandTotals,
        bool? showColumnGrandTotals)
    {
        _ = excelPath; // retained for schema compatibility (operations require open session)

        return ExcelToolsBase.ExecuteToolAction(
            "excel_pivottable",
            action.ToActionString(),
            () =>
            {
                var commands = new PivotTableCommands();

                return action switch
                {
                    PivotTableAction.List => List(commands, sessionId),
                    PivotTableAction.Read => Read(commands, sessionId, pivotTableName),
                    PivotTableAction.CreateFromRange => CreateFromRange(commands, sessionId, sheetName, range, destinationSheet, destinationCell, pivotTableName),
                    PivotTableAction.CreateFromTable => CreateFromTable(commands, sessionId, tableName, destinationSheet, destinationCell, pivotTableName),
                    PivotTableAction.CreateFromDataModel => CreateFromDataModel(commands, sessionId, dataModelTableName, destinationSheet, destinationCell, pivotTableName),
                    PivotTableAction.Delete => Delete(commands, sessionId, pivotTableName),
                    PivotTableAction.Refresh => Refresh(commands, sessionId, pivotTableName),
                    PivotTableAction.ListFields => ListFields(commands, sessionId, pivotTableName),
                    PivotTableAction.AddRowField => AddRowField(commands, sessionId, pivotTableName, fieldName, position),
                    PivotTableAction.AddColumnField => AddColumnField(commands, sessionId, pivotTableName, fieldName, position),
                    PivotTableAction.AddValueField => AddValueField(commands, sessionId, pivotTableName, fieldName, aggregationFunction, customName),
                    PivotTableAction.AddFilterField => AddFilterField(commands, sessionId, pivotTableName, fieldName),
                    PivotTableAction.RemoveField => RemoveField(commands, sessionId, pivotTableName, fieldName),
                    PivotTableAction.SetFieldFunction => SetFieldFunction(commands, sessionId, pivotTableName, fieldName, aggregationFunction),
                    PivotTableAction.SetFieldName => SetFieldName(commands, sessionId, pivotTableName, fieldName, customName),
                    PivotTableAction.SetFieldFormat => SetFieldFormat(commands, sessionId, pivotTableName, fieldName, numberFormat),
                    PivotTableAction.GetData => GetData(commands, sessionId, pivotTableName),
                    PivotTableAction.SetFieldFilter => SetFieldFilter(commands, sessionId, pivotTableName, fieldName, filterValues),
                    PivotTableAction.SortField => SortField(commands, sessionId, pivotTableName, fieldName, sortDirection),
                    PivotTableAction.GroupByDate => GroupByDate(commands, sessionId, pivotTableName, fieldName, dateGroupingInterval),
                    PivotTableAction.GroupByNumeric => GroupByNumeric(commands, sessionId, pivotTableName, fieldName, numericGroupingStart, numericGroupingEnd, numericGroupingInterval),
                    PivotTableAction.CreateCalculatedField => CreateCalculatedField(commands, sessionId, pivotTableName, fieldName, formula),
                    PivotTableAction.SetLayout => SetLayout(commands, sessionId, pivotTableName, layout),
                    PivotTableAction.SetSubtotals => SetSubtotals(commands, sessionId, pivotTableName, fieldName, subtotalsVisible),
                    PivotTableAction.SetGrandTotals => SetGrandTotals(commands, sessionId, pivotTableName, showRowGrandTotals, showColumnGrandTotals),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string List(
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

    private static string Read(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "read");

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

    private static string CreateFromRange(
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

    private static string CreateFromTable(
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

    private static string CreateFromDataModel(
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

        PivotTableCreateResult result;

        try
        {
            result = ExcelToolsBase.WithSession(sessionId,
                batch => commands.CreateFromDataModel(batch, dataModelTableName!,
                    destinationSheet!, destinationCell!, pivotTableName!));
        }
        catch (TimeoutException ex)
        {
            result = new PivotTableCreateResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                PivotTableName = pivotTableName!,
                SheetName = destinationSheet!,
                Range = string.Empty,
                SourceData = dataModelTableName!,
                SourceRowCount = 0,
                AvailableFields = []
            };
        }

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
            isError = !result.Success
        }, JsonOptions);
    }

    private static string Delete(
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

    private static string Refresh(
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

    private static string ListFields(
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

    private static string AddRowField(
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

    private static string AddColumnField(
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

    private static string AddValueField(
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

    private static string AddFilterField(
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

    private static string RemoveField(
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

    private static string SetFieldFunction(
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

    private static string SetFieldName(
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

    private static string SetFieldFormat(
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

    private static string GetData(
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

    private static string SetFieldFilter(
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

    private static string SortField(
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

    private static string GroupByDate(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName,
        string? dateGroupingInterval)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for group-by-date action", nameof(pivotTableName));

        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for group-by-date action", nameof(fieldName));

        if (string.IsNullOrEmpty(dateGroupingInterval))
            throw new ArgumentException("dateGroupingInterval is required for group-by-date action", nameof(dateGroupingInterval));

        // Parse date grouping interval
        if (!Enum.TryParse<DateGroupingInterval>(dateGroupingInterval, true, out var interval))
        {
            throw new ArgumentException(
                $"Invalid date grouping interval '{dateGroupingInterval}'. Valid values: Days, Months, Quarters, Years",
                nameof(dateGroupingInterval));
        }

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.GroupByDate(batch, pivotTableName!, fieldName!, interval));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.ErrorMessage,
            result.WorkflowHint
        }, JsonOptions);
    }

    private static string GroupByNumeric(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName,
        double? numericGroupingStart,
        double? numericGroupingEnd,
        double? numericGroupingInterval)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for group-by-numeric action", nameof(pivotTableName));

        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for group-by-numeric action", nameof(fieldName));

        if (!numericGroupingInterval.HasValue || numericGroupingInterval.Value <= 0)
            throw new ArgumentException("numericGroupingInterval is required and must be > 0 for group-by-numeric action", nameof(numericGroupingInterval));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.GroupByNumeric(batch, pivotTableName!, fieldName!, numericGroupingStart, numericGroupingEnd, numericGroupingInterval.Value));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.ErrorMessage,
            result.WorkflowHint
        }, JsonOptions);
    }

    private static string CreateCalculatedField(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, string? formula)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for create-calculated-field action", nameof(pivotTableName));

        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for create-calculated-field action", nameof(fieldName));

        if (string.IsNullOrEmpty(formula))
            throw new ArgumentException("formula is required for create-calculated-field action", nameof(formula));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.CreateCalculatedField(batch, pivotTableName!, fieldName!, formula!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Formula,
            result.ErrorMessage,
            result.WorkflowHint
        }, JsonOptions);
    }

    private static string SetLayout(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        int? layout)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for set-layout action", nameof(pivotTableName));

        if (!layout.HasValue)
            throw new ArgumentException("layout is required for set-layout action (0=Compact, 1=Tabular, 2=Outline)", nameof(layout));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetLayout(batch, pivotTableName!, layout.Value));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string SetSubtotals(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName,
        bool? subtotalsVisible)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for set-subtotals action", nameof(pivotTableName));

        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for set-subtotals action", nameof(fieldName));

        if (!subtotalsVisible.HasValue)
            throw new ArgumentException("subtotalsVisible is required for set-subtotals action", nameof(subtotalsVisible));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetSubtotals(batch, pivotTableName!, fieldName!, subtotalsVisible.Value));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.ErrorMessage,
            result.WorkflowHint
        }, JsonOptions);
    }

    private static string SetGrandTotals(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        bool? showRowGrandTotals,
        bool? showColumnGrandTotals)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for set-grand-totals action", nameof(pivotTableName));

        if (!showRowGrandTotals.HasValue)
            throw new ArgumentException("showRowGrandTotals is required for set-grand-totals action", nameof(showRowGrandTotals));

        if (!showColumnGrandTotals.HasValue)
            throw new ArgumentException("showColumnGrandTotals is required for set-grand-totals action", nameof(showColumnGrandTotals));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetGrandTotals(batch, pivotTableName!, showRowGrandTotals.Value, showColumnGrandTotals.Value));

        return JsonSerializer.Serialize(result, JsonOptions);
    }
}
