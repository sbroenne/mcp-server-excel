using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Table column, filter, and sort operations.
/// </summary>
[McpServerToolType]
public static partial class TableColumnTool
{
    /// <summary>
    /// Table column/filter/sort operations.
    /// Related: excel_table (lifecycle/data)
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="path">File path</param>
    /// <param name="sid">Session ID</param>
    /// <param name="tn">Table name</param>
    /// <param name="col">Column name</param>
    /// <param name="nn">New column name</param>
    /// <param name="crit">Filter criteria or region</param>
    /// <param name="fv">Filter values or sort columns JSON</param>
    /// <param name="fmt">Format code (US locale)</param>
    /// <param name="pos">Column position (1-based)</param>
    /// <param name="asc">Ascending order</param>
    [McpServerTool(Name = "excel_table_column", Title = "Excel Table Column Operations")]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string TableColumn(
        TableColumnAction action,
        string path,
        string sid,
        [DefaultValue(null)] string? tn,
        [DefaultValue(null)] string? col,
        [DefaultValue(null)] string? nn,
        [DefaultValue(null)] string? crit,
        [DefaultValue(null)] string? fv,
        [DefaultValue(null)] string? fmt,
        [DefaultValue(null)] string? pos,
        [DefaultValue(true)] bool asc)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_table_column",
            action.ToActionString(),
            path,
            () =>
            {
                var tableCommands = new TableCommands();

                return action switch
                {
                    TableColumnAction.ApplyFilter => ApplyFilter(tableCommands, sid, tn, col, crit),
                    TableColumnAction.ApplyFilterValues => ApplyFilterValues(tableCommands, sid, tn, col, fv),
                    TableColumnAction.ClearFilters => ClearFilters(tableCommands, sid, tn),
                    TableColumnAction.GetFilters => GetFilters(tableCommands, sid, tn),
                    TableColumnAction.AddColumn => AddColumn(tableCommands, sid, tn, col, pos),
                    TableColumnAction.RemoveColumn => RemoveColumn(tableCommands, sid, tn, col),
                    TableColumnAction.RenameColumn => RenameColumn(tableCommands, sid, tn, col, nn),
                    TableColumnAction.GetStructuredReference => GetStructuredReference(tableCommands, sid, tn, crit, col),
                    TableColumnAction.Sort => SortTable(tableCommands, sid, tn, col, asc),
                    TableColumnAction.SortMulti => SortTableMulti(tableCommands, sid, tn, fv),
                    TableColumnAction.GetColumnNumberFormat => GetColumnNumberFormat(tableCommands, sid, tn, col),
                    TableColumnAction.SetColumnNumberFormat => SetColumnNumberFormat(tableCommands, sid, tn, col, fmt),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ApplyFilter(TableCommands commands, string sid, string? tn, string? col, string? crit)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "apply-filter");
        if (string.IsNullOrWhiteSpace(col)) ExcelToolsBase.ThrowMissingParameter(nameof(col), "apply-filter");
        if (string.IsNullOrWhiteSpace(crit)) ExcelToolsBase.ThrowMissingParameter(nameof(crit), "apply-filter");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.ApplyFilter(batch, tn!, col!, crit!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Filter applied successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ApplyFilterValues(TableCommands commands, string sid, string? tn, string? col, string? fvJson)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "apply-filter-values");
        if (string.IsNullOrWhiteSpace(col)) ExcelToolsBase.ThrowMissingParameter(nameof(col), "apply-filter-values");
        if (string.IsNullOrWhiteSpace(fvJson)) ExcelToolsBase.ThrowMissingParameter(nameof(fvJson), "apply-filter-values");

        List<string> filterValues;
        try
        {
            filterValues = JsonSerializer.Deserialize<List<string>>(fvJson!) ?? [];
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid JSON array for filterValues: {ex.Message}", nameof(fvJson));
        }

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.ApplyFilter(batch, tn!, col!, filterValues);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Filter applied successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ClearFilters(TableCommands commands, string sid, string? tn)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "clear-filters");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.ClearFilters(batch, tn!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Filters cleared successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string GetFilters(TableCommands commands, string sid, string? tn)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "get-filters");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetFilters(batch, tn!));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string AddColumn(TableCommands commands, string sid, string? tn, string? col, string? pos)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "add-column");
        if (string.IsNullOrWhiteSpace(col)) ExcelToolsBase.ThrowMissingParameter(nameof(col), "add-column");

        int? position = null;
        if (!string.IsNullOrWhiteSpace(pos))
        {
            if (int.TryParse(pos, out int posVal))
            {
                position = posVal;
            }
            else
            {
                throw new ArgumentException($"Invalid position value: '{pos}'. Must be a number.", nameof(pos));
            }
        }

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.AddColumn(batch, tn!, col!, position);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Column added successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string RemoveColumn(TableCommands commands, string sid, string? tn, string? col)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "remove-column");
        if (string.IsNullOrWhiteSpace(col)) ExcelToolsBase.ThrowMissingParameter(nameof(col), "remove-column");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.RemoveColumn(batch, tn!, col!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Column removed successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string RenameColumn(TableCommands commands, string sid, string? tn, string? col, string? nn)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "rename-column");
        if (string.IsNullOrWhiteSpace(col)) ExcelToolsBase.ThrowMissingParameter(nameof(col), "rename-column");
        if (string.IsNullOrWhiteSpace(nn)) ExcelToolsBase.ThrowMissingParameter(nameof(nn), "rename-column");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.RenameColumn(batch, tn!, col!, nn!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Column renamed successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string GetStructuredReference(TableCommands commands, string sid, string? tn, string? region, string? col)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "get-structured-reference");

        var tableRegion = TableRegion.Data;
        if (!string.IsNullOrWhiteSpace(region) && !Enum.TryParse(region, true, out tableRegion))
        {
            throw new ArgumentException($"Invalid region '{region}'. Valid values: All, Data, Headers, Totals, ThisRow", nameof(region));
        }

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetStructuredReference(batch, tn!, tableRegion, col));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SortTable(TableCommands commands, string sid, string? tn, string? col, bool asc)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "sort");
        if (string.IsNullOrWhiteSpace(col)) ExcelToolsBase.ThrowMissingParameter(nameof(col), "sort");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.Sort(batch, tn!, col!, asc);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table sorted successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string SortTableMulti(TableCommands commands, string sid, string? tn, string? sortJson)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "sort-multi");
        if (string.IsNullOrWhiteSpace(sortJson)) ExcelToolsBase.ThrowMissingParameter(nameof(sortJson), "sort-multi");

        List<TableSortColumn>? sortColumns;
        try
        {
            sortColumns = JsonSerializer.Deserialize<List<TableSortColumn>>(sortJson!);
            if (sortColumns == null || sortColumns.Count == 0)
            {
                throw new ArgumentException("sortColumns JSON must be a non-empty array", nameof(sortJson));
            }
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid sortColumns JSON: {ex.Message}", nameof(sortJson));
        }

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.Sort(batch, tn!, sortColumns);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table sorted successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string GetColumnNumberFormat(TableCommands commands, string sid, string? tn, string? col)
    {
        if (string.IsNullOrEmpty(tn))
            ExcelToolsBase.ThrowMissingParameter("tn", "get-column-number-format");
        if (string.IsNullOrEmpty(col))
            ExcelToolsBase.ThrowMissingParameter("col", "get-column-number-format");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetColumnNumberFormat(batch, tn!, col!));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetColumnNumberFormat(TableCommands commands, string sid, string? tn, string? col, string? fmt)
    {
        if (string.IsNullOrEmpty(tn))
            ExcelToolsBase.ThrowMissingParameter("tn", "set-column-number-format");
        if (string.IsNullOrEmpty(col))
            ExcelToolsBase.ThrowMissingParameter("col", "set-column-number-format");
        if (string.IsNullOrEmpty(fmt))
            ExcelToolsBase.ThrowMissingParameter("fmt", "set-column-number-format");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.SetColumnNumberFormat(batch, tn!, col!, fmt!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Column number format set successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }
}
