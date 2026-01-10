using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Table lifecycle and data operations.
/// </summary>
[McpServerToolType]
public static partial class TableTool
{
    /// <summary>
    /// Excel Tables - lifecycle and data.
    ///
    /// DATA MODEL WORKFLOW: To analyze worksheet data with DAX/Power Pivot:
    /// 1. Create or identify an Excel Table on a worksheet
    /// 2. Use add-to-datamodel action to add the table to Power Pivot
    /// 3. Then use excel_datamodel to create DAX measures on it
    ///
    /// Related: excel_table_column (filter/sort/columns), excel_datamodel (DAX measures)
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="path">File path</param>
    /// <param name="sid">Session ID</param>
    /// <param name="tn">Table name</param>
    /// <param name="sn">Sheet name</param>
    /// <param name="rng">Range (A1:D10)</param>
    /// <param name="nn">New name</param>
    /// <param name="hdr">Has headers or show totals</param>
    /// <param name="style">Style name or total function or CSV data</param>
    /// <param name="vo">Visible rows only</param>
    [McpServerTool(Name = "excel_table", Title = "Excel Table Operations")]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string Table(
        TableAction action,
        string path,
        string sid,
        [DefaultValue(null)] string? tn,
        [DefaultValue(null)] string? sn,
        [DefaultValue(null)] string? rng,
        [DefaultValue(null)] string? nn,
        [DefaultValue(true)] bool hdr,
        [DefaultValue(null)] string? style,
        [DefaultValue(false)] bool vo)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_table",
            action.ToActionString(),
            path,
            () =>
            {
                var tableCommands = new TableCommands();

                return action switch
                {
                    TableAction.List => ListTables(tableCommands, sid),
                    TableAction.Create => CreateTable(tableCommands, sid, sn, tn, rng, hdr, style),
                    TableAction.Read => ReadTable(tableCommands, sid, tn),
                    TableAction.Rename => RenameTable(tableCommands, sid, tn, nn),
                    TableAction.Delete => DeleteTable(tableCommands, sid, tn),
                    TableAction.Resize => ResizeTable(tableCommands, sid, tn, rng),
                    TableAction.ToggleTotals => ToggleTotals(tableCommands, sid, tn, hdr),
                    TableAction.SetColumnTotal => SetColumnTotal(tableCommands, sid, tn, nn, style),
                    TableAction.Append => AppendRows(tableCommands, sid, tn, style),
                    TableAction.GetData => GetData(tableCommands, sid, tn, vo),
                    TableAction.SetStyle => SetTableStyle(tableCommands, sid, tn, style),
                    TableAction.AddToDataModel => AddToDataModel(tableCommands, sid, tn),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ListTables(TableCommands commands, string sid)
    {
        var result = ExcelToolsBase.WithSession(sid, batch => commands.List(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Tables,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateTable(TableCommands commands, string sid, string? sn, string? tn, string? rng, bool hdr, string? style)
    {
        if (string.IsNullOrWhiteSpace(sn)) ExcelToolsBase.ThrowMissingParameter(nameof(sn), "create");
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "create");
        if (string.IsNullOrWhiteSpace(rng)) ExcelToolsBase.ThrowMissingParameter(nameof(rng), "create");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.Create(batch, sn!, tn!, rng!, hdr, style);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table created successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ReadTable(TableCommands commands, string sid, string? tn)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "read");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.Read(batch, tn!));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string GetData(TableCommands commands, string sid, string? tn, bool vo)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "get-data");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetData(batch, tn!, vo));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string RenameTable(TableCommands commands, string sid, string? tn, string? nn)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "rename");
        if (string.IsNullOrWhiteSpace(nn)) ExcelToolsBase.ThrowMissingParameter(nameof(nn), "rename");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.Rename(batch, tn!, nn!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table renamed successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string DeleteTable(TableCommands commands, string sid, string? tn)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "delete");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.Delete(batch, tn!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table deleted successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ResizeTable(TableCommands commands, string sid, string? tn, string? rng)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "resize");
        if (string.IsNullOrWhiteSpace(rng)) ExcelToolsBase.ThrowMissingParameter(nameof(rng), "resize");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.Resize(batch, tn!, rng!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table resized successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ToggleTotals(TableCommands commands, string sid, string? tn, bool showTotals)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "toggle-totals");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.ToggleTotals(batch, tn!, showTotals);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Totals toggled successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string SetColumnTotal(TableCommands commands, string sid, string? tn, string? col, string? func)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "set-column-total");
        if (string.IsNullOrWhiteSpace(col)) ExcelToolsBase.ThrowMissingParameter(nameof(col), "set-column-total");
        if (string.IsNullOrWhiteSpace(func)) ExcelToolsBase.ThrowMissingParameter(nameof(func), "set-column-total");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.SetColumnTotal(batch, tn!, col!, func!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Column total set successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string AppendRows(TableCommands commands, string sid, string? tn, string? csv)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "append");
        if (string.IsNullOrWhiteSpace(csv)) ExcelToolsBase.ThrowMissingParameter(nameof(csv), "append");

        try
        {
            // Parse CSV data to List<List<object?>>
            var rows = ParseCsvToRows(csv!);

            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.Append(batch, tn!, rows);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Rows appended successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    /// <summary>
    /// Parse CSV data into List of List of objects for table operations.
    /// Simple CSV parser - assumes comma delimiter, handles quoted strings.
    /// </summary>
    private static List<List<object?>> ParseCsvToRows(string csvData)
    {
        var lines = csvData.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

        var rows = lines.Select(line =>
        {
            var values = line.Split(',');
            return values.Select(value =>
            {
                var trimmed = value.Trim().Trim('"');
                return string.IsNullOrEmpty(trimmed) ? null : (object?)trimmed;
            }).ToList();
        }).ToList();

        return rows;
    }

    private static string SetTableStyle(TableCommands commands, string sid, string? tn, string? style)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "set-style");
        if (string.IsNullOrWhiteSpace(style)) ExcelToolsBase.ThrowMissingParameter(nameof(style), "set-style");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.SetStyle(batch, tn!, style!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table style set successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string AddToDataModel(TableCommands commands, string sid, string? tn)
    {
        if (string.IsNullOrWhiteSpace(tn)) ExcelToolsBase.ThrowMissingParameter(nameof(tn), "add-to-datamodel");

        try
        {
            ExcelToolsBase.WithSession(
                sid,
                batch =>
                {
                    commands.AddToDataModel(batch, tn!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table added to data model successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }
}

