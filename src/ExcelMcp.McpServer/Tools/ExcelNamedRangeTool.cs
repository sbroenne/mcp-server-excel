using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel named range (parameter) operations.
/// </summary>
[McpServerToolType]
public static partial class ExcelNamedRangeTool
{
    /// <summary>
    /// Manage Excel named ranges - named cell references for reusable formulas and parameters.
    /// CREATE/UPDATE: value is a cell reference (e.g., 'Sheet1!A1' or 'Sheet1!$A$1:$B$10'). Use $ for absolute references that won't shift when copied.
    /// WRITE: value is the actual data to store in the named range's cell(s).
    /// TIP: Use excel_range with rangeAddress=namedRangeName for bulk data operations on named ranges.
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="excelPath">Excel file path (.xlsx or .xlsm)</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action</param>
    /// <param name="namedRangeName">Named range name (for read, write, create, update, delete actions)</param>
    /// <param name="value">Named range value (for write action) or cell reference (for create/update actions, e.g., 'Sheet1!A1')</param>
    [McpServerTool(Name = "excel_namedrange")]
    [McpMeta("category", "data")]
    public static partial string ExcelParameter(
        NamedRangeAction action,
        string excelPath,
        string sessionId,
        [DefaultValue(null)] string? namedRangeName,
        [DefaultValue(null)] string? value)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_namedrange",
            action.ToActionString(),
            excelPath,
            () =>
            {
                var namedRangeCommands = new NamedRangeCommands();

                // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
                return action switch
                {
                    NamedRangeAction.List => ListNamedRangesAsync(namedRangeCommands, sessionId),
                    NamedRangeAction.Read => ReadNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName),
                    NamedRangeAction.Write => WriteNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName, value),
                    NamedRangeAction.Create => CreateNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName, value),
                    NamedRangeAction.Update => UpdateNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName, value),
                    NamedRangeAction.Delete => DeleteNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ListNamedRangesAsync(NamedRangeCommands commands, string sessionId)
    {
        var namedRanges = ExcelToolsBase.WithSession(sessionId, batch => commands.List(batch));
        return JsonSerializer.Serialize(namedRanges, ExcelToolsBase.JsonOptions);
    }

    private static string ReadNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName)
    {
        if (string.IsNullOrEmpty(namedRangeName))
            throw new ArgumentException("namedRangeName is required for read action", nameof(namedRangeName));

        var namedRangeValue = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Read(batch, namedRangeName));

        return JsonSerializer.Serialize(namedRangeValue, ExcelToolsBase.JsonOptions);
    }

    private static string WriteNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName, string? value)
    {
        if (string.IsNullOrEmpty(namedRangeName) || value == null)
            throw new ArgumentException("namedRangeName and value are required for write action", "namedRangeName,value");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.Write(batch, namedRangeName, value);
            return 0; // Dummy return for WithSession
        });

        return JsonSerializer.Serialize(new
        {
            success = true,
            message = $"Named range '{namedRangeName}' value updated successfully"
        }, ExcelToolsBase.JsonOptions);
    }

    private static string UpdateNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName, string? value)
    {
        if (string.IsNullOrEmpty(namedRangeName) || string.IsNullOrEmpty(value))
            throw new ArgumentException("namedRangeName and value (cell reference) are required for update action", "namedRangeName,value");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.Update(batch, namedRangeName, value);
            return 0; // Dummy return for WithSession
        });

        return JsonSerializer.Serialize(new
        {
            success = true,
            message = $"Named range '{namedRangeName}' reference updated successfully"
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName, string? value)
    {
        if (string.IsNullOrEmpty(namedRangeName) || string.IsNullOrEmpty(value))
            throw new ArgumentException("namedRangeName and value (cell reference) are required for create action", "namedRangeName,value");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.Create(batch, namedRangeName, value);
            return 0; // Dummy return for WithSession
        });

        return JsonSerializer.Serialize(new
        {
            success = true,
            message = $"Named range '{namedRangeName}' created successfully"
        }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName)
    {
        if (string.IsNullOrEmpty(namedRangeName))
            throw new ArgumentException("namedRangeName is required for delete action", nameof(namedRangeName));

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.Delete(batch, namedRangeName);
            return 0; // Dummy return for WithSession
        });

        return JsonSerializer.Serialize(new
        {
            success = true,
            message = $"Named range '{namedRangeName}' deleted successfully"
        }, ExcelToolsBase.JsonOptions);
    }
}

