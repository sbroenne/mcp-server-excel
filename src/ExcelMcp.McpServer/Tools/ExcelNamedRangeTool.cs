using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel named range (parameter) operations.
/// </summary>
[McpServerToolType]
public static partial class ExcelNamedRangeTool
{
    /// <summary>
    /// Named ranges for formulas/parameters.
    /// CREATE/UPDATE: value is cell reference (e.g., 'Sheet1!$A$1').
    /// WRITE: value is data to store.
    /// TIP: excel_range(rangeAddress=namedRangeName) for bulk data.
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="path">Excel file path (.xlsx or .xlsm)</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action</param>
    /// <param name="namedRangeName">Named range name (for read, write, create, update, delete actions)</param>
    /// <param name="value">Named range value (for write action) or cell reference (for create/update actions, e.g., 'Sheet1!A1')</param>
    [McpServerTool(Name = "excel_namedrange", Title = "Excel Named Range Operations", Destructive = true)]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelParameter(
        NamedRangeAction action,
        string path,
        string sessionId,
        [DefaultValue(null)] string? namedRangeName,
        [DefaultValue(null)] string? value)
    {
        _ = path; // retained parameter for schema compatibility

        return ExcelToolsBase.ExecuteToolAction(
            "excel_namedrange",
            ServiceRegistry.NamedRange.ToActionString(action),
            path,
            () => ServiceRegistry.NamedRange.RouteAction(
                action,
                sessionId,
                ExcelToolsBase.ForwardToServiceFunc,
                paramName: namedRangeName,
                value: value,
                reference: value));  // value doubles as reference for create/update
    }
}




