using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel connection management tool for MCP server.
/// Manages Excel data connections (OLEDB, ODBC, Text, Web, etc.) for data refresh operations.
/// Power Query connections automatically redirect to excel_powerquery tool.
/// </summary>
[McpServerToolType]
public static partial class ExcelConnectionTool
{
    /// <summary>
    /// Data connections (OLEDB, ODBC, ODC import).
    /// TEXT/WEB/CSV: Use excel_powerquery instead.
    /// Power Query connections auto-redirect to excel_powerquery.
    /// TIMEOUT: 5 min auto-timeout for refresh/loadto.
    /// </summary>
    /// <param name="action">Action to perform (enum displayed as dropdown in MCP clients)</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action</param>
    /// <param name="connectionName">Connection name</param>
    /// <param name="connectionString">Connection string (for create or set-properties)</param>
    /// <param name="commandText">Command text/SQL query (for create or set-properties)</param>
    /// <param name="description">Connection description (for create or set-properties)</param>
    /// <param name="sheetName">Sheet name for loadto action</param>
    /// <param name="backgroundQuery">Background query setting (for set-properties)</param>
    /// <param name="refreshOnFileOpen">Refresh on file open setting (for set-properties)</param>
    /// <param name="savePassword">Save password setting (for set-properties)</param>
    /// <param name="refreshPeriod">Refresh period in minutes (for set-properties)</param>
    [McpServerTool(Name = "excel_connection", Title = "Excel Data Connection Operations", Destructive = true)]
    [McpMeta("category", "query")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelConnection(
        ConnectionAction action,
        string sessionId,
        [DefaultValue(null)] string? connectionName,
        [DefaultValue(null)] string? connectionString,
        [DefaultValue(null)] string? commandText,
        [DefaultValue(null)] string? description,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] bool? backgroundQuery,
        [DefaultValue(null)] bool? refreshOnFileOpen,
        [DefaultValue(null)] bool? savePassword,
        [DefaultValue(null)] int? refreshPeriod)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_connection",
            ServiceRegistry.Connection.ToActionString(action),
            () => ServiceRegistry.Connection.RouteAction(
                action,
                sessionId,
                ExcelToolsBase.ForwardToServiceFunc,
                connectionName: connectionName,
                connectionString: connectionString,
                commandText: commandText,
                description: description,
                sheetName: sheetName,
                timeout: null, // Use default timeout handling in service
                backgroundQuery: backgroundQuery,
                refreshOnFileOpen: refreshOnFileOpen,
                savePassword: savePassword,
                refreshPeriod: refreshPeriod));
    }
}




