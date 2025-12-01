using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;

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
    /// Manage Excel data connections (OLEDB, ODBC).
    /// CONNECTION TYPES: OLEDB (SQL Server, Access/Excel ACE, Oracle - provider must be installed), ODBC data sources, DataFeed/Model (existing Power Query/Power Pivot connections - list/view/refresh/delete only, creation routes to excel_powerquery).
    /// TEXT/WEB: Not supported via create - use excel_powerquery for CSV/text/web imports.
    /// POWER QUERY: Connections auto-redirect to excel_powerquery tool.
    /// TIMEOUT: Refresh and load-to auto-timeout after 5 minutes to prevent hanging.
    /// </summary>
    /// <param name="action">Action to perform (enum displayed as dropdown in MCP clients)</param>
    /// <param name="excelPath">Excel file path (.xlsx or .xlsm)</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action</param>
    /// <param name="connectionName">Connection name</param>
    /// <param name="connectionString">Connection string (for create action)</param>
    /// <param name="commandText">Command text/SQL query (for create action, optional)</param>
    /// <param name="description">Connection description (for create action, optional)</param>
    /// <param name="sheetName">Sheet name for loadto action</param>
    /// <param name="newConnectionString">New connection string (for set-properties, optional)</param>
    /// <param name="newCommandText">New command text/SQL query (for set-properties, optional)</param>
    /// <param name="newDescription">New connection description (for set-properties, optional)</param>
    /// <param name="backgroundQuery">Background query setting (for set-properties, optional)</param>
    /// <param name="refreshOnFileOpen">Refresh on file open setting (for set-properties, optional)</param>
    /// <param name="savePassword">Save password setting (for set-properties, optional)</param>
    /// <param name="refreshPeriod">Refresh period in minutes (for set-properties, optional)</param>
    [McpServerTool(Name = "excel_connection")]
    [McpMeta("category", "query")]
    public static partial string ExcelConnection(
        ConnectionAction action,
        string excelPath,
        string sessionId,
        [DefaultValue(null)] string? connectionName,
        [DefaultValue(null)] string? connectionString,
        [DefaultValue(null)] string? commandText,
        [DefaultValue(null)] string? description,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? newConnectionString,
        [DefaultValue(null)] string? newCommandText,
        [DefaultValue(null)] string? newDescription,
        [DefaultValue(null)] bool? backgroundQuery,
        [DefaultValue(null)] bool? refreshOnFileOpen,
        [DefaultValue(null)] bool? savePassword,
        [DefaultValue(null)] int? refreshPeriod)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_connection",
            action.ToActionString(),
            excelPath,
            () =>
            {
                var connectionCommands = new ConnectionCommands();

                // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
                return action switch
                {
                    ConnectionAction.List => ListConnectionsAsync(connectionCommands, sessionId),
                    ConnectionAction.View => ViewConnectionAsync(connectionCommands, sessionId, connectionName),
                    ConnectionAction.Create => CreateConnectionAsync(connectionCommands, sessionId, connectionName, connectionString, commandText, description),
                    ConnectionAction.Refresh => RefreshConnectionAsync(connectionCommands, excelPath, sessionId, connectionName),
                    ConnectionAction.Delete => DeleteConnectionAsync(connectionCommands, sessionId, connectionName),
                    ConnectionAction.Test => TestConnectionAsync(connectionCommands, sessionId, connectionName),
                    ConnectionAction.LoadTo => LoadToWorksheetAsync(connectionCommands, sessionId, connectionName, sheetName),
                    ConnectionAction.GetProperties => GetPropertiesAsync(connectionCommands, sessionId, connectionName),
                    ConnectionAction.SetProperties => SetPropertiesAsync(connectionCommands, sessionId, connectionName, newConnectionString, newCommandText, newDescription, backgroundQuery, refreshOnFileOpen, savePassword, refreshPeriod),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ListConnectionsAsync(ConnectionCommands commands, string sessionId)
    {
        ConnectionListResult result = ExcelToolsBase.WithSession(sessionId, batch => commands.List(batch));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Connections
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ViewConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for view action", nameof(connectionName));

        ConnectionViewResult result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.View(batch, connectionName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints for viewing connection details
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ConnectionName,
            result.ConnectionString,
            result.CommandText,
            result.IsPowerQuery
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RefreshConnectionAsync(ConnectionCommands commands, string excelPath, string sessionId, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for refresh action", nameof(connectionName));

        _ = excelPath; // retained parameter for schema compatibility

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.Refresh(batch, connectionName, TimeSpan.FromMinutes(5));
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Connection '{connectionName}' refreshed successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (TimeoutException ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string DeleteConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for delete action", nameof(connectionName));

        ExcelToolsBase.WithSession(
            sessionId,
            batch =>
            {
                commands.Delete(batch, connectionName);
                return 0;
            });

        return JsonSerializer.Serialize(new
        {
            success = true,
            message = $"Connection '{connectionName}' deleted successfully."
        }, ExcelToolsBase.JsonOptions);
    }

    private static string LoadToWorksheetAsync(ConnectionCommands commands, string sessionId, string? connectionName, string? sheetName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for loadto action", nameof(connectionName));

        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for loadto action", nameof(sheetName));

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.LoadTo(batch, connectionName, sheetName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Connection '{connectionName}' loaded to sheet '{sheetName}'."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (TimeoutException ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string GetPropertiesAsync(ConnectionCommands commands, string sessionId, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for properties action", nameof(connectionName));

        ConnectionPropertiesResult result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetProperties(batch, connectionName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            result.BackgroundQuery,
            result.RefreshOnFileOpen,
            result.SavePassword,
            result.RefreshPeriod
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetPropertiesAsync(ConnectionCommands commands, string sessionId, string? connectionName,
        string? newConnectionString, string? newCommandText, string? newDescription,
        bool? backgroundQuery, bool? refreshOnFileOpen, bool? savePassword, int? refreshPeriod)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for set-properties action", nameof(connectionName));

        ExcelToolsBase.WithSession(
            sessionId,
            batch =>
            {
                commands.SetProperties(batch, connectionName, newConnectionString, newCommandText, newDescription,
                    backgroundQuery, refreshOnFileOpen, savePassword, refreshPeriod);
                return 0;
            });

        return JsonSerializer.Serialize(new
        {
            success = true,
            message = $"Updated properties for connection '{connectionName}'."
        }, ExcelToolsBase.JsonOptions);
    }

    private static string TestConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for test action", nameof(connectionName));

        ExcelToolsBase.WithSession(
            sessionId,
            batch =>
            {
                commands.Test(batch, connectionName);
                return 0;
            });

        return JsonSerializer.Serialize(new
        {
            success = true,
            message = $"Connection '{connectionName}' is accessible."
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateConnectionAsync(
        ConnectionCommands commands,
        string sessionId,
        string? connectionName,
        string? connectionString,
        string? commandText,
        string? description)
    {
        if (string.IsNullOrWhiteSpace(connectionName))
            throw new ArgumentException("connectionName is required for create action", nameof(connectionName));

        if (string.IsNullOrWhiteSpace(connectionString))
            throw new ArgumentException("connectionString is required for create action", nameof(connectionString));

        ExcelToolsBase.WithSession(
            sessionId,
            batch =>
            {
                commands.Create(batch, connectionName, connectionString, commandText, description);
                return 0;
            });

        return JsonSerializer.Serialize(new
        {
            success = true,
            connectionName,
            message = $"Connection '{connectionName}' created successfully."
        }, ExcelToolsBase.JsonOptions);
    }
}
