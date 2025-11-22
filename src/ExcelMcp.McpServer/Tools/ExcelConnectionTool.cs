using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.McpServer.Models;

#pragma warning disable CA1861 // Avoid constant arrays as arguments - workflow hints are contextual per-call

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel connection management tool for MCP server.
/// Manages Excel data connections (OLEDB, ODBC, Text, Web, etc.) for data refresh operations.
/// Power Query connections automatically redirect to excel_powerquery tool.
/// </summary>
[McpServerToolType]
[SuppressMessage("Performance", "CA1861:Avoid constant arrays as arguments", Justification = "Conditional arrays with dynamic content")]
public static class ExcelConnectionTool
{
    /// <summary>
    /// Manage Excel data connections - OLEDB, ODBC, and ODC file imports
    /// </summary>
    [McpServerTool(Name = "excel_connection")]
    [Description(@"Manage Excel data connections (OLEDB, ODBC) and import ODC files.

CONNECTION TYPES SUPPORTED:
- OLEDB: SQL Server, Access, Oracle databases
- ODBC: ODBC data sources
- DataFeed: OData and data feeds
- Model: Data Model connections

TEXT/WEB FILE IMPORTS:
- TEXT and WEB connections are NOT supported via create action
- Use excel_powerquery tool for CSV/text file and web imports instead

POWER QUERY AUTO-REDIRECT:
- Power Query connections automatically redirect to excel_powerquery tool
- Use excel_powerquery for M code-based connections
")]
    public static string ExcelConnection(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        ConnectionAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [Required]
        [Description("Session ID from excel_file 'open' action")]
        string sessionId,

        [StringLength(255, MinimumLength = 1)]
        [Description("Connection name")]
        string? connectionName = null,

        [Description("Connection string (for create action)")]
        string? connectionString = null,

        [Description("Command text/SQL query (for create action, optional)")]
        string? commandText = null,

        [Description("Connection description (for create action, optional)")]
        string? description = null,

        [StringLength(31, MinimumLength = 1)]
        [Description("Sheet name for loadto action")]
        string? sheetName = null,

        [Description("New connection string (for set-properties, optional)")]
        string? newConnectionString = null,

        [Description("New command text/SQL query (for set-properties, optional)")]
        string? newCommandText = null,

        [Description("New connection description (for set-properties, optional)")]
        string? newDescription = null,

        [Description("Background query setting (for set-properties, optional)")]
        bool? backgroundQuery = null,

        [Description("Refresh on file open setting (for set-properties, optional)")]
        bool? refreshOnFileOpen = null,

        [Description("Save password setting (for set-properties, optional)")]
        bool? savePassword = null,

        [Description("Refresh period in minutes (for set-properties, optional)")]
        int? refreshPeriod = null)
    {
        try
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
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"{action.ToActionString()} failed: {ex.Message}",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ListConnectionsAsync(ConnectionCommands commands, string sessionId)
    {
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.List(batch));

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

        var result = ExcelToolsBase.WithSession(
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

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Refresh(batch, connectionName, null));

        if (result.Success)
        {
            return JsonSerializer.Serialize(new
            {
                result.Success
            }, ExcelToolsBase.JsonOptions);
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for delete action", nameof(connectionName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Delete(batch, connectionName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string LoadToWorksheetAsync(ConnectionCommands commands, string sessionId, string? connectionName, string? sheetName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for loadto action", nameof(connectionName));

        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for loadto action", nameof(sheetName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.LoadTo(batch, connectionName, sheetName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetPropertiesAsync(ConnectionCommands commands, string sessionId, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for properties action", nameof(connectionName));

        var result = ExcelToolsBase.WithSession(
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

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.SetProperties(batch, connectionName, newConnectionString, newCommandText, newDescription,
                backgroundQuery, refreshOnFileOpen, savePassword, refreshPeriod));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string TestConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for test action", nameof(connectionName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Test(batch, connectionName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints for connection testing
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
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

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Create(batch, connectionName, connectionString, commandText, description));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            connectionName
        }, ExcelToolsBase.JsonOptions);
    }
}
