using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
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
    /// Manage Excel data connections - OLEDB, ODBC, Text, Web, and other connection types
    /// </summary>
    [McpServerTool(Name = "excel_connection")]
    [Description(@"Manage Excel data connections (OLEDB, ODBC, Text, Web).

CONNECTION TYPES SUPPORTED:
- OLEDB: SQL Server, Access, Oracle databases
- ODBC: ODBC data sources
- Text: CSV/text file imports
- Web: Web queries and APIs
- DataFeed: OData and data feeds
- Model: Data Model connections

POWER QUERY AUTO-REDIRECT:
- Power Query connections automatically redirect to excel_powerquery tool
- Use excel_powerquery for M code-based connections
")]
    public static async Task<string> ExcelConnection(
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

        [Description("JSON file path for import/export/update, or sheet name for loadto")]
        string? targetPath = null,

        [Description("Background query setting (for set-properties)")]
        bool? backgroundQuery = null,

        [Description("Refresh on file open setting (for set-properties)")]
        bool? refreshOnFileOpen = null,

        [Description("Save password setting (for set-properties)")]
        bool? savePassword = null,

        [Description("Refresh period in minutes (for set-properties)")]
        int? refreshPeriod = null,

        [Description("Timeout in minutes for connection operations. Default: 2 minutes")]
        double? timeout = null)
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
                ConnectionAction.Import => ImportConnectionAsync(connectionCommands, sessionId, connectionName, targetPath),
                ConnectionAction.Export => ExportConnectionAsync(connectionCommands, sessionId, connectionName, targetPath),
                ConnectionAction.UpdateProperties => UpdateConnectionAsync(connectionCommands, sessionId, connectionName, targetPath),
                ConnectionAction.Refresh => RefreshConnectionAsync(connectionCommands, excelPath, sessionId, connectionName, timeout),
                ConnectionAction.Delete => DeleteConnectionAsync(connectionCommands, sessionId, connectionName),
                ConnectionAction.Test => TestConnectionAsync(connectionCommands, sessionId, connectionName),
                ConnectionAction.LoadTo => LoadToWorksheetAsync(connectionCommands, sessionId, connectionName, targetPath),
                ConnectionAction.GetProperties => GetPropertiesAsync(connectionCommands, sessionId, connectionName),
                ConnectionAction.SetProperties => SetPropertiesAsync(connectionCommands, sessionId, connectionName, backgroundQuery, refreshOnFileOpen, savePassword, refreshPeriod),
                _ => throw new ArgumentException(
                    $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
            };
        }
        catch (Exception ex)
        {
            return Task.FromResult(JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"{action.ToActionString()} failed: {ex.Message}",
                isError = true
            }, ExcelToolsBase.JsonOptions));
        }
    }

    private static string ListConnectionsAsync(ConnectionCommands commands, string sessionId)
    {
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.List(batch));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        var count = result.Connections?.Count ?? 0;
        var powerQueryCount = result.Connections?.Count(c => c.IsPowerQuery) ?? 0;

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

    private static string ImportConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName, string? jsonPath)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for import action", nameof(connectionName));

        if (string.IsNullOrEmpty(jsonPath))
            throw new ArgumentException("targetPath (JSON file path) is required for import action", nameof(jsonPath));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Import(batch, connectionName, jsonPath));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        return JsonSerializer.Serialize(new
        {
            result.Success
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ExportConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName, string? jsonPath)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for export action", nameof(connectionName));

        if (string.IsNullOrEmpty(jsonPath))
            throw new ArgumentException("connectionName is required for export action", nameof(connectionName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.View(batch, connectionName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string UpdateConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName, string? jsonPath)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for update action", nameof(connectionName));

        if (string.IsNullOrEmpty(jsonPath))
            throw new ArgumentException("targetPath (JSON file path) is required for update action", nameof(jsonPath));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.UpdateProperties(batch, connectionName, jsonPath));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RefreshConnectionAsync(ConnectionCommands commands, string excelPath, string sessionId, string? connectionName, double? timeoutMinutes)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for refresh action", nameof(connectionName));

        try
        {
            // Apply timeout parameter (default 2 minutes for connection operations)
            var timeoutSpan = timeoutMinutes.HasValue ? (TimeSpan?)TimeSpan.FromMinutes(timeoutMinutes.Value) : null;

            var result = ExcelToolsBase.WithSession(
                sessionId,
                batch => commands.Refresh(batch, connectionName, timeoutSpan));

            // Always return JSON (success or failure) - MCP clients handle the success flag
            // Add workflow hints based on actual result and operation context
            if (result.Success)
            {
                // Check if connection is connection-only (no data loaded)
                bool isConnectionOnly = result.OperationContext?.ContainsKey("IsConnectionOnly") == true &&
                                      (bool)result.OperationContext["IsConnectionOnly"];

                return JsonSerializer.Serialize(new
                {
                    result.Success
                }, ExcelToolsBase.JsonOptions);
            }
            else
            {
                // Failed refresh - check for specific error types
                bool isPowerQueryConnection = result.ErrorMessage?.Contains("Power Query connection") == true;

                return JsonSerializer.Serialize(new
                {
                    result.Success,
                    result.ErrorMessage
                }, ExcelToolsBase.JsonOptions);
            }
        }
        catch (TimeoutException ex)
        {
            // Enrich timeout error with operation-specific guidance (MCP layer responsibility)
            var result = new OperationResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = excelPath,
                Action = "refresh",

                OperationContext = new Dictionary<string, object>
                {
                    { "OperationType", "Connection.Refresh" },
                    { "ConnectionName", connectionName },
                    { "TimeoutReached", true },
                    { "UsedMaxTimeout", ex.Message.Contains("maximum timeout") }
                },

                IsRetryable = !ex.Message.Contains("maximum timeout"),

                RetryGuidance = ex.Message.Contains("maximum timeout")
                    ? "Maximum timeout (5 minutes) reached. Check data source connectivity and resolve any authentication prompts before retrying."
                    : "Retry acceptable after checking for hidden dialogs and verifying data source connectivity. Consider using 'test' action first."
            };

            // MCP layer: Add workflow guidance for LLMs
            var response = new
            {
                result.Success,
                result.ErrorMessage,
                result.FilePath,
                result.Action,
                result.OperationContext,
                result.IsRetryable,
                result.RetryGuidance
            };

            return JsonSerializer.Serialize(response, ExcelToolsBase.JsonOptions);
        }
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
            throw new ArgumentException("targetPath (sheet name) is required for loadto action", nameof(sheetName));

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
        bool? backgroundQuery, bool? refreshOnFileOpen, bool? savePassword, int? refreshPeriod)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ArgumentException("connectionName is required for set-properties action", nameof(connectionName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.SetProperties(batch, connectionName, backgroundQuery, refreshOnFileOpen, savePassword, refreshPeriod));

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

