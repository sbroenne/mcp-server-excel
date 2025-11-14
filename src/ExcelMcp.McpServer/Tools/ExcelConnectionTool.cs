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
                ConnectionAction.List => await ListConnectionsAsync(connectionCommands, sessionId),
                ConnectionAction.View => await ViewConnectionAsync(connectionCommands, sessionId, connectionName),
                ConnectionAction.Create => await CreateConnectionAsync(connectionCommands, sessionId, connectionName, connectionString, commandText, description),
                ConnectionAction.Import => await ImportConnectionAsync(connectionCommands, sessionId, connectionName, targetPath),
                ConnectionAction.Export => await ExportConnectionAsync(connectionCommands, sessionId, connectionName, targetPath),
                ConnectionAction.UpdateProperties => await UpdateConnectionAsync(connectionCommands, sessionId, connectionName, targetPath),
                ConnectionAction.Refresh => await RefreshConnectionAsync(connectionCommands, excelPath, sessionId, connectionName, timeout),
                ConnectionAction.Delete => await DeleteConnectionAsync(connectionCommands, sessionId, connectionName),
                ConnectionAction.Test => await TestConnectionAsync(connectionCommands, sessionId, connectionName),
                ConnectionAction.LoadTo => await LoadToWorksheetAsync(connectionCommands, sessionId, connectionName, targetPath),
                ConnectionAction.GetProperties => await GetPropertiesAsync(connectionCommands, sessionId, connectionName),
                ConnectionAction.SetProperties => await SetPropertiesAsync(connectionCommands, sessionId, connectionName, backgroundQuery, refreshOnFileOpen, savePassword, refreshPeriod),
                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action: {action} ({action.ToActionString()})")
            };
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

    private static async Task<string> ListConnectionsAsync(ConnectionCommands commands, string sessionId)
    {
        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ListAsync(batch));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        var count = result.Connections?.Count ?? 0;
        var powerQueryCount = result.Connections?.Count(c => c.IsPowerQuery) ?? 0;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Connections,
            workflowHint = count == 0
                ? "No connections found. Create connections via Excel UI or import from .odc files."
                : powerQueryCount > 0
                    ? $"Found {count} connection(s): {count - powerQueryCount} regular, {powerQueryCount} Power Query. Different tools needed."
                    : $"Found {count} regular connection(s). Ready for refresh or data operations."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ViewConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for view action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ViewAsync(batch, connectionName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints for viewing connection details
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ConnectionName,
            result.ConnectionString,
            result.CommandText,
            result.IsPowerQuery,
            workflowHint = result.IsPowerQuery
                ? $"Power Query connection '{connectionName}' detected. Use excel_powerquery tool for management."
                : $"Connection '{connectionName}' details retrieved. Ready for refresh or configuration."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ImportConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName, string? jsonPath)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for import action");

        if (string.IsNullOrEmpty(jsonPath))
            throw new ModelContextProtocol.McpException("targetPath (JSON file path) is required for import action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ImportAsync(batch, connectionName, jsonPath));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        return JsonSerializer.Serialize(new
        {
            result.Success,
            workflowHint = $"Connection '{connectionName}' imported successfully from {jsonPath}. Ready for use."

        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ExportConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName, string? jsonPath)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for export action");

        if (string.IsNullOrEmpty(jsonPath))
            throw new ModelContextProtocol.McpException("targetPath (JSON file path) is required for export action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ExportAsync(batch, connectionName, jsonPath));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Connection '{connectionName}' exported to {jsonPath}. Use for version control or deployment."
                : $"Failed to export connection '{connectionName}'. Verify connection exists and file path is writable."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName, string? jsonPath)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for update action");

        if (string.IsNullOrEmpty(jsonPath))
            throw new ModelContextProtocol.McpException("targetPath (JSON file path) is required for update action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.UpdatePropertiesAsync(batch, connectionName, jsonPath));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Connection '{connectionName}' properties updated from {jsonPath}. New settings applied."
                : $"Failed to update connection '{connectionName}'. Verify connection exists and JSON file format is valid."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshConnectionAsync(ConnectionCommands commands, string excelPath, string sessionId, string? connectionName, double? timeoutMinutes)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for refresh action");

        try
        {
            // Apply timeout parameter (default 2 minutes for connection operations)
            var timeoutSpan = timeoutMinutes.HasValue ? (TimeSpan?)TimeSpan.FromMinutes(timeoutMinutes.Value) : null;

            var result = await ExcelToolsBase.WithSessionAsync(
                sessionId,
                async batch => await commands.RefreshAsync(batch, connectionName, timeoutSpan));

            // Always return JSON (success or failure) - MCP clients handle the success flag
            // Add workflow hints based on actual result and operation context
            if (result.Success)
            {
                // Check if connection is connection-only (no data loaded)
                bool isConnectionOnly = result.OperationContext?.ContainsKey("IsConnectionOnly") == true &&
                                      (bool)result.OperationContext["IsConnectionOnly"];

                return JsonSerializer.Serialize(new
                {
                    result.Success,
                    workflowHint = isConnectionOnly
                        ? $"Connection '{connectionName}' validated successfully. Connection is working but no data is loaded to worksheets."
                        : $"Connection '{connectionName}' refreshed successfully. External data has been updated in the workbook.",
                    suggestedNextActions = isConnectionOnly ? new[]
                    {
                        "Connection validation confirmed - data source is accessible",
                        "Use excel_connection 'loadto' to load data to a specific worksheet",
                        "Use excel_connection 'view' to see connection details and last refresh time",
                        "Connection-only means no QueryTables exist - data source ready for use",
                        "Validate more connections in this session"
                    } :
                    [
                        "Data refresh completed - external data source has been queried",
                        "Use excel_range 'get-values' or 'get-used-range' to examine refreshed data",
                        "Use excel_connection 'view' to verify last refresh timestamp",
                        "Use excel_connection 'properties' to check auto-refresh settings",
                        "Refresh more connections in this session"
                    ]
                }, ExcelToolsBase.JsonOptions);
            }
            else
            {
                // Failed refresh - check for specific error types
                bool isPowerQueryConnection = result.ErrorMessage?.Contains("Power Query connection") == true;

                return JsonSerializer.Serialize(new
                {
                    result.Success,
                    result.ErrorMessage,
                    workflowHint = $"Connection '{connectionName}' refresh failed - data source issue detected."
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
                result.RetryGuidance,

                // Workflow hints - MCP Server layer responsibility
                WorkflowHint = "Connection refresh timeout - verify data source accessibility",
                SuggestedNextActions = new[]
                {
                    "Connection refresh timed out - check for blocking dialogs in Excel",
                    "Verify the data source is responsive (database server, network share, web service)",
                    "For OLEDB/ODBC connections, test connectivity using Windows ODBC Data Source Administrator",
                    "Check firewall rules and network connectivity to remote data sources",
                    "Look for credential prompts or authentication dialogs that may be hidden",
                    "Large datasets may require longer refresh times - consider data filtering at source"
                }
            };

            return JsonSerializer.Serialize(response, ExcelToolsBase.JsonOptions);
        }
    }

    private static async Task<string> DeleteConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for delete action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.DeleteAsync(batch, connectionName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Connection '{connectionName}' deleted successfully. QueryTables using this connection may need cleanup."
                : $"Failed to delete connection '{connectionName}'. Verify connection exists and is not in use."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> LoadToWorksheetAsync(ConnectionCommands commands, string sessionId, string? connectionName, string? sheetName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for loadto action");

        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("targetPath (sheet name) is required for loadto action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.LoadToAsync(batch, connectionName, sheetName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Connection '{connectionName}' data loaded to worksheet '{sheetName}'. Data table created."
                : $"Failed to load connection '{connectionName}' to '{sheetName}'. Verify connection and sheet exist."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetPropertiesAsync(ConnectionCommands commands, string sessionId, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for properties action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetPropertiesAsync(batch, connectionName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            result.BackgroundQuery,
            result.RefreshOnFileOpen,
            result.SavePassword,
            result.RefreshPeriod,
            workflowHint = result.Success
                ? $"Connection '{connectionName}' properties retrieved. Review refresh settings and background query status."
                : $"Failed to retrieve properties for connection '{connectionName}'. Verify connection exists."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetPropertiesAsync(ConnectionCommands commands, string sessionId, string? connectionName,
        bool? backgroundQuery, bool? refreshOnFileOpen, bool? savePassword, int? refreshPeriod)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for set-properties action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SetPropertiesAsync(batch, connectionName, backgroundQuery, refreshOnFileOpen, savePassword, refreshPeriod));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Connection '{connectionName}' properties updated. New settings will apply on next refresh."
                : $"Failed to update properties for connection '{connectionName}'. Verify connection exists and property values are valid."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> TestConnectionAsync(ConnectionCommands commands, string sessionId, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for test action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.TestAsync(batch, connectionName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints for connection testing
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Connection '{connectionName}' test successful - data source is accessible and responding."
                : $"Connection '{connectionName}' test failed - data source connectivity issue detected."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateConnectionAsync(
        ConnectionCommands commands,
        string sessionId,
        string? connectionName,
        string? connectionString,
        string? commandText,
        string? description)
    {
        if (string.IsNullOrWhiteSpace(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for create action");

        if (string.IsNullOrWhiteSpace(connectionString))
            throw new ModelContextProtocol.McpException("connectionString is required for create action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.CreateAsync(batch, connectionName, connectionString, commandText, description));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            connectionName,
            workflowHint = result.Success
                ? $"Connection '{connectionName}' created successfully and ready for use."
                : $"Connection '{connectionName}' creation failed - check connection string format and parameters."
        }, ExcelToolsBase.JsonOptions);
    }
}
