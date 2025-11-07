using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel connection management tool for MCP server.
/// Manages Excel data connections (OLEDB, ODBC, Text, Web, etc.) for data refresh operations.
/// Power Query connections automatically redirect to excel_powerquery tool.
/// </summary>
[McpServerToolType]
public static class ExcelConnectionTool
{
    /// <summary>
    /// Manage Excel data connections - OLEDB, ODBC, Text, Web, and other connection types
    /// </summary>
    [McpServerTool(Name = "excel_connection")]
    [Description("Manage Excel data connections. Supports: list, view, create, import, export, update, refresh, delete, loadto, properties, set-properties, test.")]
    public static async Task<string> ExcelConnection(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        ConnectionAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

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

        [Description("Optional batch ID for grouping operations")]
        string? batchId = null,

        [Description("Timeout in minutes for connection operations. Default: 2 minutes")]
        double? timeout = null)
    {
        try
        {
            var connectionCommands = new ConnectionCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                ConnectionAction.List => await ListConnectionsAsync(connectionCommands, excelPath, batchId),
                ConnectionAction.View => await ViewConnectionAsync(connectionCommands, excelPath, connectionName, batchId),
                ConnectionAction.Create => await CreateConnectionAsync(connectionCommands, excelPath, connectionName, connectionString, commandText, description, batchId),
                ConnectionAction.Import => await ImportConnectionAsync(connectionCommands, excelPath, connectionName, targetPath, batchId),
                ConnectionAction.Export => await ExportConnectionAsync(connectionCommands, excelPath, connectionName, targetPath, batchId),
                ConnectionAction.UpdateProperties => await UpdateConnectionAsync(connectionCommands, excelPath, connectionName, targetPath, batchId),
                ConnectionAction.Refresh => await RefreshConnectionAsync(connectionCommands, excelPath, connectionName, timeout, batchId),
                ConnectionAction.Delete => await DeleteConnectionAsync(connectionCommands, excelPath, connectionName, batchId),
                ConnectionAction.Test => await TestConnectionAsync(connectionCommands, excelPath, connectionName, batchId),
                ConnectionAction.LoadTo => await LoadToWorksheetAsync(connectionCommands, excelPath, connectionName, targetPath, batchId),
                ConnectionAction.GetProperties => await GetPropertiesAsync(connectionCommands, excelPath, connectionName, batchId),
                ConnectionAction.SetProperties => await SetPropertiesAsync(connectionCommands, excelPath, connectionName, backgroundQuery, refreshOnFileOpen, savePassword, refreshPeriod, batchId),
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

    private static async Task<string> ListConnectionsAsync(ConnectionCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ListAsync(batch));

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
                    : $"Found {count} regular connection(s). Ready for refresh or data operations.",
            suggestedNextActions = count == 0
                ? new[] {
                    "Use 'import' to add connections from .odc files",
                    "Use excel_powerquery for M code connections",
                    "Create connections via Excel UI (Data → Get Data)"
                }
                : new[]
                {
                    powerQueryCount > 0 ? "Use excel_powerquery tool for Power Query connections" : null,
                    "Use 'refresh' to update data from external sources",
                    "Use 'view' to inspect connection details and credentials",
                    "Use 'properties' to check refresh settings and background query status",
                    "Use 'export' to backup connection definitions as JSON"
                }.Where(s => s != null).ToArray()!
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ViewConnectionAsync(ConnectionCommands commands, string filePath, string? connectionName, string? batchId)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for view action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ViewAsync(batch, connectionName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints for viewing connection details
        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ConnectionName,
            result.ConnectionString,
            result.CommandText,
            result.IsPowerQuery,
            workflowHint = result.IsPowerQuery
                ? $"Power Query connection '{connectionName}' detected. Use excel_powerquery tool for management."
                : $"Connection '{connectionName}' details retrieved. Ready for refresh or configuration.",
            suggestedNextActions = result.IsPowerQuery ? new[]
            {
                "Use excel_powerquery 'view' to see the M code for this Power Query connection",
                "Use excel_powerquery 'refresh' to update this Power Query data",
                "Use excel_powerquery 'list' to see all Power Query connections"
            } : new[]
            {
                "Use excel_connection 'refresh' to update data from this connection",
                "Use excel_connection 'test' to validate connection without refreshing data",
                "Use excel_connection 'properties' to check refresh settings and background query status",
                "Use excel_connection 'export' to backup this connection definition",
                inBatch ? "View more connections in this batch" : "Need to check multiple connections? Use excel_batch for efficiency"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ImportConnectionAsync(ConnectionCommands commands, string filePath, string? connectionName, string? jsonPath, string? batchId)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for import action");

        if (string.IsNullOrEmpty(jsonPath))
            throw new ModelContextProtocol.McpException("targetPath (JSON file path) is required for import action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.ImportAsync(batch, connectionName, jsonPath));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            workflowHint = $"Connection '{connectionName}' imported successfully from {jsonPath}. Ready for use.",
            suggestedNextActions = new[]
            {
                "Use excel_connection 'test' to verify the imported connection works",
                "Use excel_connection 'refresh' to load latest data from the data source",
                "Use excel_connection 'loadto' to load connection data to a specific worksheet",
                "Use excel_connection 'view' to inspect the imported connection details",
                inBatch ? "Import more connections in this batch" : "Importing multiple connections? Use excel_batch for efficiency"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ExportConnectionAsync(ConnectionCommands commands, string filePath, string? connectionName, string? jsonPath, string? batchId)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for export action");

        if (string.IsNullOrEmpty(jsonPath))
            throw new ModelContextProtocol.McpException("targetPath (JSON file path) is required for export action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ExportAsync(batch, connectionName, jsonPath));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateConnectionAsync(ConnectionCommands commands, string filePath, string? connectionName, string? jsonPath, string? batchId)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for update action");

        if (string.IsNullOrEmpty(jsonPath))
            throw new ModelContextProtocol.McpException("targetPath (JSON file path) is required for update action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.UpdatePropertiesAsync(batch, connectionName, jsonPath));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshConnectionAsync(ConnectionCommands commands, string filePath, string? connectionName, double? timeoutMinutes, string? batchId)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for refresh action");

        try
        {
            // Apply timeout parameter (default 2 minutes for connection operations)
            var timeoutSpan = timeoutMinutes.HasValue ? (TimeSpan?)TimeSpan.FromMinutes(timeoutMinutes.Value) : null;

            var result = await ExcelToolsBase.WithBatchAsync(
                batchId,
                filePath,
                save: true,
                async (batch) => await commands.RefreshAsync(batch, connectionName, timeoutSpan));

            // Always return JSON (success or failure) - MCP clients handle the success flag
            // Add workflow hints based on actual result and operation context
            var inBatch = !string.IsNullOrEmpty(batchId);

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
                        inBatch ? "Validate more connections in this batch" : "Testing multiple connections? Use excel_batch for efficiency"
                    } : new[]
                    {
                        "Data refresh completed - external data source has been queried",
                        "Use excel_range 'get-values' or 'get-used-range' to examine refreshed data",
                        "Use excel_connection 'view' to verify last refresh timestamp",
                        "Use excel_connection 'properties' to check auto-refresh settings",
                        inBatch ? "Refresh more connections in this batch" : "Refreshing multiple connections? Use excel_batch for better performance"
                    }
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
                    workflowHint = $"Connection '{connectionName}' refresh failed - data source issue detected.",
                    suggestedNextActions = isPowerQueryConnection ? new[]
                    {
                        "Power Query connections detected - use excel_powerquery 'refresh' instead",
                        "Use excel_powerquery 'list' to see all Power Query connections",
                        "Use excel_connection 'list' to see regular data connections only",
                        "Power Query connections require different refresh mechanism"
                    } : new[]
                    {
                        "Check if data source is accessible (database server, file share, web service)",
                        "Use excel_connection 'view' to inspect connection string and credentials",
                        "Verify network connectivity and firewall rules for external data sources",
                        "Use excel_connection 'test' to validate connection without refreshing data",
                        "Check if credentials have expired or need updating"
                    }
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
                FilePath = filePath,
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

    private static async Task<string> DeleteConnectionAsync(ConnectionCommands commands, string filePath, string? connectionName, string? batchId)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for delete action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.DeleteAsync(batch, connectionName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> LoadToWorksheetAsync(ConnectionCommands commands, string filePath, string? connectionName, string? sheetName, string? batchId)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for loadto action");

        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("targetPath (sheet name) is required for loadto action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.LoadToAsync(batch, connectionName, sheetName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetPropertiesAsync(ConnectionCommands commands, string filePath, string? connectionName, string? batchId)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for properties action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetPropertiesAsync(batch, connectionName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetPropertiesAsync(ConnectionCommands commands, string filePath, string? connectionName,
        bool? backgroundQuery, bool? refreshOnFileOpen, bool? savePassword, int? refreshPeriod, string? batchId)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for set-properties action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.SetPropertiesAsync(batch, connectionName, backgroundQuery, refreshOnFileOpen, savePassword, refreshPeriod));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> TestConnectionAsync(ConnectionCommands commands, string filePath, string? connectionName, string? batchId)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for test action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.TestAsync(batch, connectionName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints for connection testing
        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Connection '{connectionName}' test successful - data source is accessible and responding."
                : $"Connection '{connectionName}' test failed - data source connectivity issue detected.",
            suggestedNextActions = result.Success ? new[]
            {
                "Connection is working - use excel_connection 'refresh' to load actual data",
                "Use excel_connection 'loadto' to load connection data to a specific worksheet",
                "Use excel_connection 'properties' to configure refresh settings",
                "Use excel_connection 'view' to inspect connection details",
                inBatch ? "Test more connections in this batch" : "Testing multiple connections? Use excel_batch for efficiency"
            } : new[]
            {
                "Connection test failed - check if data source is accessible",
                "Use excel_connection 'view' to inspect connection string and credentials",
                "Verify network connectivity and firewall rules for external data sources",
                "Check if credentials have expired or need updating",
                "For OLEDB/ODBC connections, test using Windows ODBC Data Source Administrator"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateConnectionAsync(
        ConnectionCommands commands,
        string excelPath,
        string? connectionName,
        string? connectionString,
        string? commandText,
        string? description,
        string? batchId)
    {
        if (string.IsNullOrWhiteSpace(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for create action");

        if (string.IsNullOrWhiteSpace(connectionString))
            throw new ModelContextProtocol.McpException("connectionString is required for create action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.CreateAsync(batch, connectionName, connectionString, commandText, description));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"create action failed for '{connectionName}': {result.ErrorMessage}");
        }

        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            connectionName,
            workflowHint = result.Success
                ? $"Connection '{connectionName}' created successfully and ready for use."
                : $"Connection '{connectionName}' creation failed - check connection string format and parameters.",
            suggestedNextActions = result.Success ? new[]
            {
                "Test the new connection with excel_connection 'test' to verify connectivity",
                "Use excel_connection 'refresh' to load data from the connection",
                "Use excel_connection 'loadto' to load connection data to a specific worksheet",
                "Use excel_connection 'properties' to configure refresh settings (background query, auto-refresh)",
                "Use excel_connection 'view' to inspect the created connection details",
                inBatch ? "Create more connections in this batch" : "Creating multiple connections? Use excel_batch for efficiency"
            } : new[]
            {
                "Connection creation failed - verify connection string format is correct",
                "For TEXT connections, use format: 'TEXT;C:\\path\\to\\file.csv'",
                "For OLEDB connections, include Provider and connection parameters",
                "For ODBC connections, reference a valid DSN or use connection string format",
                "Use excel_connection 'view' on existing connections to see working examples"
            }
        }, ExcelToolsBase.JsonOptions);
    }
}
