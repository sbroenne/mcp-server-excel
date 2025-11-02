using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel connection management tool for MCP server.
/// Handles data connections (OLEDB, ODBC, Text, Web, etc.) for Excel automation.
///
/// LLM Usage Patterns:
/// - Use "list" to see all connections in a workbook
/// - Use "view" to inspect connection details (connection string, command text)
/// - Use "export" to save connection definitions to JSON for version control
/// - Use "update" to modify existing connections from JSON definitions
/// - Use "refresh" to update data from external sources
/// - Use "loadto" to load connection data to a worksheet
/// - Use "properties" to check connection configuration (background query, refresh settings)
/// - Use "set-properties" to configure connection behavior
/// - Use "test" to validate connection without refreshing data
/// - Use "delete" to remove obsolete connections
///
/// Note: Power Query connections are detected and users are redirected to excel_powerquery tool.
/// Regular connections (OLEDB, ODBC, Text, Web) use standard connection strings.
/// Password sanitization is applied automatically for security.
/// </summary>
[McpServerToolType]
public static class ExcelConnectionTool
{
    /// <summary>
    /// Manage Excel data connections - OLEDB, ODBC, Text, Web, and other connection types
    /// </summary>
    [McpServerTool(Name = "excel_connection")]
    [Description("Manage Excel data connections. Supports: list, view, import, export, update, refresh, delete, loadto, properties, set-properties, test.")]
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
        string? batchId = null)
    {
        try
        {
            var connectionCommands = new ConnectionCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                ConnectionAction.List => await ListConnectionsAsync(connectionCommands, excelPath, batchId),
                ConnectionAction.View => await ViewConnectionAsync(connectionCommands, excelPath, connectionName, batchId),
                ConnectionAction.Import => await ImportConnectionAsync(connectionCommands, excelPath, connectionName, targetPath, batchId),
                ConnectionAction.Export => await ExportConnectionAsync(connectionCommands, excelPath, connectionName, targetPath, batchId),
                ConnectionAction.UpdateProperties => await UpdateConnectionAsync(connectionCommands, excelPath, connectionName, targetPath, batchId),
                ConnectionAction.Refresh => await RefreshConnectionAsync(connectionCommands, excelPath, connectionName, batchId),
                ConnectionAction.Delete => await DeleteConnectionAsync(connectionCommands, excelPath, connectionName, batchId),
                ConnectionAction.Test => await TestConnectionAsync(connectionCommands, excelPath, connectionName, batchId),
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"list failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"view failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"import failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"export failed for '{filePath}': {result.ErrorMessage}");
        }

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
            async (batch) => await commands.UpdateAsync(batch, connectionName, jsonPath));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"update failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshConnectionAsync(ConnectionCommands commands, string filePath, string? connectionName, string? batchId)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for refresh action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.RefreshAsync(batch, connectionName));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"refresh failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"delete failed for '{filePath}': {result.ErrorMessage}");
        }

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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"loadto failed for '{filePath}': {result.ErrorMessage}");
        }

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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"properties failed for '{filePath}': {result.ErrorMessage}");
        }

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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-properties failed for '{filePath}': {result.ErrorMessage}");
        }

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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"test failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
