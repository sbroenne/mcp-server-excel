using Sbroenne.ExcelMcp.Core.Commands;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;

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
    [McpServerTool(Name = "connection")]
    [Description("Manage Excel data connections. Supports: list, view, import, export, update, refresh, delete, loadto, properties, set-properties, test.")]
    public static string Connection(
        [Required]
        [RegularExpression("^(list|view|import|export|update|refresh|delete|loadto|properties|set-properties|test)$")]
        [Description("Action: list, view, import, export, update, refresh, delete, loadto, properties, set-properties, test")] 
        string action,
        
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
        int? refreshPeriod = null)
    {
        try
        {
            var connectionCommands = new ConnectionCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => ListConnections(connectionCommands, excelPath),
                "view" => ViewConnection(connectionCommands, excelPath, connectionName),
                "import" => ImportConnection(connectionCommands, excelPath, connectionName, targetPath),
                "export" => ExportConnection(connectionCommands, excelPath, connectionName, targetPath),
                "update" => UpdateConnection(connectionCommands, excelPath, connectionName, targetPath),
                "refresh" => RefreshConnection(connectionCommands, excelPath, connectionName),
                "delete" => DeleteConnection(connectionCommands, excelPath, connectionName),
                "loadto" => LoadToWorksheet(connectionCommands, excelPath, connectionName, targetPath),
                "properties" => GetProperties(connectionCommands, excelPath, connectionName),
                "set-properties" => SetProperties(connectionCommands, excelPath, connectionName, 
                    backgroundQuery, refreshOnFileOpen, savePassword, refreshPeriod),
                "test" => TestConnection(connectionCommands, excelPath, connectionName),
                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action '{action}'. Supported: list, view, import, export, update, refresh, delete, loadto, properties, set-properties, test")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action, excelPath);
            throw; // Unreachable but satisfies compiler
        }
    }

    private static string ListConnections(ConnectionCommands commands, string filePath)
    {
        var result = commands.List(filePath);
        
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"list failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ViewConnection(ConnectionCommands commands, string filePath, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for view action");

        var result = commands.View(filePath, connectionName);
        
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"view failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ImportConnection(ConnectionCommands commands, string filePath, string? connectionName, string? jsonPath)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for import action");
        
        if (string.IsNullOrEmpty(jsonPath))
            throw new ModelContextProtocol.McpException("targetPath (JSON file path) is required for import action");

        var result = commands.Import(filePath, connectionName, jsonPath);
        
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"import failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ExportConnection(ConnectionCommands commands, string filePath, string? connectionName, string? jsonPath)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for export action");
        
        if (string.IsNullOrEmpty(jsonPath))
            throw new ModelContextProtocol.McpException("targetPath (JSON file path) is required for export action");

        var result = commands.Export(filePath, connectionName, jsonPath);
        
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"export failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string UpdateConnection(ConnectionCommands commands, string filePath, string? connectionName, string? jsonPath)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for update action");
        
        if (string.IsNullOrEmpty(jsonPath))
            throw new ModelContextProtocol.McpException("targetPath (JSON file path) is required for update action");

        var result = commands.Update(filePath, connectionName, jsonPath);
        
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"update failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string RefreshConnection(ConnectionCommands commands, string filePath, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for refresh action");

        var result = commands.Refresh(filePath, connectionName);
        
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"refresh failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteConnection(ConnectionCommands commands, string filePath, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for delete action");

        var result = commands.Delete(filePath, connectionName);
        
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"delete failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string LoadToWorksheet(ConnectionCommands commands, string filePath, string? connectionName, string? sheetName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for loadto action");
        
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("targetPath (sheet name) is required for loadto action");

        var result = commands.LoadTo(filePath, connectionName, sheetName);
        
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"loadto failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string GetProperties(ConnectionCommands commands, string filePath, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for properties action");

        var result = commands.GetProperties(filePath, connectionName);
        
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"properties failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetProperties(ConnectionCommands commands, string filePath, string? connectionName,
        bool? backgroundQuery, bool? refreshOnFileOpen, bool? savePassword, int? refreshPeriod)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for set-properties action");

        var result = commands.SetProperties(filePath, connectionName, backgroundQuery, refreshOnFileOpen, savePassword, refreshPeriod);
        
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-properties failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string TestConnection(ConnectionCommands commands, string filePath, string? connectionName)
    {
        if (string.IsNullOrEmpty(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for test action");

        var result = commands.Test(filePath, connectionName);
        
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"test failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
