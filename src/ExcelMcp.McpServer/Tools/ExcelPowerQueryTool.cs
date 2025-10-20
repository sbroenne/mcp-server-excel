using Sbroenne.ExcelMcp.Core.Commands;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel Power Query management tool for MCP server.
/// Handles M code operations, query management, and data loading configurations.
/// 
/// LLM Usage Patterns:
/// - Use "list" to see all Power Queries in a workbook
/// - Use "view" to examine M code for a specific query
/// - Use "import" to add new queries from .pq files
/// - Use "export" to save M code to files for version control
/// - Use "update" to modify existing query M code
/// - Use "delete" to remove queries
/// - Use "set-load-to-table" to load query data to worksheet
/// - Use "set-load-to-data-model" to load to Excel's data model
/// - Use "set-load-to-both" to load to both table and data model
/// - Use "set-connection-only" to prevent data loading
/// - Use "get-load-config" to check current loading configuration
/// </summary>
public static class ExcelPowerQueryTool
{
    /// <summary>
    /// Manage Power Query operations - M code, data loading, and query lifecycle
    /// </summary>
    [McpServerTool(Name = "excel_powerquery")]
    [Description("Manage Power Query M code and data loading. Supports: list, view, import, export, update, delete, set-load-to-table, set-load-to-data-model, set-load-to-both, set-connection-only, get-load-config.")]
    public static string ExcelPowerQuery(
        [Description("Action: list, view, import, export, update, delete, set-load-to-table, set-load-to-data-model, set-load-to-both, set-connection-only, get-load-config")] string action,
        [Description("Excel file path (.xlsx or .xlsm)")] string filePath,
        [Description("Power Query name (required for most actions)")] string? queryName = null,
        [Description("Source .pq file path (for import/update) or target file path (for export)")] string? sourceOrTargetPath = null,
        [Description("Target worksheet name (for set-load-to-table action)")] string? targetSheet = null)
    {
        try
        {
            var powerQueryCommands = new PowerQueryCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => ListPowerQueries(powerQueryCommands, filePath),
                "view" => ViewPowerQuery(powerQueryCommands, filePath, queryName),
                "import" => ImportPowerQuery(powerQueryCommands, filePath, queryName, sourceOrTargetPath),
                "export" => ExportPowerQuery(powerQueryCommands, filePath, queryName, sourceOrTargetPath),
                "update" => UpdatePowerQuery(powerQueryCommands, filePath, queryName, sourceOrTargetPath),
                "delete" => DeletePowerQuery(powerQueryCommands, filePath, queryName),
                "set-load-to-table" => SetLoadToTable(powerQueryCommands, filePath, queryName, targetSheet),
                "set-load-to-data-model" => SetLoadToDataModel(powerQueryCommands, filePath, queryName),
                "set-load-to-both" => SetLoadToBoth(powerQueryCommands, filePath, queryName, targetSheet),
                "set-connection-only" => SetConnectionOnly(powerQueryCommands, filePath, queryName),
                "get-load-config" => GetLoadConfig(powerQueryCommands, filePath, queryName),
                _ => ExcelToolsBase.CreateUnknownActionError(action, 
                    "list", "view", "import", "export", "update", "delete",
                    "set-load-to-table", "set-load-to-data-model", "set-load-to-both", 
                    "set-connection-only", "get-load-config")
            };
        }
        catch (Exception ex)
        {
            return ExcelToolsBase.CreateExceptionError(ex, action, filePath);
        }
    }

    private static string ListPowerQueries(PowerQueryCommands commands, string filePath)
    {
        var result = commands.List(filePath);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ViewPowerQuery(PowerQueryCommands commands, string filePath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for view action" }, ExcelToolsBase.JsonOptions);

        var result = commands.View(filePath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ImportPowerQuery(PowerQueryCommands commands, string filePath, string? queryName, string? sourceOrTargetPath)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourceOrTargetPath))
            return JsonSerializer.Serialize(new { error = "queryName and sourceOrTargetPath are required for import action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Import(filePath, queryName, sourceOrTargetPath).GetAwaiter().GetResult();
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ExportPowerQuery(PowerQueryCommands commands, string filePath, string? queryName, string? sourceOrTargetPath)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourceOrTargetPath))
            return JsonSerializer.Serialize(new { error = "queryName and sourceOrTargetPath are required for export action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Export(filePath, queryName, sourceOrTargetPath).GetAwaiter().GetResult();
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string UpdatePowerQuery(PowerQueryCommands commands, string filePath, string? queryName, string? sourceOrTargetPath)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourceOrTargetPath))
            return JsonSerializer.Serialize(new { error = "queryName and sourceOrTargetPath are required for update action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Update(filePath, queryName, sourceOrTargetPath).GetAwaiter().GetResult();
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string DeletePowerQuery(PowerQueryCommands commands, string filePath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for delete action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Delete(filePath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetLoadToTable(PowerQueryCommands commands, string filePath, string? queryName, string? targetSheet)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for set-load-to-table action" }, ExcelToolsBase.JsonOptions);

        var result = commands.SetLoadToTable(filePath, queryName, targetSheet ?? "");
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetLoadToDataModel(PowerQueryCommands commands, string filePath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for set-load-to-data-model action" }, ExcelToolsBase.JsonOptions);

        var result = commands.SetLoadToDataModel(filePath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetLoadToBoth(PowerQueryCommands commands, string filePath, string? queryName, string? targetSheet)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for set-load-to-both action" }, ExcelToolsBase.JsonOptions);

        var result = commands.SetLoadToBoth(filePath, queryName, targetSheet ?? "");
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetConnectionOnly(PowerQueryCommands commands, string filePath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for set-connection-only action" }, ExcelToolsBase.JsonOptions);

        var result = commands.SetConnectionOnly(filePath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string GetLoadConfig(PowerQueryCommands commands, string filePath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for get-load-config action" }, ExcelToolsBase.JsonOptions);

        var result = commands.GetLoadConfig(filePath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}