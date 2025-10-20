using Sbroenne.ExcelMcp.Core.Commands;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
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
/// - Use "refresh" to refresh query data from source
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
    [Description("Manage Power Query M code and data loading. Supports: list, view, import, export, update, refresh, delete, set-load-to-table, set-load-to-data-model, set-load-to-both, set-connection-only, get-load-config.")]
    public static string ExcelPowerQuery(
        [Required]
        [RegularExpression("^(list|view|import|export|update|refresh|delete|set-load-to-table|set-load-to-data-model|set-load-to-both|set-connection-only|get-load-config)$")]
        [Description("Action: list, view, import, export, update, refresh, delete, set-load-to-table, set-load-to-data-model, set-load-to-both, set-connection-only, get-load-config")] 
        string action,
        
        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")] 
        string excelPath,
        
        [StringLength(255, MinimumLength = 1)]
        [Description("Power Query name (required for most actions)")] 
        string? queryName = null,
        
        [FileExtensions(Extensions = "pq,txt,m")]
        [Description("Source .pq file path (for import/update actions)")] 
        string? sourcePath = null,
        
        [FileExtensions(Extensions = "pq,txt,m")]
        [Description("Target file path (for export action)")] 
        string? targetPath = null,
        
        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("Target worksheet name (for set-load-to-table action)")] 
        string? targetSheet = null)
    {
        try
        {
            var powerQueryCommands = new PowerQueryCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => ListPowerQueries(powerQueryCommands, excelPath),
                "view" => ViewPowerQuery(powerQueryCommands, excelPath, queryName),
                "import" => ImportPowerQuery(powerQueryCommands, excelPath, queryName, sourcePath),
                "export" => ExportPowerQuery(powerQueryCommands, excelPath, queryName, targetPath),
                "update" => UpdatePowerQuery(powerQueryCommands, excelPath, queryName, sourcePath),
                "refresh" => RefreshPowerQuery(powerQueryCommands, excelPath, queryName),
                "delete" => DeletePowerQuery(powerQueryCommands, excelPath, queryName),
                "set-load-to-table" => SetLoadToTable(powerQueryCommands, excelPath, queryName, targetSheet),
                "set-load-to-data-model" => SetLoadToDataModel(powerQueryCommands, excelPath, queryName),
                "set-load-to-both" => SetLoadToBoth(powerQueryCommands, excelPath, queryName, targetSheet),
                "set-connection-only" => SetConnectionOnly(powerQueryCommands, excelPath, queryName),
                "get-load-config" => GetLoadConfig(powerQueryCommands, excelPath, queryName),
                _ => ExcelToolsBase.CreateUnknownActionError(action, 
                    "list", "view", "import", "export", "update", "refresh", "delete",
                    "set-load-to-table", "set-load-to-data-model", "set-load-to-both", 
                    "set-connection-only", "get-load-config")
            };
        }
        catch (Exception ex)
        {
            return ExcelToolsBase.CreateExceptionError(ex, action, excelPath);
        }
    }

    private static string ListPowerQueries(PowerQueryCommands commands, string excelPath)
    {
        var result = commands.List(excelPath);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ViewPowerQuery(PowerQueryCommands commands, string excelPath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for view action" }, ExcelToolsBase.JsonOptions);

        var result = commands.View(excelPath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ImportPowerQuery(PowerQueryCommands commands, string excelPath, string? queryName, string? sourcePath)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourcePath))
            return JsonSerializer.Serialize(new { error = "queryName and sourcePath are required for import action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Import(excelPath, queryName, sourcePath).GetAwaiter().GetResult();
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ExportPowerQuery(PowerQueryCommands commands, string excelPath, string? queryName, string? targetPath)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(targetPath))
            return JsonSerializer.Serialize(new { error = "queryName and targetPath are required for export action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Export(excelPath, queryName, targetPath).GetAwaiter().GetResult();
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string UpdatePowerQuery(PowerQueryCommands commands, string excelPath, string? queryName, string? sourcePath)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourcePath))
            return JsonSerializer.Serialize(new { error = "queryName and sourcePath are required for update action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Update(excelPath, queryName, sourcePath).GetAwaiter().GetResult();
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string RefreshPowerQuery(PowerQueryCommands commands, string excelPath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for refresh action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Refresh(excelPath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string DeletePowerQuery(PowerQueryCommands commands, string excelPath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for delete action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Delete(excelPath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetLoadToTable(PowerQueryCommands commands, string excelPath, string? queryName, string? targetSheet)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for set-load-to-table action" }, ExcelToolsBase.JsonOptions);

        var result = commands.SetLoadToTable(excelPath, queryName, targetSheet ?? "");
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetLoadToDataModel(PowerQueryCommands commands, string excelPath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for set-load-to-data-model action" }, ExcelToolsBase.JsonOptions);

        var result = commands.SetLoadToDataModel(excelPath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetLoadToBoth(PowerQueryCommands commands, string excelPath, string? queryName, string? targetSheet)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for set-load-to-both action" }, ExcelToolsBase.JsonOptions);

        var result = commands.SetLoadToBoth(excelPath, queryName, targetSheet ?? "");
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetConnectionOnly(PowerQueryCommands commands, string excelPath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for set-connection-only action" }, ExcelToolsBase.JsonOptions);

        var result = commands.SetConnectionOnly(excelPath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string GetLoadConfig(PowerQueryCommands commands, string excelPath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            return JsonSerializer.Serialize(new { error = "queryName is required for get-load-config action" }, ExcelToolsBase.JsonOptions);

        var result = commands.GetLoadConfig(excelPath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}