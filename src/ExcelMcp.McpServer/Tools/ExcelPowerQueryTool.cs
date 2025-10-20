using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
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
[McpServerToolType]
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
        string? targetSheet = null,
        
        [RegularExpression("^(None|Private|Organizational|Public)$")]
        [Description("Privacy level for Power Query data combining (optional). If not specified and privacy error occurs, LLM must ask user to choose: None (least secure), Private (most secure), Organizational (internal data), or Public (public data)")]
        string? privacyLevel = null)
    {
        try
        {
            var powerQueryCommands = new PowerQueryCommands();

            // Parse privacy level if provided
            PowerQueryPrivacyLevel? parsedPrivacyLevel = null;
            if (!string.IsNullOrEmpty(privacyLevel))
            {
                if (Enum.TryParse<PowerQueryPrivacyLevel>(privacyLevel, ignoreCase: true, out var level))
                {
                    parsedPrivacyLevel = level;
                }
            }

            return action.ToLowerInvariant() switch
            {
                "list" => ListPowerQueries(powerQueryCommands, excelPath),
                "view" => ViewPowerQuery(powerQueryCommands, excelPath, queryName),
                "import" => ImportPowerQuery(powerQueryCommands, excelPath, queryName, sourcePath, parsedPrivacyLevel),
                "export" => ExportPowerQuery(powerQueryCommands, excelPath, queryName, targetPath),
                "update" => UpdatePowerQuery(powerQueryCommands, excelPath, queryName, sourcePath, parsedPrivacyLevel),
                "refresh" => RefreshPowerQuery(powerQueryCommands, excelPath, queryName),
                "delete" => DeletePowerQuery(powerQueryCommands, excelPath, queryName),
                "set-load-to-table" => SetLoadToTable(powerQueryCommands, excelPath, queryName, targetSheet, parsedPrivacyLevel),
                "set-load-to-data-model" => SetLoadToDataModel(powerQueryCommands, excelPath, queryName, parsedPrivacyLevel),
                "set-load-to-both" => SetLoadToBoth(powerQueryCommands, excelPath, queryName, targetSheet, parsedPrivacyLevel),
                "set-connection-only" => SetConnectionOnly(powerQueryCommands, excelPath, queryName),
                "get-load-config" => GetLoadConfig(powerQueryCommands, excelPath, queryName),
                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action '{action}'. Supported: list, view, import, export, update, refresh, delete, set-load-to-table, set-load-to-data-model, set-load-to-both, set-connection-only, get-load-config")
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

    private static string ListPowerQueries(PowerQueryCommands commands, string excelPath)
    {
        var result = commands.List(excelPath);
        
        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"list failed for '{excelPath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ViewPowerQuery(PowerQueryCommands commands, string excelPath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for view action");

        var result = commands.View(excelPath, queryName);
        
        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"view failed for '{excelPath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ImportPowerQuery(PowerQueryCommands commands, string excelPath, string? queryName, string? sourcePath, PowerQueryPrivacyLevel? privacyLevel)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("queryName and sourcePath are required for import action");

        var result = commands.Import(excelPath, queryName, sourcePath, privacyLevel).GetAwaiter().GetResult();
        
        // Return result as JSON (including PowerQueryPrivacyErrorResult if privacy error occurred)
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ExportPowerQuery(PowerQueryCommands commands, string excelPath, string? queryName, string? targetPath)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(targetPath))
            throw new ModelContextProtocol.McpException("queryName and targetPath are required for export action");

        var result = commands.Export(excelPath, queryName, targetPath).GetAwaiter().GetResult();
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string UpdatePowerQuery(PowerQueryCommands commands, string excelPath, string? queryName, string? sourcePath, PowerQueryPrivacyLevel? privacyLevel)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("queryName and sourcePath are required for update action");

        var result = commands.Update(excelPath, queryName, sourcePath, privacyLevel).GetAwaiter().GetResult();
        
        // Return result as JSON (including PowerQueryPrivacyErrorResult if privacy error occurred)
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string RefreshPowerQuery(PowerQueryCommands commands, string excelPath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for refresh action");

        var result = commands.Refresh(excelPath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string DeletePowerQuery(PowerQueryCommands commands, string excelPath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for delete action");

        var result = commands.Delete(excelPath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetLoadToTable(PowerQueryCommands commands, string excelPath, string? queryName, string? targetSheet, PowerQueryPrivacyLevel? privacyLevel)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for set-load-to-table action");

        var result = commands.SetLoadToTable(excelPath, queryName, targetSheet ?? "", privacyLevel);
        
        // Return result as JSON (including PowerQueryPrivacyErrorResult if privacy error occurred)
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetLoadToDataModel(PowerQueryCommands commands, string excelPath, string? queryName, PowerQueryPrivacyLevel? privacyLevel)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for set-load-to-data-model action");

        var result = commands.SetLoadToDataModel(excelPath, queryName, privacyLevel);
        
        // Return result as JSON (including PowerQueryPrivacyErrorResult if privacy error occurred)
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetLoadToBoth(PowerQueryCommands commands, string excelPath, string? queryName, string? targetSheet, PowerQueryPrivacyLevel? privacyLevel)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for set-load-to-both action");

        var result = commands.SetLoadToBoth(excelPath, queryName, targetSheet ?? "", privacyLevel);
        
        // Return result as JSON (including PowerQueryPrivacyErrorResult if privacy error occurred)
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetConnectionOnly(PowerQueryCommands commands, string excelPath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for set-connection-only action");

        var result = commands.SetConnectionOnly(excelPath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string GetLoadConfig(PowerQueryCommands commands, string excelPath, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for get-load-config action");

        var result = commands.GetLoadConfig(excelPath, queryName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}