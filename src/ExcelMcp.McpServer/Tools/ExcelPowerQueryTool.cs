using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel Power Query management tool for MCP server.
/// Handles M code operations, query management, and data loading configurations.
///
/// LLM Usage Patterns:
/// - Use "list" to see all Power Queries in a workbook
/// - Use "view" to examine M code for a specific query
/// - Use "import" to add new queries from .pq files (DEFAULT: auto-loads to worksheet for validation)
/// - Use "export" to save M code to files for version control
/// - Use "update" to modify existing query M code (preserves existing load configuration)
/// - Use "refresh" to refresh query data from source
/// - Use "delete" to remove queries
/// - Use "set-load-to-table" to load query data to worksheet (validates M code via execution)
/// - Use "set-load-to-data-model" to load to Excel's data model
/// - Use "set-load-to-both" to load to both table and data model
/// - Use "set-connection-only" to prevent data loading (M code not validated)
/// - Use "get-load-config" to check current loading configuration
///
/// IMPORTANT:
/// - Import DEFAULT behavior: Automatically loads to worksheet (validates M code by executing it)
/// - Validation = Execution: Power Query M code is only validated when data is actually loaded/refreshed
/// - Connection-only queries are NOT validated until first execution via set-load-to-table or refresh
/// </summary>
[McpServerToolType]
public static class ExcelPowerQueryTool
{
    /// <summary>
    /// Manage Power Query operations - M code, data loading, and query lifecycle
    /// </summary>
    [McpServerTool(Name = "excel_powerquery")]
    [Description("Manage Power Query M code and data loading. Supports: list, view, import, export, update, refresh, delete, set-load-to-table, set-load-to-data-model, set-load-to-both, set-connection-only, get-load-config. Optional batchId for batch sessions.")]
    public static async Task<string> ExcelPowerQuery(
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
        string? privacyLevel = null,
        
        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            // Create commands
            var dataModelCommands = new DataModelCommands();
            var powerQueryCommands = new PowerQueryCommands(dataModelCommands);
            var parameterCommands = new ParameterCommands();

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
                "list" => await ListPowerQueriesAsync(powerQueryCommands, excelPath, batchId),
                "view" => await ViewPowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
                "import" => await ImportPowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, parsedPrivacyLevel, batchId),
                "export" => await ExportPowerQueryAsync(powerQueryCommands, excelPath, queryName, targetPath, batchId),
                "update" => await UpdatePowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, parsedPrivacyLevel, batchId),
                "refresh" => await RefreshPowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
                "delete" => await DeletePowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
                "set-load-to-table" => await SetLoadToTableAsync(powerQueryCommands, excelPath, queryName, targetSheet, parsedPrivacyLevel, batchId),
                "set-load-to-data-model" => await SetLoadToDataModelAsync(powerQueryCommands, excelPath, queryName, parsedPrivacyLevel, batchId),
                "set-load-to-both" => await SetLoadToBothAsync(powerQueryCommands, excelPath, queryName, targetSheet, parsedPrivacyLevel, batchId),
                "set-connection-only" => await SetConnectionOnlyAsync(powerQueryCommands, excelPath, queryName, batchId),
                "get-load-config" => await GetLoadConfigAsync(powerQueryCommands, excelPath, queryName, batchId),
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

    private static async Task<string> ListPowerQueriesAsync(PowerQueryCommands commands, string excelPath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.ListAsync(batch));
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the Excel file exists and is accessible",
                "Verify the file path and try again"
            };
            result.WorkflowHint = "List failed. Ensure the file exists and retry.";
            throw new ModelContextProtocol.McpException($"list failed for '{excelPath}': {result.ErrorMessage}");
        }
        else
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'view' to inspect a query's M code",
                "Use 'import' to add a new Power Query",
                "Use 'delete' to remove a query"
            };
            result.WorkflowHint = "Power Queries listed. Next, view, import, or delete queries as needed.";
        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ViewPowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for view action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.ViewAsync(batch, queryName));
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the query name is correct",
                "Use 'list' to see available queries"
            };
            result.WorkflowHint = "View failed. Ensure the query exists and retry.";
            throw new ModelContextProtocol.McpException($"view failed for '{excelPath}': {result.ErrorMessage}");
        }
        else
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'update' to modify the query's M code",
                "Use 'set-load-to-table' to load data to worksheet",
                "Use 'refresh' to update query data"
            };
            result.WorkflowHint = "Query M code viewed. Next, update, load, or refresh as needed.";
        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ImportPowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? sourcePath, PowerQueryPrivacyLevel? privacyLevel, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("queryName and sourcePath are required for import action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.ImportAsync(batch, queryName, sourcePath, privacyLevel));

        // Always provide actionable next steps and workflow hint for LLM guidance
        if (result.Success)
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'set-load-to-table' to load query data to worksheet",
                "Use 'refresh' to validate the query and update data",
                "Use 'view' to inspect the imported M code"
            };
            result.WorkflowHint = "Query imported successfully. Next, load data to worksheet or refresh to validate.";
        }
        else
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the source M code file exists and is accessible",
                "Verify the file path and try again",
                "Use 'list' to see available queries"
            };
            result.WorkflowHint = "Import failed due to missing M code file. Ensure the file exists and retry.";
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ExportPowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? targetPath, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(targetPath))
            throw new ModelContextProtocol.McpException("queryName and targetPath are required for export action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.ExportAsync(batch, queryName, targetPath));
        if (result.Success)
        {
            result.SuggestedNextActions = new List<string>
            {
                "Edit the exported M code file as needed",
                "Use 'update' to re-import modified code",
                "Use 'refresh' to validate changes"
            };
            result.WorkflowHint = "Query exported. Next, edit and update as needed.";
        }
        else
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the target path is valid and writable",
                "Verify the query name and try again"
            };
            result.WorkflowHint = "Export failed. Ensure the path and query are correct.";
        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdatePowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? sourcePath, PowerQueryPrivacyLevel? privacyLevel, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("queryName and sourcePath are required for update action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.UpdateAsync(batch, queryName, sourcePath, privacyLevel));
        if (result.Success)
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'refresh' to validate the updated query",
                "Use 'view' to inspect the new M code",
                "Use 'set-load-to-table' to load updated data"
            };
            result.WorkflowHint = "Query updated. Next, refresh, view, or load data.";
        }
        else
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the source M code file exists and is accessible",
                "Verify the query name and try again"
            };
            result.WorkflowHint = "Update failed. Ensure the file and query are correct.";
        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshPowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for refresh action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.RefreshAsync(batch, queryName));
        if (result.Success)
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'view' to inspect the query M code",
                "Use worksheet 'read' to verify loaded data",
                "Use 'get-load-config' to check load settings"
            };
            result.WorkflowHint = "Query refreshed successfully. Next, view code or verify data.";
        }
        else
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check the query M code for errors using 'view'",
                "Verify data source connectivity",
                "Review privacy level settings if needed"
            };
            result.WorkflowHint = "Refresh failed. Check query code and data source.";
        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeletePowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for delete action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.DeleteAsync(batch, queryName));
        if (result.Success)
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'list' to verify query was removed",
                "Use 'import' to add a new query",
                "Review remaining queries with 'list'"
            };
            result.WorkflowHint = "Query deleted successfully. Next, verify with list.";
        }
        else
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the query name is correct",
                "Use 'list' to see available queries",
                "Verify the file is not read-only"
            };
            result.WorkflowHint = "Delete failed. Ensure the query exists and file is writable.";
        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetLoadToTableAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? targetSheet, PowerQueryPrivacyLevel? privacyLevel, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for set-load-to-table action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.SetLoadToTableAsync(batch, queryName, targetSheet ?? "", privacyLevel));
        if (result.Success)
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'refresh' to load data to the worksheet",
                "Use worksheet 'read' to verify loaded data",
                "Use 'get-load-config' to confirm load settings"
            };
            result.WorkflowHint = "Load-to-table configured. Next, refresh to load data.";
        }
        else
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the query exists using 'list'",
                "Verify the target sheet name is correct",
                "Review privacy level settings if needed"
            };
            result.WorkflowHint = "Set-load-to-table failed. Check query and sheet names.";
        }

        // Return result as JSON (including PowerQueryPrivacyErrorResult if privacy error occurred)
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetLoadToDataModelAsync(PowerQueryCommands commands, string excelPath, string? queryName, PowerQueryPrivacyLevel? privacyLevel, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for set-load-to-data-model action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.SetLoadToDataModelAsync(batch, queryName, privacyLevel));
        
        // Result now includes verification metrics: RowsLoaded, TablesInDataModel, WorkflowStatus
        // WorkflowHint and SuggestedNextActions are set by Core layer based on verification outcome
        // Do NOT overwrite these values - they reflect actual operation results
        
        // Return result as JSON with all verification details
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetLoadToBothAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? targetSheet, PowerQueryPrivacyLevel? privacyLevel, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for set-load-to-both action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.SetLoadToBothAsync(batch, queryName, targetSheet ?? "", privacyLevel));
        if (result.Success)
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'refresh' to load data to both worksheet and data model",
                "Use worksheet 'read' to verify worksheet data",
                "Use 'get-load-config' to confirm dual-load settings"
            };
            result.WorkflowHint = "Load-to-both configured. Next, refresh to load data.";
        }
        else
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the query exists using 'list'",
                "Verify the target sheet name is correct",
                "Review privacy level settings if needed"
            };
            result.WorkflowHint = "Set-load-to-both failed. Check query and sheet names.";
        }

        // Return result as JSON (including PowerQueryPrivacyErrorResult if privacy error occurred)
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetConnectionOnlyAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for set-connection-only action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.SetConnectionOnlyAsync(batch, queryName));
        if (result.Success)
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'get-load-config' to confirm connection-only setting",
                "Use 'set-load-to-table' to load data to worksheet later",
                "Use 'view' to inspect the query M code"
            };
            result.WorkflowHint = "Connection-only configured. Query will not load data automatically.";
        }
        else
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the query exists using 'list'",
                "Verify the file is not read-only"
            };
            result.WorkflowHint = "Set-connection-only failed. Ensure the query exists.";
        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetLoadConfigAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for get-load-config action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.GetLoadConfigAsync(batch, queryName));
        if (result.Success)
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'set-load-to-table' to change load settings",
                "Use 'refresh' to update query data",
                "Use 'view' to inspect the query M code"
            };
            result.WorkflowHint = "Load configuration retrieved. Next, modify settings or refresh data.";
        }
        else
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the query exists using 'list'",
                "Verify the query name is correct"
            };
            result.WorkflowHint = "Get-load-config failed. Ensure the query exists.";
        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
