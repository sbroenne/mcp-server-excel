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
/// - Use "set-load-to-table" to load query data to worksheet (visible to users, NOT in Data Model)
/// - Use "set-load-to-data-model" to load to Excel's Power Pivot Data Model (ready for DAX measures)
/// - Use "set-load-to-both" to load to BOTH worksheet AND Power Pivot Data Model
/// - Use "set-connection-only" to prevent data loading (M code not validated)
/// - Use "get-load-config" to check current loading configuration
///
/// IMPORTANT FOR DATA MODEL WORKFLOWS:
/// - "set-load-to-table" loads data to WORKSHEET ONLY (users see formatted table, but NOT in Power Pivot)
/// - For Data Model/DAX workflows: Use "set-load-to-data-model" or "set-load-to-both" actions
/// - Cannot directly add worksheet-only query to Data Model via excel_table tool
/// - If query is already loaded to worksheet: Use set-load-to-data-model to add to Power Pivot
///
/// VALIDATION & EXECUTION:
/// - Import DEFAULT behavior: Automatically loads to worksheet (validates M code by executing it)
/// - Validation = Execution: Power Query M code is only validated when data is actually loaded/refreshed
/// - Connection-only queries are NOT validated until first execution via set-load-to-table or refresh
/// - For Power Pivot operations beyond loading data (DAX measures, relationships), use excel_datamodel or excel_powerpivot tools
/// </summary>
[McpServerToolType]
public static class ExcelPowerQueryTool
{
    /// <summary>
    /// Manage Power Query operations - M code, data loading, and query lifecycle
    /// </summary>
    [McpServerTool(Name = "excel_powerquery")]
    [Description("Manage Power Query M code and data loading. Primary tool for loading data into Power Pivot: use 'set-load-to-data-model' action to add data to Power Pivot (Data Model). Supports: list, view, import, export, update, refresh, delete, set-load-to-table, set-load-to-data-model, set-load-to-both, set-connection-only, get-load-config. After loading data to Power Pivot, use excel_datamodel or excel_powerpivot tools for DAX measures and relationships. Optional batchId for batch sessions.")]
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

        [Description("Automatically load query data to worksheet for validation (default: true). When false, creates connection-only query without validation.")]
        bool? loadToWorksheet = null,

        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            // Create commands
            var dataModelCommands = new DataModelCommands();
            var powerQueryCommands = new PowerQueryCommands(dataModelCommands);
            var parameterCommands = new ParameterCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => await ListPowerQueriesAsync(powerQueryCommands, excelPath, batchId),
                "view" => await ViewPowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
                "import" => await ImportPowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, loadToWorksheet, batchId),
                "export" => await ExportPowerQueryAsync(powerQueryCommands, excelPath, queryName, targetPath, batchId),
                "update" => await UpdatePowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, loadToWorksheet, batchId),
                "refresh" => await RefreshPowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
                "delete" => await DeletePowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
                "set-load-to-table" => await SetLoadToTableAsync(powerQueryCommands, excelPath, queryName, targetSheet, batchId),
                "set-load-to-data-model" => await SetLoadToDataModelAsync(powerQueryCommands, excelPath, queryName, batchId),
                "set-load-to-both" => await SetLoadToBothAsync(powerQueryCommands, excelPath, queryName, targetSheet, batchId),
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
            result.SuggestedNextActions =
            [
                "Check that the Excel file exists and is accessible",
                "Verify the file path and try again"
            ];
            result.WorkflowHint = "List failed. Ensure the file exists and retry.";
            throw new ModelContextProtocol.McpException($"list failed for '{excelPath}': {result.ErrorMessage}");
        }
        else
        {
            result.SuggestedNextActions =
            [
                "Use 'view' to inspect a query's M code",
                "Use 'import' to add a new Power Query",
                "Use 'delete' to remove a query"
            ];
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
            result.SuggestedNextActions =
            [
                "Check that the query name is correct",
                "Use 'list' to see available queries"
            ];
            result.WorkflowHint = "View failed. Ensure the query exists and retry.";
            throw new ModelContextProtocol.McpException($"view failed for '{excelPath}': {result.ErrorMessage}");
        }
        else
        {
            result.SuggestedNextActions =
            [
                "Use 'update' to modify the query's M code",
                "Use 'get-load-config' to check if query is loaded anywhere",
                "Use 'set-load-to-table' to load data to worksheet"
            ];
            result.WorkflowHint = "Query M code viewed. Check load config or load to destination.";
        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ImportPowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? sourcePath, bool? loadToWorksheet, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("queryName and sourcePath are required for import action");

        // Default to true if not specified (auto-load for validation)
        bool shouldLoad = loadToWorksheet ?? true;

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.ImportAsync(batch, queryName, sourcePath, shouldLoad));

        // Core already sets appropriate workflow guidance based on actual load outcome
        // Only enhance guidance if in batch mode
        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        if (result.Success && usedBatchMode)
        {
            // Enhance guidance for batch mode (Core doesn't know about batch mode)
            bool isConnectionOnly = !shouldLoad;

            result.SuggestedNextActions = PowerQueryWorkflowGuidance.GetNextStepsAfterImport(
                isConnectionOnly: isConnectionOnly,
                hasErrors: false,
                usedBatchMode: usedBatchMode);

            result.WorkflowHint = isConnectionOnly
                ? "Query imported as connection-only in batch mode. Use set-load-to-table to load data or continue adding operations."
                : "Query imported in batch mode. Continue adding operations to this batch.";
        }
        // Otherwise, Core's guidance is already correct (for both success and failure cases) - don't overwrite it!

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
            result.SuggestedNextActions =
            [
                "Edit the exported M code file as needed",
                "Use 'update' to re-import modified code",
                "Use 'get-load-config' to check current load configuration"
            ];
            result.WorkflowHint = "Query exported. Edit file, then use 'update' to apply changes.";
        }
        else
        {
            result.SuggestedNextActions =
            [
                "Check that the target path is valid and writable",
                "Verify the query name and try again"
            ];
            result.WorkflowHint = "Export failed. Ensure the path and query are correct.";
        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdatePowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? sourcePath, bool? loadToWorksheet, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("queryName and sourcePath are required for update action");

        // Note: loadToWorksheet parameter is ignored for update - Core preserves existing load configuration
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.UpdateAsync(batch, queryName, sourcePath));

        // Use workflow guidance with batch mode awareness
        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        if (result.Success)
        {
            result.SuggestedNextActions = PowerQueryWorkflowGuidance.GetNextStepsAfterUpdate(
                configPreserved: true,
                hasErrors: false,
                usedBatchMode: usedBatchMode);

            result.WorkflowHint = usedBatchMode
                ? "Query updated in batch mode. Configuration preserved. Continue with more operations."
                : "Query updated successfully. Configuration preserved. For multiple updates, use begin_excel_batch.";
        }
        else
        {
            result.SuggestedNextActions =
            [
                "Check that the source M code file exists and is accessible",
                "Verify the query name and try again"
            ];
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
            result.SuggestedNextActions =
            [
                "Use 'view' to inspect the query M code",
                "Use worksheet 'read' to verify loaded data",
                "Use 'get-load-config' to check load settings"
            ];
            result.WorkflowHint = "Query refreshed successfully. Next, view code or verify data.";
        }
        else
        {
            result.SuggestedNextActions =
            [
                "Check the query M code for errors using 'view'",
                "Verify data source connectivity",
                "Review privacy level settings if needed"
            ];
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
            result.SuggestedNextActions =
            [
                "Use 'list' to verify query was removed",
                "Use 'import' to add a new query",
                "Review remaining queries with 'list'"
            ];
            result.WorkflowHint = "Query deleted successfully. Next, verify with list.";
        }
        else
        {
            result.SuggestedNextActions =
            [
                "Check that the query name is correct",
                "Use 'list' to see available queries",
                "Verify the file is not read-only"
            ];
            result.WorkflowHint = "Delete failed. Ensure the query exists and file is writable.";
        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetLoadToTableAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? targetSheet, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for set-load-to-table action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.SetLoadToTableAsync(batch, queryName, targetSheet ?? ""));

        // Use workflow guidance with batch mode awareness
        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        if (result.Success)
        {
            result.SuggestedNextActions = PowerQueryWorkflowGuidance.GetNextStepsAfterLoadConfig(
                loadMode: "LoadToTable",
                usedBatchMode: usedBatchMode);

            result.WorkflowHint = usedBatchMode
                ? "Load-to-table configured in batch mode. Continue configuring other queries."
                : "Load-to-table configured. For configuring multiple queries, use begin_excel_batch.";
        }
        else
        {
            result.SuggestedNextActions =
            [
                "Check that the query exists using 'list'",
                "Verify the target sheet name is correct",
                "Review privacy level settings if needed"
            ];
            result.WorkflowHint = "Set-load-to-table failed. Check query and sheet names.";
        }

        // Return result as JSON (including PowerQueryPrivacyErrorResult if privacy error occurred)
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetLoadToDataModelAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for set-load-to-data-model action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.SetLoadToDataModelAsync(batch, queryName));

        // Use workflow guidance with batch mode awareness
        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        if (result.Success)
        {
            result.SuggestedNextActions = PowerQueryWorkflowGuidance.GetNextStepsAfterLoadConfig(
                loadMode: "LoadToDataModel",
                usedBatchMode: usedBatchMode);

            result.WorkflowHint = usedBatchMode
                ? "Load-to-data-model configured in batch mode. Continue configuring other queries."
                : "Load-to-data-model configured. For configuring multiple queries, use begin_excel_batch.";
        }
        else
        {
            result.SuggestedNextActions =
            [
                "Check that the query exists using 'list'",
                "Review query M code for errors using 'view'",
                "Verify query data loads successfully with 'refresh'"
            ];
            result.WorkflowHint = "Set-load-to-data-model failed. Check query exists and has valid M code.";
        }

        // Return result as JSON (including PowerQueryPrivacyErrorResult if privacy error occurred)
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetLoadToBothAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? targetSheet, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for set-load-to-both action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.SetLoadToBothAsync(batch, queryName, targetSheet ?? ""));

        // Result now includes dual atomic operation verification metrics:
        // RowsLoadedToTable, RowsLoadedToModel, TablesInDataModel, WorkflowStatus (Complete/Partial/Failed)
        // DataLoadedToTable, DataLoadedToModel, ConfigurationApplied
        // WorkflowHint and SuggestedNextActions are set by Core layer based on verification outcome
        // Do NOT overwrite these values - they reflect actual operation results

        // Return result as JSON with all verification details
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
            result.SuggestedNextActions =
            [
                "Use 'get-load-config' to confirm connection-only setting",
                "Use 'set-load-to-table' to load data to worksheet later",
                "Use 'view' to inspect the query M code"
            ];
            result.WorkflowHint = "Connection-only configured. Query will not load data automatically.";
        }
        else
        {
            result.SuggestedNextActions =
            [
                "Check that the query exists using 'list'",
                "Verify the file is not read-only"
            ];
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
            result.SuggestedNextActions =
            [
                "Use 'set-load-to-table' or 'set-load-to-data-model' to change load destination",
                "Use 'refresh' to update data (only works if loaded to table/model)",
                "Use 'view' to inspect the query M code"
            ];
            result.WorkflowHint = "Load configuration retrieved. Modify settings or refresh if already loaded.";
        }
        else
        {
            result.SuggestedNextActions =
            [
                "Check that the query exists using 'list'",
                "Verify the query name is correct"
            ];
            result.WorkflowHint = "Get-load-config failed. Ensure the query exists.";
        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
