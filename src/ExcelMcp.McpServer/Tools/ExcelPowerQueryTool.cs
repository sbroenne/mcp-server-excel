using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel Power Query management tool for MCP server.
/// Handles M code operations, query management, and data loading configurations.
///
/// LLM Usage Patterns:
/// - Use "list" to see all Power Queries in a workbook
/// - Use "view" to examine M code for a specific query
/// - Use "import" to add new queries from .pq files (use loadDestination parameter: worksheet|data-model|both|connection-only)
/// - Use "export" to save M code to files for version control
/// - Use "update" to modify existing query M code (preserves existing load configuration)
/// - Use "refresh" to refresh query data from source (optionally specify loadDestination to apply load config while refreshing)
/// - Use "delete" to remove queries
/// - Use "set-load-to-table" to load query data to worksheet (visible to users, NOT in Data Model)
/// - Use "set-load-to-data-model" to load to Excel's Power Pivot Data Model (ready for DAX measures)
/// - Use "set-load-to-both" to load to BOTH worksheet AND Power Pivot Data Model
/// - Use "set-connection-only" to prevent data loading (M code not validated)
/// - Use "get-load-config" to check current loading configuration
///
/// REFRESH WITH LOAD DESTINATION (NEW):
/// - Refresh connection-only query AND apply load config in one call: refresh(loadDestination: 'worksheet')
/// - Eliminates need for two calls: set-load-to-table + refresh
/// - Example: excel_powerquery(action: 'refresh', queryName: 'Sales', loadDestination: 'data-model')
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
    [Description(@"Manage Power Query M code and data loading.

⚡ PERFORMANCE: For 2+ operations on same file, use begin_excel_batch FIRST (75-90% faster):
  1. batch = begin_excel_batch(excelPath: 'file.xlsx')
  2. excel_powerquery(..., batchId: batch.batchId)  // repeat for each operation
  3. commit_excel_batch(batchId: batch.batchId, save: true)

PRIMARY ACTIONS:
- Import: Add Power Query from .pq file (use loadDestination parameter for data model workflows)
- SetLoadToDataModel: Load query data to Power Pivot Data Model (ready for DAX measures)
- SetLoadToTable: Load query data to worksheet (visible to users, NOT in data model)
- SetLoadToBoth: Load to BOTH worksheet AND data model

After loading to Data Model, use excel_datamodel tool for DAX measures and relationships.")]
    public static async Task<string> ExcelPowerQuery(
        [Required]
        [Description("Action to perform (enum values displayed as dropdown in MCP clients)")]
        PowerQueryAction action,

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
        [Description("Target worksheet name (when loadDestination is 'worksheet' or 'both', or for set-load-to-table action)")]
        string? targetSheet = null,

        [RegularExpression("^(worksheet|data-model|both|connection-only)$")]
        [Description(@"Load destination for query (for import/refresh actions). Options:
  - 'worksheet': Load to worksheet as table (DEFAULT - users can see/validate data)
  - 'data-model': Load to Power Pivot Data Model (for DAX measures/relationships)
  - 'both': Load to both worksheet AND Data Model
  - 'connection-only': Don't load data (M code imported but not executed)
For import: DEFAULT is 'worksheet'. For refresh: applies load config if query is connection-only.")]
        string? loadDestination = null,

        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            // Create commands
            var dataModelCommands = new DataModelCommands();
            var powerQueryCommands = new PowerQueryCommands(dataModelCommands);
            var NamedRangeCommands = new NamedRangeCommands();

            // Convert enum to action string
            var actionString = action.ToActionString();
            
            return actionString switch
            {
                "list" => await ListPowerQueriesAsync(powerQueryCommands, excelPath, batchId),
                "view" => await ViewPowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
                "import" => await ImportPowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, loadDestination, batchId),
                "export" => await ExportPowerQueryAsync(powerQueryCommands, excelPath, queryName, targetPath, batchId),
                "update" => await UpdatePowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, loadDestination, batchId),
                "refresh" => await RefreshPowerQueryAsync(powerQueryCommands, excelPath, queryName, loadDestination, targetSheet, batchId),
                "delete" => await DeletePowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
                "set-load-to-table" => await SetLoadToTableAsync(powerQueryCommands, excelPath, queryName, targetSheet, batchId),
                "set-load-to-data-model" => await SetLoadToDataModelAsync(powerQueryCommands, excelPath, queryName, batchId),
                "set-load-to-both" => await SetLoadToBothAsync(powerQueryCommands, excelPath, queryName, targetSheet, batchId),
                "set-connection-only" => await SetConnectionOnlyAsync(powerQueryCommands, excelPath, queryName, batchId),
                "get-load-config" => await GetLoadConfigAsync(powerQueryCommands, excelPath, queryName, batchId),
                _ => throw new ModelContextProtocol.McpException($"Unknown action '{actionString}'")
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

    private static async Task<string> ListPowerQueriesAsync(PowerQueryCommands commands, string excelPath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.ListAsync(batch));
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {

            throw new ModelContextProtocol.McpException($"list failed for '{excelPath}': {result.ErrorMessage}");
        }
        else
        {

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

            throw new ModelContextProtocol.McpException($"view failed for '{excelPath}': {result.ErrorMessage}");
        }
        else
        {

        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ImportPowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? sourcePath, string? loadDestination, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("queryName and sourcePath are required for import action");

        // Default to "worksheet" if not specified
        string destination = loadDestination ?? "worksheet";

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.ImportAsync(batch, queryName, sourcePath, destination));

        // Core already sets appropriate workflow guidance based on actual load outcome
        // Only enhance guidance if in batch mode
        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        if (result.Success && usedBatchMode)
        {
            // Enhance guidance for batch mode (Core doesn't know about batch mode)
            bool isConnectionOnly = destination.ToLowerInvariant() == "connection-only";

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

        }
        else
        {

        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdatePowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? sourcePath, string? loadDestination, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("queryName and sourcePath are required for update action");

        // Note: loadDestination parameter is ignored for update - Core preserves existing load configuration
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.UpdateAsync(batch, queryName, sourcePath));

        // Use workflow guidance with batch mode awareness
        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        if (result.Success)
        {

        }
        else
        {

        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshPowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? loadDestination, string? targetSheet, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for refresh action");

        // If loadDestination is specified, apply load configuration first
        if (!string.IsNullOrEmpty(loadDestination))
        {
            var destination = loadDestination.ToLowerInvariant();
            OperationResult? loadResult = null;

            // Apply load configuration before refresh
            switch (destination)
            {
                case "worksheet":
                    string sheet = targetSheet ?? queryName;
                    loadResult = await ExcelToolsBase.WithBatchAsync(
                        batchId,
                        excelPath,
                        save: true,
                        async (batch) => await commands.SetLoadToTableAsync(batch, queryName, sheet));
                    break;

                case "data-model":
                    var dmResult = await ExcelToolsBase.WithBatchAsync(
                        batchId,
                        excelPath,
                        save: true,
                        async (batch) => await commands.SetLoadToDataModelAsync(batch, queryName));
                    loadResult = new OperationResult
                    {
                        Success = dmResult.Success,
                        ErrorMessage = dmResult.ErrorMessage,
                        FilePath = dmResult.FilePath
                    };
                    break;

                case "both":
                    string sheetBoth = targetSheet ?? queryName;
                    var bothResult = await ExcelToolsBase.WithBatchAsync(
                        batchId,
                        excelPath,
                        save: true,
                        async (batch) => await commands.SetLoadToBothAsync(batch, queryName, sheetBoth));
                    loadResult = new OperationResult
                    {
                        Success = bothResult.Success,
                        ErrorMessage = bothResult.ErrorMessage,
                        FilePath = bothResult.FilePath
                    };
                    break;

                case "connection-only":
                    loadResult = await ExcelToolsBase.WithBatchAsync(
                        batchId,
                        excelPath,
                        save: true,
                        async (batch) => await commands.SetConnectionOnlyAsync(batch, queryName));
                    break;
            }

            // If load configuration failed, return error
            if (loadResult != null && !loadResult.Success)
            {
                var errorResult = new
                {
                    success = false,
                    errorMessage = $"Failed to apply load configuration '{destination}': {loadResult.ErrorMessage}",
                    filePath = excelPath,
                    queryName,
                    refreshTime = DateTime.Now,
                    suggestedNextActions = new[]
                    {
                        "Check that the query exists using 'list'",
                        "Use 'view' to verify query M code is valid",
                        $"Try 'set-load-to-{destination}' separately to diagnose the issue"
                    },
                    workflowHint = $"Failed to configure query to load to {destination}"
                };
                return JsonSerializer.Serialize(errorResult, ExcelToolsBase.JsonOptions);
            }

            // Load configuration applied successfully, now the refresh will use it
            // Continue to refresh below (SetLoadToTable/DataModel already refreshes, but we'll ensure it's fresh)
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.RefreshAsync(batch, queryName));
        if (result.Success)
        {
            // Update workflow hints based on whether loadDestination was applied
            bool loadDestinationApplied = !string.IsNullOrEmpty(loadDestination);

        }
        else
        {

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

        }
        else
        {

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

        }
        else
        {

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

        }
        else
        {

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

        }
        else
        {

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

        }
        else
        {

        }
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
