using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Security;
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
        string? batchId = null,

        [Description("Timeout in minutes for Power Query operations. Default: 5 minutes for refresh operations, 2 minutes for others")]
        double? timeout = null)
    {
        try
        {
            // Create commands
            var dataModelCommands = new DataModelCommands();
            var powerQueryCommands = new PowerQueryCommands(dataModelCommands);

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                PowerQueryAction.List => await ListPowerQueriesAsync(powerQueryCommands, excelPath, batchId),
                PowerQueryAction.View => await ViewPowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
                PowerQueryAction.Import => await ImportPowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, loadDestination, batchId),
                PowerQueryAction.Export => await ExportPowerQueryAsync(powerQueryCommands, excelPath, queryName, targetPath, batchId),
                PowerQueryAction.Update => await UpdatePowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, loadDestination, batchId),
                PowerQueryAction.Refresh => await RefreshPowerQueryAsync(powerQueryCommands, excelPath, queryName, loadDestination, targetSheet, timeout, batchId),
                PowerQueryAction.Delete => await DeletePowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
                PowerQueryAction.SetLoadToTable => await SetLoadToTableAsync(powerQueryCommands, excelPath, queryName, targetSheet, batchId),
                PowerQueryAction.SetLoadToDataModel => await SetLoadToDataModelAsync(powerQueryCommands, excelPath, queryName, batchId),
                PowerQueryAction.SetLoadToBoth => await SetLoadToBothAsync(powerQueryCommands, excelPath, queryName, targetSheet, batchId),
                PowerQueryAction.SetConnectionOnly => await SetConnectionOnlyAsync(powerQueryCommands, excelPath, queryName, batchId),
                PowerQueryAction.GetLoadConfig => await GetLoadConfigAsync(powerQueryCommands, excelPath, queryName, batchId),
                PowerQueryAction.Errors => await ErrorsPowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
                PowerQueryAction.LoadTo => await LoadToPowerQueryAsync(powerQueryCommands, excelPath, queryName, targetSheet, batchId),
                PowerQueryAction.ListExcelSources => await ListExcelSourcesAsync(powerQueryCommands, excelPath, batchId),
                PowerQueryAction.Eval => await EvalPowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, batchId),
                
                // Phase 1: Atomic Operations
                PowerQueryAction.Create => await CreatePowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, loadDestination, targetSheet, batchId),
                PowerQueryAction.UpdateMCode => await UpdateMCodePowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, batchId),
                PowerQueryAction.Unload => await UnloadPowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
                PowerQueryAction.ValidateSyntax => await ValidateSyntaxPowerQueryAsync(powerQueryCommands, excelPath, sourcePath, batchId),
                PowerQueryAction.UpdateAndRefresh => await UpdateAndRefreshPowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, batchId),
                PowerQueryAction.RefreshAll => await RefreshAllPowerQueriesAsync(powerQueryCommands, excelPath, batchId),
                
                _ => throw new ModelContextProtocol.McpException($"Unknown action: {action} ({action.ToActionString()})")
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
        // Always return JSON (success or failure) - MCP clients handle the success flag
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

        // Always return JSON (success or failure) - MCP clients handle the success flag
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

        // Always return JSON (success or failure) - MCP clients handle the success flag
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

        // Always return JSON (success or failure) - MCP clients handle the success flag
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

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshPowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? loadDestination, string? targetSheet, double? timeoutMinutes, string? batchId)
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

            // If load configuration failed, throw exception
            if (loadResult != null && !loadResult.Success)
            {
                throw new ModelContextProtocol.McpException($"Failed to apply load configuration '{destination}' for query '{queryName}' in '{excelPath}': {loadResult.ErrorMessage}");
            }

            // Load configuration applied successfully, now the refresh will use it
            // Continue to refresh below (SetLoadToTable/DataModel already refreshes, but we'll ensure it's fresh)
        }

        try
        {
            // Apply operation-specific timeout default (5 minutes for refresh)
            var timeoutSpan = timeoutMinutes.HasValue ? (TimeSpan?)TimeSpan.FromMinutes(timeoutMinutes.Value) : null;
            
            var result = await ExcelToolsBase.WithBatchAsync(
                batchId,
                excelPath,
                save: true,
                async (batch) => await commands.RefreshAsync(batch, queryName, timeoutSpan));

            // Always return JSON (success or failure) - MCP clients handle the success flag
            return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
        }
        catch (TimeoutException ex)
        {
            // Enrich timeout error with operation-specific guidance
            var result = new PowerQueryRefreshResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                QueryName = queryName,
                FilePath = excelPath,
                RefreshTime = DateTime.Now,

                SuggestedNextActions = new List<string>
                {
                    "Check if Excel is showing a 'Privacy Level' dialog or credential prompt",
                    "Verify the data source is accessible (network connection, database availability)",
                    "For large datasets, consider filtering data at source or breaking into smaller queries",
                    "Use batch mode (begin_excel_batch) if not already using it to optimize multiple operations"
                },

                OperationContext = new Dictionary<string, object>
                {
                    { "OperationType", "PowerQuery.Refresh" },
                    { "QueryName", queryName },
                    { "TimeoutReached", true },
                    { "UsedMaxTimeout", ex.Message.Contains("maximum timeout") }
                },

                IsRetryable = !ex.Message.Contains("maximum timeout"),

                RetryGuidance = ex.Message.Contains("maximum timeout")
                    ? "Operation reached maximum timeout (5 minutes). Do not retry automatically - manual intervention needed to check Excel state and data source."
                    : "Operation can be retried if transient data source issue suspected. Current timeout is already at maximum (5 minutes)."
            };

            return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
        }
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

        // Always return JSON (success or failure) - MCP clients handle the success flag
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

        // Always return JSON (success or failure) - MCP clients handle the success flag
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

        // Always return JSON (success or failure) - MCP clients handle the success flag
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

        // Always return JSON (success or failure) - MCP clients handle the success flag
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

        // Always return JSON (success or failure) - MCP clients handle the success flag
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

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ErrorsPowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for errors action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.ErrorsAsync(batch, queryName));
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> LoadToPowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? targetSheet, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(targetSheet))
            throw new ModelContextProtocol.McpException("queryName and targetSheet are required for load-to action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.LoadToAsync(batch, queryName, targetSheet));
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListExcelSourcesAsync(PowerQueryCommands commands, string excelPath, string? batchId)
    {
        // list-excel-sources action lists all available sources (doesn't require queryName)
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.ListExcelSourcesAsync(batch));
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> EvalPowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, string? sourcePath, string? batchId)
    {
        if (string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("sourcePath is required for eval action (M code file to evaluate)");

        // Validate and read M code from file
        string mExpression;
        try
        {
            sourcePath = PathValidator.ValidateExistingFile(sourcePath, nameof(sourcePath));
            mExpression = File.ReadAllText(sourcePath);
        }
        catch (Exception ex)
        {
            throw new ModelContextProtocol.McpException($"Failed to read M code from '{sourcePath}': {ex.Message}");
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.EvalAsync(batch, mExpression));
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // =========================================================================
    // PHASE 1 HANDLERS - Atomic Operations
    // =========================================================================

    private static async Task<string> CreatePowerQueryAsync(
        PowerQueryCommands commands, 
        string excelPath, 
        string? queryName, 
        string? sourcePath,
        string? loadDestination,
        string? targetSheet,
        string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for create action");
        if (string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("sourcePath is required for create action (.pq file)");

        sourcePath = PathValidator.ValidateExistingFile(sourcePath, nameof(sourcePath));
        
        // Parse loadDestination to PowerQueryLoadMode enum
        var loadMode = ParseLoadMode(loadDestination ?? "worksheet");
        
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.CreateAsync(batch, queryName, sourcePath, loadMode, targetSheet));
        
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateMCodePowerQueryAsync(
        PowerQueryCommands commands,
        string excelPath,
        string? queryName,
        string? sourcePath,
        string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for update-mcode action");
        if (string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("sourcePath is required for update-mcode action (.pq file)");

        sourcePath = PathValidator.ValidateExistingFile(sourcePath, nameof(sourcePath));
        
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.UpdateMCodeAsync(batch, queryName, sourcePath));
        
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UnloadPowerQueryAsync(
        PowerQueryCommands commands,
        string excelPath,
        string? queryName,
        string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for unload action");
        
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.UnloadAsync(batch, queryName));
        
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ValidateSyntaxPowerQueryAsync(
        PowerQueryCommands commands,
        string excelPath,
        string? sourcePath,
        string? batchId)
    {
        if (string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("sourcePath is required for validate-syntax action (.pq file)");

        sourcePath = PathValidator.ValidateExistingFile(sourcePath, nameof(sourcePath));
        
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,  // Validation doesn't modify workbook permanently
            async (batch) => await commands.ValidateSyntaxAsync(batch, sourcePath));
        
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateAndRefreshPowerQueryAsync(
        PowerQueryCommands commands,
        string excelPath,
        string? queryName,
        string? sourcePath,
        string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for update-and-refresh action");
        if (string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("sourcePath is required for update-and-refresh action (.pq file)");

        sourcePath = PathValidator.ValidateExistingFile(sourcePath, nameof(sourcePath));
        
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.UpdateAndRefreshAsync(batch, queryName, sourcePath));
        
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshAllPowerQueriesAsync(
        PowerQueryCommands commands,
        string excelPath,
        string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.RefreshAllAsync(batch));
        
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static PowerQueryLoadMode ParseLoadMode(string loadDestination)
    {
        return loadDestination.ToLowerInvariant() switch
        {
            "worksheet" => PowerQueryLoadMode.LoadToTable,
            "data-model" => PowerQueryLoadMode.LoadToDataModel,
            "both" => PowerQueryLoadMode.LoadToBoth,
            "connection-only" => PowerQueryLoadMode.ConnectionOnly,
            _ => throw new ModelContextProtocol.McpException($"Invalid loadDestination: '{loadDestination}'. Valid values: worksheet, data-model, both, connection-only")
        };
    }
}

