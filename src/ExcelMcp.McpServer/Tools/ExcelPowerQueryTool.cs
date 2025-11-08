using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Security;
using Sbroenne.ExcelMcp.McpServer.Models;

#pragma warning disable CA1861 // Avoid constant arrays as arguments - workflow hints are contextual per-call

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel Power Query management tool for MCP server.
/// Handles M code operations, query management, and data loading configurations.
///
/// LLM Usage Patterns:
/// - Use "list" to see all Power Queries in a workbook
/// - Use "view" to examine M code for a specific query
/// - Use "create" to add new queries from .pq files (atomic: import + load in one call)
/// - Use "export" to save M code to files for version control
/// - Use "update-mcode" to modify M code only (no refresh)
/// - Use "update-and-refresh" to update M code and refresh data atomically
/// - Use "refresh" to refresh query data from source
/// - Use "unload" to convert query to connection-only (inverse of load-to)
/// - Use "refresh-all" to refresh all queries in workbook
/// - Use "delete" to remove queries
/// - Use "get-load-config" to check current loading configuration
///
/// ATOMIC OPERATIONS:
/// - create: Import + load in one atomic operation (replaces import + load-to)
/// - update-mcode: Update M code without refresh (for staging changes)
/// - update-and-refresh: Update M code + refresh in one atomic operation
/// - unload: Convert to connection-only (inverse of load-to)
/// - refresh-all: Refresh all queries in workbook
///
/// IMPORTANT FOR DATA MODEL WORKFLOWS:
/// - "create" with loadDestination='data-model' loads to Power Pivot Data Model (ready for DAX measures)
/// - "create" with loadDestination='worksheet' loads to worksheet (users see formatted table)
/// - "create" with loadDestination='both' loads to BOTH worksheet AND Power Pivot
/// - For Power Pivot operations beyond loading data (DAX measures, relationships), use excel_datamodel or excel_powerpivot tools
///
/// VALIDATION AND EXECUTION:
/// - Create DEFAULT behavior: Automatically loads to worksheet (validates M code by executing it)
/// - Validation = Execution: Power Query M code is only validated when data is actually loaded/refreshed
/// - Connection-only queries are NOT validated until first execution
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

LOAD DESTINATIONS (loadDestination parameter):
- 'worksheet': Load to worksheet as table (DEFAULT - users can see/validate data)
- 'data-model': Load to Power Pivot Data Model (ready for DAX measures/relationships)
- 'both': Load to BOTH worksheet AND Data Model
- 'connection-only': Don't load data (M code imported but not executed)

OPERATIONS GUIDANCE:
- Create: Import M code from .pq file AND optionally load data in ONE operation
- UpdateMCode: Update M code ONLY (no refresh) - use when staging changes before refresh
- UpdateAndRefresh: Update M code AND refresh data in ONE operation
- LoadTo: Apply load destination to connection-only query (make it load data to a worksheet or Data Model)
- Unload: Convert query to connection-only (remove data, keep M code definition)
- RefreshAll: Refresh ALL Power Queries in workbook (batch refresh)

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
        [Description("Target worksheet name (when loadDestination is 'worksheet' or 'both')")]
        string? targetSheet = null,

        [RegularExpression("^(worksheet|data-model|both|connection-only)$")]
        [Description(@"Load destination for query (for create action). Options:
  - 'worksheet': Load to worksheet as table (DEFAULT - users can see/validate data)
  - 'data-model': Load to Power Pivot Data Model (for DAX measures/relationships)
  - 'both': Load to both worksheet AND Data Model
  - 'connection-only': Don't load data (M code imported but not executed)")]
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
                PowerQueryAction.Export => await ExportPowerQueryAsync(powerQueryCommands, excelPath, queryName, targetPath, batchId),
                PowerQueryAction.Refresh => await RefreshPowerQueryAsync(powerQueryCommands, excelPath, queryName, timeout, batchId),
                PowerQueryAction.Delete => await DeletePowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
                PowerQueryAction.GetLoadConfig => await GetLoadConfigAsync(powerQueryCommands, excelPath, queryName, batchId),
                PowerQueryAction.ListExcelSources => await ListExcelSourcesAsync(powerQueryCommands, excelPath, batchId),
                PowerQueryAction.Eval => await EvalPowerQueryAsync(powerQueryCommands, excelPath, sourcePath, batchId),

                // Atomic Operations
                PowerQueryAction.Create => await CreatePowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, loadDestination, targetSheet, batchId),
                PowerQueryAction.UpdateMCode => await UpdateMCodePowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, batchId),
                PowerQueryAction.LoadTo => await LoadToPowerQueryAsync(powerQueryCommands, excelPath, queryName, loadDestination, targetSheet, batchId),
                PowerQueryAction.Unload => await UnloadPowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Queries,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Found {result.Queries.Count} Power Queries. Review M code and refresh configurations."
                : "Failed to list queries. Verify workbook has Power Query data connections.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'view' to examine M code for specific queries", "Use 'get-load-config' to check data loading settings", "Use 'refresh' to reload data from sources" }
                : ["Verify workbook has Power Query connections", "Check if workbook is macro-enabled if needed", "Use excel_connection list to see all connections"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryName,
            result.MCode,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"M code retrieved for '{queryName}'. Review transformations and data source connections."
                : $"Failed to view '{queryName}'. Verify query name is correct.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'export' to save M code for version control", "Use 'update-mcode' to modify transformations", "Use 'refresh' to reload with current M code" }
                : ["Use 'list' to see all available query names", "Check for typos in query name", "Verify query exists in workbook"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FilePath,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"M code exported to '{targetPath}'. Store in version control for change tracking."
                : $"Failed to export '{queryName}'. Verify query exists and target path is writable.",
            suggestedNextActions = result.Success
                ? new[] { "Commit .pq file to version control system", "Use 'update-mcode' to import modified M code", "Share .pq file with team for reuse" }
                : ["Use 'list' to verify query name", "Check directory permissions for target path", "Ensure parent directory exists"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshPowerQueryAsync(PowerQueryCommands commands, string excelPath, string? queryName, double? timeoutMinutes, string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for refresh action");

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
            // Enrich timeout error with operation-specific guidance (MCP layer responsibility)
            var result = new PowerQueryRefreshResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                QueryName = queryName,
                FilePath = excelPath,
                RefreshTime = DateTime.Now,

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

            // MCP layer: Add workflow guidance for LLMs
            var response = new
            {
                result.Success,
                result.ErrorMessage,
                result.QueryName,
                result.FilePath,
                result.RefreshTime,
                result.OperationContext,
                result.IsRetryable,
                result.RetryGuidance,

                // Workflow hints - MCP Server layer responsibility
                WorkflowHint = "Power Query refresh timeout - check data source and Excel dialogs",
                SuggestedNextActions = new[]
                {
                    "Check if Excel is showing a 'Privacy Level' dialog or credential prompt",
                    "Verify the data source is accessible (network connection, database availability)",
                    "For large datasets, consider filtering data at source or breaking into smaller queries",
                    "Use batch mode (begin_excel_batch) if not already using it to optimize multiple operations"
                }
            };

            return JsonSerializer.Serialize(response, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Query '{queryName}' deleted successfully. QueryTables referencing it may need cleanup."
                : $"Failed to delete '{queryName}'. Verify query exists and is not actively refreshing.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'list' to verify deletion", "Check for orphaned QueryTables with excel_querytable list", "Export backup before deletion (if not already done)" }
                : ["Use 'list' to verify query name", "Stop any active refresh operations", "Check if query is in use by PivotTables"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryName,
            result.LoadMode,
            result.TargetSheet,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Load configuration: {result.LoadMode}. Data is {(result.LoadMode == PowerQueryLoadMode.ConnectionOnly ? "not loaded" : $"loaded to {result.TargetSheet ?? "worksheet/data-model"}")}."
                : $"Failed to get load config for '{queryName}'. Verify query exists.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'load-to' to change data loading destination", "Use 'unload' to convert to connection-only", "Use 'refresh' to reload with current configuration" }
                : ["Use 'list' to verify query name", "Check query exists in workbook", "Verify query is not corrupted"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListExcelSourcesAsync(PowerQueryCommands commands, string excelPath, string? batchId)
    {
        // list-excel-sources action lists all available sources (doesn't require queryName)
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.ListExcelSourcesAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Worksheets,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Found {result.Worksheets.Count} available Excel sources. Use in M code: Excel.CurrentWorkbook(){{[Name=\"SourceName\"]}}[Content]."
                : "Failed to list Excel sources. Verify workbook has tables or named ranges.",
            suggestedNextActions = result.Success
                ? new[] { "Reference sources in M code transformations", "Use excel_table list to see structured tables", "Use excel_namedrange list to see named ranges" }
                : ["Create Excel Tables with excel_table create", "Create named ranges with excel_namedrange create", "Verify workbook structure"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> EvalPowerQueryAsync(PowerQueryCommands commands, string excelPath, string? sourcePath, string? batchId)
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryName,
            result.MCode,
            result.ErrorMessage,
            workflowHint = result.Success
                ? "M expression evaluated successfully. Use for testing transformations."
                : "M expression evaluation failed. Review syntax and data source availability.",
            suggestedNextActions = result.Success
                ? new[] { "Integrate evaluated M code into queries", "Test with different data sources", "Use 'create' to make permanent query from working code" }
                : ["Check M syntax for errors", "Verify data sources are accessible", "Test with simpler M expressions first"]
        }, ExcelToolsBase.JsonOptions);
    }

    // =========================================================================
    // ATOMIC OPERATIONS HANDLERS
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
        // Validate ALL required parameters first so error message lists every missing one
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourcePath))
        {
            var missing = new List<string>();
            if (string.IsNullOrEmpty(queryName)) missing.Add("queryName");
            if (string.IsNullOrEmpty(sourcePath)) missing.Add("sourcePath");
            var plural = missing.Count > 1 ? "are" : "is";
            throw new ModelContextProtocol.McpException($"{string.Join(" and ", missing)} {plural} required for create action (.pq file required)");
        }

        sourcePath = PathValidator.ValidateExistingFile(sourcePath, nameof(sourcePath));

        // Parse loadDestination to PowerQueryLoadMode enum
        var loadMode = ParseLoadMode(loadDestination ?? "worksheet");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.CreateAsync(batch, queryName, sourcePath, loadMode, targetSheet));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryName,
            result.LoadDestination,
            result.WorksheetName,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Query '{queryName}' created and data loaded to {loadMode}. M code imported from .pq file."
                : $"Failed to create '{queryName}'. Check M code syntax and data source connectivity.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'refresh' to reload data from source", "Use 'view' to inspect M code", "Use 'get-load-config' to verify loading settings" }
                : ["Verify M code syntax in .pq file", "Check data source connectivity", "Use 'eval' to test M code before creating query"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"M code updated for '{queryName}'. Data NOT refreshed - use 'refresh' or 'update-and-refresh' to reload."
                : $"Failed to update M code for '{queryName}'. Verify query exists and M syntax is valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'refresh' to reload data with new M code", "Use 'update-and-refresh' to update and refresh atomically next time", "Use 'view' to verify M code changes" }
                : ["Use 'list' to verify query name", "Check M syntax in .pq file", "Use 'eval' to test M code before updating"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> LoadToPowerQueryAsync(
        PowerQueryCommands commands,
        string excelPath,
        string? queryName,
        string? loadDestination,
        string? targetSheet,
        string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for load-to action");

        // Parse loadDestination to PowerQueryLoadMode enum
        var loadMode = ParseLoadMode(loadDestination ?? "worksheet");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.LoadToAsync(batch, queryName, loadMode, targetSheet));

        // Add workflow hints
        var inBatch = !string.IsNullOrEmpty(batchId);
        var destinationName = loadMode switch
        {
            PowerQueryLoadMode.LoadToTable => "worksheet",
            PowerQueryLoadMode.LoadToDataModel => "Data Model",
            PowerQueryLoadMode.LoadToBoth => "worksheet and Data Model",
            _ => "connection-only"
        };

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            result.QueryName,
            result.LoadDestination,
            result.WorksheetName,
            result.ConfigurationApplied,
            result.DataRefreshed,
            result.RowsLoaded,
            workflowHint = result.Success
                ? $"Query '{queryName}' now loaded to {destinationName}. Data refreshed with {result.RowsLoaded} rows."
                : $"Failed to load query '{queryName}': {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? (loadMode == PowerQueryLoadMode.LoadToDataModel || loadMode == PowerQueryLoadMode.LoadToBoth
                    ? new[]
                    {
                        "Use excel_datamodel 'list-tables' to verify query appears in Data Model",
                        "Use excel_datamodel 'create-measure' to add DAX calculations",
                        "Use excel_datamodel 'list-relationships' to check table relationships",
                        inBatch ? "Load more queries in this batch" : "Loading multiple queries? Use excel_batch for efficiency"
                    }
                    :
                    [
                        $"Use excel_range 'get-values' to read data from worksheet '{targetSheet ?? queryName}'",
                        "Use excel_powerquery 'refresh' to update data from source",
                        "Use excel_table 'create' to convert range to Excel Table for filtering/sorting",
                        inBatch ? "Load more queries in this batch" : "Loading multiple queries? Use excel_batch for efficiency"
                    ])
                :
                [
                    "Check if query name is correct with excel_powerquery 'list'",
                    "Verify query is connection-only with excel_powerquery 'get-load-config'",
                    "Review error message for specific issue"
                ]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Query '{queryName}' converted to connection-only. Data removed from worksheet/data-model."
                : $"Failed to unload '{queryName}'. Verify query exists and is currently loaded.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-load-config' to verify connection-only status", "Use 'load-to' to reload data to worksheet/data-model", "Use 'list' to see updated query status" }
                : ["Use 'get-load-config' to check current load status", "Verify query is not already connection-only", "Use 'list' to verify query name"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"M code updated and data refreshed for '{queryName}' atomically. Changes applied and loaded."
                : $"Failed atomic update for '{queryName}'. M code may be partially updated - verify with 'view'.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'view' to verify M code changes", "Use excel_range to inspect loaded data", "Use 'get-load-config' to see loading configuration" }
                : ["Use 'view' to check if M code was updated", "Use 'update-mcode' then 'refresh' separately if atomic fails", "Check M syntax and data source connectivity"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? "All Power Queries refreshed successfully. All data reloaded from sources."
                : "Failed to refresh all queries. Some queries may have connectivity issues.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'list' to verify all queries refreshed", "Use excel_range to inspect updated data", "Check refresh timestamps with 'list' action" }
                : ["Use 'refresh' on individual queries to isolate failures", "Check data source connectivity for failed queries", "Review error messages for specific query issues"]
        }, ExcelToolsBase.JsonOptions);
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
