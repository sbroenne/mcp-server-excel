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
/// MCP tool for Power Query M code and data loading operations.
/// </summary>
[McpServerToolType]
public static class ExcelPowerQueryTool
{
    [McpServerTool(Name = "excel_powerquery")]
    [Description(@"Manage Power Query M code and data loading.

⚠️ SHEET NAME CONFLICTS (LoadTo action):
- If a worksheet with the target name already exists, LoadTo returns an error
- User must delete the existing sheet first using excel_worksheet action='Delete'
- Then retry LoadTo - this ensures explicit user control over data deletion

WHEN TO USE CREATE vs UPDATE:
- Create: For NEW queries only (FAILS with 'already exists' error if query exists)
- Update: For EXISTING queries (updates M code + refreshes data)
- Not sure? Use List action first to check if query exists

OPERATIONS GUIDANCE:
- Create: Import M code from .pq file AND optionally load data (NEW queries only)
- Update: Update M code AND refresh data in ONE operation (EXISTING queries only)
- LoadTo: Apply load destination to connection-only query (make it load data to a worksheet or Data Model)
- Unload: Convert query to connection-only (remove data, keep M code definition)
- RefreshAll: Refresh ALL Power Queries in workbook (batch refresh)
")]
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

                // Atomic Operations
                PowerQueryAction.Create => await CreatePowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, loadDestination, targetSheet, batchId),
                PowerQueryAction.Update => await UpdatePowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, batchId),
                PowerQueryAction.LoadTo => await LoadToPowerQueryAsync(powerQueryCommands, excelPath, queryName, loadDestination, targetSheet, batchId),
                PowerQueryAction.Unload => await UnloadPowerQueryAsync(powerQueryCommands, excelPath, queryName, batchId),
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

    private static async Task<string> UpdatePowerQueryAsync(
        PowerQueryCommands commands,
        string excelPath,
        string? queryName,
        string? sourcePath,
        string? batchId)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for update action");
        if (string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("sourcePath is required for update action (.pq file)");

        sourcePath = PathValidator.ValidateExistingFile(sourcePath, nameof(sourcePath));

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.UpdateAsync(batch, queryName, sourcePath));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"M code updated and data refreshed for '{queryName}' atomically. Query is current."
                : $"Failed to update '{queryName}'. {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[] { "Use 'view' to verify M code changes", "Use excel_range to inspect refreshed data", "Use 'get-load-config' to see loading configuration" }
                : ["Use 'list' to verify query name", "Check M syntax in .pq file", "Verify data source connectivity"]
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

        // Detect sheet conflict error and provide specific guidance
        var isSheetConflict = !result.Success &&
                             result.ErrorMessage?.Contains("worksheet already exists", StringComparison.OrdinalIgnoreCase) == true;

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
                : isSheetConflict
                    ? $"Cannot load query '{queryName}': sheet '{targetSheet ?? queryName}' already exists. Delete it first, then retry LoadTo."
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
                : isSheetConflict
                    ?
                    [
                        $"Use excel_worksheet action='Delete' sheetName='{targetSheet ?? queryName}' to delete the existing sheet",
                        $"Then retry: excel_powerquery action='LoadTo' queryName='{queryName}' loadDestination='{loadDestination ?? "worksheet"}' targetSheet='{targetSheet ?? queryName}'",
                        "Or rename the target sheet if you want to keep existing data"
                    ]
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
