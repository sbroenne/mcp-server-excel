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
/// PHASE 1 ATOMIC OPERATIONS:
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

PRIMARY ACTIONS (PHASE 1 ATOMIC OPERATIONS):
- Create: Import Power Query from .pq file + load data in one atomic operation
- UpdateMCode: Update M code only (no refresh - for staging changes)
- UpdateAndRefresh: Update M code + refresh data in one atomic operation
- Unload: Convert query to connection-only (remove data, keep definition)
- RefreshAll: Refresh all Power Queries in workbook

LOAD DESTINATIONS:
- 'worksheet': Load to worksheet as table (users can see/validate data)
- 'data-model': Load to Power Pivot Data Model (ready for DAX measures/relationships)
- 'both': Load to BOTH worksheet AND Data Model
- 'connection-only': Don't load data (M code imported but not executed)

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

                // Phase 1: Atomic Operations
                PowerQueryAction.Create => await CreatePowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, loadDestination, targetSheet, batchId),
                PowerQueryAction.UpdateMCode => await UpdateMCodePowerQueryAsync(powerQueryCommands, excelPath, queryName, sourcePath, batchId),
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
