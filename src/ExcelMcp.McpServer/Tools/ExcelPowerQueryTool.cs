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
        [Description("Session ID from excel_file 'open' action. Required for all Power Query operations.")]
        string sessionId,

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
                PowerQueryAction.List => ListPowerQueriesAsync(powerQueryCommands, sessionId),
                PowerQueryAction.View => ViewPowerQueryAsync(powerQueryCommands, sessionId, queryName),
                PowerQueryAction.Export => ExportPowerQueryAsync(powerQueryCommands, sessionId, queryName, targetPath),
                PowerQueryAction.Refresh => RefreshPowerQueryAsync(powerQueryCommands, sessionId, queryName, timeout),
                PowerQueryAction.Delete => DeletePowerQueryAsync(powerQueryCommands, sessionId, queryName),
                PowerQueryAction.GetLoadConfig => GetLoadConfigAsync(powerQueryCommands, sessionId, queryName),
                PowerQueryAction.ListExcelSources => ListExcelSourcesAsync(powerQueryCommands, sessionId),

                // Atomic Operations
                PowerQueryAction.Create => CreatePowerQueryAsync(powerQueryCommands, sessionId, queryName, sourcePath, loadDestination, targetSheet),
                PowerQueryAction.Update => UpdatePowerQueryAsync(powerQueryCommands, sessionId, queryName, sourcePath),
                PowerQueryAction.LoadTo => LoadToPowerQueryAsync(powerQueryCommands, sessionId, queryName, loadDestination, targetSheet),
                PowerQueryAction.Unload => UnloadPowerQueryAsync(powerQueryCommands, sessionId, queryName),
                PowerQueryAction.RefreshAll => RefreshAllPowerQueriesAsync(powerQueryCommands, sessionId),

                _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
            };
        }
        catch (Exception ex)
        {
            return Task.FromResult(JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"{action.ToActionString()} failed: {ex.Message}",
                isError = true
            }, ExcelToolsBase.JsonOptions));
        }
    }

    private static string ListPowerQueriesAsync(PowerQueryCommands commands, string sessionId)
    {
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.List(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Queries,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ViewPowerQueryAsync(PowerQueryCommands commands, string sessionId, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ArgumentException("queryName is required for view action", nameof(queryName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.View(batch, queryName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryName,
            result.MCode,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ExportPowerQueryAsync(PowerQueryCommands commands, string sessionId, string? queryName, string? targetPath)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ArgumentException("queryName is required for export action", nameof(queryName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.View(batch, queryName));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string RefreshPowerQueryAsync(PowerQueryCommands commands, string sessionId, string? queryName, double? timeoutMinutes)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ArgumentException("queryName is required for refresh action", nameof(queryName));

        try
        {
            // Apply operation-specific timeout default (5 minutes for refresh)
            var timeoutSpan = timeoutMinutes.HasValue ? (TimeSpan?)TimeSpan.FromMinutes(timeoutMinutes.Value) : null;

            var result = ExcelToolsBase.WithSession(sessionId,
                batch => commands.Refresh(batch, queryName, timeoutSpan));

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
                FilePath = null,
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
                    "For large datasets, consider filtering data at source or breaking into smaller queries"
                }
            };

            return JsonSerializer.Serialize(response, ExcelToolsBase.JsonOptions);
        }
    }

    private static string DeletePowerQueryAsync(PowerQueryCommands commands, string sessionId, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ArgumentException("queryName is required for delete action", nameof(queryName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Delete(batch, queryName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? "QueryTables referencing this query may need cleanup."
                : null
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetLoadConfigAsync(PowerQueryCommands commands, string sessionId, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ArgumentException("queryName is required for get-load-config action", nameof(queryName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.GetLoadConfig(batch, queryName));

        return Task.FromResult(JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryName,
            result.LoadMode,
            result.TargetSheet,
            result.ErrorMessage,
            workflowHint = result.Success && result.LoadMode == PowerQueryLoadMode.ConnectionOnly
                ? "Query is connection-only (M code defined but data not loaded)."
                : null
        }, ExcelToolsBase.JsonOptions));
    }

    private static string ListExcelSourcesAsync(PowerQueryCommands commands, string sessionId)
    {
        // list-excel-sources action lists all available sources (doesn't require queryName)
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.ListExcelSources(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Worksheets,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // =========================================================================
    // ATOMIC OPERATIONS HANDLERS
    // =========================================================================

    private static string CreatePowerQueryAsync(
        PowerQueryCommands commands,
        string sessionId,
        string? queryName,
        string? sourcePath,
        string? loadDestination,
        string? targetSheet)
    {
        // Validate ALL required parameters first so error message lists every missing one
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrEmpty(sourcePath))
        {
            var missing = new List<string>();
            if (string.IsNullOrEmpty(queryName)) missing.Add("queryName");
            if (string.IsNullOrEmpty(sourcePath)) missing.Add("sourcePath");
            var plural = missing.Count > 1 ? "are" : "is";
            throw new ArgumentException($"{string.Join(" and ", missing)} {plural} required for create action (.pq file required)", string.Join(",", missing));
        }

        sourcePath = PathValidator.ValidateExistingFile(sourcePath, nameof(sourcePath));

        // Parse loadDestination to PowerQueryLoadMode enum
        var loadMode = ParseLoadMode(loadDestination ?? "worksheet");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Create(batch, queryName, sourcePath, loadMode, targetSheet));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryName,
            result.LoadDestination,
            result.WorksheetName,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string UpdatePowerQueryAsync(
        PowerQueryCommands commands,
        string sessionId,
        string? queryName,
        string? sourcePath)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ArgumentException("queryName is required for update action", nameof(queryName));
        if (string.IsNullOrEmpty(sourcePath))
            throw new ArgumentException("sourcePath is required for update action (.pq file)", nameof(sourcePath));

        sourcePath = PathValidator.ValidateExistingFile(sourcePath, nameof(sourcePath));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Update(batch, queryName, sourcePath));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string LoadToPowerQueryAsync(
        PowerQueryCommands commands,
        string sessionId,
        string? queryName,
        string? loadDestination,
        string? targetSheet)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ArgumentException("queryName is required for load-to action", nameof(queryName));

        // Parse loadDestination to PowerQueryLoadMode enum
        var loadMode = ParseLoadMode(loadDestination ?? "worksheet");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.LoadTo(batch, queryName, loadMode, targetSheet));

        // Detect sheet conflict error and provide specific guidance
        var isSheetConflict = !result.Success &&
                             result.ErrorMessage?.Contains("worksheet already exists", StringComparison.OrdinalIgnoreCase) == true;

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
            workflowHint = isSheetConflict
                ? $"Sheet '{targetSheet ?? queryName}' already exists. Delete it first with excel_worksheet action='Delete', then retry LoadTo."
                : null
        }, ExcelToolsBase.JsonOptions);
    }

    private static string UnloadPowerQueryAsync(
        PowerQueryCommands commands,
        string sessionId,
        string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ArgumentException("queryName is required for unload action", nameof(queryName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Unload(batch, queryName));

        return Task.FromResult(JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? "Query converted to connection-only (M code preserved, data removed)."
                : null
        }, ExcelToolsBase.JsonOptions));
    }

    private static string RefreshAllPowerQueriesAsync(
        PowerQueryCommands commands,
        string sessionId)
    {
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.RefreshAll(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
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
            _ => throw new ArgumentException($"Invalid loadDestination: '{loadDestination}'. Valid values: worksheet, data-model, both, connection-only", nameof(loadDestination))
        };
    }
}

