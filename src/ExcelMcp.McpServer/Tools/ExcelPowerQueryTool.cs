using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
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

    ⚠️ INLINE M CODE
    - Provide raw M code via mCode parameter

    ⚠️ TARGET SHEETS (Create & LoadTo actions):
    - targetCellAddress works for BOTH create and load-to to place tables without clearing other content
    - If targetCellAddress is omitted and sheet already contains data, server returns guidance instead of deleting it
    - When re-using an existing QueryTable, LoadTo refreshes data in-place without recreating the table

    ⏱️ TIMEOUT SAFEGUARD
    - Long-running refresh/load operations auto-timeout after 5 minutes
    - On timeout the tool returns SuggestedNextActions instead of hanging the session
")]
    public static string ExcelPowerQuery(
        [Required]
        [Description("Action to perform (enum values displayed as dropdown in MCP clients)")]
        PowerQueryAction action,

        [Required]
        [Description("Session ID from excel_file 'open' action. Required for all Power Query operations.")]
        string sessionId,

        [StringLength(255, MinimumLength = 1)]
        [Description("Power Query name (required for most actions)")]
        string? queryName = null,

        [Description("Raw Power Query M code (inline string). Required for create and update actions.")]
        string? mCode = null,


        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("Target worksheet name (when loadDestination is 'worksheet' or 'both')")]
        string? targetSheet = null,

        [RegularExpression(@"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$")]
        [Description("Top-left cell for create/load-to actions when placing data on an existing worksheet (e.g., 'B5'). Only used when load destination is 'worksheet' or 'both'.")]
        string? targetCellAddress = null,

        [RegularExpression("^(worksheet|data-model|both|connection-only)$")]
        [Description(@"Load destination for query (for create action). Options:
      - 'worksheet': Load to worksheet as table (DEFAULT - users can see/validate data)
      - 'data-model': Load to Power Pivot Data Model (for DAX measures/relationships)
      - 'both': Load to both worksheet AND Data Model
      - 'connection-only': Don't load data (M code imported but not executed)")]
        string? loadDestination = null,

        [Range(60, 600)]
        [Description("Timeout in seconds for refresh action (60-600 seconds / 1-10 minutes). Required when action is 'refresh'.")]
        int? refreshTimeoutSeconds = null)
    {
        return ExcelToolsBase.ExecuteToolAction(
            action.ToActionString(),
            () =>
            {
                var dataModelCommands = new DataModelCommands();
                var powerQueryCommands = new PowerQueryCommands(dataModelCommands);

                return action switch
                {
                    PowerQueryAction.List => ListPowerQueriesAsync(powerQueryCommands, sessionId),
                    PowerQueryAction.View => ViewPowerQueryAsync(powerQueryCommands, sessionId, queryName),
                    PowerQueryAction.Refresh => RefreshPowerQueryAsync(powerQueryCommands, sessionId, queryName, refreshTimeoutSeconds),
                    PowerQueryAction.Delete => DeletePowerQueryAsync(powerQueryCommands, sessionId, queryName),
                    PowerQueryAction.GetLoadConfig => GetLoadConfigAsync(powerQueryCommands, sessionId, queryName),

                    // Atomic Operations
                    PowerQueryAction.Create => CreatePowerQueryAsync(powerQueryCommands, sessionId, queryName, mCode, loadDestination, targetSheet, targetCellAddress),
                    PowerQueryAction.Update => UpdatePowerQueryAsync(powerQueryCommands, sessionId, queryName, mCode),
                    PowerQueryAction.LoadTo => LoadToPowerQueryAsync(powerQueryCommands, sessionId, queryName, loadDestination, targetSheet, targetCellAddress),
                    PowerQueryAction.RefreshAll => RefreshAllPowerQueriesAsync(powerQueryCommands, sessionId),

                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
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

    private static string RefreshPowerQueryAsync(PowerQueryCommands commands, string sessionId, string? queryName, int? refreshTimeoutSeconds)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ArgumentException("queryName is required for refresh action", nameof(queryName));

        if (refreshTimeoutSeconds is null)
            throw new ArgumentException("refreshTimeoutSeconds is required for refresh action", nameof(refreshTimeoutSeconds));

        const int minSeconds = 60;
        const int maxSeconds = 600;
        if (refreshTimeoutSeconds < minSeconds || refreshTimeoutSeconds > maxSeconds)
        {
            throw new ArgumentOutOfRangeException(
                nameof(refreshTimeoutSeconds),
                $"refreshTimeoutSeconds must be between {minSeconds} seconds (1 minute) and {maxSeconds} seconds (10 minutes). For longer operations, ask the user to run the refresh manually in Excel.");
        }

        var timeout = TimeSpan.FromSeconds(refreshTimeoutSeconds.Value);

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Refresh(batch, queryName, timeout));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryName,
            result.LoadMode,
            result.TargetSheet,
            result.ErrorMessage,
            workflowHint = result.Success && result.LoadMode == PowerQueryLoadMode.ConnectionOnly
                ? "Query is connection-only (M code defined but data not loaded)."
                : null
        }, ExcelToolsBase.JsonOptions);
    }

    // =========================================================================
    // ATOMIC OPERATIONS HANDLERS
    // =========================================================================

    private static string CreatePowerQueryAsync(
        PowerQueryCommands commands,
        string sessionId,
        string? queryName,
        string? mCode,
        string? loadDestination,
        string? targetSheet,
        string? targetCellAddress)
    {
        // Validate ALL required parameters first so error message lists every missing one
        if (string.IsNullOrEmpty(queryName) || string.IsNullOrWhiteSpace(mCode))
        {
            var missing = new List<string>();
            if (string.IsNullOrEmpty(queryName)) missing.Add("queryName");
            if (string.IsNullOrWhiteSpace(mCode)) missing.Add("mCode");
            var plural = missing.Count > 1 ? "are" : "is";
            throw new ArgumentException($"{string.Join(" and ", missing)} {plural} required for create action", string.Join(",", missing));
        }

        // Parse loadDestination to PowerQueryLoadMode enum
        var loadMode = ParseLoadMode(loadDestination ?? "worksheet");

        if (!RequiresWorksheet(loadMode) && !string.IsNullOrWhiteSpace(targetCellAddress))
        {
            throw new ArgumentException("targetCellAddress can only be used when loadDestination is 'worksheet' or 'both'", nameof(targetCellAddress));
        }

        var resolvedTargetSheet = RequiresWorksheet(loadMode)
            ? (string.IsNullOrWhiteSpace(targetSheet) ? queryName : targetSheet)
            : targetSheet;

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Create(batch, queryName, mCode, loadMode, resolvedTargetSheet, targetCellAddress));

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
        string? mCode)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ArgumentException("queryName is required for update action", nameof(queryName));
        if (string.IsNullOrWhiteSpace(mCode))
            throw new ArgumentException("mCode is required for update action", nameof(mCode));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Update(batch, queryName, mCode));

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
        string? targetSheet,
        string? targetCellAddress)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ArgumentException("queryName is required for load-to action", nameof(queryName));

        // Parse loadDestination to PowerQueryLoadMode enum
        var loadMode = ParseLoadMode(loadDestination ?? "worksheet");

        var requiresWorksheet = RequiresWorksheet(loadMode);

        if (!requiresWorksheet && !string.IsNullOrWhiteSpace(targetCellAddress))
            throw new ArgumentException("targetCellAddress can only be used when loadDestination is 'worksheet' or 'both'", nameof(targetCellAddress));

        var resolvedTargetSheet = requiresWorksheet
            ? (string.IsNullOrWhiteSpace(targetSheet) ? queryName : targetSheet)
            : targetSheet;

        PowerQueryLoadResult result;
        bool isTimeout = false;
        string[]? suggestedNextActions = null;
        Dictionary<string, object>? operationContext = null;
        string? retryGuidance = null;

        try
        {
            result = ExcelToolsBase.WithSession(
                sessionId,
                batch => commands.LoadTo(batch, queryName, loadMode, resolvedTargetSheet, targetCellAddress));
        }
        catch (TimeoutException ex)
        {
            isTimeout = true;

            result = new PowerQueryLoadResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                QueryName = queryName!,
                LoadDestination = loadMode,
                WorksheetName = resolvedTargetSheet,
                TargetCellAddress = targetCellAddress,
                ConfigurationApplied = false,
                DataRefreshed = false,
                RowsLoaded = 0
            };

            bool usedMaxTimeout = ex.Message.Contains("maximum timeout", StringComparison.OrdinalIgnoreCase);

            suggestedNextActions =
            [
                "Check if Excel is showing a 'Privacy Level' or credential dialog and dismiss it.",
                "Verify the data source is reachable and credentials are valid.",
                "If the dataset is large, load to worksheet first or break the query into smaller sources.",
                "Use begin_excel_batch to keep the Excel session open while iterating on load configuration."
            ];

            operationContext = new Dictionary<string, object>
            {
                ["OperationType"] = "PowerQuery.LoadTo",
                ["QueryName"] = queryName!,
                ["TimeoutReached"] = true,
                ["UsedMaxTimeout"] = usedMaxTimeout
            };

            retryGuidance = usedMaxTimeout
                ? "Maximum timeout reached. Resolve Excel dialogs or reduce the amount of data before retrying."
                : "After verifying the data source, you can retry this operation within the 5 minute timeout limit.";
        }

        // Detect sheet conflict error and provide specific guidance
        var isSheetConflict = !result.Success &&
                             result.ErrorMessage?.Contains("worksheet already exists", StringComparison.OrdinalIgnoreCase) == true;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            result.QueryName,
            result.LoadDestination,
            result.WorksheetName,
            result.TargetCellAddress,
            result.ConfigurationApplied,
            result.DataRefreshed,
            result.RowsLoaded,
            isError = !result.Success || isTimeout,
            suggestedNextActions,
            retryGuidance,
            operationContext,
            workflowHint = isSheetConflict
                ? $"Sheet '{targetSheet ?? queryName}' already contains data. Provide targetCellAddress (e.g., \"B5\") to place the table without deleting the sheet."
                : isTimeout
                    ? "Excel may be waiting on a modal dialog (privacy levels, credentials, etc.). Check Excel before retrying."
                    : null
        }, ExcelToolsBase.JsonOptions);
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

    private static bool RequiresWorksheet(PowerQueryLoadMode loadMode) =>
        loadMode == PowerQueryLoadMode.LoadToTable || loadMode == PowerQueryLoadMode.LoadToBoth;
}

