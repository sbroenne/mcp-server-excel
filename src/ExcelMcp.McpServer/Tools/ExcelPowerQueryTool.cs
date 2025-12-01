using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Power Query M code and data loading operations.
/// </summary>
[McpServerToolType]
public static partial class ExcelPowerQueryTool
{
    /// <summary>
    /// Manage Power Query M code and data loading.
    /// INLINE M CODE: Provide raw M code via mCode parameter.
    /// TARGET SHEETS (Create and LoadTo actions): targetCellAddress works for BOTH create and load-to to place tables without clearing other content. If targetCellAddress is omitted and sheet already contains data, server returns guidance instead of deleting it. When re-using an existing QueryTable, LoadTo refreshes data in-place without recreating the table.
    /// TIMEOUT SAFEGUARD: Long-running refresh/load operations auto-timeout after 5 minutes. On timeout the tool returns SuggestedNextActions instead of hanging the session.
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action. Required for all Power Query operations.</param>
    /// <param name="queryName">Power Query name (required for most actions)</param>
    /// <param name="mCode">Raw Power Query M code (inline string). Required for create and update actions.</param>
    /// <param name="targetSheet">Target worksheet name (when loadDestination is 'worksheet' or 'both')</param>
    /// <param name="targetCellAddress">Top-left cell for create/load-to actions when placing data on an existing worksheet (e.g., 'B5'). Only used when load destination is 'worksheet' or 'both'.</param>
    /// <param name="loadDestination">Load destination for query: 'worksheet' (DEFAULT - load to worksheet as table), 'data-model' (load to Power Pivot), 'both' (load to both), 'connection-only' (don't load data)</param>
    /// <param name="refreshTimeoutSeconds">Timeout in seconds for refresh action (60-600 seconds / 1-10 minutes). Required when action is 'refresh'.</param>
    [McpServerTool(Name = "excel_powerquery")]
    public static partial string ExcelPowerQuery(
        PowerQueryAction action,
        string sessionId,
        string? queryName,
        string? mCode,
        string? targetSheet,
        string? targetCellAddress,
        string? loadDestination,
        int? refreshTimeoutSeconds)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_powerquery",
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

        try
        {
            ExcelToolsBase.WithSession(sessionId,
                batch =>
                {
                    commands.Delete(batch, queryName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Query '{queryName}' deleted successfully",
                workflowHint = "QueryTables referencing this query may need cleanup."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(sessionId,
                batch =>
                {
                    commands.Create(batch, queryName, mCode, loadMode, resolvedTargetSheet, targetCellAddress);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                queryName,
                loadDestination = loadMode.ToString(),
                worksheetName = resolvedTargetSheet,
                message = $"Query '{queryName}' created successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(sessionId,
                batch =>
                {
                    commands.Update(batch, queryName, mCode);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Query '{queryName}' updated successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.LoadTo(batch, queryName, loadMode, resolvedTargetSheet, targetCellAddress);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                queryName,
                loadDestination = loadMode.ToString(),
                worksheetName = resolvedTargetSheet,
                message = $"Query '{queryName}' load configuration applied successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (TimeoutException ex)
        {
            var suggestedNextActions = new[]
            {
                "Check if Excel is showing a 'Privacy Level' or credential dialog and dismiss it.",
                "Verify the data source is reachable and credentials are valid.",
                "If the dataset is large, load to worksheet first or break the query into smaller sources.",
                "Use begin_excel_batch to keep the Excel session open while iterating on load configuration."
            };

            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true,
                suggestedNextActions,
                retryGuidance = ex.Message.Contains("maximum timeout", StringComparison.OrdinalIgnoreCase)
                    ? "Maximum timeout reached. Resolve Excel dialogs or reduce the amount of data before retrying."
                    : "After verifying the data source, you can retry this operation within the 5 minute timeout limit."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string RefreshAllPowerQueriesAsync(
        PowerQueryCommands commands,
        string sessionId)
    {
        try
        {
            ExcelToolsBase.WithSession(sessionId,
                batch =>
                {
                    commands.RefreshAll(batch);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = "All queries refreshed successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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

