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

        [RegularExpression(@"^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$")]
        [Description("Top-left cell for create/load-to actions when placing data on an existing worksheet (e.g., 'B5'). Only used when load destination is 'worksheet' or 'both'.")]
        string? targetCellAddress = null,

        [RegularExpression("^(worksheet|data-model|both|connection-only)$")]
        [Description(@"Load destination for query (for create action). Options:
  - 'worksheet': Load to worksheet as table (DEFAULT - users can see/validate data)
  - 'data-model': Load to Power Pivot Data Model (for DAX measures/relationships)
  - 'both': Load to both worksheet AND Data Model
  - 'connection-only': Don't load data (M code imported but not executed)")]
        string? loadDestination = null)
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
                PowerQueryAction.Refresh => RefreshPowerQueryAsync(powerQueryCommands, sessionId, queryName),
                PowerQueryAction.Delete => DeletePowerQueryAsync(powerQueryCommands, sessionId, queryName),
                PowerQueryAction.GetLoadConfig => GetLoadConfigAsync(powerQueryCommands, sessionId, queryName),
                PowerQueryAction.ListExcelSources => ListExcelSourcesAsync(powerQueryCommands, sessionId),

                // Atomic Operations
                PowerQueryAction.Create => CreatePowerQueryAsync(powerQueryCommands, sessionId, queryName, mCode, loadDestination, targetSheet, targetCellAddress),
                PowerQueryAction.Update => UpdatePowerQueryAsync(powerQueryCommands, sessionId, queryName, mCode),
                PowerQueryAction.LoadTo => LoadToPowerQueryAsync(powerQueryCommands, sessionId, queryName, loadDestination, targetSheet, targetCellAddress),
                PowerQueryAction.Unload => UnloadPowerQueryAsync(powerQueryCommands, sessionId, queryName),
                PowerQueryAction.RefreshAll => RefreshAllPowerQueriesAsync(powerQueryCommands, sessionId),

                _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"{action.ToActionString()} failed: {ex.Message}",
                isError = true
            }, ExcelToolsBase.JsonOptions);
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

    private static string RefreshPowerQueryAsync(PowerQueryCommands commands, string sessionId, string? queryName)
    {
        if (string.IsNullOrEmpty(queryName))
            throw new ArgumentException("queryName is required for refresh action", nameof(queryName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Refresh(batch, queryName, null));

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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.LoadTo(batch, queryName, loadMode, resolvedTargetSheet, targetCellAddress));

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
            result.TargetCellAddress,
            result.ConfigurationApplied,
            result.DataRefreshed,
            result.RowsLoaded,
            workflowHint = isSheetConflict
                ? $"Sheet '{targetSheet ?? queryName}' already contains data. Provide targetCellAddress (e.g., \"B5\") to place the table without deleting the sheet."
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? "Query converted to connection-only (M code preserved, data removed)."
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

