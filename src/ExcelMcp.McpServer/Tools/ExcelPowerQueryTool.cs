using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Power Query M code and data loading operations.
/// </summary>
[McpServerToolType]
public static partial class ExcelPowerQueryTool
{
    /// <summary>
    /// Power Query M code and data loading.
    ///
    /// TEST-FIRST DEVELOPMENT WORKFLOW (BEST PRACTICE):
    /// 1. evaluate → Test M code WITHOUT persisting (catches syntax errors, validates sources, shows data preview)
    /// 2. create/update → Store VALIDATED query in workbook
    /// 3. refresh/load-to → Load data to destination
    /// Skip evaluate only for trivial literal tables (#table with hardcoded values).
    ///
    /// IF CREATE/UPDATE FAILS: Use evaluate to get detailed Power Query error message, fix code, retry.
    /// This is the RECOMMENDED way to debug M code issues - evaluate gives you the actual M engine error.
    ///
    /// DATETIME COLUMNS: Always include Table.TransformColumnTypes() in M code to set column types explicitly.
    /// Without explicit types, dates may be stored as numbers and Data Model relationships may fail.
    /// Example: Table.TransformColumnTypes(Source, {{"OrderDate", type datetime}, {"Amount", type number}})
    ///
    /// M-CODE FORMATTING: Create and Update automatically format M code using powerqueryformatter.com API.
    ///
    /// DESTINATIONS: worksheet (default), data-model (for DAX), both, connection-only.
    /// Use 'data-model' to load directly to Power Pivot, then use excel_datamodel to create DAX measures.
    ///
    /// TARGET CELL: targetCellAddress places tables without clearing sheet.
    /// TIMEOUT: 5 min auto-timeout for refresh/load.
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action. Required for all Power Query operations.</param>
    /// <param name="queryName">Power Query name (required for most actions)</param>
    /// <param name="mCode">Raw Power Query M code (inline string). Required for create, update, evaluate actions (unless mCodeFile is provided).</param>
    /// <param name="mCodeFile">Full path to .m or .pq file containing M code. Alternative to mCode parameter - use for large/complex queries.</param>
    /// <param name="targetSheet">Target worksheet name (when loadDestination is 'worksheet' or 'both')</param>
    /// <param name="targetCellAddress">Top-left cell for create/load-to actions when placing data on an existing worksheet (e.g., 'B5'). Only used when load destination is 'worksheet' or 'both'.</param>
    /// <param name="loadDestination">Load destination for query: 'worksheet' (DEFAULT - load to worksheet as table), 'data-model' (load to Power Pivot), 'both' (load to both), 'connection-only' (don't load data)</param>
    /// <param name="refreshTimeoutSeconds">Timeout in seconds for refresh action (60-600 seconds / 1-10 minutes). Required when action is 'refresh'.</param>
    /// <param name="newName">New name for query (required for rename action). Names are trimmed; case-insensitive uniqueness is enforced.</param>
    [McpServerTool(Name = "excel_powerquery", Title = "Excel Power Query Operations", Destructive = true)]
    [McpMeta("category", "query")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelPowerQuery(
        PowerQueryAction action,
        string sessionId,
        [DefaultValue(null)] string? queryName,
        [DefaultValue(null)] string? mCode,
        [DefaultValue(null)] string? mCodeFile,
        [DefaultValue(null)] string? targetSheet,
        [DefaultValue(null)] string? targetCellAddress,
        [DefaultValue(null)] string? loadDestination,
        [DefaultValue(null)] int? refreshTimeoutSeconds,
        [DefaultValue(null)] string? newName)
    {
        // Convert refreshTimeoutSeconds to TimeSpan for the generated code
        TimeSpan? timeout = refreshTimeoutSeconds.HasValue
            ? TimeSpan.FromSeconds(refreshTimeoutSeconds.Value)
            : null;

        // Map queryName to oldName for rename action
        string? oldName = action == PowerQueryAction.Rename ? queryName : null;

        return ExcelToolsBase.ExecuteToolAction(
            "excel_powerquery",
            ServiceRegistry.PowerQuery.ToActionString(action),
            () => ServiceRegistry.PowerQuery.RouteAction(
                action,
                sessionId,
                ExcelToolsBase.ForwardToServiceFunc,
                queryName: queryName,
                timeout: timeout,
                mCode: mCode,
                mCodeFile: mCodeFile,
                loadDestination: loadDestination,
                targetSheet: targetSheet,
                targetCellAddress: targetCellAddress,
                refresh: true,
                oldName: oldName,
                newName: newName
            ));
    }
}





