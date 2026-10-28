using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query M code and data loading.
///
/// Prefer evaluate for new or materially changed M code before persisting it.
/// Create loads its selected destination (worksheet by default); connection-only stores without loading.
/// Update refreshes by default unless refresh=false. Load-to changes destinations; refresh updates loaded data.
/// Trivial literal tables or already-validated code with unchanged dependencies need no redundant evaluation.
///
/// IF CREATE/UPDATE FAILS: Changes are not rolled back. Formula.Firewall failures use category PrivacyFirewall and point to set-privacy-settings.
/// Errors after storing the definition include partialState with the surviving query/load configuration.
/// Inspect list/view/get-load-config; correct an existing query with update rather than repeating create.
///
/// DATETIME COLUMNS: Always include Table.TransformColumnTypes() in M code to set column types explicitly.
/// Without explicit types, dates may be stored as numbers and Data Model relationships may fail.
///
/// DESTINATIONS: 'worksheet' (default), 'data-model' (for DAX), 'both', 'connection-only'.
/// Values are case-insensitive and unknown values are rejected.
/// Use 'data-model' to load to Power Pivot, then use datamodel to create DAX measures.
///
/// M-CODE: Preserved exactly by default. Set formatMCode=true only with explicit user consent;
/// it sends M code to powerqueryformatter.com.
///
/// TARGET CELL: targetCellAddress places tables without clearing sheet.
/// TIMEOUT: Create, load-to, and update-with-refresh use a 30-minute wait; refresh accepts an override. Expiry ends the caller wait, not Excel execution.
/// READS: list returns compact metadata and exact load state without formula previews.
/// Use view for one query's full M code.
/// DIAGNOSTICS: Loads/refreshes report durationMs and verified rowsLoaded,
/// worksheetRowsLoaded and modelRowsLoaded where available. Null/omitted counts
/// are unknown, not zero; both destinations are not double-counted.
/// These are loaded-row counts, not cell-error or warning counts.
/// Wrong credential errors include Authentication guidance; unsupported Analysis
/// Services cultures include a Culture hint without rewriting M code.
/// </summary>
[ServiceCategory("PowerQuery")]
[McpTool("powerquery", Title = "Power Query Operations", Destructive = true, Category = "query",
    Description = "Create, change, evaluate, refresh, and remove Power Query M code and data loads. Prefer evaluate for new or changed M code before saving it; it reports engine errors. " +
        "LOADING: create loads to its destination (worksheet by default): worksheet, data-model, both, or connection-only. target_cell_address places tables without clearing other content. update refreshes unless refresh=false. load-to changes destinations. Set explicit column types for dates and Data Model relationships. " +
        "REFRESH: refresh updates one query (query_name). refresh-all refreshes every loaded query, skips parameter and connection-only queries (skippedQueries), then dependent PivotTables unless refresh_pivot_tables=false. It continues past failures: success is false, failures are listed in failedQueries/failedPivotTables, and refreshed data is not rolled back. trace=true adds a trace summary; use powerquery_trace_read for detail. Use workbook refresh-all to refresh all workbook connections the way Excel does. " +
        "TIMEOUTS: create, load-to, and update-with-refresh wait up to 30 minutes; refresh accepts a timeout. A timeout or cancellation ends only the wait and Excel work may continue; inspect session status before further work. " +
        "PRIVACY: Formula.Firewall errors can be resolved with set-privacy-settings ignore_privacy_levels=true, which may allow data from different sources to be combined and sent between them; it is stored in the workbook when saved. " +
        "format_m_code=true sends M code to powerqueryformatter.com and requires user consent.")]
[McpReadOnlyActions("list", "view", "get-load-config", "get-privacy-settings")]
public interface IPowerQueryCommands
{
    /// <summary>
    /// Lists compact metadata for all Power Queries, including exact worksheet/Data Model
    /// load state. Formula previews are omitted from CLI and MCP output;
    /// use view for a single query's complete formula. Inspection failures fail the action
    /// rather than silently omitting a query.
    /// </summary>
    [ServiceAction("list")]
    PowerQueryListResult List(IExcelBatch batch);

    /// <summary>
    /// Views the full M code and exact worksheet/Data Model load state of a Power Query
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to view</param>
    [ServiceAction("view")]
    PowerQueryViewResult View(IExcelBatch batch, [RequiredParameter] string queryName);

    /// <summary>
    /// Refreshes a Power Query to update its data with error detection using a caller-specified timeout
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to refresh</param>
    /// <param name="timeout">Public input is whole seconds from 0 through 2147483. Omitted or 0 uses the 30-minute data-operation default.</param>
    /// <param name="trace">Include a Power Query trace summary for this refresh</param>
    /// <param name="progress">Optional progress reporter</param>
    [ServiceAction("refresh")]
    PowerQueryRefreshResult Refresh(IExcelBatch batch, [RequiredParameter] string queryName, TimeSpan timeout = default, bool trace = false, IProgress<ProgressInfo>? progress = null);

    /// <summary>
    /// Gets the current load configuration of a Power Query
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    [ServiceAction("get-load-config")]
    PowerQueryLoadConfigResult GetLoadConfig(IExcelBatch batch, [RequiredParameter] string queryName);

    /// <summary>
    /// Reads the workbook's configured FastCombine setting. Excel's global privacy option may override
    /// whether the setting is applied when combining data sources.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    [ServiceAction("get-privacy-settings")]
    PowerQueryPrivacySettingsResult GetPrivacySettings(IExcelBatch batch);

    /// <summary>
    /// Sets the workbook's configured FastCombine value, which controls whether Power Query ignores
    /// privacy levels when combining data sources. Excel's global privacy option may override whether
    /// the setting is applied. Setting it to true may allow data from different sources to be combined
    /// and sent between them.
    /// The change is stored in the workbook when it is saved.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="ignorePrivacyLevels">Configured workbook FastCombine value: true to ignore privacy levels, false not to ignore them. Excel's global privacy option may override whether the setting is applied.</param>
    [ServiceAction("set-privacy-settings")]
    PowerQueryPrivacySettingsResult SetPrivacySettings(
        IExcelBatch batch,
        [RequiredParameter] bool ignorePrivacyLevels);

    /// <summary>
    /// Deletes a Power Query and only the destinations with the exact mashup Location identity.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to delete</param>
    /// <exception cref="InvalidOperationException">Thrown when the Power Query is not found or cannot be deleted</exception>
    [ServiceAction("delete")]
    OperationResult Delete(IExcelBatch batch, [RequiredParameter] string queryName);

    /// <summary>
    /// Creates a new Power Query and loads its selected destination in one call.
    /// This is not a transaction: load failure keeps the query and any created destinations.
    /// The error includes partialState with the inspected query/load state; nothing is rolled back.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name for the new query</param>
    /// <param name="mCode">Raw M code. Public callers must supply either inline mCode or a readable mCodeFile, not both.</param>
    /// <param name="loadMode">Load destination mode</param>
    /// <param name="targetSheet">Target worksheet name (required for LoadToTable and LoadToBoth; defaults to query name when omitted)</param>
    /// <param name="targetCellAddress">Optional target cell address for worksheet loads (e.g., "B5"). Required when loading to an existing worksheet with other data.</param>
    /// <param name="formatMCode">Whether to send M code to the remote powerqueryformatter.com service before saving. Defaults to false to preserve privacy.</param>
    /// <exception cref="InvalidOperationException">Thrown when query cannot be created, M code is invalid, or load operation fails</exception>
    PowerQueryCreateResult Create(
        IExcelBatch batch,
        [RequiredParameter] string queryName,
        [RequiredParameter][FileOrValue] string mCode,
        [FromString("loadDestination")] PowerQueryLoadMode loadMode = PowerQueryLoadMode.LoadToTable,
        string? targetSheet = null,
        string? targetCellAddress = null,
        bool formatMCode = false);

    /// <summary>
    /// Updates M code. Optionally refreshes loaded data.
    /// A refresh failure keeps the saved M code and reports partialState; loaded data may remain stale.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to update</param>
    /// <param name="mCode">Raw M code. Public callers must supply either inline mCode or a readable mCodeFile, not both.</param>
    /// <param name="refresh">Whether to refresh data after update (default: true)</param>
    /// <param name="formatMCode">Whether to send M code to the remote powerqueryformatter.com service before saving. Defaults to false to preserve privacy.</param>
    /// <exception cref="InvalidOperationException">Thrown when the query is not found, M code is invalid, or refresh fails</exception>
    PowerQueryOperationResult Update(IExcelBatch batch, [RequiredParameter] string queryName, [RequiredParameter][FileOrValue] string mCode, bool refresh = true, bool formatMCode = false);

    /// <summary>
    /// Atomically sets load destination and refreshes data
    /// Replaces multi-step workflow (configure + refresh in ONE operation)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="loadMode">Load destination mode</param>
    /// <param name="targetSheet">Target worksheet name (required for LoadToTable and LoadToBoth)</param>
    /// <param name="targetCellAddress">Optional target cell address (e.g., "B5"). Required when loading to an existing worksheet to avoid clearing other content.</param>
    /// <exception cref="InvalidOperationException">Thrown when the query is not found, load destination is invalid, or refresh fails</exception>
    PowerQueryLoadResult LoadTo(
        IExcelBatch batch,
        [RequiredParameter] string queryName,
        [FromString("loadDestination")] PowerQueryLoadMode loadMode,
        string? targetSheet = null,
        string? targetCellAddress = null);

    // ValidateSyntaxAsync removed - Excel doesn't validate M code syntax at query creation time.
    // Validation only happens during refresh, making syntax-only validation unreliable.

    /// <summary>
    /// Refreshes every stored query that has a worksheet, Data Model, or workbook
    /// connection load. Parameter and connection-only staging queries have nothing to
    /// refresh on their own and are listed in skippedQueries; the loaded queries that
    /// use them re-evaluate them. A failing query is listed in failedQueries and the
    /// remaining queries are still refreshed. PivotTables backed by successfully refreshed
    /// query loads are then refreshed by default. Set refreshPivotTables=false to skip them.
    /// success is false when any query or PivotTable failed; successful refreshes are not
    /// rolled back. Timeouts, cancellation, and Excel disconnects stop the whole operation.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="timeout">Public input is whole seconds from 0 through 2147483. Omitted or 0 uses the 30-minute data-operation default.</param>
    /// <param name="refreshPivotTables">Refresh PivotTables dependent on successfully refreshed queries (default true).</param>
    /// <param name="trace">Include a Power Query trace summary for this refresh</param>
    /// <param name="progress">Optional progress reporter</param>
    PowerQueryRefreshAllResult RefreshAll(
        IExcelBatch batch,
        TimeSpan timeout = default,
        bool refreshPivotTables = true,
        bool trace = false,
        IProgress<ProgressInfo>? progress = null);

    /// <summary>
    /// Renames a Power Query using trim + case-insensitive uniqueness semantics.
    /// - Names are normalized (trimmed) before comparison.
    /// - No-op success when normalized names are equal.
    /// - Case-only rename attempts COM rename (Excel decides outcome).
    /// - No auto-save.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="oldName">Current name of the query</param>
    /// <param name="newName">New name for the query</param>
    /// <returns>Result with objectType=power-query and normalized names</returns>
    [ServiceAction("rename")]
    RenameResult Rename(IExcelBatch batch, [RequiredParameter] string oldName, [RequiredParameter] string newName);

    /// <summary>
    /// Converts query to connection-only by removing data from its exact destinations.
    /// Removes worksheet ListObjects AND Data Model connections, but keeps the query definition.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to unload</param>
    /// <returns>Operation result</returns>
    [ServiceAction("unload")]
    OperationResult Unload(IExcelBatch batch, [RequiredParameter] string queryName);

    /// <summary>
    /// Evaluates M code and returns the result data without creating a permanent query.
    /// Creates a temporary query, executes it, reads the results, then cleans up every
    /// exact temporary mashup connection regardless of Excel's display name.
    /// Cleanup failures are surfaced and include recovery guidance.
    /// Useful for testing M code snippets and getting preview data.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="mCode">M code to evaluate. Public callers must supply either inline mCode or a readable mCodeFile, not both.</param>
    /// <returns>Result containing evaluated data as columns/rows</returns>
    /// <exception cref="InvalidOperationException">Thrown when M code has errors</exception>
    [ServiceAction("evaluate")]
    PowerQueryEvaluateResult Evaluate(IExcelBatch batch, [RequiredParameter][FileOrValue] string mCode);
}
