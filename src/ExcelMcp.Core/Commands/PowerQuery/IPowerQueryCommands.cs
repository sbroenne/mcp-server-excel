using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query (M code) management - create, edit, execute, and load queries.
/// Use for ETL operations, data transformation, and connecting to external data sources.
/// </summary>
[ServiceCategory("powerquery", "PowerQuery")]
[McpTool("excel_powerquery")]
public interface IPowerQueryCommands
{
    /// <summary>
    /// Lists all Power Query queries in the workbook
    /// </summary>
    [ServiceAction("list")]
    PowerQueryListResult List(IExcelBatch batch);

    /// <summary>
    /// Views the M code of a Power Query
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
    /// <param name="timeout">Maximum time to wait for refresh</param>
    [ServiceAction("refresh")]
    PowerQueryRefreshResult Refresh(IExcelBatch batch, [RequiredParameter] string queryName, TimeSpan timeout);

    /// <summary>
    /// Gets the current load configuration of a Power Query
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    [ServiceAction("get-load-config")]
    PowerQueryLoadConfigResult GetLoadConfig(IExcelBatch batch, [RequiredParameter] string queryName);

    /// <summary>
    /// Deletes a Power Query from the workbook
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to delete</param>
    /// <exception cref="InvalidOperationException">Thrown when the Power Query is not found or cannot be deleted</exception>
    [ServiceAction("delete")]
    void Delete(IExcelBatch batch, [RequiredParameter] string queryName);

    /// <summary>
    /// Creates a new Power Query by importing M code and loading data atomically
    /// Replaces multi-step workflow (import + configure + refresh in ONE operation)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name for the new query</param>
    /// <param name="mCode">Raw M code (inline string)</param>
    /// <param name="loadMode">Load destination mode</param>
    /// <param name="targetSheet">Target worksheet name (required for LoadToTable and LoadToBoth; defaults to query name when omitted)</param>
    /// <param name="targetCellAddress">Optional target cell address for worksheet loads (e.g., "B5"). Required when loading to an existing worksheet with other data.</param>
    /// <exception cref="InvalidOperationException">Thrown when query cannot be created, M code is invalid, or load operation fails</exception>
    void Create(
        IExcelBatch batch,
        [RequiredParameter] string queryName,
        [RequiredParameter][FileOrValue] string mCode,
        [FromString("loadDestination")] PowerQueryLoadMode loadMode = PowerQueryLoadMode.LoadToTable,
        string? targetSheet = null,
        string? targetCellAddress = null);

    /// <summary>
    /// Updates M code. Optionally refreshes loaded data.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to update</param>
    /// <param name="mCode">Raw M code (inline string)</param>
    /// <param name="refresh">Whether to refresh data after update (default: true)</param>
    /// <exception cref="InvalidOperationException">Thrown when the query is not found, M code is invalid, or refresh fails</exception>
    void Update(IExcelBatch batch, [RequiredParameter] string queryName, [RequiredParameter][FileOrValue] string mCode, bool refresh = true);

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
    void LoadTo(
        IExcelBatch batch,
        [RequiredParameter] string queryName,
        [FromString("loadDestination")] PowerQueryLoadMode loadMode,
        string? targetSheet = null,
        string? targetCellAddress = null);

    // ValidateSyntaxAsync removed - Excel doesn't validate M code syntax at query creation time.
    // Validation only happens during refresh, making syntax-only validation unreliable.

    /// <summary>
    /// Refreshes all Power Queries in the workbook
    /// Batch refresh with error tracking
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <exception cref="InvalidOperationException">Thrown when any Power Query fails to refresh</exception>
    void RefreshAll(IExcelBatch batch);

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
    /// Converts query to connection-only by removing data from all destinations.
    /// Removes worksheet ListObjects AND Data Model connections, but keeps the query definition.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to unload</param>
    /// <returns>Operation result</returns>
    [ServiceAction("unload")]
    OperationResult Unload(IExcelBatch batch, [RequiredParameter] string queryName);

    /// <summary>
    /// Evaluates M code and returns the result data without creating a permanent query.
    /// Creates a temporary query, executes it, reads the results, then cleans up.
    /// Useful for testing M code snippets and getting preview data.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="mCode">M code to evaluate</param>
    /// <returns>Result containing evaluated data as columns/rows</returns>
    /// <exception cref="InvalidOperationException">Thrown when M code has errors</exception>
    [ServiceAction("evaluate")]
    PowerQueryEvaluateResult Evaluate(IExcelBatch batch, [RequiredParameter][FileOrValue] string mCode);
}
