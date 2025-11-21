using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query management commands
/// </summary>
public interface IPowerQueryCommands
{
    /// <summary>
    /// Lists all Power Query queries in the workbook
    /// </summary>
    PowerQueryListResult List(IExcelBatch batch);

    /// <summary>
    /// Views the M code of a Power Query
    /// </summary>
    PowerQueryViewResult View(IExcelBatch batch, string queryName);

    /// <summary>
    /// Refreshes a Power Query to update its data with error detection using a caller-specified timeout
    /// </summary>
    PowerQueryRefreshResult Refresh(IExcelBatch batch, string queryName, TimeSpan timeout);

    /// <summary>
    /// Gets the current load configuration of a Power Query
    /// </summary>
    PowerQueryLoadConfigResult GetLoadConfig(IExcelBatch batch, string queryName);

    /// <summary>
    /// Deletes a Power Query from the workbook
    /// </summary>
    OperationResult Delete(IExcelBatch batch, string queryName);

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
    /// <returns>PowerQueryCreateResult with creation and load tracking</returns>
    PowerQueryCreateResult Create(
        IExcelBatch batch,
        string queryName,
        string mCode,
        PowerQueryLoadMode loadMode = PowerQueryLoadMode.LoadToTable,
        string? targetSheet = null,
        string? targetCellAddress = null);

    /// <summary>
    /// Updates M code and refreshes data atomically
    /// Complete operation: Updates query formula AND reloads fresh data (no stale data footgun)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to update</param>
    /// <param name="mCode">Raw M code (inline string)</param>
    /// <returns>OperationResult with update and refresh status</returns>
    OperationResult Update(IExcelBatch batch, string queryName, string mCode);

    /// <summary>
    /// Atomically sets load destination and refreshes data
    /// Replaces multi-step workflow (configure + refresh in ONE operation)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="loadMode">Load destination mode</param>
    /// <param name="targetSheet">Target worksheet name (required for LoadToTable and LoadToBoth)</param>
    /// <param name="targetCellAddress">Optional target cell address (e.g., "B5"). Required when loading to an existing worksheet to avoid clearing other content.</param>
    /// <returns>PowerQueryLoadResult with configuration and refresh tracking</returns>
    PowerQueryLoadResult LoadTo(
        IExcelBatch batch,
        string queryName,
        PowerQueryLoadMode loadMode,
        string? targetSheet = null,
        string? targetCellAddress = null);

    // ValidateSyntaxAsync removed - Excel doesn't validate M code syntax at query creation time.
    // Validation only happens during refresh, making syntax-only validation unreliable.

    /// <summary>
    /// Refreshes all Power Queries in the workbook
    /// Batch refresh with error tracking
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <returns>OperationResult with batch refresh summary</returns>
    OperationResult RefreshAll(IExcelBatch batch);
}


