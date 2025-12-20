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
    /// <exception cref="InvalidOperationException">Thrown when the Power Query is not found or cannot be deleted</exception>
    void Delete(IExcelBatch batch, string queryName);

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
        string queryName,
        string mCode,
        PowerQueryLoadMode loadMode = PowerQueryLoadMode.LoadToTable,
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
    void Update(IExcelBatch batch, string queryName, string mCode, bool refresh = true);

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
    /// <param name="oldName">Existing query name</param>
    /// <param name="newName">Desired new name</param>
    /// <returns>Result with objectType=power-query and normalized names</returns>
    RenameResult Rename(IExcelBatch batch, string oldName, string newName);
}


