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
    Task<PowerQueryListResult> ListAsync(IExcelBatch batch);

    /// <summary>
    /// Views the M code of a Power Query
    /// </summary>
    Task<PowerQueryViewResult> ViewAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Exports a Power Query's M code to a file
    /// </summary>
    Task<OperationResult> ExportAsync(IExcelBatch batch, string queryName, string outputFile);

    /// <summary>
    /// Refreshes a Power Query to update its data with error detection
    /// </summary>
    Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Refreshes a Power Query to update its data with error detection and timeout
    /// </summary>
    Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName, TimeSpan? timeout);

    /// <summary>
    /// Gets the current load configuration of a Power Query
    /// </summary>
    Task<PowerQueryLoadConfigResult> GetLoadConfigAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Deletes a Power Query from the workbook
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Lists available data sources (Excel.CurrentWorkbook() sources: tables and named ranges)
    /// </summary>
    Task<WorksheetListResult> ListExcelSourcesAsync(IExcelBatch batch);

    // =========================================================================
    // ATOMIC OPERATIONS - Improved Workflows
    // =========================================================================

    /// <summary>
    /// Creates a new Power Query by importing M code and loading data atomically
    /// Replaces multi-step workflow (import + configure + refresh in ONE operation)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name for the new query</param>
    /// <param name="mCodeFile">Path to M code file</param>
    /// <param name="loadMode">Load destination mode</param>
    /// <param name="targetSheet">Target worksheet name (required for LoadToTable and LoadToBoth)</param>
    /// <returns>PowerQueryCreateResult with creation and load tracking</returns>
    Task<PowerQueryCreateResult> CreateAsync(IExcelBatch batch, string queryName, string mCodeFile, PowerQueryLoadMode loadMode = PowerQueryLoadMode.LoadToTable, string? targetSheet = null);

    /// <summary>
    /// Updates only the M code formula of an existing query (no refresh)
    /// Explicit separation of update vs refresh for atomic workflows
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to update</param>
    /// <param name="mCodeFile">Path to M code file</param>
    /// <returns>OperationResult with update status</returns>
    Task<OperationResult> UpdateMCodeAsync(IExcelBatch batch, string queryName, string mCodeFile);

    /// <summary>
    /// Atomically sets load destination and refreshes data
    /// Replaces multi-step workflow (configure + refresh in ONE operation)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="loadMode">Load destination mode</param>
    /// <param name="targetSheet">Target worksheet name (required for LoadToTable and LoadToBoth)</param>
    /// <returns>PowerQueryLoadResult with configuration and refresh tracking</returns>
    Task<PowerQueryLoadResult> LoadToAsync(IExcelBatch batch, string queryName, PowerQueryLoadMode loadMode, string? targetSheet = null);

    /// <summary>
    /// Converts a query to connection-only mode (removes all data loads)
    /// Explicit unload operation (inverse of LoadToAsync)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <returns>OperationResult with unload status</returns>
    Task<OperationResult> UnloadAsync(IExcelBatch batch, string queryName);

    // ValidateSyntaxAsync removed - Excel doesn't validate M code syntax at query creation time.
    // Validation only happens during refresh, making syntax-only validation unreliable.

    /// <summary>
    /// Convenience method: Updates M code then refreshes data
    /// Common workflow as single operation (UpdateMCodeAsync + RefreshAsync)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="mCodeFile">Path to M code file</param>
    /// <returns>OperationResult with combined update and refresh status</returns>
    Task<OperationResult> UpdateAndRefreshAsync(IExcelBatch batch, string queryName, string mCodeFile);

    /// <summary>
    /// Refreshes all Power Queries in the workbook
    /// Batch refresh with error tracking
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <returns>OperationResult with batch refresh summary</returns>
    Task<OperationResult> RefreshAllAsync(IExcelBatch batch);
}

