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
    /// Lists all Power Query queries in the workbook (filePath-based API)
    /// </summary>
    Task<PowerQueryListResult> ListAsync(string filePath);

    /// <summary>
    /// Views the M code of a Power Query
    /// </summary>
    Task<PowerQueryViewResult> ViewAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Views the M code of a Power Query (filePath-based API)
    /// </summary>
    Task<PowerQueryViewResult> ViewAsync(string filePath, string queryName);

    /// <summary>
    /// Exports a Power Query's M code to a file
    /// </summary>
    Task<OperationResult> ExportAsync(IExcelBatch batch, string queryName, string outputFile);

    /// <summary>
    /// Exports a Power Query's M code to a file (filePath-based API)
    /// </summary>
    Task<OperationResult> ExportAsync(string filePath, string queryName, string outputFile);

    /// <summary>
    /// Refreshes a Power Query to update its data with error detection
    /// </summary>
    Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Refreshes a Power Query to update its data with error detection (filePath-based API)
    /// </summary>
    Task<PowerQueryRefreshResult> RefreshAsync(string filePath, string queryName);

    /// <summary>
    /// Refreshes a Power Query to update its data with error detection and timeout
    /// </summary>
    Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName, TimeSpan? timeout);

    /// <summary>
    /// Refreshes a Power Query to update its data with error detection and timeout (filePath-based API)
    /// </summary>
    Task<PowerQueryRefreshResult> RefreshAsync(string filePath, string queryName, TimeSpan? timeout);

    /// <summary>
    /// Gets the current load configuration of a Power Query
    /// </summary>
    Task<PowerQueryLoadConfigResult> GetLoadConfigAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Gets the current load configuration of a Power Query (filePath-based API)
    /// </summary>
    Task<PowerQueryLoadConfigResult> GetLoadConfigAsync(string filePath, string queryName);

    /// <summary>
    /// Deletes a Power Query from the workbook
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Deletes a Power Query from the workbook (filePath-based API)
    /// </summary>
    Task<OperationResult> DeleteAsync(string filePath, string queryName);

    /// <summary>
    /// Lists available data sources (Excel.CurrentWorkbook() sources: tables and named ranges)
    /// </summary>
    Task<WorksheetListResult> ListExcelSourcesAsync(IExcelBatch batch);

    /// <summary>
    /// Lists available data sources (Excel.CurrentWorkbook() sources: tables and named ranges) (filePath-based API)
    /// </summary>
    Task<WorksheetListResult> ListExcelSourcesAsync(string filePath);

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
    /// Creates a new Power Query by importing M code and loading data atomically (filePath-based API)
    /// </summary>
    Task<PowerQueryCreateResult> CreateAsync(string filePath, string queryName, string mCodeFile, PowerQueryLoadMode loadMode = PowerQueryLoadMode.LoadToTable, string? targetSheet = null);

    /// <summary>
    /// Updates M code and refreshes data atomically
    /// Complete operation: Updates query formula AND reloads fresh data (no stale data footgun)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to update</param>
    /// <param name="mCodeFile">Path to M code file</param>
    /// <returns>OperationResult with update and refresh status</returns>
    Task<OperationResult> UpdateAsync(IExcelBatch batch, string queryName, string mCodeFile);

    /// <summary>
    /// Updates M code and refreshes data atomically (filePath-based API)
    /// </summary>
    Task<OperationResult> UpdateAsync(string filePath, string queryName, string mCodeFile);

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
    /// Atomically sets load destination and refreshes data (filePath-based API)
    /// </summary>
    Task<PowerQueryLoadResult> LoadToAsync(string filePath, string queryName, PowerQueryLoadMode loadMode, string? targetSheet = null);

    /// <summary>
    /// Converts a query to connection-only mode (removes all data loads)
    /// Explicit unload operation (inverse of LoadToAsync)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <returns>OperationResult with unload status</returns>
    Task<OperationResult> UnloadAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Converts a query to connection-only mode (removes all data loads) (filePath-based API)
    /// </summary>
    Task<OperationResult> UnloadAsync(string filePath, string queryName);

    // ValidateSyntaxAsync removed - Excel doesn't validate M code syntax at query creation time.
    // Validation only happens during refresh, making syntax-only validation unreliable.

    /// <summary>
    /// Refreshes all Power Queries in the workbook
    /// Batch refresh with error tracking
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <returns>OperationResult with batch refresh summary</returns>
    Task<OperationResult> RefreshAllAsync(IExcelBatch batch);

    /// <summary>
    /// Refreshes all Power Queries in the workbook (filePath-based API)
    /// </summary>
    Task<OperationResult> RefreshAllAsync(string filePath);
}

