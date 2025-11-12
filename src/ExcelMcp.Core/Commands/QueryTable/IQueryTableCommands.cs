using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.QueryTable;

/// <summary>
/// QueryTable management commands for Excel automation
/// Provides CRUD operations for Excel QueryTables - simple data imports with reliable persistence
/// </summary>
public interface IQueryTableCommands
{
    // FilePath-based API (new pattern)
    /// <summary>
    /// Lists all QueryTables in the workbook with connection and range information
    /// </summary>
    Task<QueryTableListResult> ListAsync(string filePath);

    /// <summary>
    /// Gets detailed information about a specific QueryTable
    /// </summary>
    Task<QueryTableInfoResult> GetAsync(string filePath, string queryTableName);

    /// <summary>
    /// Deletes a QueryTable from the workbook
    /// </summary>
    Task<OperationResult> DeleteAsync(string filePath, string queryTableName);

    // Batch-based API (existing - will be removed in final cleanup)
    /// <summary>
    /// Lists all QueryTables in the workbook with connection and range information
    /// </summary>
    Task<QueryTableListResult> ListAsync(IExcelBatch batch);

    /// <summary>
    /// Gets detailed information about a specific QueryTable
    /// </summary>
    Task<QueryTableInfoResult> GetAsync(IExcelBatch batch, string queryTableName);

    /// <summary>
    /// Deletes a QueryTable from the workbook
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string queryTableName);

    // Complex operations still use batch-based API (will be converted later with ExecuteAsync pattern)
    /// <summary>
    /// Creates a QueryTable from an existing connection
    /// </summary>
    Task<OperationResult> CreateFromConnectionAsync(IExcelBatch batch, string sheetName,
        string queryTableName, string connectionName, string range = "A1",
        QueryTableCreateOptions? options = null);

    /// <summary>
    /// Creates a QueryTable from a Power Query (leverages existing PowerQueryHelpers)
    /// </summary>
    Task<OperationResult> CreateFromQueryAsync(IExcelBatch batch, string sheetName,
        string queryTableName, string queryName, string range = "A1",
        QueryTableCreateOptions? options = null);

    /// <summary>
    /// Refreshes a QueryTable using synchronous pattern for guaranteed persistence
    /// </summary>
    Task<OperationResult> RefreshAsync(IExcelBatch batch, string queryTableName, TimeSpan? timeout = null);

    /// <summary>
    /// Updates QueryTable properties (refresh settings, formatting options)
    /// </summary>
    Task<OperationResult> UpdatePropertiesAsync(IExcelBatch batch, string queryTableName,
        QueryTableUpdateOptions options);

    /// <summary>
    /// Refreshes all QueryTables in the workbook using synchronous pattern
    /// </summary>
    Task<OperationResult> RefreshAllAsync(IExcelBatch batch, TimeSpan? timeout = null);
}
