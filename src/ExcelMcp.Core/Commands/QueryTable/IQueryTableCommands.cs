using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.QueryTable;

/// <summary>
/// QueryTable management commands for Excel automation
/// Provides CRUD operations for Excel QueryTables - simple data imports with reliable persistence
/// </summary>
public interface IQueryTableCommands
{
    /// <summary>
    /// Lists all QueryTables in the workbook with connection and range information
    /// </summary>
    Task<QueryTableListResult> ListAsync(IExcelBatch batch);

    /// <summary>
    /// Lists all QueryTables in the workbook with connection and range information (filePath-based API)
    /// </summary>
    Task<QueryTableListResult> ListAsync(string filePath);

    /// <summary>
    /// Gets detailed information about a specific QueryTable
    /// </summary>
    Task<QueryTableInfoResult> GetAsync(IExcelBatch batch, string queryTableName);

    /// <summary>
    /// Gets detailed information about a specific QueryTable (filePath-based API)
    /// </summary>
    Task<QueryTableInfoResult> GetAsync(string filePath, string queryTableName);

    /// <summary>
    /// Creates a QueryTable from an existing connection
    /// </summary>
    Task<OperationResult> CreateFromConnectionAsync(IExcelBatch batch, string sheetName,
        string queryTableName, string connectionName, string range = "A1",
        QueryTableCreateOptions? options = null);

    /// <summary>
    /// Creates a QueryTable from an existing connection (filePath-based API)
    /// </summary>
    Task<OperationResult> CreateFromConnectionAsync(string filePath, string sheetName,
        string queryTableName, string connectionName, string range = "A1",
        QueryTableCreateOptions? options = null);

    /// <summary>
    /// Creates a QueryTable from a Power Query (leverages existing PowerQueryHelpers)
    /// </summary>
    Task<OperationResult> CreateFromQueryAsync(IExcelBatch batch, string sheetName,
        string queryTableName, string queryName, string range = "A1",
        QueryTableCreateOptions? options = null);

    /// <summary>
    /// Creates a QueryTable from a Power Query (filePath-based API)
    /// </summary>
    Task<OperationResult> CreateFromQueryAsync(string filePath, string sheetName,
        string queryTableName, string queryName, string range = "A1",
        QueryTableCreateOptions? options = null);

    /// <summary>
    /// Refreshes a QueryTable using synchronous pattern for guaranteed persistence
    /// </summary>
    Task<OperationResult> RefreshAsync(IExcelBatch batch, string queryTableName, TimeSpan? timeout = null);

    /// <summary>
    /// Refreshes a QueryTable using synchronous pattern for guaranteed persistence (filePath-based API)
    /// </summary>
    Task<OperationResult> RefreshAsync(string filePath, string queryTableName, TimeSpan? timeout = null);

    /// <summary>
    /// Updates QueryTable properties (refresh settings, formatting options)
    /// </summary>
    Task<OperationResult> UpdatePropertiesAsync(IExcelBatch batch, string queryTableName,
        QueryTableUpdateOptions options);

    /// <summary>
    /// Updates QueryTable properties (refresh settings, formatting options) (filePath-based API)
    /// </summary>
    Task<OperationResult> UpdatePropertiesAsync(string filePath, string queryTableName,
        QueryTableUpdateOptions options);

    /// <summary>
    /// Deletes a QueryTable from the workbook
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string queryTableName);

    /// <summary>
    /// Deletes a QueryTable from the workbook (filePath-based API)
    /// </summary>
    Task<OperationResult> DeleteAsync(string filePath, string queryTableName);

    /// <summary>
    /// Refreshes all QueryTables in the workbook using synchronous pattern
    /// </summary>
    Task<OperationResult> RefreshAllAsync(IExcelBatch batch, TimeSpan? timeout = null);

    /// <summary>
    /// Refreshes all QueryTables in the workbook using synchronous pattern (filePath-based API)
    /// </summary>
    Task<OperationResult> RefreshAllAsync(string filePath, TimeSpan? timeout = null);
}
