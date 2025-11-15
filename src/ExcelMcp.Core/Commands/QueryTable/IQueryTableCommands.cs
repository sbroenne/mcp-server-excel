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
    QueryTableListResult List(IExcelBatch batch);

    /// <summary>
    /// Gets detailed information about a specific QueryTable
    /// </summary>
    QueryTableInfoResult Read(IExcelBatch batch, string queryTableName);

    /// <summary>
    /// Creates a QueryTable from an existing connection
    /// </summary>
    OperationResult CreateFromConnection(IExcelBatch batch, string sheetName,
        string queryTableName, string connectionName, string range = "A1",
        QueryTableCreateOptions? options = null);

    /// <summary>
    /// Creates a QueryTable from a Power Query (leverages existing PowerQueryHelpers)
    /// </summary>
    OperationResult CreateFromQuery(IExcelBatch batch, string sheetName,
        string queryTableName, string queryName, string range = "A1",
        QueryTableCreateOptions? options = null);

    /// <summary>
    /// Refreshes a QueryTable using synchronous pattern for guaranteed persistence
    /// </summary>
    OperationResult Refresh(IExcelBatch batch, string queryTableName, TimeSpan? timeout = null);

    /// <summary>
    /// Updates QueryTable properties (refresh settings, formatting options)
    /// </summary>
    OperationResult UpdateProperties(IExcelBatch batch, string queryTableName,
        QueryTableUpdateOptions options);

    /// <summary>
    /// Deletes a QueryTable from the workbook
    /// </summary>
    OperationResult Delete(IExcelBatch batch, string queryTableName);

    /// <summary>
    /// Refreshes all QueryTables in the workbook using synchronous pattern
    /// </summary>
    OperationResult RefreshAll(IExcelBatch batch, TimeSpan? timeout = null);
}
