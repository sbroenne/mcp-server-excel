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
    /// Updates an existing Power Query with new M code
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to update</param>
    /// <param name="mCodeFile">Path to M code file</param>
    Task<OperationResult> UpdateAsync(IExcelBatch batch, string queryName, string mCodeFile);

    /// <summary>
    /// Exports a Power Query's M code to a file
    /// </summary>
    Task<OperationResult> ExportAsync(IExcelBatch batch, string queryName, string outputFile);

    /// <summary>
    /// Imports M code from a file to create a new Power Query
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name for the new query</param>
    /// <param name="mCodeFile">Path to M code file</param>
    /// <param name="loadDestination">Where to load query data: "worksheet" (default), "data-model", "both", or "connection-only"</param>
    /// <param name="worksheetName">Optional worksheet name when loadDestination is "worksheet" or "both". If not specified, uses query name</param>
    Task<OperationResult> ImportAsync(IExcelBatch batch, string queryName, string mCodeFile, string loadDestination = "worksheet", string? worksheetName = null);

    /// <summary>
    /// Refreshes a Power Query to update its data with error detection
    /// </summary>
    Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Refreshes a Power Query to update its data with error detection and timeout
    /// </summary>
    Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName, TimeSpan? timeout);

    /// <summary>
    /// Shows errors from Power Query operations
    /// </summary>
    Task<PowerQueryViewResult> ErrorsAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Loads a connection-only Power Query to a worksheet
    /// </summary>
    Task<OperationResult> LoadToAsync(IExcelBatch batch, string queryName, string sheetName);

    /// <summary>
    /// Sets a Power Query to Connection Only mode (no data loaded to worksheet)
    /// </summary>
    Task<OperationResult> SetConnectionOnlyAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Sets a Power Query to Load to Table mode (data loaded to worksheet)
    /// ATOMIC OPERATION: Configures query AND refreshes to ensure data persists
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="sheetName">Target worksheet name</param>
    /// <returns>PowerQueryLoadToTableResult with verification of data actually loaded</returns>
    Task<PowerQueryLoadToTableResult> SetLoadToTableAsync(IExcelBatch batch, string queryName, string sheetName);

    /// <summary>
    /// Sets a Power Query to Load to Data Model mode (data loaded to PowerPivot)
    /// ATOMIC OPERATION: Configures query AND refreshes to ensure data persists
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <returns>PowerQueryLoadToDataModelResult with verification of data actually loaded</returns>
    Task<PowerQueryLoadToDataModelResult> SetLoadToDataModelAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Sets a Power Query to Load to Both modes (table + data model)
    /// ATOMIC OPERATION: Configures query AND refreshes to ensure data persists to both destinations
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="sheetName">Target worksheet name</param>
    /// <returns>PowerQueryLoadToBothResult with verification of data loaded to both table and Data Model</returns>
    Task<PowerQueryLoadToBothResult> SetLoadToBothAsync(IExcelBatch batch, string queryName, string sheetName);

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

    /// <summary>
    /// Evaluates M code expressions interactively
    /// </summary>
    Task<PowerQueryViewResult> EvalAsync(IExcelBatch batch, string mExpression);

    // =========================================================================
    // PHASE 1 METHODS - Atomic Operations for Improved Workflows
    // =========================================================================

    /// <summary>
    /// Creates a new Power Query by importing M code and loading data atomically
    /// PHASE 1: Replaces ImportAsync workflow (import + configure + refresh in ONE operation)
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
    /// PHASE 1: Explicit separation of update vs refresh for atomic workflows
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to update</param>
    /// <param name="mCodeFile">Path to M code file</param>
    /// <returns>OperationResult with update status</returns>
    Task<OperationResult> UpdateMCodeAsync(IExcelBatch batch, string queryName, string mCodeFile);

    /// <summary>
    /// Atomically sets load destination and refreshes data
    /// PHASE 1: Replaces SetLoadTo* + RefreshAsync workflow (configure + refresh in ONE operation)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="loadMode">Load destination mode</param>
    /// <param name="targetSheet">Target worksheet name (required for LoadToTable and LoadToBoth)</param>
    /// <returns>PowerQueryLoadResult with configuration and refresh tracking</returns>
    Task<PowerQueryLoadResult> LoadToAsync(IExcelBatch batch, string queryName, PowerQueryLoadMode loadMode, string? targetSheet = null);

    /// <summary>
    /// Converts a query to connection-only mode (removes all data loads)
    /// PHASE 1: Explicit unload operation (inverse of LoadToAsync)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <returns>OperationResult with unload status</returns>
    Task<OperationResult> UnloadAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Validates M code syntax without creating a permanent query
    /// PHASE 1: Pre-flight validation for safer imports
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="mCodeFile">Path to M code file</param>
    /// <returns>PowerQueryValidationResult with syntax validation details</returns>
    Task<PowerQueryValidationResult> ValidateSyntaxAsync(IExcelBatch batch, string mCodeFile);

    /// <summary>
    /// Convenience method: Updates M code then refreshes data
    /// PHASE 1: Common workflow as single operation (UpdateMCodeAsync + RefreshAsync)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="mCodeFile">Path to M code file</param>
    /// <returns>OperationResult with combined update and refresh status</returns>
    Task<OperationResult> UpdateAndRefreshAsync(IExcelBatch batch, string queryName, string mCodeFile);

    /// <summary>
    /// Refreshes all Power Queries in the workbook
    /// PHASE 1: Batch refresh with error tracking
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <returns>OperationResult with batch refresh summary</returns>
    Task<OperationResult> RefreshAllAsync(IExcelBatch batch);
}

