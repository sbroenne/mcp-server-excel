using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;

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
    /// <param name="privacyLevel">Optional privacy level for data combining. If not specified and privacy error occurs, operation returns PowerQueryPrivacyErrorResult for user to choose.</param>
    Task<OperationResult> UpdateAsync(IExcelBatch batch, string queryName, string mCodeFile, PowerQueryPrivacyLevel? privacyLevel = null);

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
    /// <param name="privacyLevel">Optional privacy level for data combining. If not specified and privacy error occurs, operation returns PowerQueryPrivacyErrorResult for user to choose.</param>
    /// <param name="loadToWorksheet">Automatically load query data to a worksheet (default: true). When true, validates query by executing it.</param>
    /// <param name="worksheetName">Optional worksheet name to load data to. If not specified, uses query name</param>
    Task<OperationResult> ImportAsync(IExcelBatch batch, string queryName, string mCodeFile, PowerQueryPrivacyLevel? privacyLevel = null, bool loadToWorksheet = true, string? worksheetName = null);

    /// <summary>
    /// Refreshes a Power Query to update its data with error detection
    /// </summary>
    Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName);

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
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="sheetName">Target worksheet name</param>
    /// <param name="privacyLevel">Optional privacy level for data combining. If not specified and privacy error occurs, operation returns PowerQueryPrivacyErrorResult for user to choose.</param>
    Task<OperationResult> SetLoadToTableAsync(IExcelBatch batch, string queryName, string sheetName, PowerQueryPrivacyLevel? privacyLevel = null);

    /// <summary>
    /// Sets a Power Query to Load to Data Model mode (data loaded to PowerPivot)
    /// ATOMIC OPERATION: Configures query AND refreshes to ensure data persists
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="privacyLevel">Optional privacy level for data combining. If not specified and privacy error occurs, operation returns PowerQueryPrivacyErrorResult for user to choose.</param>
    /// <returns>PowerQueryLoadToDataModelResult with verification of data actually loaded</returns>
    Task<PowerQueryLoadToDataModelResult> SetLoadToDataModelAsync(IExcelBatch batch, string queryName, PowerQueryPrivacyLevel? privacyLevel = null);

    /// <summary>
    /// Sets a Power Query to Load to Both modes (table + data model)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="sheetName">Target worksheet name</param>
    /// <param name="privacyLevel">Optional privacy level for data combining. If not specified and privacy error occurs, operation returns PowerQueryPrivacyErrorResult for user to choose.</param>
    Task<OperationResult> SetLoadToBothAsync(IExcelBatch batch, string queryName, string sheetName, PowerQueryPrivacyLevel? privacyLevel = null);

    /// <summary>
    /// Gets the current load configuration of a Power Query
    /// </summary>
    Task<PowerQueryLoadConfigResult> GetLoadConfigAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Deletes a Power Query from the workbook
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string queryName);

    /// <summary>
    /// Lists available data sources (Excel.CurrentWorkbook() sources)
    /// </summary>
    Task<WorksheetListResult> SourcesAsync(IExcelBatch batch);

    /// <summary>
    /// Tests connectivity to a Power Query data source
    /// </summary>
    Task<OperationResult> TestAsync(IExcelBatch batch, string sourceName);

    /// <summary>
    /// Previews sample data from a Power Query data source
    /// </summary>
    Task<WorksheetDataResult> PeekAsync(IExcelBatch batch, string sourceName);

    /// <summary>
    /// Evaluates M code expressions interactively
    /// </summary>
    Task<PowerQueryViewResult> EvalAsync(IExcelBatch batch, string mExpression);
}
