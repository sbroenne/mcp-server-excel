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
    PowerQueryListResult List(string filePath);

    /// <summary>
    /// Views the M code of a Power Query
    /// </summary>
    PowerQueryViewResult View(string filePath, string queryName);

    /// <summary>
    /// Updates an existing Power Query with new M code
    /// </summary>
    /// <param name="filePath">Excel file path</param>
    /// <param name="queryName">Name of the query to update</param>
    /// <param name="mCodeFile">Path to M code file</param>
    /// <param name="privacyLevel">Optional privacy level for data combining. If not specified and privacy error occurs, operation returns PowerQueryPrivacyErrorResult for user to choose.</param>
    Task<OperationResult> Update(string filePath, string queryName, string mCodeFile, PowerQueryPrivacyLevel? privacyLevel = null);

    /// <summary>
    /// Exports a Power Query's M code to a file
    /// </summary>
    Task<OperationResult> Export(string filePath, string queryName, string outputFile);

    /// <summary>
    /// Imports M code from a file to create a new Power Query
    /// </summary>
    /// <param name="filePath">Excel file path</param>
    /// <param name="queryName">Name for the new query</param>
    /// <param name="mCodeFile">Path to M code file</param>
    /// <param name="privacyLevel">Optional privacy level for data combining. If not specified and privacy error occurs, operation returns PowerQueryPrivacyErrorResult for user to choose.</param>
    /// <param name="loadToWorksheet">Automatically load query data to a worksheet (default: true). When true, validates query by executing it.</param>
    /// <param name="worksheetName">Optional worksheet name to load data to. If not specified, uses query name</param>
    Task<OperationResult> Import(string filePath, string queryName, string mCodeFile, PowerQueryPrivacyLevel? privacyLevel = null, bool loadToWorksheet = true, string? worksheetName = null);

    /// <summary>
    /// Refreshes a Power Query to update its data with error detection
    /// </summary>
    PowerQueryRefreshResult Refresh(string filePath, string queryName);

    /// <summary>
    /// Shows errors from Power Query operations
    /// </summary>
    PowerQueryViewResult Errors(string filePath, string queryName);

    /// <summary>
    /// Loads a connection-only Power Query to a worksheet
    /// </summary>
    OperationResult LoadTo(string filePath, string queryName, string sheetName);

    /// <summary>
    /// Sets a Power Query to Connection Only mode (no data loaded to worksheet)
    /// </summary>
    OperationResult SetConnectionOnly(string filePath, string queryName);

    /// <summary>
    /// Sets a Power Query to Load to Table mode (data loaded to worksheet)
    /// </summary>
    /// <param name="filePath">Excel file path</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="sheetName">Target worksheet name</param>
    /// <param name="privacyLevel">Optional privacy level for data combining. If not specified and privacy error occurs, operation returns PowerQueryPrivacyErrorResult for user to choose.</param>
    OperationResult SetLoadToTable(string filePath, string queryName, string sheetName, PowerQueryPrivacyLevel? privacyLevel = null);

    /// <summary>
    /// Sets a Power Query to Load to Data Model mode (data loaded to PowerPivot)
    /// </summary>
    /// <param name="filePath">Excel file path</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="privacyLevel">Optional privacy level for data combining. If not specified and privacy error occurs, operation returns PowerQueryPrivacyErrorResult for user to choose.</param>
    OperationResult SetLoadToDataModel(string filePath, string queryName, PowerQueryPrivacyLevel? privacyLevel = null);

    /// <summary>
    /// Sets a Power Query to Load to Both modes (table + data model)
    /// </summary>
    /// <param name="filePath">Excel file path</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="sheetName">Target worksheet name</param>
    /// <param name="privacyLevel">Optional privacy level for data combining. If not specified and privacy error occurs, operation returns PowerQueryPrivacyErrorResult for user to choose.</param>
    OperationResult SetLoadToBoth(string filePath, string queryName, string sheetName, PowerQueryPrivacyLevel? privacyLevel = null);

    /// <summary>
    /// Gets the current load configuration of a Power Query
    /// </summary>
    PowerQueryLoadConfigResult GetLoadConfig(string filePath, string queryName);

    /// <summary>
    /// Deletes a Power Query from the workbook
    /// </summary>
    OperationResult Delete(string filePath, string queryName);

    /// <summary>
    /// Lists available data sources (Excel.CurrentWorkbook() sources)
    /// </summary>
    WorksheetListResult Sources(string filePath);

    /// <summary>
    /// Tests connectivity to a Power Query data source
    /// </summary>
    OperationResult Test(string filePath, string sourceName);

    /// <summary>
    /// Previews sample data from a Power Query data source
    /// </summary>
    WorksheetDataResult Peek(string filePath, string sourceName);

    /// <summary>
    /// Evaluates M code expressions interactively
    /// </summary>
    PowerQueryViewResult Eval(string filePath, string mExpression);
}
