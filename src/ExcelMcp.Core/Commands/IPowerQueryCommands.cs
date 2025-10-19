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
    Task<OperationResult> Update(string filePath, string queryName, string mCodeFile);
    
    /// <summary>
    /// Exports a Power Query's M code to a file
    /// </summary>
    Task<OperationResult> Export(string filePath, string queryName, string outputFile);
    
    /// <summary>
    /// Imports M code from a file to create a new Power Query
    /// </summary>
    Task<OperationResult> Import(string filePath, string queryName, string mCodeFile);
    
    /// <summary>
    /// Refreshes a Power Query to update its data
    /// </summary>
    OperationResult Refresh(string filePath, string queryName);
    
    /// <summary>
    /// Shows errors from Power Query operations
    /// </summary>
    PowerQueryViewResult Errors(string filePath, string queryName);
    
    /// <summary>
    /// Loads a connection-only Power Query to a worksheet
    /// </summary>
    OperationResult LoadTo(string filePath, string queryName, string sheetName);
    
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
