using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Excel Table (ListObject) management commands
/// </summary>
public interface ITableCommands
{
    /// <summary>
    /// Lists all Excel Tables in the workbook
    /// </summary>
    TableListResult List(string filePath);

    /// <summary>
    /// Creates a new Excel Table from a range
    /// </summary>
    OperationResult Create(string filePath, string sheetName, string tableName, string range, bool hasHeaders = true, string? tableStyle = null);

    /// <summary>
    /// Renames an Excel Table
    /// </summary>
    OperationResult Rename(string filePath, string tableName, string newName);

    /// <summary>
    /// Deletes an Excel Table (converts back to range)
    /// </summary>
    OperationResult Delete(string filePath, string tableName);

    /// <summary>
    /// Gets detailed information about an Excel Table
    /// </summary>
    TableInfoResult GetInfo(string filePath, string tableName);
}
