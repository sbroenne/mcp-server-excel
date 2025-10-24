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

    /// <summary>
    /// Resizes an Excel Table to a new range
    /// </summary>
    OperationResult Resize(string filePath, string tableName, string newRange);

    /// <summary>
    /// Toggles the totals row for an Excel Table
    /// </summary>
    OperationResult ToggleTotals(string filePath, string tableName, bool showTotals);

    /// <summary>
    /// Sets the totals function for a specific column in an Excel Table
    /// </summary>
    OperationResult SetColumnTotal(string filePath, string tableName, string columnName, string totalFunction);

    /// <summary>
    /// Reads data from an Excel Table
    /// </summary>
    TableDataResult ReadData(string filePath, string tableName);

    /// <summary>
    /// Appends rows to an Excel Table
    /// </summary>
    OperationResult AppendRows(string filePath, string tableName, string csvData);

    /// <summary>
    /// Changes the style of an Excel Table
    /// </summary>
    OperationResult SetStyle(string filePath, string tableName, string tableStyle);

    /// <summary>
    /// Adds an Excel Table to the Power Pivot Data Model
    /// </summary>
    OperationResult AddToDataModel(string filePath, string tableName);
}
