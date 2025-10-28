using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Excel Table (ListObject) management commands
/// </summary>
public interface ITableCommands
{
    /// <summary>
    /// Lists all Excel Tables in the workbook
    /// </summary>
    Task<TableListResult> ListAsync(IExcelBatch batch);

    /// <summary>
    /// Creates a new Excel Table from a range
    /// </summary>
    Task<OperationResult> CreateAsync(IExcelBatch batch, string sheetName, string tableName, string range, bool hasHeaders = true, string? tableStyle = null);

    /// <summary>
    /// Renames an Excel Table
    /// </summary>
    Task<OperationResult> RenameAsync(IExcelBatch batch, string tableName, string newName);

    /// <summary>
    /// Deletes an Excel Table (converts back to range)
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string tableName);

    /// <summary>
    /// Gets detailed information about an Excel Table
    /// </summary>
    Task<TableInfoResult> GetInfoAsync(IExcelBatch batch, string tableName);

    /// <summary>
    /// Resizes an Excel Table to a new range
    /// </summary>
    Task<OperationResult> ResizeAsync(IExcelBatch batch, string tableName, string newRange);

    /// <summary>
    /// Toggles the totals row for an Excel Table
    /// </summary>
    Task<OperationResult> ToggleTotalsAsync(IExcelBatch batch, string tableName, bool showTotals);

    /// <summary>
    /// Sets the totals function for a specific column in an Excel Table
    /// </summary>
    Task<OperationResult> SetColumnTotalAsync(IExcelBatch batch, string tableName, string columnName, string totalFunction);

    /// <summary>
    /// Reads data from an Excel Table
    /// </summary>
    Task<TableDataResult> ReadDataAsync(IExcelBatch batch, string tableName);

    /// <summary>
    /// Appends rows to an Excel Table
    /// </summary>
    Task<OperationResult> AppendRowsAsync(IExcelBatch batch, string tableName, string csvData);

    /// <summary>
    /// Changes the style of an Excel Table
    /// </summary>
    Task<OperationResult> SetStyleAsync(IExcelBatch batch, string tableName, string tableStyle);

    /// <summary>
    /// Adds an Excel Table to the Power Pivot Data Model
    /// </summary>
    Task<OperationResult> AddToDataModelAsync(IExcelBatch batch, string tableName);
}
