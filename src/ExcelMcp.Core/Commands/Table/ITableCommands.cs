using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Excel Table (ListObject) management commands
/// </summary>
public interface ITableCommands
{
    /// <summary>
    /// Lists all Excel Tables in the workbook
    /// </summary>
    TableListResult List(IExcelBatch batch);

    /// <summary>
    /// Creates a new Excel Table from a range
    /// </summary>
    OperationResult Create(IExcelBatch batch, string sheetName, string tableName, string range, bool hasHeaders = true, string? tableStyle = null);

    /// <summary>
    /// Renames an Excel Table
    /// </summary>
    OperationResult Rename(IExcelBatch batch, string tableName, string newName);

    /// <summary>
    /// Deletes an Excel Table (converts back to range)
    /// </summary>
    OperationResult Delete(IExcelBatch batch, string tableName);

    /// <summary>
    /// Gets detailed information about an Excel Table
    /// </summary>
    TableInfoResult Read(IExcelBatch batch, string tableName);

    /// <summary>
    /// Resizes an Excel Table to a new range
    /// </summary>
    OperationResult Resize(IExcelBatch batch, string tableName, string newRange);

    /// <summary>
    /// Toggles the totals row for an Excel Table
    /// </summary>
    OperationResult ToggleTotals(IExcelBatch batch, string tableName, bool showTotals);

    /// <summary>
    /// Sets the totals function for a specific column in an Excel Table
    /// </summary>
    OperationResult SetColumnTotal(IExcelBatch batch, string tableName, string columnName, string totalFunction);

    /// <summary>
    /// Appends rows to an Excel Table (table auto-expands)
    /// </summary>
    OperationResult Append(IExcelBatch batch, string tableName, List<List<object?>> rows);

    /// <summary>
    /// Changes the style of an Excel Table
    /// </summary>
    OperationResult SetStyle(IExcelBatch batch, string tableName, string tableStyle);

    /// <summary>
    /// Adds an Excel Table to the Power Pivot Data Model
    /// </summary>
    OperationResult AddToDataModel(IExcelBatch batch, string tableName);

    // === FILTER OPERATIONS ===

    /// <summary>
    /// Applies a filter to a table column with single criteria
    /// </summary>
    OperationResult ApplyFilter(IExcelBatch batch, string tableName, string columnName, string criteria);

    /// <summary>
    /// Applies a filter to a table column with multiple values
    /// </summary>
    OperationResult ApplyFilter(IExcelBatch batch, string tableName, string columnName, List<string> values);

    /// <summary>
    /// Clears all filters from a table
    /// </summary>
    OperationResult ClearFilters(IExcelBatch batch, string tableName);

    /// <summary>
    /// Gets current filter state for all columns in a table
    /// </summary>
    TableFilterResult GetFilters(IExcelBatch batch, string tableName);

    // === COLUMN OPERATIONS ===

    /// <summary>
    /// Adds a new column to a table
    /// </summary>
    OperationResult AddColumn(IExcelBatch batch, string tableName, string columnName, int? position = null);

    /// <summary>
    /// Removes a column from a table
    /// </summary>
    OperationResult RemoveColumn(IExcelBatch batch, string tableName, string columnName);

    /// <summary>
    /// Renames a column in a table
    /// </summary>
    OperationResult RenameColumn(IExcelBatch batch, string tableName, string oldName, string newName);

    // === STRUCTURED REFERENCE OPERATIONS ===

    /// <summary>
    /// Gets structured reference information for a table region or column
    /// </summary>
    TableStructuredReferenceResult GetStructuredReference(IExcelBatch batch, string tableName, TableRegion region, string? columnName = null);

    // === SORT OPERATIONS ===

    /// <summary>
    /// Sorts a table by a single column
    /// </summary>
    OperationResult Sort(IExcelBatch batch, string tableName, string columnName, bool ascending = true);

    /// <summary>
    /// Sorts a table by multiple columns
    /// </summary>
    OperationResult Sort(IExcelBatch batch, string tableName, List<TableSortColumn> sortColumns);

    // === NUMBER FORMATTING ===

    /// <summary>
    /// Gets number formats for a table column
    /// Delegates to RangeCommands.GetNumberFormatsAsync() on column range
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Table name</param>
    /// <param name="columnName">Column name</param>
    RangeNumberFormatResult GetColumnNumberFormat(IExcelBatch batch, string tableName, string columnName);

    /// <summary>
    /// Sets uniform number format for entire table column
    /// Delegates to RangeCommands.SetNumberFormatAsync() on column data range (excludes header)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Table name</param>
    /// <param name="columnName">Column name</param>
    /// <param name="formatCode">Excel format code (e.g., "$#,##0.00", "0.00%")</param>
    OperationResult SetColumnNumberFormat(IExcelBatch batch, string tableName, string columnName, string formatCode);
}

