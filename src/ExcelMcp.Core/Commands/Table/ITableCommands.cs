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
    Task<TableInfoResult> GetAsync(IExcelBatch batch, string tableName);

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
    /// Appends rows to an Excel Table (table auto-expands)
    /// </summary>
    Task<OperationResult> AppendAsync(IExcelBatch batch, string tableName, List<List<object?>> rows);

    /// <summary>
    /// Changes the style of an Excel Table
    /// </summary>
    Task<OperationResult> SetStyleAsync(IExcelBatch batch, string tableName, string tableStyle);

    /// <summary>
    /// Adds an Excel Table to the Power Pivot Data Model
    /// </summary>
    Task<OperationResult> AddToDataModelAsync(IExcelBatch batch, string tableName);

    // === FILTER OPERATIONS ===

    /// <summary>
    /// Applies a filter to a table column with single criteria
    /// </summary>
    Task<OperationResult> ApplyFilterAsync(IExcelBatch batch, string tableName, string columnName, string criteria);

    /// <summary>
    /// Applies a filter to a table column with multiple values
    /// </summary>
    Task<OperationResult> ApplyFilterAsync(IExcelBatch batch, string tableName, string columnName, List<string> values);

    /// <summary>
    /// Clears all filters from a table
    /// </summary>
    Task<OperationResult> ClearFiltersAsync(IExcelBatch batch, string tableName);

    /// <summary>
    /// Gets current filter state for all columns in a table
    /// </summary>
    Task<TableFilterResult> GetFiltersAsync(IExcelBatch batch, string tableName);

    // === COLUMN OPERATIONS ===

    /// <summary>
    /// Adds a new column to a table
    /// </summary>
    Task<OperationResult> AddColumnAsync(IExcelBatch batch, string tableName, string columnName, int? position = null);

    /// <summary>
    /// Removes a column from a table
    /// </summary>
    Task<OperationResult> RemoveColumnAsync(IExcelBatch batch, string tableName, string columnName);

    /// <summary>
    /// Renames a column in a table
    /// </summary>
    Task<OperationResult> RenameColumnAsync(IExcelBatch batch, string tableName, string oldName, string newName);

    // === STRUCTURED REFERENCE OPERATIONS ===

    /// <summary>
    /// Gets structured reference information for a table region or column
    /// </summary>
    Task<TableStructuredReferenceResult> GetStructuredReferenceAsync(IExcelBatch batch, string tableName, TableRegion region, string? columnName = null);

    // === SORT OPERATIONS ===

    /// <summary>
    /// Sorts a table by a single column
    /// </summary>
    Task<OperationResult> SortAsync(IExcelBatch batch, string tableName, string columnName, bool ascending = true);

    /// <summary>
    /// Sorts a table by multiple columns
    /// </summary>
    Task<OperationResult> SortAsync(IExcelBatch batch, string tableName, List<TableSortColumn> sortColumns);

    // === NUMBER FORMATTING ===

    /// <summary>
    /// Gets number formats for a table column
    /// Delegates to RangeCommands.GetNumberFormatsAsync() on column range
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Table name</param>
    /// <param name="columnName">Column name</param>
    Task<RangeNumberFormatResult> GetColumnNumberFormatAsync(IExcelBatch batch, string tableName, string columnName);

    /// <summary>
    /// Sets uniform number format for entire table column
    /// Delegates to RangeCommands.SetNumberFormatAsync() on column data range (excludes header)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Table name</param>
    /// <param name="columnName">Column name</param>
    /// <param name="formatCode">Excel format code (e.g., "$#,##0.00", "0.00%")</param>
    Task<OperationResult> SetColumnNumberFormatAsync(IExcelBatch batch, string tableName, string columnName, string formatCode);
}
