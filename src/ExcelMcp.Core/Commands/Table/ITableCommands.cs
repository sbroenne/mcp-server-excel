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
    /// <exception cref="InvalidOperationException">Sheet not found, table name already exists, or range invalid</exception>
    void Create(IExcelBatch batch, string sheetName, string tableName, string range, bool hasHeaders = true, string? tableStyle = null);

    /// <summary>
    /// Renames an Excel Table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found or new name already exists</exception>
    void Rename(IExcelBatch batch, string tableName, string newName);

    /// <summary>
    /// Deletes an Excel Table (converts back to range)
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found</exception>
    void Delete(IExcelBatch batch, string tableName);

    /// <summary>
    /// Gets detailed information about an Excel Table
    /// </summary>
    TableInfoResult Read(IExcelBatch batch, string tableName);

    /// <summary>
    /// Resizes an Excel Table to a new range
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found or new range invalid</exception>
    void Resize(IExcelBatch batch, string tableName, string newRange);

    /// <summary>
    /// Toggles the totals row for an Excel Table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found</exception>
    void ToggleTotals(IExcelBatch batch, string tableName, bool showTotals);

    /// <summary>
    /// Sets the totals function for a specific column in an Excel Table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    void SetColumnTotal(IExcelBatch batch, string tableName, string columnName, string totalFunction);

    /// <summary>
    /// Appends rows to an Excel Table (table auto-expands)
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found or append failed</exception>
    void Append(IExcelBatch batch, string tableName, List<List<object?>> rows);

    /// <summary>
    /// Changes the style of an Excel Table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found or invalid style</exception>
    void SetStyle(IExcelBatch batch, string tableName, string tableStyle);

    /// <summary>
    /// Adds an Excel Table to the Power Pivot Data Model
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found or model not available</exception>
    void AddToDataModel(IExcelBatch batch, string tableName);

    // === FILTER OPERATIONS ===

    /// <summary>
    /// Applies a filter to a table column with single criteria
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    void ApplyFilter(IExcelBatch batch, string tableName, string columnName, string criteria);

    /// <summary>
    /// Applies a filter to a table column with multiple values
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    void ApplyFilter(IExcelBatch batch, string tableName, string columnName, List<string> values);

    /// <summary>
    /// Clears all filters from a table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found</exception>
    void ClearFilters(IExcelBatch batch, string tableName);

    /// <summary>
    /// Gets current filter state for all columns in a table
    /// </summary>
    TableFilterResult GetFilters(IExcelBatch batch, string tableName);

    // === COLUMN OPERATIONS ===

    /// <summary>
    /// Adds a new column to a table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found or position invalid</exception>
    void AddColumn(IExcelBatch batch, string tableName, string columnName, int? position = null);

    /// <summary>
    /// Removes a column from a table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    void RemoveColumn(IExcelBatch batch, string tableName, string columnName);

    /// <summary>
    /// Renames a column in a table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    void RenameColumn(IExcelBatch batch, string tableName, string oldName, string newName);

    // === STRUCTURED REFERENCE OPERATIONS ===

    /// <summary>
    /// Gets structured reference information for a table region or column
    /// </summary>
    TableStructuredReferenceResult GetStructuredReference(IExcelBatch batch, string tableName, TableRegion region, string? columnName = null);

    // === SORT OPERATIONS ===

    /// <summary>
    /// Sorts a table by a single column
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    void Sort(IExcelBatch batch, string tableName, string columnName, bool ascending = true);

    /// <summary>
    /// Sorts a table by multiple columns
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    void Sort(IExcelBatch batch, string tableName, List<TableSortColumn> sortColumns);

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
    /// <exception cref="InvalidOperationException">Table or column not found, or format code invalid</exception>
    void SetColumnNumberFormat(IExcelBatch batch, string tableName, string columnName, string formatCode);
}

