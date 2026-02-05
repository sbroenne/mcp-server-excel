using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Excel Table column operations - filtering, column management, sorting, number formatting
/// </summary>
[ServiceCategory("tablecolumn", "TableColumn")]
[McpTool("excel_table_column")]
public interface ITableColumnCommands
{
    // === FILTER OPERATIONS ===

    /// <summary>
    /// Applies a filter to a table column with single criteria
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    [ServiceAction("apply-filter")]
    void ApplyFilter(IExcelBatch batch, string tableName, string columnName, string criteria);

    /// <summary>
    /// Applies a filter to a table column with multiple values
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    [ServiceAction("apply-filter-values")]
    void ApplyFilterValues(IExcelBatch batch, string tableName, string columnName, List<string> values);

    /// <summary>
    /// Clears all filters from a table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found</exception>
    [ServiceAction("clear-filters")]
    void ClearFilters(IExcelBatch batch, string tableName);

    /// <summary>
    /// Gets current filter state for all columns in a table
    /// </summary>
    [ServiceAction("get-filters")]
    TableFilterResult GetFilters(IExcelBatch batch, string tableName);

    // === COLUMN OPERATIONS ===

    /// <summary>
    /// Adds a new column to a table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found or position invalid</exception>
    [ServiceAction("add-column")]
    void AddColumn(IExcelBatch batch, string tableName, string columnName, int? position = null);

    /// <summary>
    /// Removes a column from a table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    [ServiceAction("remove-column")]
    void RemoveColumn(IExcelBatch batch, string tableName, string columnName);

    /// <summary>
    /// Renames a column in a table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    [ServiceAction("rename-column")]
    void RenameColumn(IExcelBatch batch, string tableName, string oldName, string newName);

    // === STRUCTURED REFERENCE OPERATIONS ===

    /// <summary>
    /// Gets structured reference information for a table region or column
    /// </summary>
    [ServiceAction("get-structured-reference")]
    TableStructuredReferenceResult GetStructuredReference(IExcelBatch batch, string tableName, TableRegion region, string? columnName = null);

    // === SORT OPERATIONS ===

    /// <summary>
    /// Sorts a table by a single column
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    [ServiceAction("sort")]
    void Sort(IExcelBatch batch, string tableName, string columnName, bool ascending = true);

    /// <summary>
    /// Sorts a table by multiple columns
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    [ServiceAction("sort-multi")]
    void SortMulti(IExcelBatch batch, string tableName, List<TableSortColumn> sortColumns);

    // === NUMBER FORMATTING ===

    /// <summary>
    /// Gets number formats for a table column
    /// Delegates to RangeCommands.GetNumberFormatsAsync() on column range
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Table name</param>
    /// <param name="columnName">Column name</param>
    [ServiceAction("get-column-number-format")]
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
    [ServiceAction("set-column-number-format")]
    void SetColumnNumberFormat(IExcelBatch batch, string tableName, string columnName, string formatCode);
}
