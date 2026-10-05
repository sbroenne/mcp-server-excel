using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Commands.Filtering;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table column, filtering, and sorting operations for Excel Tables (ListObjects).
/// Use table for table-level lifecycle and data operations.
///
/// FILTERING:
/// - 'apply-filter': Typed native comparison, value/date, top/bottom, color, icon, or dynamic filter
/// - 'clear-filters': Remove all active filters
/// - 'get-filters': See current filter state
///
/// SORTING:
/// - 'sort': Single column sort (ascending/descending)
/// - 'sort-multi': Multi-column sort (provide list of {columnName, ascending} objects)
///
/// COLUMN MANAGEMENT:
/// - 'add-column'/'remove-column'/'rename-column': Modify table structure
///
/// NUMBER FORMATS: Use US locale format codes (e.g., '#,##0.00', '0%', 'yyyy-mm-dd')
/// </summary>
[ServiceCategory("tablecolumn", "TableColumn")]
[McpTool("table_column", Title = "Table Column Operations", Destructive = true, Category = "data",
    Description = "Change table columns, filters, and sorting. FILTERING: apply-filter uses typed options for comparisons, AND/OR, values, date groups, top/bottom, colors, icons, and dynamic filters. Nested options use camelCase. clear-filters affects only the selected table. apply-filter-values is removed, not an alias. SORTING: sort (single column), sort-multi (JSON array of {columnName, ascending}). COLUMNS: add-column, remove-column, rename-column. NUMBER FORMATS: US locale codes (#,##0.00, 0%, yyyy-mm-dd). Use range_edit for ordinary-range filters and table for lifecycle/data.")]
[McpReadOnlyActions("get-filters", "get-structured-reference", "get-column-number-format")]
public interface ITableColumnCommands
{
    // === FILTER OPERATIONS ===

    /// <summary>
    /// Applies a typed native filter to one table column while preserving other column filters
    /// </summary>
    /// <param name="tableName">Name of the Excel table</param>
    /// <param name="columnName">Name of the column to filter</param>
    /// <param name="options">Native operator and its applicable criteria; nested keys remain camelCase</param>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    [ServiceAction("apply-filter")]
    OperationResult ApplyFilter(IExcelBatch batch, string tableName, string columnName, [RequiredParameter] FilterOptions options);

    /// <summary>
    /// Clears all filters from a table
    /// </summary>
    /// <param name="tableName">Name of the Excel table</param>
    /// <exception cref="InvalidOperationException">Table not found</exception>
    [ServiceAction("clear-filters")]
    OperationResult ClearFilters(IExcelBatch batch, string tableName);

    /// <summary>
    /// Gets current filter state for all columns in a table
    /// </summary>
    /// <param name="tableName">Name of the Excel table</param>
    [ServiceAction("get-filters")]
    TableFilterResult GetFilters(IExcelBatch batch, string tableName);

    // === COLUMN OPERATIONS ===

    /// <summary>
    /// Adds a new column to a table
    /// </summary>
    /// <param name="tableName">Name of the Excel table</param>
    /// <param name="columnName">Name for the new column</param>
    /// <param name="position">1-based column position (optional, defaults to end of table)</param>
    /// <exception cref="InvalidOperationException">Table not found or position invalid</exception>
    [ServiceAction("add-column")]
    OperationResult AddColumn(IExcelBatch batch, string tableName, string columnName, int? position = null);

    /// <summary>
    /// Removes a column from a table
    /// </summary>
    /// <param name="tableName">Name of the Excel table</param>
    /// <param name="columnName">Name of the column to remove</param>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    [ServiceAction("remove-column")]
    OperationResult RemoveColumn(IExcelBatch batch, string tableName, string columnName);

    /// <summary>
    /// Renames a column in a table
    /// </summary>
    /// <param name="tableName">Name of the Excel table</param>
    /// <param name="oldName">Current column name</param>
    /// <param name="newName">New column name</param>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    [ServiceAction("rename-column")]
    OperationResult RenameColumn(IExcelBatch batch, string tableName, string oldName, string newName);

    // === STRUCTURED REFERENCE OPERATIONS ===

    /// <summary>
    /// Gets structured reference information for a table region or column
    /// </summary>
    /// <param name="tableName">Name of the Excel table</param>
    /// <param name="region">Table region: 'Data', 'Headers', 'Totals', or 'All'</param>
    /// <param name="columnName">Optional column name for column-specific reference</param>
    [ServiceAction("get-structured-reference")]
    TableStructuredReferenceResult GetStructuredReference(IExcelBatch batch, string tableName, [FromString] TableRegion region, string? columnName = null);

    // === SORT OPERATIONS ===

    /// <summary>
    /// Sorts a table by a single column
    /// </summary>
    /// <param name="tableName">Name of the Excel table</param>
    /// <param name="columnName">Column to sort by</param>
    /// <param name="ascending">Sort order: true = ascending (A-Z, 0-9), false = descending (default: true)</param>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    [ServiceAction("sort")]
    OperationResult Sort(IExcelBatch batch, string tableName, string columnName, bool ascending = true);

    /// <summary>
    /// Sorts a table by multiple columns
    /// </summary>
    /// <param name="tableName">Name of the Excel table</param>
    /// <param name="sortColumns">List of sort specifications: [{columnName: 'Col1', ascending: true}, ...] - applied in order</param>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    [ServiceAction("sort-multi")]
    OperationResult SortMulti(IExcelBatch batch, string tableName, List<TableSortColumn> sortColumns);

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
    /// <param name="tableName">Name of the Excel table</param>
    /// <param name="columnName">Name of the column to format</param>
    /// <param name="formatCode">Number format code in US locale (e.g., '#,##0.00', '0%', 'yyyy-mm-dd')</param>
    /// <exception cref="InvalidOperationException">Table or column not found, or format code invalid</exception>
    [ServiceAction("set-column-number-format")]
    OperationResult SetColumnNumberFormat(IExcelBatch batch, string tableName, string columnName, string formatCode);
}
