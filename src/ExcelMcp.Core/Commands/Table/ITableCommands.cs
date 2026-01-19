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
    /// Retrieves data rows from a table, optionally limited to currently visible rows.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Table name</param>
    /// <param name="visibleOnly">If true, only rows not hidden by filters are returned</param>
    /// <exception cref="InvalidOperationException">Table not found</exception>
    TableDataResult GetData(IExcelBatch batch, string tableName, bool visibleOnly = false);

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

    // === DAX-BACKED TABLE OPERATIONS ===

    /// <summary>
    /// Creates a new Excel Table backed by a DAX EVALUATE query.
    /// The table will be connected to the Data Model and refresh when the model refreshes.
    /// Uses Model.CreateModelWorkbookConnection + xlCmdDAX + ListObjects.Add pattern.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Target worksheet name</param>
    /// <param name="tableName">Name for the new table</param>
    /// <param name="daxQuery">DAX EVALUATE query (e.g., "EVALUATE 'TableName'" or "EVALUATE SUMMARIZE(...)")</param>
    /// <param name="targetCell">Target cell address for table placement (default: "A1")</param>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing</exception>
    /// <exception cref="InvalidOperationException">Sheet not found, table name exists, or no Data Model</exception>
    void CreateFromDax(IExcelBatch batch, string sheetName, string tableName, string daxQuery, string? targetCell = null);

    /// <summary>
    /// Updates the DAX query for an existing DAX-backed Excel Table.
    /// The table must have been created with CreateFromDax or manually connected to a DAX query.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Name of the DAX-backed table to update</param>
    /// <param name="daxQuery">New DAX EVALUATE query</param>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing</exception>
    /// <exception cref="InvalidOperationException">Table not found or table is not DAX-backed</exception>
    void UpdateDax(IExcelBatch batch, string tableName, string daxQuery);

    /// <summary>
    /// Gets the DAX query and connection information for a DAX-backed Excel Table.
    /// Returns empty query info if table is not backed by a DAX query.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Name of the table</param>
    /// <returns>Result containing DAX query info (if any)</returns>
    /// <exception cref="ArgumentException">Thrown when tableName is missing</exception>
    /// <exception cref="InvalidOperationException">Table not found</exception>
    TableDaxInfoResult GetDax(IExcelBatch batch, string tableName);

    // === SLICER OPERATIONS ===

    /// <summary>
    /// Creates a slicer for an Excel Table column.
    /// Slicers provide visual filtering for Table data.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Name of the Excel Table to create slicer for</param>
    /// <param name="columnName">Name of the column to use for the slicer</param>
    /// <param name="slicerName">Name for the new slicer</param>
    /// <param name="destinationSheet">Worksheet where slicer will be placed</param>
    /// <param name="position">Top-left cell position for the slicer (e.g., "H2")</param>
    /// <returns>Created slicer details with available items</returns>
    /// <remarks>
    /// TABLE SLICER BEHAVIOR:
    /// - Slicers are visual filter controls that filter the connected Table
    /// - One SlicerCache is created per column, which can have multiple visual Slicers
    /// - Unlike PivotTable slicers, Table slicers can only filter one Table
    /// </remarks>
    SlicerResult CreateTableSlicer(IExcelBatch batch, string tableName,
        string columnName, string slicerName, string destinationSheet, string position);

    /// <summary>
    /// Lists all slicers in the workbook connected to Tables, optionally filtered by Table name.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Optional Table name to filter slicers (null = all Table slicers)</param>
    /// <returns>List of slicers with names, columns, positions, and selections</returns>
    /// <remarks>
    /// Returns only Table slicers (not PivotTable slicers).
    /// When tableName is specified, only slicers connected to that Table are returned.
    /// </remarks>
    SlicerListResult ListTableSlicers(IExcelBatch batch, string? tableName = null);

    /// <summary>
    /// Sets the selection for a Table slicer, filtering the connected Table.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the slicer to modify</param>
    /// <param name="selectedItems">Items to select (show in Table)</param>
    /// <param name="clearFirst">If true, clears existing selection before setting new items (default: true)</param>
    /// <returns>Updated slicer state with current selection</returns>
    /// <remarks>
    /// SELECTION BEHAVIOR:
    /// - Only selected items are visible in connected Table
    /// - Empty selectedItems list shows all items (clears filter)
    /// - Invalid item names are ignored with a warning
    /// </remarks>
    SlicerResult SetTableSlicerSelection(IExcelBatch batch, string slicerName,
        List<string> selectedItems, bool clearFirst = true);

    /// <summary>
    /// Deletes a Table slicer from the workbook.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the slicer to delete</param>
    /// <returns>Operation result indicating success or failure</returns>
    /// <remarks>
    /// DELETION BEHAVIOR:
    /// - Deletes the visual Slicer object
    /// - If this is the last Slicer using the SlicerCache, the cache is also deleted
    /// - Connected Table filter is cleared when slicer is deleted
    /// </remarks>
    OperationResult DeleteTableSlicer(IExcelBatch batch, string slicerName);
}

