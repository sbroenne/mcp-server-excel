using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Excel Tables (ListObjects) - lifecycle and data operations.
/// Tables provide structured references, automatic formatting, and Data Model integration.
///
/// BEST PRACTICE: Use 'list' to check existing tables before creating.
/// Prefer 'append'/'resize'/'rename' over delete+recreate to preserve references.
///
/// WARNING: Deleting tables used as PivotTable sources or in Data Model relationships will break those objects.
///
/// DATA MODEL WORKFLOW: To analyze worksheet data with DAX/Power Pivot:
/// 1. Create or identify an Excel Table on a worksheet
/// 2. Use 'add-to-datamodel' to add the table to Power Pivot
/// 3. Then use datamodel to create DAX measures on it
///
/// DAX-BACKED TABLES: Create tables populated by DAX EVALUATE queries:
/// - 'create-from-dax': Create a new table backed by a DAX query (e.g., SUMMARIZE, FILTER)
/// - 'update-dax': Update the DAX query for an existing DAX-backed table
/// - 'get-dax': Get the DAX query info for a table (check if it's DAX-backed)
///
/// Related: tablecolumn (filter/sort/columns), datamodel (DAX measures, evaluate queries)
/// </summary>
[ServiceCategory("table", "Table")]
[McpTool("excel_table", Title = "Excel Table Operations", Destructive = true, Category = "data",
    Description = "Excel Tables (ListObjects) - lifecycle and data operations. BEST PRACTICE: List before creating, prefer append/resize/rename over delete+recreate. WARNING: Deleting tables used as PivotTable sources or in Data Model breaks those objects. DATA MODEL: add-to-datamodel to load into Power Pivot, then excel_datamodel for DAX measures. DAX-BACKED TABLES: create-from-dax, update-dax, get-dax. APPEND: csvData in CSV format (comma-separated, newline-separated rows). Use excel_table_column for filtering/sorting/columns.")]
public interface ITableCommands
{
    /// <summary>
    /// Lists all Excel Tables in the workbook
    /// </summary>
    [ServiceAction("list")]
    TableListResult List(IExcelBatch batch);

    /// <summary>
    /// Creates a new Excel Table from a range
    /// </summary>
    /// <param name="sheetName">Name of the worksheet to create the table on</param>
    /// <param name="tableName">Name for the new table (must be unique in workbook)</param>
    /// <param name="range">Cell range address for the table (e.g., 'A1:D10')</param>
    /// <param name="hasHeaders">True if first row contains column headers (default: true)</param>
    /// <param name="tableStyle">Table style name (e.g., 'TableStyleMedium2', 'TableStyleLight1'). Optional.</param>
    /// <exception cref="InvalidOperationException">Sheet not found, table name already exists, or range invalid</exception>
    [ServiceAction("create")]
    void Create(IExcelBatch batch, string sheetName, string tableName, string range, bool hasHeaders = true, string? tableStyle = null);

    /// <summary>
    /// Renames an Excel Table
    /// </summary>
    /// <param name="tableName">Current name of the table</param>
    /// <param name="newName">New name for the table (must be unique in workbook)</param>
    /// <exception cref="InvalidOperationException">Table not found or new name already exists</exception>
    [ServiceAction("rename")]
    void Rename(IExcelBatch batch, string tableName, string newName);

    /// <summary>
    /// Deletes an Excel Table (converts back to range)
    /// </summary>
    /// <param name="tableName">Name of the table to delete</param>
    /// <exception cref="InvalidOperationException">Table not found</exception>
    [ServiceAction("delete")]
    void Delete(IExcelBatch batch, string tableName);

    /// <summary>
    /// Gets detailed information about an Excel Table
    /// </summary>
    /// <param name="tableName">Name of the table</param>
    [ServiceAction("read")]
    TableInfoResult Read(IExcelBatch batch, string tableName);

    /// <summary>
    /// Resizes an Excel Table to a new range
    /// </summary>
    /// <param name="tableName">Name of the table to resize</param>
    /// <param name="newRange">New range address (e.g., 'A1:F20')</param>
    /// <exception cref="InvalidOperationException">Table not found or new range invalid</exception>
    [ServiceAction("resize")]
    void Resize(IExcelBatch batch, string tableName, string newRange);

    /// <summary>
    /// Toggles the totals row for an Excel Table
    /// </summary>
    /// <param name="tableName">Name of the table</param>
    /// <param name="showTotals">True to show totals row, false to hide</param>
    /// <exception cref="InvalidOperationException">Table not found</exception>
    [ServiceAction("toggle-totals")]
    void ToggleTotals(IExcelBatch batch, string tableName, bool showTotals);

    /// <summary>
    /// Sets the totals function for a specific column in an Excel Table
    /// </summary>
    /// <param name="tableName">Name of the table</param>
    /// <param name="columnName">Name of the column to set total function on</param>
    /// <param name="totalFunction">Totals function name: Sum, Count, Average, Min, Max, CountNums, StdDev, Var, None</param>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    [ServiceAction("set-column-total")]
    void SetColumnTotal(IExcelBatch batch, string tableName, string columnName, string totalFunction);

    /// <summary>
    /// Appends rows to an Excel Table (table auto-expands)
    /// </summary>
    /// <param name="tableName">Name of the table to append to (table auto-expands)</param>
    /// <param name="rows">2D array of row data to append - column order must match table columns</param>
    /// <exception cref="InvalidOperationException">Table not found or append failed</exception>
    [ServiceAction("append")]
    void Append(IExcelBatch batch, string tableName, List<List<object?>> rows);

    /// <summary>
    /// Retrieves data rows from a table, optionally limited to currently visible rows.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Name of the table to read data from</param>
    /// <param name="visibleOnly">True to return only visible (non-filtered) rows; false for all rows (default: false)</param>
    /// <exception cref="InvalidOperationException">Table not found</exception>
    [ServiceAction("get-data")]
    TableDataResult GetData(IExcelBatch batch, string tableName, bool visibleOnly = false);

    /// <summary>
    /// Changes the style of an Excel Table
    /// </summary>
    /// <param name="tableName">Name of the table to style</param>
    /// <param name="tableStyle">Table style name (e.g., 'TableStyleMedium2', 'TableStyleLight1', 'TableStyleDark1')</param>
    /// <exception cref="InvalidOperationException">Table not found or invalid style</exception>
    [ServiceAction("set-style")]
    void SetStyle(IExcelBatch batch, string tableName, string tableStyle);

    /// <summary>
    /// Adds an Excel Table to the Power Pivot Data Model
    /// </summary>
    /// <param name="tableName">Name of the table to add</param>
    /// <exception cref="InvalidOperationException">Table not found or model not available</exception>
    [ServiceAction("add-to-data-model")]
    void AddToDataModel(IExcelBatch batch, string tableName);

    // === DAX-BACKED TABLE OPERATIONS ===

    /// <summary>
    /// Creates a new Excel Table backed by a DAX EVALUATE query.
    /// The table will be connected to the Data Model and refresh when the model refreshes.
    /// Uses Model.CreateModelWorkbookConnection + xlCmdDAX + ListObjects.Add pattern.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Target worksheet name for the new table</param>
    /// <param name="tableName">Name for the new DAX-backed table</param>
    /// <param name="daxQuery">DAX EVALUATE query (e.g., 'EVALUATE Sales' or 'EVALUATE SUMMARIZE(...)')</param>
    /// <param name="targetCell">Target cell address for table placement (default: 'A1')</param>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing</exception>
    /// <exception cref="InvalidOperationException">Sheet not found, table name exists, or no Data Model</exception>
    [ServiceAction("create-from-dax")]
    void CreateFromDax(IExcelBatch batch, string sheetName, string tableName, string daxQuery, string? targetCell = null);

    /// <summary>
    /// Updates the DAX query for an existing DAX-backed Excel Table.
    /// The table must have been created with CreateFromDax or manually connected to a DAX query.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Name of the DAX-backed table to update</param>
    /// <param name="daxQuery">New DAX EVALUATE query (e.g., 'EVALUATE SUMMARIZE(...)')</param>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing</exception>
    /// <exception cref="InvalidOperationException">Table not found or table is not DAX-backed</exception>
    [ServiceAction("update-dax")]
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
    [ServiceAction("get-dax")]
    TableDaxInfoResult GetDax(IExcelBatch batch, string tableName);
}
