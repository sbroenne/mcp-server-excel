using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Excel Table lifecycle and data operations
/// </summary>
[ServiceCategory("table", "Table")]
[McpTool("excel_table")]
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
    /// <exception cref="InvalidOperationException">Sheet not found, table name already exists, or range invalid</exception>
    [ServiceAction("create")]
    void Create(IExcelBatch batch, string sheetName, string tableName, string range, bool hasHeaders = true, string? tableStyle = null);

    /// <summary>
    /// Renames an Excel Table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found or new name already exists</exception>
    [ServiceAction("rename")]
    void Rename(IExcelBatch batch, string tableName, string newName);

    /// <summary>
    /// Deletes an Excel Table (converts back to range)
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found</exception>
    [ServiceAction("delete")]
    void Delete(IExcelBatch batch, string tableName);

    /// <summary>
    /// Gets detailed information about an Excel Table
    /// </summary>
    [ServiceAction("read")]
    TableInfoResult Read(IExcelBatch batch, string tableName);

    /// <summary>
    /// Resizes an Excel Table to a new range
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found or new range invalid</exception>
    [ServiceAction("resize")]
    void Resize(IExcelBatch batch, string tableName, string newRange);

    /// <summary>
    /// Toggles the totals row for an Excel Table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found</exception>
    [ServiceAction("toggle-totals")]
    void ToggleTotals(IExcelBatch batch, string tableName, bool showTotals);

    /// <summary>
    /// Sets the totals function for a specific column in an Excel Table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table or column not found</exception>
    [ServiceAction("set-column-total")]
    void SetColumnTotal(IExcelBatch batch, string tableName, string columnName, string totalFunction);

    /// <summary>
    /// Appends rows to an Excel Table (table auto-expands)
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found or append failed</exception>
    [ServiceAction("append")]
    void Append(IExcelBatch batch, string tableName, List<List<object?>> rows);

    /// <summary>
    /// Retrieves data rows from a table, optionally limited to currently visible rows.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Table name</param>
    /// <param name="visibleOnly">If true, only rows not hidden by filters are returned</param>
    /// <exception cref="InvalidOperationException">Table not found</exception>
    [ServiceAction("get-data")]
    TableDataResult GetData(IExcelBatch batch, string tableName, bool visibleOnly = false);

    /// <summary>
    /// Changes the style of an Excel Table
    /// </summary>
    /// <exception cref="InvalidOperationException">Table not found or invalid style</exception>
    [ServiceAction("set-style")]
    void SetStyle(IExcelBatch batch, string tableName, string tableStyle);

    /// <summary>
    /// Adds an Excel Table to the Power Pivot Data Model
    /// </summary>
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
    /// <param name="sheetName">Target worksheet name</param>
    /// <param name="tableName">Name for the new table</param>
    /// <param name="daxQuery">DAX EVALUATE query (e.g., "EVALUATE 'TableName'" or "EVALUATE SUMMARIZE(...)")</param>
    /// <param name="targetCell">Target cell address for table placement (default: "A1")</param>
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
    /// <param name="daxQuery">New DAX EVALUATE query</param>
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
