using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable lifecycle operations (create, list, read, refresh, delete).
/// Use pivottablefield for field operations, pivottablecalc for calculated fields and layout.
/// </summary>
[ServiceCategory("pivottable", "PivotTable")]
[McpTool("excel_pivottable")]
public interface IPivotTableCommands
{
    // === LIFECYCLE OPERATIONS ===

    /// <summary>
    /// Lists all PivotTables in workbook with structure details
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <returns>List of PivotTables with names, sheets, ranges, source data, field counts, last refresh</returns>
    [ServiceAction("list")]
    PivotTableListResult List(IExcelBatch batch);

    /// <summary>
    /// Gets complete PivotTable configuration and current layout
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <returns>All fields with positions, aggregation functions, filter states</returns>
    [ServiceAction("read")]
    PivotTableInfoResult Read(IExcelBatch batch, string pivotTableName);

    /// <summary>
    /// Creates PivotTable from Excel range with auto-detection of headers
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sourceSheet">Source worksheet name</param>
    /// <param name="sourceRange">Source range address (e.g., "A1:F100")</param>
    /// <param name="destinationSheet">Destination worksheet name</param>
    /// <param name="destinationCell">Destination cell address (e.g., "A1")</param>
    /// <param name="pivotTableName">Name for the new PivotTable</param>
    /// <returns>Created PivotTable name and initial field list</returns>
    [ServiceAction("create-from-range")]
    PivotTableCreateResult CreateFromRange(IExcelBatch batch,
        string sourceSheet, string sourceRange,
        string destinationSheet, string destinationCell,
        string pivotTableName);

    /// <summary>
    /// Creates PivotTable from Excel Table (ListObject)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Name of the Excel Table</param>
    /// <param name="destinationSheet">Destination worksheet name</param>
    /// <param name="destinationCell">Destination cell address (e.g., "A1")</param>
    /// <param name="pivotTableName">Name for the new PivotTable</param>
    /// <returns>Created PivotTable name and available fields</returns>
    [ServiceAction("create-from-table")]
    PivotTableCreateResult CreateFromTable(IExcelBatch batch,
        string tableName,
        string destinationSheet, string destinationCell,
        string pivotTableName);

    /// <summary>
    /// Creates PivotTable from Power Pivot Data Model table
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Name of the Data Model table</param>
    /// <param name="destinationSheet">Destination worksheet name</param>
    /// <param name="destinationCell">Destination cell address (e.g., "A1")</param>
    /// <param name="pivotTableName">Name for the new PivotTable</param>
    /// <returns>Created PivotTable name and available fields</returns>
    [ServiceAction("create-from-datamodel")]
    PivotTableCreateResult CreateFromDataModel(IExcelBatch batch,
        string tableName,
        string destinationSheet, string destinationCell,
        string pivotTableName);

    /// <summary>
    /// Deletes PivotTable completely
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable to delete</param>
    /// <returns>Operation result</returns>
    [ServiceAction("delete")]
    OperationResult Delete(IExcelBatch batch, string pivotTableName);

    /// <summary>
    /// Refreshes PivotTable data from source and returns updated info
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable to refresh</param>
    /// <param name="timeout">Optional timeout for the refresh operation</param>
    /// <returns>Refresh timestamp, record count, any structural changes</returns>
    [ServiceAction("refresh")]
    PivotTableRefreshResult Refresh(IExcelBatch batch, string pivotTableName, TimeSpan? timeout = null);
}


