using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable management commands for Excel automation
/// Provides complete PivotTable lifecycle, field management, and analysis capabilities
/// </summary>
public interface IPivotTableCommands
{
    // === LIFECYCLE OPERATIONS ===

    /// <summary>
    /// Lists all PivotTables in workbook with structure details
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <returns>List of PivotTables with names, sheets, ranges, source data, field counts, last refresh</returns>
    PivotTableListResult List(IExcelBatch batch);

    /// <summary>
    /// Gets complete PivotTable configuration and current layout
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <returns>All fields with positions, aggregation functions, filter states</returns>
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
    OperationResult Delete(IExcelBatch batch, string pivotTableName);

    /// <summary>
    /// Refreshes PivotTable data from source and returns updated info
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable to refresh</param>
    /// <param name="timeout">Optional timeout for the refresh operation</param>
    /// <returns>Refresh timestamp, record count, any structural changes</returns>
    PivotTableRefreshResult Refresh(IExcelBatch batch, string pivotTableName, TimeSpan? timeout = null);

    // === FIELD MANAGEMENT (WITH IMMEDIATE VALIDATION) ===

    /// <summary>
    /// Lists all available fields and their current placement
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <returns>Field names, data types, current areas, aggregation functions</returns>
    PivotFieldListResult ListFields(IExcelBatch batch, string pivotTableName);

    /// <summary>
    /// Adds field to Row area with position validation
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field to add</param>
    /// <param name="position">Optional position in row area (1-based)</param>
    /// <returns>Updated field layout with new position</returns>
    PivotFieldResult AddRowField(IExcelBatch batch, string pivotTableName,
        string fieldName, int? position = null);

    /// <summary>
    /// Adds field to Column area with position validation
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field to add</param>
    /// <param name="position">Optional position in column area (1-based)</param>
    /// <returns>Updated field layout</returns>
    PivotFieldResult AddColumnField(IExcelBatch batch, string pivotTableName,
        string fieldName, int? position = null);

    /// <summary>
    /// Adds field to Values area with aggregation function.
    /// 
    /// For OLAP PivotTables, supports TWO modes:
    /// 1. Pre-existing measure: fieldName = "Total Sales" or "[Measures].[Total Sales]"
    ///    - Adds existing DAX measure without creating duplicate
    ///    - aggregationFunction ignored (measure formula defines aggregation)
    /// 2. Auto-create measure: fieldName = "Sales" (column name)
    ///    - Creates new DAX measure with specified aggregation function
    ///    - customName becomes the measure name
    /// 
    /// For Regular PivotTables: Adds column with aggregation function
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Field/column name OR existing measure name (OLAP)</param>
    /// <param name="aggregationFunction">Aggregation function (for Regular and OLAP auto-create mode)</param>
    /// <param name="customName">Optional custom name for the field/measure</param>
    /// <returns>Field configuration with applied function and custom name</returns>
    PivotFieldResult AddValueField(IExcelBatch batch, string pivotTableName,
        string fieldName, AggregationFunction aggregationFunction = AggregationFunction.Sum,
        string? customName = null);

    /// <summary>
    /// Adds field to Filter area (Page field)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field to add</param>
    /// <returns>Field configuration with available filter items</returns>
    PivotFieldResult AddFilterField(IExcelBatch batch, string pivotTableName,
        string fieldName);

    /// <summary>
    /// Removes field from any area
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field to remove</param>
    /// <returns>Updated layout after removal</returns>
    PivotFieldResult RemoveField(IExcelBatch batch, string pivotTableName,
        string fieldName);

    // === FIELD CONFIGURATION (WITH RESULT VERIFICATION) ===

    /// <summary>
    /// Sets aggregation function for value field
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field</param>
    /// <param name="aggregationFunction">Aggregation function to set</param>
    /// <returns>Applied function and sample calculation result</returns>
    PivotFieldResult SetFieldFunction(IExcelBatch batch, string pivotTableName,
        string fieldName, AggregationFunction aggregationFunction);

    /// <summary>
    /// Sets custom name for field in any area
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field</param>
    /// <param name="customName">Custom name to set</param>
    /// <returns>Applied name and field reference</returns>
    PivotFieldResult SetFieldName(IExcelBatch batch, string pivotTableName,
        string fieldName, string customName);

    /// <summary>
    /// Sets number format for value field
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field</param>
    /// <param name="numberFormat">Number format string</param>
    /// <returns>Applied format with sample formatted value</returns>
    PivotFieldResult SetFieldFormat(IExcelBatch batch, string pivotTableName,
        string fieldName, string numberFormat);

    // === ANALYSIS OPERATIONS (WITH DATA VALIDATION) ===

    /// <summary>
    /// Gets current PivotTable data as 2D array for LLM analysis
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <returns>Values with headers, row/column labels, formatted numbers</returns>
    PivotTableDataResult GetData(IExcelBatch batch, string pivotTableName);

    /// <summary>
    /// Sets filter for field with validation of filter items
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field to filter</param>
    /// <param name="selectedValues">Values to show (others will be hidden)</param>
    /// <returns>Applied filter state and affected row count</returns>
    PivotFieldFilterResult SetFieldFilter(IExcelBatch batch, string pivotTableName,
        string fieldName, List<string> selectedValues);

    /// <summary>
    /// Sorts field with immediate layout update
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field to sort</param>
    /// <param name="direction">Sort direction</param>
    /// <returns>Applied sort configuration and preview of changes</returns>
    PivotFieldResult SortField(IExcelBatch batch, string pivotTableName,
        string fieldName, SortDirection direction = SortDirection.Ascending);
}

