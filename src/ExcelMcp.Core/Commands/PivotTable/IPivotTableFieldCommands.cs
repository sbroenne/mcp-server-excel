using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable field management: add/remove/configure fields, filtering, sorting, and grouping.
/// Use pivottable for lifecycle, pivottablecalc for calculated fields and layout.
///
/// IMPORTANT: Field operations modify structure only. Call pivottable refresh after
/// configuring fields to update the visual display, especially for OLAP/Data Model PivotTables.
///
/// FIELD AREAS:
/// - Row fields: Group data by categories (add-row-field)
/// - Column fields: Create column headers (add-column-field)
/// - Value fields: Aggregate numeric data with Sum, Count, Average, etc. (add-value-field)
/// - Filter fields: Add report-level filters (add-filter-field)
///
/// AGGREGATION FUNCTIONS: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP
///
/// GROUPING:
/// - Date fields: Group by Days, Months, Quarters, Years (group-by-date)
/// - Numeric fields: Group by ranges with start/end/interval (group-by-numeric)
///
/// NUMBER FORMAT: Use US format codes like '#,##0.00' for currency or '0.00%' for percentages.
/// </summary>
[ServiceCategory("pivottablefield", "PivotTableField")]
[McpTool("excel_pivottable_field", Title = "Excel PivotTable Field Operations", Destructive = true, Category = "analysis",
    Description = "PivotTable field management: add/remove/configure fields, filtering, sorting, and grouping. IMPORTANT: Field operations modify structure only - call excel_pivottable(refresh) after configuring, especially for OLAP/Data Model PivotTables. FIELD AREAS: Row (categories), Column (headers), Value (aggregation: Sum/Count/Average/Max/Min/etc.), Filter (report-level). GROUPING: date (Days/Months/Quarters/Years), numeric (start/end/interval). NUMBER FORMAT: US format codes. Use excel_pivottable for lifecycle, excel_pivottable_calc for calculated fields.")]
public interface IPivotTableFieldCommands
{
    // === FIELD MANAGEMENT (WITH IMMEDIATE VALIDATION) ===

    /// <summary>
    /// Lists all available fields and their current placement
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <returns>Field names, data types, current areas, aggregation functions</returns>
    [ServiceAction("list-fields")]
    PivotFieldListResult ListFields(IExcelBatch batch, string pivotTableName);

    /// <summary>
    /// Adds field to Row area with position validation
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field to add</param>
    /// <param name="position">Optional position in row area (1-based)</param>
    /// <returns>Updated field layout with new position</returns>
    [ServiceAction("add-row-field")]
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
    [ServiceAction("add-column-field")]
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
    [ServiceAction("add-value-field")]
    PivotFieldResult AddValueField(IExcelBatch batch, string pivotTableName,
        string fieldName, [FromString] AggregationFunction aggregationFunction = AggregationFunction.Sum,
        string? customName = null);

    /// <summary>
    /// Adds field to Filter area (Page field)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field to add</param>
    /// <returns>Field configuration with available filter items</returns>
    [ServiceAction("add-filter-field")]
    PivotFieldResult AddFilterField(IExcelBatch batch, string pivotTableName,
        string fieldName);

    /// <summary>
    /// Removes field from any area
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field to remove</param>
    /// <returns>Updated layout after removal</returns>
    [ServiceAction("remove-field")]
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
    [ServiceAction("set-field-function")]
    PivotFieldResult SetFieldFunction(IExcelBatch batch, string pivotTableName,
        string fieldName, [FromString] AggregationFunction aggregationFunction);

    /// <summary>
    /// Sets custom name for field in any area
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field</param>
    /// <param name="customName">Custom name to set</param>
    /// <returns>Applied name and field reference</returns>
    [ServiceAction("set-field-name")]
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
    [ServiceAction("set-field-format")]
    PivotFieldResult SetFieldFormat(IExcelBatch batch, string pivotTableName,
        string fieldName, string numberFormat);

    // === ANALYSIS OPERATIONS (WITH DATA VALIDATION) ===

    /// <summary>
    /// Sets filter for field with validation of filter items
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the field to filter</param>
    /// <param name="selectedValues">Values to show (others will be hidden)</param>
    /// <returns>Applied filter state and affected row count</returns>
    [ServiceAction("set-field-filter")]
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
    [ServiceAction("sort-field")]
    PivotFieldResult SortField(IExcelBatch batch, string pivotTableName,
        string fieldName, [FromString] SortDirection direction = SortDirection.Ascending);

    // === GROUPING OPERATIONS (DATE AND NUMERIC) ===

    /// <summary>
    /// Groups date/time field by specified interval (Month, Quarter, Year)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the date/time field to group</param>
    /// <param name="interval">Grouping interval (Months, Quarters, Years)</param>
    /// <returns>Applied grouping configuration and resulting group count</returns>
    /// <remarks>
    /// Creates automatic date hierarchy in PivotTable (e.g., Years > Quarters > Months).
    /// Works for both regular and OLAP PivotTables.
    /// Example: Group "OrderDate" by Months to see monthly sales trends.
    /// </remarks>
    [ServiceAction("group-by-date")]
    PivotFieldResult GroupByDate(IExcelBatch batch, string pivotTableName,
        string fieldName, [FromString] DateGroupingInterval interval);

    /// <summary>
    /// Groups a numeric field by specified interval (e.g., 0-100, 100-200, 200-300).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of PivotTable</param>
    /// <param name="fieldName">Field to group</param>
    /// <param name="start">Starting value (null = use field minimum)</param>
    /// <param name="endValue">Ending value (null = use field maximum)</param>
    /// <param name="intervalSize">Size of each group (e.g., 100 for groups of 100)</param>
    /// <returns>Grouping result with created groups</returns>
    /// <remarks>
    /// Creates numeric range groups in PivotTable for analysis.
    /// Use cases: Age groups (0-20, 20-40), price ranges (0-100, 100-200), score bands (0-50, 50-100).
    /// Works for regular PivotTables. OLAP PivotTables require grouping in Data Model.
    /// Example: Group "Sales" by 100 to analyze sales distribution across price ranges.
    /// </remarks>
    [ServiceAction("group-by-numeric")]
    PivotFieldResult GroupByNumeric(IExcelBatch batch, string pivotTableName,
        string fieldName, double? start, double? endValue, double intervalSize);
}
