using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Strategy interface for PivotTable field operations.
/// Handles different PivotTable types (Regular vs OLAP/Data Model).
/// </summary>
public interface IPivotTableFieldStrategy
{
    /// <summary>
    /// Determines if this strategy can handle the given PivotTable
    /// </summary>
    bool CanHandle(dynamic pivot);

    /// <summary>
    /// Gets a field for manipulation from the PivotTable.
    /// Returns CubeField for OLAP, PivotField for regular.
    /// </summary>
    dynamic GetFieldForManipulation(dynamic pivot, string fieldName);

    /// <summary>
    /// Lists all fields in the PivotTable
    /// </summary>
    PivotFieldListResult ListFields(dynamic pivot, string workbookPath);

    /// <summary>
    /// Adds a field to the Row area
    /// </summary>
    PivotFieldResult AddRowField(dynamic pivot, string fieldName, int? position, string workbookPath);

    /// <summary>
    /// Adds a field to the Column area
    /// </summary>
    PivotFieldResult AddColumnField(dynamic pivot, string fieldName, int? position, string workbookPath);

    /// <summary>
    /// Adds a field to the Values area with aggregation
    /// </summary>
    PivotFieldResult AddValueField(dynamic pivot, string fieldName, AggregationFunction aggregationFunction, string? customName, string workbookPath);

    /// <summary>
    /// Adds a field to the Filter area
    /// </summary>
    PivotFieldResult AddFilterField(dynamic pivot, string fieldName, string workbookPath);

    /// <summary>
    /// Removes a field from any area
    /// </summary>
    PivotFieldResult RemoveField(dynamic pivot, string fieldName, string workbookPath);

    /// <summary>
    /// Sets custom name for a field
    /// </summary>
    PivotFieldResult SetFieldName(dynamic pivot, string fieldName, string customName, string workbookPath);

    /// <summary>
    /// Sets aggregation function for a value field
    /// </summary>
    PivotFieldResult SetFieldFunction(dynamic pivot, string fieldName, AggregationFunction aggregationFunction, string workbookPath);

    /// <summary>
    /// Sets format for a value field
    /// </summary>
    PivotFieldResult SetFieldFormat(dynamic pivot, string fieldName, string numberFormat, string workbookPath);

    /// <summary>
    /// Sets filter for a field
    /// </summary>
    PivotFieldFilterResult SetFieldFilter(dynamic pivot, string fieldName, List<string> filterValues, string workbookPath);

    /// <summary>
    /// Sorts a field
    /// </summary>
    PivotFieldResult SortField(dynamic pivot, string fieldName, SortDirection direction, string workbookPath);

    /// <summary>
    /// Groups a date/time field by specified interval (Days, Months, Quarters, Years).
    /// </summary>
    /// <remarks>
    /// CRITICAL REQUIREMENT: Source data MUST be formatted with date NumberFormat BEFORE creating the PivotTable.
    /// Excel stores dates as serial numbers (e.g., 45672 = 2025-01-15). Without proper date formatting,
    /// Excel treats these as plain numbers and grouping silently fails.
    ///
    /// Example:
    /// <code>
    /// // Format source data BEFORE creating PivotTable
    /// sheet.Range["D2:D6"].NumberFormat = "m/d/yyyy";
    /// </code>
    /// </remarks>
    PivotFieldResult GroupByDate(dynamic pivot, string fieldName, DateGroupingInterval interval, string workbookPath, Microsoft.Extensions.Logging.ILogger? logger = null);

    /// <summary>
    /// Groups a numeric field by specified interval (e.g., 0-10, 10-20, 20-30).
    /// </summary>
    /// <param name="pivot">The PivotTable object</param>
    /// <param name="fieldName">Field to group</param>
    /// <param name="start">Starting value (null = use field minimum)</param>
    /// <param name="endValue">Ending value (null = use field maximum)</param>
    /// <param name="intervalSize">Size of each group (e.g., 10 for groups of 10)</param>
    /// <param name="workbookPath">Path to workbook for error reporting</param>
    /// <param name="logger">Optional logger for diagnostics</param>
    /// <returns>Result indicating success or failure</returns>
    /// <remarks>
    /// Use cases: Age groups (0-20, 20-40), price ranges (0-100, 100-200), score bands (0-50, 50-100).
    /// Source data should be formatted with numeric NumberFormat for reliable grouping.
    /// If start/end are null, Excel automatically uses the field's minimum/maximum values.
    /// </remarks>
    PivotFieldResult GroupByNumeric(dynamic pivot, string fieldName, double? start, double? endValue, double intervalSize, string workbookPath, Microsoft.Extensions.Logging.ILogger? logger = null);

    /// <summary>
    /// Creates a calculated field with a custom formula.
    /// </summary>
    /// <param name="pivot">The PivotTable object</param>
    /// <param name="fieldName">Name for the calculated field</param>
    /// <param name="formula">Formula using field references (e.g., "=Revenue-Cost" or "=Profit/Revenue")</param>
    /// <param name="workbookPath">Path to workbook for error reporting</param>
    /// <param name="logger">Optional logger for diagnostics</param>
    /// <returns>Result indicating success or failure</returns>
    /// <remarks>
    /// FORMULA SYNTAX:
    /// - Use field names in formulas: =Revenue-Cost, =Profit/Revenue*100
    /// - Operators: + - * / ^ () for basic arithmetic
    /// - Excel will auto-convert field names to proper references
    /// - Example: "Profit" field formula "=Revenue-Cost"
    /// - Example: "Margin%" field formula "=Profit/Revenue"
    ///
    /// IMPORTANT LIMITATIONS:
    /// - Regular PivotTables: Full support via CalculatedFields collection
    /// - OLAP PivotTables: NOT SUPPORTED (use CalculatedMembers with MDX/DAX instead)
    /// - For OLAP, use Data Model DAX measures via excel_datamodel tool
    ///
    /// COMMON USE CASES:
    /// - Financial: Profit = Revenue - Cost, Margin% = Profit/Revenue
    /// - Variance: Actual - Budget, (Actual-Budget)/Budget
    /// - Ratios: Cost/Unit, Revenue/Customer
    /// </remarks>
    PivotFieldResult CreateCalculatedField(dynamic pivot, string fieldName, string formula, string workbookPath, Microsoft.Extensions.Logging.ILogger? logger = null);

    /// <summary>
    /// Sets the row layout form for the PivotTable.
    /// </summary>
    /// <param name="pivot">The PivotTable object</param>
    /// <param name="layoutType">Layout form: 0=Compact, 1=Tabular, 2=Outline</param>
    /// <param name="workbookPath">Path to workbook for error reporting</param>
    /// <param name="logger">Optional logger for diagnostics</param>
    /// <returns>Result indicating success or failure</returns>
    /// <remarks>
    /// LAYOUT FORMS:
    /// - Compact (0): All row fields in single column with indentation (Excel default)
    /// - Tabular (1): Each field in separate column, subtotals at bottom
    /// - Outline (2): Each field in separate column, subtotals at top
    ///
    /// SUPPORT:
    /// - Regular PivotTables: Full support for all three forms
    /// - OLAP PivotTables: Full support for all three forms
    /// </remarks>
    OperationResult SetLayout(dynamic pivot, int layoutType, string workbookPath, Microsoft.Extensions.Logging.ILogger? logger = null);

    /// <summary>
    /// Shows or hides subtotals for a specific row field.
    /// </summary>
    /// <param name="pivot">The PivotTable object</param>
    /// <param name="fieldName">Name of the row field</param>
    /// <param name="showSubtotals">True to show automatic subtotals, false to hide</param>
    /// <param name="workbookPath">Path to workbook for error reporting</param>
    /// <param name="logger">Optional logger for diagnostics</param>
    /// <returns>Result with updated field configuration</returns>
    /// <remarks>
    /// SUBTOTALS BEHAVIOR:
    /// - When enabled: Shows Automatic subtotals (uses appropriate function based on data)
    /// - When disabled: Hides all subtotals for the field
    ///
    /// OLAP LIMITATION:
    /// - OLAP PivotTables only support Automatic subtotals
    /// - Regular PivotTables can choose Sum, Count, Average, etc. (future enhancement)
    /// </remarks>
    PivotFieldResult SetSubtotals(dynamic pivot, string fieldName, bool showSubtotals, string workbookPath, Microsoft.Extensions.Logging.ILogger? logger = null);

    /// <summary>
    /// Shows or hides grand totals for rows and/or columns in the PivotTable.
    /// </summary>
    /// <param name="pivot">The PivotTable object</param>
    /// <param name="showRowGrandTotals">True to show row grand totals, false to hide</param>
    /// <param name="showColumnGrandTotals">True to show column grand totals, false to hide</param>
    /// <param name="workbookPath">Path to workbook for error reporting</param>
    /// <param name="logger">Optional logger for diagnostics</param>
    /// <returns>Result indicating success or failure</returns>
    /// <remarks>
    /// GRAND TOTALS:
    /// - Row Grand Totals: Summary row at the bottom showing totals across all rows
    /// - Column Grand Totals: Summary column at the right showing totals across all columns
    /// - Independent control: Can show/hide row and column totals separately
    ///
    /// SUPPORT:
    /// - Regular PivotTables: Full support
    /// - OLAP PivotTables: Full support
    /// </remarks>
    OperationResult SetGrandTotals(dynamic pivot, bool showRowGrandTotals, bool showColumnGrandTotals, string workbookPath, Microsoft.Extensions.Logging.ILogger? logger = null);
}


