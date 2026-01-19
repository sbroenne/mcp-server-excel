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
    PivotFieldResult GroupByDate(IExcelBatch batch, string pivotTableName,
        string fieldName, DateGroupingInterval interval);

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
    PivotFieldResult GroupByNumeric(IExcelBatch batch, string pivotTableName,
        string fieldName, double? start, double? endValue, double intervalSize);

    /// <summary>
    /// Creates a calculated field with a custom formula.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name for the calculated field</param>
    /// <param name="formula">Formula using field references (e.g., "=Revenue-Cost")</param>
    /// <returns>Result with calculated field details</returns>
    /// <remarks>
    /// Formula examples:
    /// - "=Revenue-Cost" creates Profit field
    /// - "=Profit/Revenue" creates Margin field
    /// - "=(Actual-Budget)/Budget" creates Variance% field
    ///
    /// NOTE: OLAP PivotTables do not support CalculatedFields.
    /// For OLAP, use Data Model DAX measures instead.
    /// </remarks>
    PivotFieldResult CreateCalculatedField(IExcelBatch batch, string pivotTableName,
        string fieldName, string formula);

    /// <summary>
    /// Lists all calculated fields in a regular PivotTable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <returns>List of calculated fields with names and formulas</returns>
    /// <remarks>
    /// NOTE: OLAP PivotTables do not support CalculatedFields.
    /// Use ListCalculatedMembers for OLAP PivotTables instead.
    /// </remarks>
    CalculatedFieldListResult ListCalculatedFields(IExcelBatch batch, string pivotTableName);

    /// <summary>
    /// Deletes a calculated field from a regular PivotTable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the calculated field to delete</param>
    /// <returns>Result indicating success or failure</returns>
    /// <remarks>
    /// NOTE: OLAP PivotTables do not support CalculatedFields.
    /// Use DeleteCalculatedMember for OLAP PivotTables instead.
    /// </remarks>
    OperationResult DeleteCalculatedField(IExcelBatch batch, string pivotTableName, string fieldName);

    /// <summary>
    /// Sets the row layout form for a PivotTable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="layoutType">Layout form: 0=Compact, 1=Tabular, 2=Outline</param>
    /// <returns>Result indicating success or failure</returns>
    /// <remarks>
    /// LAYOUT FORMS:
    /// - Compact (0): All row fields in single column with indentation (Excel default)
    /// - Tabular (1): Each field in separate column, subtotals at bottom
    /// - Outline (2): Each field in separate column, subtotals at top
    ///
    /// Supported by both regular and OLAP PivotTables.
    /// </remarks>
    OperationResult SetLayout(IExcelBatch batch, string pivotTableName, int layoutType);

    /// <summary>
    /// Shows or hides subtotals for a specific row field.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the row field</param>
    /// <param name="showSubtotals">True to show automatic subtotals, false to hide</param>
    /// <returns>Result with updated field configuration</returns>
    /// <remarks>
    /// SUBTOTALS:
    /// - Enabled: Shows automatic subtotals (Sum for numbers, Count for text)
    /// - Disabled: Hides all subtotals, shows only detail rows
    ///
    /// OLAP PivotTables only support Automatic subtotals.
    /// </remarks>
    PivotFieldResult SetSubtotals(IExcelBatch batch, string pivotTableName,
        string fieldName, bool showSubtotals);

    /// <summary>
    /// Shows or hides grand totals for rows and/or columns in the PivotTable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable to configure</param>
    /// <param name="showRowGrandTotals">Show row grand totals (bottom summary row)</param>
    /// <param name="showColumnGrandTotals">Show column grand totals (right summary column)</param>
    /// <returns>Operation result indicating success or failure</returns>
    /// <remarks>
    /// GRAND TOTALS:
    /// - Row Grand Totals: Summary row at bottom of PivotTable
    /// - Column Grand Totals: Summary column at right of PivotTable
    /// - Independent control: Can show/hide row and column separately
    ///
    /// SUPPORT:
    /// - Regular PivotTables: Full support
    /// - OLAP PivotTables: Full support
    /// </remarks>
    OperationResult SetGrandTotals(IExcelBatch batch, string pivotTableName,
        bool showRowGrandTotals, bool showColumnGrandTotals);

    // === CALCULATED MEMBERS (OLAP ONLY) ===

    /// <summary>
    /// Lists all calculated members in an OLAP PivotTable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <returns>List of calculated members with names, formulas, types</returns>
    /// <remarks>
    /// OLAP ONLY: Calculated members are only available for OLAP PivotTables (Data Model-based).
    /// Regular PivotTables use calculated fields instead (see CreateCalculatedField).
    ///
    /// CALCULATED MEMBER TYPES:
    /// - Member: Custom MDX formula creating a new member in a hierarchy
    /// - Set: Named set of members for filtering/grouping
    /// - Measure: DAX-like calculated measure for Data Model
    /// </remarks>
    CalculatedMemberListResult ListCalculatedMembers(IExcelBatch batch, string pivotTableName);

    /// <summary>
    /// Creates a calculated member (MDX formula) in an OLAP PivotTable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="memberName">Name for the calculated member (MDX naming format)</param>
    /// <param name="formula">MDX formula for the calculated member</param>
    /// <param name="type">Type of calculated member (Member, Set, or Measure)</param>
    /// <param name="solveOrder">Solve order for calculation precedence (default: 0)</param>
    /// <param name="displayFolder">Display folder path for organizing measures (optional)</param>
    /// <param name="numberFormat">Number format code for the calculated member (optional)</param>
    /// <returns>Result with created calculated member details</returns>
    /// <remarks>
    /// OLAP ONLY: Works only with OLAP PivotTables (Data Model-based).
    /// Regular PivotTables should use CreateCalculatedField instead.
    ///
    /// MDX FORMULA EXAMPLES:
    /// - Measure: "[Measures].[Profit]" formula = "[Measures].[Revenue] - [Measures].[Cost]"
    /// - Member: "[Product].[Category].[All].[High Margin]" formula = "Aggregate({[Product].[Category].&amp;[A], [Product].[Category].&amp;[B]})"
    ///
    /// SOLVE ORDER:
    /// - Higher solve order = calculated later (can reference lower solve order members)
    /// - Default is 0, use higher values for dependent calculations
    /// </remarks>
    CalculatedMemberResult CreateCalculatedMember(IExcelBatch batch, string pivotTableName,
        string memberName, string formula, CalculatedMemberType type = CalculatedMemberType.Measure,
        int solveOrder = 0, string? displayFolder = null, string? numberFormat = null);

    /// <summary>
    /// Deletes a calculated member from an OLAP PivotTable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="memberName">Name of the calculated member to delete</param>
    /// <returns>Operation result indicating success or failure</returns>
    /// <remarks>
    /// OLAP ONLY: Works only with OLAP PivotTables (Data Model-based).
    /// </remarks>
    OperationResult DeleteCalculatedMember(IExcelBatch batch, string pivotTableName, string memberName);

    // === SLICER OPERATIONS ===

    /// <summary>
    /// Creates a slicer for a PivotTable field.
    /// Slicers provide visual filtering for PivotTable data.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable to create slicer for</param>
    /// <param name="fieldName">Name of the field to use for the slicer</param>
    /// <param name="slicerName">Name for the new slicer</param>
    /// <param name="destinationSheet">Worksheet where slicer will be placed</param>
    /// <param name="position">Top-left cell position for the slicer (e.g., "H2")</param>
    /// <returns>Created slicer details with available items</returns>
    /// <remarks>
    /// SLICER BEHAVIOR:
    /// - Slicers are visual filter controls that can filter one or more PivotTables
    /// - One SlicerCache is created per field, which can have multiple visual Slicers
    /// - Multiple PivotTables can be connected to the same SlicerCache
    /// 
    /// SUPPORTED:
    /// - Regular PivotTables: Full support
    /// - OLAP PivotTables: Full support (Data Model-based)
    /// </remarks>
    SlicerResult CreateSlicer(IExcelBatch batch, string pivotTableName,
        string fieldName, string slicerName, string destinationSheet, string position);

    /// <summary>
    /// Lists all slicers in the workbook, optionally filtered by PivotTable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Optional PivotTable name to filter slicers (null = all slicers)</param>
    /// <returns>List of slicers with names, fields, positions, and selections</returns>
    /// <remarks>
    /// Returns slicers from all SlicerCaches in the workbook.
    /// When pivotTableName is specified, only slicers connected to that PivotTable are returned.
    /// </remarks>
    SlicerListResult ListSlicers(IExcelBatch batch, string? pivotTableName = null);

    /// <summary>
    /// Sets the selection for a slicer, filtering the connected PivotTable(s).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the slicer to modify</param>
    /// <param name="selectedItems">Items to select (show in PivotTable)</param>
    /// <param name="clearFirst">If true, clears existing selection before setting new items (default: true)</param>
    /// <returns>Updated slicer state with current selection</returns>
    /// <remarks>
    /// SELECTION BEHAVIOR:
    /// - Only selected items are visible in connected PivotTable(s)
    /// - Empty selectedItems list shows all items (clears filter)
    /// - Invalid item names are ignored with a warning
    /// 
    /// MULTI-PIVOTTABLE:
    /// - Selection change affects ALL PivotTables connected to this slicer
    /// </remarks>
    SlicerResult SetSlicerSelection(IExcelBatch batch, string slicerName,
        List<string> selectedItems, bool clearFirst = true);

    /// <summary>
    /// Deletes a slicer from the workbook.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the slicer to delete</param>
    /// <returns>Operation result indicating success or failure</returns>
    /// <remarks>
    /// DELETION BEHAVIOR:
    /// - Deletes the visual Slicer object
    /// - If this is the last Slicer using the SlicerCache, the cache is also deleted
    /// - Connected PivotTable filters are cleared when slicer is deleted
    /// </remarks>
    OperationResult DeleteSlicer(IExcelBatch batch, string slicerName);
}
