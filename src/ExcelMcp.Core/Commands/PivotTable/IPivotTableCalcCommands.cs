using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable calculated fields/members, layout configuration, and data extraction.
/// Use pivottable for lifecycle, pivottablefield for field placement.
///
/// CALCULATED FIELDS (for regular PivotTables):
/// - Create custom fields using formulas like '=Revenue-Cost' or '=Quantity*UnitPrice'
/// - Can reference existing fields by name
/// - After creating, use pivottablefield add-value-field to add to Values area
/// - For complex multi-table calculations, prefer DAX measures with datamodel
///
/// CALCULATED MEMBERS (for OLAP/Data Model PivotTables only):
/// - Create using MDX expressions
/// - Member types: Member, Set, Measure
///
/// LAYOUT OPTIONS:
/// - 0 = Compact (default, fields in single column)
/// - 1 = Tabular (each field in separate column - best for export/analysis)
/// - 2 = Outline (hierarchical with expand/collapse)
/// </summary>
[ServiceCategory("pivottablecalc", "PivotTableCalc")]
[McpTool("excel_pivottable_calc", Title = "Excel PivotTable Calc Operations", Destructive = true, Category = "analysis",
    Description = "PivotTable calculated fields/members, layout configuration, and data extraction. CALCULATED FIELDS: Create formulas like =Revenue-Cost, then add to Values with excel_pivottable_field. CALCULATED MEMBERS: MDX expressions (OLAP/Data Model only). LAYOUT: 0=Compact, 1=Tabular, 2=Outline. Use excel_pivottable for lifecycle, excel_pivottable_field for field management.")]
public interface IPivotTableCalcCommands
{
    // === ANALYSIS OPERATIONS (WITH DATA VALIDATION) ===

    /// <summary>
    /// Gets current PivotTable data as 2D array for LLM analysis
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <returns>Values with headers, row/column labels, formatted numbers</returns>
    [ServiceAction("get-data")]
    PivotTableDataResult GetData(IExcelBatch batch, string pivotTableName);

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
    [ServiceAction("create-calculated-field")]
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
    [ServiceAction("list-calculated-fields")]
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
    [ServiceAction("delete-calculated-field")]
    OperationResult DeleteCalculatedField(IExcelBatch batch, string pivotTableName, string fieldName);

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
    [ServiceAction("list-calculated-members")]
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
    [ServiceAction("create-calculated-member")]
    CalculatedMemberResult CreateCalculatedMember(IExcelBatch batch, string pivotTableName,
        string memberName, string formula, [FromString] CalculatedMemberType type = CalculatedMemberType.Measure,
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
    [ServiceAction("delete-calculated-member")]
    OperationResult DeleteCalculatedMember(IExcelBatch batch, string pivotTableName, string memberName);

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
    [ServiceAction("set-layout")]
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
    [ServiceAction("set-subtotals")]
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
    [ServiceAction("set-grand-totals")]
    OperationResult SetGrandTotals(IExcelBatch batch, string pivotTableName,
        bool showRowGrandTotals, bool showColumnGrandTotals);
}
