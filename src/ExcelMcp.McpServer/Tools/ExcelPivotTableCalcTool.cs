using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for PivotTable calculated fields, members, layout, and data operations
/// </summary>
[McpServerToolType]
public static partial class ExcelPivotTableCalcTool
{
    /// <summary>
    /// PivotTable calculated fields/members, layout configuration, and data extraction.
    ///
    /// CALCULATED FIELDS (for regular PivotTables):
    /// - Create custom fields using formulas like '=Revenue-Cost' or '=Quantity*UnitPrice'
    /// - Can reference existing fields by name
    /// - After creating, use excel_pivottable_field AddValueField to add to Values area
    /// - For complex multi-table calculations, prefer DAX measures with excel_datamodel
    ///
    /// CALCULATED MEMBERS (for OLAP/Data Model PivotTables only):
    /// - Create using MDX expressions
    /// - Member types: Member, Set, Measure
    ///
    /// LAYOUT OPTIONS:
    /// - 0 = Compact (default, fields in single column)
    /// - 1 = Tabular (each field in separate column)
    /// - 2 = Outline (hierarchical with expand/collapse)
    ///
    /// RELATED TOOLS:
    /// - excel_pivottable: Create/delete/refresh PivotTables
    /// - excel_pivottable_field: Add/remove/configure fields
    /// </summary>
    /// <param name="action">The calculation/layout operation to perform</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action</param>
    /// <param name="pivotTableName">Name of the PivotTable to modify</param>
    /// <param name="fieldName">Name of the calculated field or member</param>
    /// <param name="formula">Formula for calculated field (e.g., '=Revenue-Cost') or MDX expression for calculated member</param>
    /// <param name="memberType">Calculated member type: Member, Set, or Measure</param>
    /// <param name="numberFormat">Number format code in US format, e.g., '#,##0.00' for currency</param>
    /// <param name="layoutStyle">Layout style: 0=Compact, 1=Tabular, 2=Outline</param>
    /// <param name="subtotalsVisible">Whether to show subtotals for the specified field</param>
    /// <param name="showRowGrandTotals">Whether to show grand totals for rows</param>
    /// <param name="showColumnGrandTotals">Whether to show grand totals for columns</param>
    [McpServerTool(Name = "excel_pivottable_calc", Title = "Excel PivotTable Calc Operations", Destructive = true)]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelPivotTableCalc(
        PivotTableCalcAction action,
        string sessionId,
        string pivotTableName,
        [DefaultValue(null)] string? fieldName,
        [DefaultValue(null)] string? formula,
        [DefaultValue(null)] string? memberType,
        [DefaultValue(null)] string? numberFormat,
        [DefaultValue(null)] int? layoutStyle,
        [DefaultValue(null)] bool? subtotalsVisible,
        [DefaultValue(null)] bool? showRowGrandTotals,
        [DefaultValue(null)] bool? showColumnGrandTotals)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_pivottable_calc",
            ServiceRegistry.PivotTableCalc.ToActionString(action),
            () => action switch
            {
                PivotTableCalcAction.ListCalculatedFields => ForwardListCalculatedFields(sessionId, pivotTableName),
                PivotTableCalcAction.CreateCalculatedField => ForwardCreateCalculatedField(sessionId, pivotTableName, fieldName, formula),
                PivotTableCalcAction.DeleteCalculatedField => ForwardDeleteCalculatedField(sessionId, pivotTableName, fieldName),
                PivotTableCalcAction.ListCalculatedMembers => ForwardListCalculatedMembers(sessionId, pivotTableName),
                PivotTableCalcAction.CreateCalculatedMember => ForwardCreateCalculatedMember(sessionId, pivotTableName, fieldName, formula, memberType, numberFormat),
                PivotTableCalcAction.DeleteCalculatedMember => ForwardDeleteCalculatedMember(sessionId, pivotTableName, fieldName),
                PivotTableCalcAction.SetLayout => ForwardSetLayout(sessionId, pivotTableName, layoutStyle),
                PivotTableCalcAction.SetSubtotals => ForwardSetSubtotals(sessionId, pivotTableName, fieldName, subtotalsVisible),
                PivotTableCalcAction.SetGrandTotals => ForwardSetGrandTotals(sessionId, pivotTableName, showRowGrandTotals, showColumnGrandTotals),
                PivotTableCalcAction.GetData => ForwardGetData(sessionId, pivotTableName),
                _ => throw new ArgumentException($"Unknown action: {action} ({ServiceRegistry.PivotTableCalc.ToActionString(action)})", nameof(action))
            });
    }

    // === SERVICE FORWARDING METHODS ===

    private static string ForwardListCalculatedFields(string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for list-calculated-fields action", nameof(pivotTableName));

        return ExcelToolsBase.ForwardToService("pivottablecalc.list-calculated-fields", sessionId, new { pivotTableName });
    }

    private static string ForwardCreateCalculatedField(string sessionId, string? pivotTableName, string? fieldName, string? formula)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for create-calculated-field action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for create-calculated-field action", nameof(fieldName));
        if (string.IsNullOrEmpty(formula))
            throw new ArgumentException("formula is required for create-calculated-field action", nameof(formula));

        return ExcelToolsBase.ForwardToService("pivottablecalc.create-calculated-field", sessionId, new
        {
            pivotTableName,
            fieldName,
            formula
        });
    }

    private static string ForwardDeleteCalculatedField(string sessionId, string? pivotTableName, string? fieldName)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for delete-calculated-field action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for delete-calculated-field action", nameof(fieldName));

        return ExcelToolsBase.ForwardToService("pivottablecalc.delete-calculated-field", sessionId, new
        {
            pivotTableName,
            fieldName
        });
    }

    private static string ForwardListCalculatedMembers(string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for list-calculated-members action", nameof(pivotTableName));

        return ExcelToolsBase.ForwardToService("pivottablecalc.list-calculated-members", sessionId, new { pivotTableName });
    }

    private static string ForwardCreateCalculatedMember(string sessionId, string? pivotTableName, string? memberName, string? formula, string? memberType, string? numberFormat)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for create-calculated-member action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(memberName))
            throw new ArgumentException("fieldName is required for create-calculated-member action", nameof(memberName));
        if (string.IsNullOrEmpty(formula))
            throw new ArgumentException("formula is required for create-calculated-member action", nameof(formula));

        return ExcelToolsBase.ForwardToService("pivottablecalc.create-calculated-member", sessionId, new
        {
            pivotTableName,
            memberName,
            formula,
            memberType,
            numberFormat
        });
    }

    private static string ForwardDeleteCalculatedMember(string sessionId, string? pivotTableName, string? memberName)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for delete-calculated-member action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(memberName))
            throw new ArgumentException("fieldName is required for delete-calculated-member action", nameof(memberName));

        return ExcelToolsBase.ForwardToService("pivottablecalc.delete-calculated-member", sessionId, new
        {
            pivotTableName,
            memberName
        });
    }

    private static string ForwardSetLayout(string sessionId, string? pivotTableName, int? layoutStyle)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for set-layout action", nameof(pivotTableName));
        if (!layoutStyle.HasValue)
            throw new ArgumentException("layoutStyle is required for set-layout action (0=Compact, 1=Tabular, 2=Outline)", nameof(layoutStyle));

        return ExcelToolsBase.ForwardToService("pivottablecalc.set-layout", sessionId, new
        {
            pivotTableName,
            layoutType = layoutStyle
        });
    }

    private static string ForwardSetSubtotals(string sessionId, string? pivotTableName, string? fieldName, bool? subtotalsVisible)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for set-subtotals action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for set-subtotals action", nameof(fieldName));
        if (!subtotalsVisible.HasValue)
            throw new ArgumentException("subtotalsVisible is required for set-subtotals action", nameof(subtotalsVisible));

        return ExcelToolsBase.ForwardToService("pivottablecalc.set-subtotals", sessionId, new
        {
            pivotTableName,
            fieldName,
            showSubtotals = subtotalsVisible
        });
    }

    private static string ForwardSetGrandTotals(string sessionId, string? pivotTableName, bool? showRowGrandTotals, bool? showColumnGrandTotals)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for set-grand-totals action", nameof(pivotTableName));
        if (!showRowGrandTotals.HasValue)
            throw new ArgumentException("showRowGrandTotals is required for set-grand-totals action", nameof(showRowGrandTotals));
        if (!showColumnGrandTotals.HasValue)
            throw new ArgumentException("showColumnGrandTotals is required for set-grand-totals action", nameof(showColumnGrandTotals));

        return ExcelToolsBase.ForwardToService("pivottablecalc.set-grand-totals", sessionId, new
        {
            pivotTableName,
            showRowGrandTotals,
            showColumnGrandTotals
        });
    }

    private static string ForwardGetData(string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "get-data");

        return ExcelToolsBase.ForwardToService("pivottablecalc.get-data", sessionId, new { pivotTableName });
    }
}




