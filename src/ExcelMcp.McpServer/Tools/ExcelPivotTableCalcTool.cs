using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for PivotTable calculated fields, members, layout, and data operations
/// </summary>
[McpServerToolType]
public static partial class ExcelPivotTableCalcTool
{
    private static readonly JsonSerializerOptions JsonOptions = ExcelToolsBase.JsonOptions;

    /// <summary>
    /// PivotTable calculated fields/members, layout configuration, and data extraction.
    ///
    /// CALCULATED FIELDS (for regular PivotTables):
    /// - Create custom fields using formulas like '=Revenue-Cost'
    /// - Can reference existing fields by name
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
            action.ToActionString(),
            () =>
            {
                var commands = new PivotTableCommands();

                return action switch
                {
                    PivotTableCalcAction.ListCalculatedFields => ListCalculatedFields(commands, sessionId, pivotTableName),
                    PivotTableCalcAction.CreateCalculatedField => CreateCalculatedField(commands, sessionId, pivotTableName, fieldName, formula),
                    PivotTableCalcAction.DeleteCalculatedField => DeleteCalculatedField(commands, sessionId, pivotTableName, fieldName),
                    PivotTableCalcAction.ListCalculatedMembers => ListCalculatedMembers(commands, sessionId, pivotTableName),
                    PivotTableCalcAction.CreateCalculatedMember => CreateCalculatedMember(commands, sessionId, pivotTableName, fieldName, formula, memberType, numberFormat),
                    PivotTableCalcAction.DeleteCalculatedMember => DeleteCalculatedMember(commands, sessionId, pivotTableName, fieldName),
                    PivotTableCalcAction.SetLayout => SetLayout(commands, sessionId, pivotTableName, layoutStyle),
                    PivotTableCalcAction.SetSubtotals => SetSubtotals(commands, sessionId, pivotTableName, fieldName, subtotalsVisible),
                    PivotTableCalcAction.SetGrandTotals => SetGrandTotals(commands, sessionId, pivotTableName, showRowGrandTotals, showColumnGrandTotals),
                    PivotTableCalcAction.GetData => GetData(commands, sessionId, pivotTableName),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ListCalculatedFields(PivotTableCommands commands, string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for list-calculated-fields action", nameof(pivotTableName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.ListCalculatedFields(batch, pivotTableName!));

        return JsonSerializer.Serialize(new { result.Success, result.CalculatedFields, result.ErrorMessage }, JsonOptions);
    }

    private static string CreateCalculatedField(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, string? formula)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for create-calculated-field action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for create-calculated-field action", nameof(fieldName));
        if (string.IsNullOrEmpty(formula))
            throw new ArgumentException("formula is required for create-calculated-field action", nameof(formula));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.CreateCalculatedField(batch, pivotTableName!, fieldName!, formula!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Formula,
            result.ErrorMessage,
            result.WorkflowHint
        }, JsonOptions);
    }

    private static string DeleteCalculatedField(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for delete-calculated-field action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for delete-calculated-field action", nameof(fieldName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.DeleteCalculatedField(batch, pivotTableName!, fieldName!));

        return JsonSerializer.Serialize(new { result.Success, result.ErrorMessage }, JsonOptions);
    }

    private static string ListCalculatedMembers(PivotTableCommands commands, string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for list-calculated-members action", nameof(pivotTableName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.ListCalculatedMembers(batch, pivotTableName!));

        return JsonSerializer.Serialize(new { result.Success, result.CalculatedMembers, result.ErrorMessage }, JsonOptions);
    }

    private static string CreateCalculatedMember(PivotTableCommands commands, string sessionId, string? pivotTableName, string? memberName, string? formula, string? memberType, string? numberFormat)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for create-calculated-member action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(memberName))
            throw new ArgumentException("fieldName is required for create-calculated-member action", nameof(memberName));
        if (string.IsNullOrEmpty(formula))
            throw new ArgumentException("formula is required for create-calculated-member action", nameof(formula));

        CalculatedMemberType type = CalculatedMemberType.Measure;
        if (!string.IsNullOrEmpty(memberType) && !Enum.TryParse(memberType, true, out type))
        {
            throw new ArgumentException($"Invalid memberType '{memberType}'. Valid: Member, Set, Measure", nameof(memberType));
        }

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.CreateCalculatedMember(batch, pivotTableName!, memberName!, formula!, type, 0, null, numberFormat));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Name,
            result.Formula,
            result.Type,
            result.SolveOrder,
            result.IsValid,
            result.DisplayFolder,
            result.NumberFormat,
            result.WorkflowHint,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string DeleteCalculatedMember(PivotTableCommands commands, string sessionId, string? pivotTableName, string? memberName)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for delete-calculated-member action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(memberName))
            throw new ArgumentException("fieldName is required for delete-calculated-member action", nameof(memberName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.DeleteCalculatedMember(batch, pivotTableName!, memberName!));

        return JsonSerializer.Serialize(new { result.Success, result.ErrorMessage }, JsonOptions);
    }

    private static string SetLayout(PivotTableCommands commands, string sessionId, string? pivotTableName, int? layoutStyle)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for set-layout action", nameof(pivotTableName));
        if (!layoutStyle.HasValue)
            throw new ArgumentException("layoutStyle is required for set-layout action (0=Compact, 1=Tabular, 2=Outline)", nameof(layoutStyle));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetLayout(batch, pivotTableName!, layoutStyle.Value));

        return JsonSerializer.Serialize(new { result.Success, result.ErrorMessage }, JsonOptions);
    }

    private static string SetSubtotals(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, bool? subtotalsVisible)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for set-subtotals action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for set-subtotals action", nameof(fieldName));
        if (!subtotalsVisible.HasValue)
            throw new ArgumentException("subtotalsVisible is required for set-subtotals action", nameof(subtotalsVisible));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetSubtotals(batch, pivotTableName!, fieldName!, subtotalsVisible.Value));

        return JsonSerializer.Serialize(new { result.Success, result.FieldName, result.ErrorMessage, result.WorkflowHint }, JsonOptions);
    }

    private static string SetGrandTotals(PivotTableCommands commands, string sessionId, string? pivotTableName, bool? showRowGrandTotals, bool? showColumnGrandTotals)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for set-grand-totals action", nameof(pivotTableName));
        if (!showRowGrandTotals.HasValue)
            throw new ArgumentException("showRowGrandTotals is required for set-grand-totals action", nameof(showRowGrandTotals));
        if (!showColumnGrandTotals.HasValue)
            throw new ArgumentException("showColumnGrandTotals is required for set-grand-totals action", nameof(showColumnGrandTotals));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetGrandTotals(batch, pivotTableName!, showRowGrandTotals.Value, showColumnGrandTotals.Value));

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static string GetData(PivotTableCommands commands, string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "get-data");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.GetData(batch, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTableName,
            result.Values,
            result.ColumnHeaders,
            result.RowHeaders,
            result.DataRowCount,
            result.DataColumnCount,
            result.GrandTotals,
            result.ErrorMessage
        }, JsonOptions);
    }
}
