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
    /// PivotTable calculated fields/members, layout, data extraction.
    /// CALCULATED FIELDS: For regular PivotTables (formula='=Revenue-Cost').
    /// CALCULATED MEMBERS: For OLAP/Data Model PivotTables (MDX expressions).
    /// Related: excel_pivottable (lifecycle), excel_pivottable_field (fields)
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="sid">Session ID</param>
    /// <param name="ptn">PivotTable name</param>
    /// <param name="fn">Field/member name</param>
    /// <param name="formula">Formula (calculated field: '=Revenue-Cost', calculated member: MDX expression)</param>
    /// <param name="mt">Member type: Member, Set, Measure</param>
    /// <param name="nf">Number format</param>
    /// <param name="layout">Layout: 0=Compact, 1=Tabular, 2=Outline</param>
    /// <param name="stv">Subtotals visible (true/false)</param>
    /// <param name="rgt">Show row grand totals</param>
    /// <param name="cgt">Show column grand totals</param>
    [McpServerTool(Name = "excel_pivottable_calc", Title = "Excel PivotTable Calc Operations")]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelPivotTableCalc(
        PivotTableCalcAction action,
        string sid,
        string ptn,
        [DefaultValue(null)] string? fn,
        [DefaultValue(null)] string? formula,
        [DefaultValue(null)] string? mt,
        [DefaultValue(null)] string? nf,
        [DefaultValue(null)] int? layout,
        [DefaultValue(null)] bool? stv,
        [DefaultValue(null)] bool? rgt,
        [DefaultValue(null)] bool? cgt)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_pivottable_calc",
            action.ToActionString(),
            () =>
            {
                var commands = new PivotTableCommands();

                return action switch
                {
                    PivotTableCalcAction.ListCalculatedFields => ListCalculatedFields(commands, sid, ptn),
                    PivotTableCalcAction.CreateCalculatedField => CreateCalculatedField(commands, sid, ptn, fn, formula),
                    PivotTableCalcAction.DeleteCalculatedField => DeleteCalculatedField(commands, sid, ptn, fn),
                    PivotTableCalcAction.ListCalculatedMembers => ListCalculatedMembers(commands, sid, ptn),
                    PivotTableCalcAction.CreateCalculatedMember => CreateCalculatedMember(commands, sid, ptn, fn, formula, mt, nf),
                    PivotTableCalcAction.DeleteCalculatedMember => DeleteCalculatedMember(commands, sid, ptn, fn),
                    PivotTableCalcAction.SetLayout => SetLayout(commands, sid, ptn, layout),
                    PivotTableCalcAction.SetSubtotals => SetSubtotals(commands, sid, ptn, fn, stv),
                    PivotTableCalcAction.SetGrandTotals => SetGrandTotals(commands, sid, ptn, rgt, cgt),
                    PivotTableCalcAction.GetData => GetData(commands, sid, ptn),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ListCalculatedFields(PivotTableCommands commands, string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("ptn is required for list-calculated-fields action", nameof(pivotTableName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.ListCalculatedFields(batch, pivotTableName!));

        return JsonSerializer.Serialize(new { result.Success, result.CalculatedFields, result.ErrorMessage }, JsonOptions);
    }

    private static string CreateCalculatedField(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, string? formula)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("ptn is required for create-calculated-field action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fn is required for create-calculated-field action", nameof(fieldName));
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
            throw new ArgumentException("ptn is required for delete-calculated-field action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fn is required for delete-calculated-field action", nameof(fieldName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.DeleteCalculatedField(batch, pivotTableName!, fieldName!));

        return JsonSerializer.Serialize(new { result.Success, result.ErrorMessage }, JsonOptions);
    }

    private static string ListCalculatedMembers(PivotTableCommands commands, string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("ptn is required for list-calculated-members action", nameof(pivotTableName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.ListCalculatedMembers(batch, pivotTableName!));

        return JsonSerializer.Serialize(new { result.Success, result.CalculatedMembers, result.ErrorMessage }, JsonOptions);
    }

    private static string CreateCalculatedMember(PivotTableCommands commands, string sessionId, string? pivotTableName, string? memberName, string? formula, string? memberType, string? numberFormat)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("ptn is required for create-calculated-member action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(memberName))
            throw new ArgumentException("fn is required for create-calculated-member action", nameof(memberName));
        if (string.IsNullOrEmpty(formula))
            throw new ArgumentException("formula is required for create-calculated-member action", nameof(formula));

        CalculatedMemberType type = CalculatedMemberType.Measure;
        if (!string.IsNullOrEmpty(memberType))
        {
            if (!Enum.TryParse(memberType, true, out type))
            {
                throw new ArgumentException($"Invalid member type '{memberType}'. Valid: Member, Set, Measure", nameof(memberType));
            }
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
            throw new ArgumentException("ptn is required for delete-calculated-member action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(memberName))
            throw new ArgumentException("fn is required for delete-calculated-member action", nameof(memberName));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.DeleteCalculatedMember(batch, pivotTableName!, memberName!));

        return JsonSerializer.Serialize(new { result.Success, result.ErrorMessage }, JsonOptions);
    }

    private static string SetLayout(PivotTableCommands commands, string sessionId, string? pivotTableName, int? layout)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("ptn is required for set-layout action", nameof(pivotTableName));
        if (!layout.HasValue)
            throw new ArgumentException("layout is required for set-layout action (0=Compact, 1=Tabular, 2=Outline)", nameof(layout));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetLayout(batch, pivotTableName!, layout.Value));

        return JsonSerializer.Serialize(new { result.Success, result.ErrorMessage }, JsonOptions);
    }

    private static string SetSubtotals(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, bool? subtotalsVisible)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("ptn is required for set-subtotals action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fn is required for set-subtotals action", nameof(fieldName));
        if (!subtotalsVisible.HasValue)
            throw new ArgumentException("stv is required for set-subtotals action", nameof(subtotalsVisible));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetSubtotals(batch, pivotTableName!, fieldName!, subtotalsVisible.Value));

        return JsonSerializer.Serialize(new { result.Success, result.FieldName, result.ErrorMessage, result.WorkflowHint }, JsonOptions);
    }

    private static string SetGrandTotals(PivotTableCommands commands, string sessionId, string? pivotTableName, bool? showRowGrandTotals, bool? showColumnGrandTotals)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("ptn is required for set-grand-totals action", nameof(pivotTableName));
        if (!showRowGrandTotals.HasValue)
            throw new ArgumentException("rgt is required for set-grand-totals action", nameof(showRowGrandTotals));
        if (!showColumnGrandTotals.HasValue)
            throw new ArgumentException("cgt is required for set-grand-totals action", nameof(showColumnGrandTotals));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetGrandTotals(batch, pivotTableName!, showRowGrandTotals.Value, showColumnGrandTotals.Value));

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static string GetData(PivotTableCommands commands, string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("ptn", "get-data");

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
