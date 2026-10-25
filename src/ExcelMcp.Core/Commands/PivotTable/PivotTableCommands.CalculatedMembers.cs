using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Calculated Members operations for PivotTableCommands.
/// Creates, lists, and deletes calculated members for OLAP PivotTables.
/// Note: Only works with OLAP (Data Model) PivotTables. Regular PivotTables use CalculatedFields.
/// </summary>
public partial class PivotTableCommands
{
    /// <inheritdoc/>
    public CalculatedMemberListResult ListCalculatedMembers(IExcelBatch batch, string pivotTableName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.CalculatedMembers? calculatedMembers = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Check if this is an OLAP PivotTable
                if (!PivotTableHelpers.IsOlapPivotTable(pivot))
                {
                    return new CalculatedMemberListResult
                    {
                        Success = false,
                        ErrorMessage = $"PivotTable '{pivotTableName}' is not an OLAP PivotTable. Calculated members are only available for OLAP (Data Model) PivotTables. Use create-calculated-field for regular PivotTables."
                    };
                }

                calculatedMembers = pivot.CalculatedMembers;
                var result = new CalculatedMemberListResult { Success = true };

                for (int i = 1; i <= calculatedMembers.Count; i++)
                {
                    Excel.CalculatedMember? member = null;
                    try
                    {
                        member = calculatedMembers.Item[i];
                        var memberInfo = new CalculatedMemberInfo
                        {
                            Name = member.Name?.ToString() ?? string.Empty,
                            Formula = member.Formula?.ToString() ?? string.Empty,
                            Type = GetCalculatedMemberType((int)member.Type),
                            SolveOrder = Convert.ToInt32(member.SolveOrder),
                            IsValid = member.IsValid
                        };

                        // Try to get optional properties (may not exist on all calculated member types)
                        if (memberInfo.Type == CalculatedMemberType.Measure)
                        {
                            memberInfo.DisplayFolder = member.DisplayFolder?.ToString();
                            memberInfo.NumberFormat = ((int)member.NumberFormat).ToString(System.Globalization.CultureInfo.InvariantCulture);
                        }

                        result.CalculatedMembers.Add(memberInfo);
                    }
                    finally
                    {
                        ComUtilities.Release(ref member);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref calculatedMembers);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <inheritdoc/>
    public CalculatedMemberResult CreateCalculatedMember(IExcelBatch batch, string pivotTableName,
        string memberName, string formula, CalculatedMemberType type = CalculatedMemberType.Measure,
        int solveOrder = 0, string? displayFolder = null, string? numberFormat = null)
    {
        if (!Enum.IsDefined(type))
        {
            throw new ArgumentOutOfRangeException(nameof(type), type, "Unknown calculated member type.");
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.CalculatedMembers? calculatedMembers = null;
            Excel.CalculatedMember? newMember = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Check if this is an OLAP PivotTable
                if (!PivotTableHelpers.IsOlapPivotTable(pivot))
                {
                    return new CalculatedMemberResult
                    {
                        Success = false,
                        ErrorMessage = $"PivotTable '{pivotTableName}' is not an OLAP PivotTable. Calculated members are only available for OLAP (Data Model) PivotTables. Use create-calculated-field for regular PivotTables."
                    };
                }

                calculatedMembers = pivot.CalculatedMembers;

                // Convert type to COM constant
                int comType = type switch
                {
                    CalculatedMemberType.Member => XlCalculatedMemberType.xlCalculatedMember,
                    CalculatedMemberType.Set => XlCalculatedMemberType.xlCalculatedSet,
                    CalculatedMemberType.Measure => XlCalculatedMemberType.xlCalculatedMeasure,
                    _ => throw new ArgumentOutOfRangeException(nameof(type), type, "Unknown calculated member type.")
                };

                // Use AddCalculatedMember (Excel 2013+) for full feature support
                // Parameters: Name, Formula, SolveOrder, Type, DisplayFolder, MeasureGroup, ParentHierarchy, ParentMember, NumberFormat
                try
                {
                    newMember = calculatedMembers.AddCalculatedMember(
                        memberName,
                        formula,
                        solveOrder,
                        comType,
                        displayFolder ?? Type.Missing,
                        Type.Missing,  // MeasureGroup - auto-detect
                        Type.Missing,  // ParentHierarchy - not needed for measures
                        Type.Missing,  // ParentMember - not needed for measures
                        numberFormat ?? Type.Missing
                    );
                }
                catch (COMException comEx)
                {
                    ct.ThrowIfCancellationRequested();
                    throw DescribeCalculatedMemberFailure(comEx, type);
                }

                var result = new CalculatedMemberResult
                {
                    Success = true,
                    Name = newMember.Name?.ToString() ?? memberName,
                    Formula = newMember.Formula?.ToString() ?? formula,
                    Type = GetCalculatedMemberType((int)newMember.Type),
                    SolveOrder = Convert.ToInt32(newMember.SolveOrder),
                    IsValid = newMember.IsValid,
                    WorkflowHint = GetCalculatedMemberWorkflowHint(memberName, type)
                };

                // Try to get optional properties (may not exist on all calculated member types)
                if (result.Type == CalculatedMemberType.Measure)
                {
                    result.DisplayFolder = newMember.DisplayFolder?.ToString();
                    result.NumberFormat = ((int)newMember.NumberFormat).ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref newMember);
                ComUtilities.Release(ref calculatedMembers);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <inheritdoc/>
    public OperationResult DeleteCalculatedMember(IExcelBatch batch, string pivotTableName, string memberName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.CalculatedMembers? calculatedMembers = null;
            Excel.CalculatedMember? member = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Check if this is an OLAP PivotTable
                if (!PivotTableHelpers.IsOlapPivotTable(pivot))
                {
                    return new OperationResult
                    {
                        Success = false,
                        ErrorMessage = $"PivotTable '{pivotTableName}' is not an OLAP PivotTable. Calculated members are only available for OLAP (Data Model) PivotTables."
                    };
                }

                calculatedMembers = pivot.CalculatedMembers;

                // Enumeration distinguishes absence from a failing native indexer.
                for (int index = 1; index <= Convert.ToInt32(calculatedMembers.Count); index++)
                {
                    ct.ThrowIfCancellationRequested();
                    member = calculatedMembers.Item[index];
                    if (string.Equals(Convert.ToString(member.Name), memberName, StringComparison.OrdinalIgnoreCase))
                        break;
                    ComUtilities.Release(ref member);
                }
                if (member is null)
                {
                    return new OperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Calculated member '{memberName}' not found in PivotTable '{pivotTableName}'. Use list-calculated-members to see available members."
                    };
                }

                member.Delete();

                return new OperationResult
                {
                    Success = true
                };
            }
            finally
            {
                ComUtilities.Release(ref member);
                ComUtilities.Release(ref calculatedMembers);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    internal static bool IsCalculatedMemberFormulaError(COMException error) =>
        error.Message.Contains("syntax", StringComparison.OrdinalIgnoreCase);

    internal static string GetCalculatedMemberWorkflowHint(string memberName, CalculatedMemberType type)
    {
        var created = $"Created calculated {type} '{memberName}'.";
        if (type != CalculatedMemberType.Measure)
        {
            return $"{created} Use list-calculated-members to inspect its name, formula, type, and validity.";
        }

        var fieldName = memberName.StartsWith("[Measures].", StringComparison.OrdinalIgnoreCase)
            ? memberName
            : $"[Measures].[{memberName.Replace("]", "]]", StringComparison.Ordinal)}]";
        return $"{created} Use pivottable refresh to expose the measure in CubeFields, then pivottable_field (CLI: pivottablefield) add-value-field with field_name='{fieldName}' (CLI: --field-name) to add it to the PivotTable values area.";
    }

    internal static InvalidOperationException DescribeCalculatedMemberFailure(
        COMException error, CalculatedMemberType type) =>
        new(
            (IsCalculatedMemberFormulaError(error)
                ? $"Invalid formula syntax for calculated {type}: "
                : $"Excel failed to create calculated {type}: ") +
            $"{error.Message}. Inspect list-calculated-members before retrying; no changes have been rolled back.",
            error);

    /// <summary>
    /// Converts COM calculated member type constant to enum
    /// </summary>
    private static CalculatedMemberType GetCalculatedMemberType(int comType)
    {
        return comType switch
        {
            XlCalculatedMemberType.xlCalculatedMember => CalculatedMemberType.Member,
            XlCalculatedMemberType.xlCalculatedSet => CalculatedMemberType.Set,
            XlCalculatedMemberType.xlCalculatedMeasure => CalculatedMemberType.Measure,
            _ => CalculatedMemberType.Member
        };
    }
}
