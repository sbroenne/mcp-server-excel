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
            dynamic? pivot = null;
            dynamic? calculatedMembers = null;

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
                    dynamic? member = null;
                    try
                    {
                        member = calculatedMembers.Item(i);
                        var memberInfo = new CalculatedMemberInfo
                        {
                            Name = member.Name?.ToString() ?? string.Empty,
                            Formula = member.Formula?.ToString() ?? string.Empty,
                            Type = GetCalculatedMemberType(Convert.ToInt32(member.Type)),
                            SolveOrder = Convert.ToInt32(member.SolveOrder),
                            IsValid = member.IsValid
                        };

                        // Try to get optional properties (may not exist on all calculated member types)
                        try { memberInfo.DisplayFolder = member.DisplayFolder?.ToString(); } catch (System.Runtime.InteropServices.COMException) { /* Property not available */ }
                        try { memberInfo.NumberFormat = member.NumberFormat?.ToString(); } catch (System.Runtime.InteropServices.COMException) { /* Property not available */ }

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
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? calculatedMembers = null;
            dynamic? newMember = null;

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
                    _ => XlCalculatedMemberType.xlCalculatedMeasure
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
                catch (System.Runtime.InteropServices.COMException comEx)
                {
                    // MDX/DAX syntax errors or invalid formulas return specific COM errors
                    // Convert these to user-friendly error messages
                    string errorDetail = comEx.Message;
                    bool isFormulaError = errorDetail.Contains("Query") ||
                        errorDetail.Contains("syntax") ||
                        comEx.HResult == unchecked((int)0x800A03EC);

                    if (isFormulaError)
                    {
                        return new CalculatedMemberResult
                        {
                            Success = false,
                            ErrorMessage = $"Invalid formula syntax for calculated {type}: {errorDetail}. Check MDX/DAX syntax and ensure referenced measures/dimensions exist."
                        };
                    }
                    // Re-throw unknown COM errors
                    throw;
                }

                var result = new CalculatedMemberResult
                {
                    Success = true,
                    Name = newMember.Name?.ToString() ?? memberName,
                    Formula = newMember.Formula?.ToString() ?? formula,
                    Type = GetCalculatedMemberType(Convert.ToInt32(newMember.Type)),
                    SolveOrder = Convert.ToInt32(newMember.SolveOrder),
                    IsValid = newMember.IsValid,
                    WorkflowHint = $"Created calculated {type} '{memberName}'. Use add-value-field with fieldName='[Measures].[{memberName}]' to add it to the PivotTable values area."
                };

                // Try to get optional properties (may not exist on all calculated member types)
                try { result.DisplayFolder = newMember.DisplayFolder?.ToString(); } catch (System.Runtime.InteropServices.COMException) { /* Property not available */ }
                try { result.NumberFormat = newMember.NumberFormat?.ToString(); } catch (System.Runtime.InteropServices.COMException) { /* Property not available */ }

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
            dynamic? pivot = null;
            dynamic? calculatedMembers = null;
            dynamic? member = null;

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

                // Find the member by name
                try
                {
                    member = calculatedMembers.Item(memberName);
                }
                catch (COMException)
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


