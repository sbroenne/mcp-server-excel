using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Feature", "PivotTables")]
[Trait("Layer", "Core")]
[Trait("Speed", "Fast")]
public sealed class PivotTableCalculatedMemberTypeTests
{
    [Theory]
    [InlineData("Application-defined or object-defined error")]
    [InlineData("Query execution failed because the provider is unavailable")]
    public void CreateCalculatedMember_GenericExcelError_IsNotFormulaSyntax(string message)
    {
#pragma warning disable CA2201 // Exercise Excel's COM error classification without starting Excel.
        var error = new COMException(message, unchecked((int)0x800A03EC));
#pragma warning restore CA2201

        Assert.False(PivotTableCommands.IsCalculatedMemberFormulaError(error));
        var failure = PivotTableCommands.DescribeCalculatedMemberFailure(error, CalculatedMemberType.Measure);
        Assert.Same(error, failure.InnerException);
        Assert.Equal(unchecked((int)0x800A03EC), failure.InnerException!.HResult);
        Assert.Contains(message, failure.Message);
        Assert.Contains("Excel failed to create calculated Measure", failure.Message);
        Assert.DoesNotContain("Invalid formula syntax", failure.Message);
        Assert.Contains("list-calculated-members", failure.Message);
        Assert.Contains("no changes have been rolled back", failure.Message);
    }

    [Fact]
    public void CreateCalculatedMember_ExplicitSyntaxDiagnostic_RetainsNativeException()
    {
#pragma warning disable CA2201 // Simulate a provider's explicit syntax diagnostic.
        var error = new COMException("SYNTAX error in expression", unchecked((int)0x800A03EC));
#pragma warning restore CA2201

        var failure = PivotTableCommands.DescribeCalculatedMemberFailure(error, CalculatedMemberType.Set);

        Assert.True(PivotTableCommands.IsCalculatedMemberFormulaError(error));
        Assert.Same(error, failure.InnerException);
        Assert.Contains("Invalid formula syntax for calculated Set", failure.Message);
        Assert.Contains(error.Message, failure.Message);
    }

    [Fact]
    public void CreateCalculatedMember_UnknownType_FailsBeforeBatchExecution()
    {
        var commands = new PivotTableCommands();

        var error = Assert.Throws<ArgumentOutOfRangeException>(() =>
            commands.CreateCalculatedMember(
                null!,
                "Pivot",
                "[Measures].[Invalid]",
                "1",
                (CalculatedMemberType)int.MaxValue));

        Assert.Equal("type", error.ParamName);
    }
}
