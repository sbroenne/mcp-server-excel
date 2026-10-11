using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "PivotTables")]
[Trait("Layer", "Core")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class OlapRecoveryGuidanceTests
{
    [Theory]
    [InlineData("[Measures].[Profit]", "[Measures].[Profit]")]
    [InlineData("Profit", "[Measures].[Profit]")]
    public void MeasureHint_UsesOneQualifiedName(string name, string expected)
    {
        var hint = PivotTableCommands.GetCalculatedMemberWorkflowHint(name, CalculatedMemberType.Measure);
        Assert.Contains($"'{expected}'", hint);
        Assert.DoesNotContain("[Measures].[[Measures]", hint);
        Assert.Contains("add-value-field", hint);
        Assert.Contains("pivottable refresh", hint);
        Assert.Contains("field_name=", hint);
        Assert.Contains("--field-name", hint);
    }

    [Theory]
    [InlineData(CalculatedMemberType.Member)]
    [InlineData(CalculatedMemberType.Set)]
    public void NonMeasureHint_DoesNotInventValuesField(CalculatedMemberType type)
    {
        var hint = PivotTableCommands.GetCalculatedMemberWorkflowHint("[Regions]", type);
        Assert.DoesNotContain("add-value-field", hint);
        Assert.DoesNotContain("[Measures]", hint);
        Assert.Contains("list-calculated-members", hint);
    }

    [Fact]
    public void PartialChartFailure_PreservesIdentityAndCauseWithoutRollbackClaim()
    {
        var cause = new InvalidOperationException("Source binding failed");
        var failure = ChartCommands.DescribePivotChartCreationFailure("Chart 7", "Dashboard", "Binding the source", cause);
        Assert.Same(cause, failure.InnerException);
        Assert.Contains("Chart 7", failure.Message);
        Assert.Contains("Dashboard", failure.Message);
        Assert.Contains("remains", failure.Message);
        Assert.DoesNotContain("No chart was kept", failure.Message);
    }

}
