using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
public sealed class OlapFieldFailureTests
{
    [Theory]
    [InlineData("SUM('O''Brien'[Net]]Amount])", "O'Brien", "Net]Amount")]
    [InlineData("COUNT('O''''Brien'[Net]]]]Amount])", "O''Brien", "Net]]Amount")]
    public void SingleAggregation_DecodesTableAndColumnEscapes(string formula, string table, string column)
    {
        Assert.Equal((table, column), OlapPivotTableFieldStrategy.ParseDaxFormula(formula));
    }

    [Theory]
    [InlineData("[Measures].[Margin]]Pct]", "Margin]Pct")]
    [InlineData("[Measures].[Margin]]]]Pct]", "Margin]]Pct")]
    [InlineData("[Measures].[Amount.SUM]", "Amount.SUM")]
    [InlineData("Margin]Pct", "Margin]Pct")]
    public void BareMeasureName_DecodesOnlyQualifiedIdentifierEscapes(string qualified, string bare)
    {
        Assert.Equal(bare, OlapPivotTableFieldStrategy.GetBareMeasureName(qualified));
    }

    [Theory]
    [InlineData("Margin]Pct", "[Measures].[Margin]]Pct]")]
    [InlineData("Margin]]Pct", "[Measures].[Margin]]]]Pct]")]
    [InlineData("[Measures].[Margin]]Pct]", "[Measures].[Margin]]Pct]")]
    [InlineData("Amount.SUM", "[Measures].[Amount.SUM]")]
    public void QualifiedMeasureName_EscapesBracketsWithoutDoubleQualification(string name, string qualified)
    {
        Assert.Equal(qualified, OlapPivotTableFieldStrategy.GetQualifiedMeasureName(name));
    }

    [Fact]
    public void SubtotalFailure_IdentifiesCompletedAndUncertainLevelsWithoutClaimingRollback()
    {
        var error = OlapPivotTableFieldStrategy.DescribeSubtotalFailure(
            new InvalidOperationException("Provider rejected Month"), ["[Calendar].[Year]"], "[Calendar].[Month]");

        Assert.Contains("Provider rejected Month", error);
        Assert.Contains("Levels updated: [Calendar].[Year]", error);
        Assert.Contains("Attempted level: [Calendar].[Month]", error);
        Assert.Contains("may also have changed", error);
        Assert.Contains("No changes have been rolled back", error);
        Assert.Contains("Inspect the hierarchy", error);
    }

    [Fact]
    public void SubtotalFailure_BeforeAnyWriteDoesNotInventUpdatedLevels()
    {
        var error = OlapPivotTableFieldStrategy.DescribeSubtotalFailure(
            new InvalidOperationException("Missing hierarchy"), [], null);

        Assert.Contains("Levels updated: none", error);
        Assert.Contains("Attempted level: none", error);
        Assert.Contains("No changes have been rolled back", error);
    }
}
