using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "ChartDepth")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ChartTextValidationTests
{
    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void NonFiniteSize_IsRejectedBeforeExcelAccess(double size)
    {
        var commands = new ChartCommands();
        var error = Assert.Throws<ArgumentOutOfRangeException>(() =>
            commands.SetTextFormat(null!, "Sales", ChartTextElement.Title, fontSize: size));
        Assert.Equal("fontSize", error.ParamName);
    }

    [Fact]
    public void UnknownElement_IsRejectedBeforeExcelAccess()
    {
        var commands = new ChartCommands();
        var error = Assert.Throws<ArgumentOutOfRangeException>(() =>
            commands.GetTextFormat(null!, "Sales", (ChartTextElement)999));
        Assert.Equal("element", error.ParamName);
    }

    [Fact]
    public void UnknownAxis_IsRejectedBeforeExcelAccess()
    {
        var commands = new ChartCommands();
        var error = Assert.Throws<ArgumentOutOfRangeException>(() =>
            commands.GetTextFormat(null!, "Sales", ChartTextElement.AxisTitle, (ChartAxisType)999));
        Assert.Equal("axis", error.ParamName);
    }
}
