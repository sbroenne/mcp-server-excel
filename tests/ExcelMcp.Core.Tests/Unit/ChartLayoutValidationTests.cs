using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "ChartDepth")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ChartLayoutValidationTests
{
    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void NonFiniteTitlePosition_IsRejectedBeforeExcelAccess(double position)
    {
        var commands = new ChartCommands();
        Assert.Equal("left", Assert.Throws<ArgumentOutOfRangeException>(() =>
            commands.SetTitle(null!, "Sales", "Revenue", left: position)).ParamName);
        Assert.Equal("top", Assert.Throws<ArgumentOutOfRangeException>(() =>
            commands.SetTitle(null!, "Sales", "Revenue", top: position)).ParamName);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    [InlineData(double.MaxValue)]
    [InlineData(0)]
    [InlineData(-1)]
    public void InvalidLineWeight_IsRejectedBeforeExcelAccess(double weight)
    {
        var commands = new ChartCommands();
        Assert.Equal("lineWeight", Assert.Throws<ArgumentOutOfRangeException>(() =>
            commands.SetAxisFormat(null!, "Sales", ChartAxisType.Value, lineWeight: weight)).ParamName);
        Assert.Equal("lineWeight", Assert.Throws<ArgumentOutOfRangeException>(() =>
            commands.SetAreaFormat(null!, "Sales", ChartAreaTarget.Chart, lineWeight: weight)).ParamName);
    }
}
