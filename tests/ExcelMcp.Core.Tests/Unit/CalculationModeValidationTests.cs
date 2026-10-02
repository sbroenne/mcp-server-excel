using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "CalculationMode")]
[Trait("RequiresExcel", "false")]
public sealed class CalculationModeValidationTests
{
    private readonly CalculationModeCommands _commands = new();

    [Fact]
    public void SetMode_UnknownMode_ThrowsBeforeBatchExecution()
    {
        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            _commands.SetSettings(null!, (CalculationMode)int.MaxValue));

        Assert.Equal("mode", exception.ParamName);
        Assert.Contains("Unknown calculation mode", exception.Message);
    }

    [Fact]
    public void Calculate_UnknownScope_ThrowsBeforeBatchExecution()
    {
        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            _commands.Calculate(null!, (CalculationScope)int.MaxValue));

        Assert.Equal("scope", exception.ParamName);
        Assert.Contains("Unknown calculation scope", exception.Message);
    }

    [Theory]
    [InlineData(CalculationScope.Sheet, null)]
    [InlineData(CalculationScope.Sheet, "")]
    [InlineData(CalculationScope.Sheet, " \t ")]
    [InlineData(CalculationScope.Range, null)]
    [InlineData(CalculationScope.Range, "")]
    [InlineData(CalculationScope.Range, " \t ")]
    public void Calculate_MissingSheetName_ThrowsBeforeBatchExecution(
        CalculationScope scope, string? sheetName)
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.Calculate(null!, scope, sheetName,
                scope == CalculationScope.Range ? "A1" : null));

        Assert.Equal("sheetName", exception.ParamName);
        Assert.Contains("required", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" \t ")]
    public void Calculate_MissingRangeAddress_ThrowsBeforeBatchExecution(string? rangeAddress)
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.Calculate(null!, CalculationScope.Range, "Sheet1", rangeAddress));

        Assert.Equal("rangeAddress", exception.ParamName);
        Assert.Contains("required", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
