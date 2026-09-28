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
            _commands.SetMode(null!, (CalculationMode)int.MaxValue));

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
}
