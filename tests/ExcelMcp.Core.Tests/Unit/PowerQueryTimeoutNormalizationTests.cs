using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "false")]
public sealed class PowerQueryTimeoutNormalizationTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void NormalizeRefreshTimeout_NonPositiveValue_UsesDataOperationDefault(int seconds)
    {
        Assert.Equal(
            ComInteropConstants.DataOperationTimeout,
            PowerQueryCommands.NormalizeRefreshTimeout(TimeSpan.FromSeconds(seconds)));
    }

    [Fact]
    public void NormalizeRefreshTimeout_AtCancellationTokenLimit_PreservesValue()
    {
        TimeSpan maximum = TimeSpan.FromMilliseconds(uint.MaxValue - 1);

        Assert.Equal(maximum, PowerQueryCommands.NormalizeRefreshTimeout(maximum));
    }

    [Theory]
    [InlineData(4294967295d)]
    [InlineData(155520000000d)]
    public void NormalizeRefreshTimeout_AboveCancellationTokenLimit_Clamps(double milliseconds)
    {
        Assert.Equal(
            TimeSpan.FromMilliseconds(uint.MaxValue - 1),
            PowerQueryCommands.NormalizeRefreshTimeout(TimeSpan.FromMilliseconds(milliseconds)));
    }
}
