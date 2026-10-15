using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

// Native getter failures cannot be induced reliably in Excel; inject them at this boundary.
[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "Chart")]
[Trait("RequiresExcel", "false")]
public sealed class ChartPropertyFailureTests
{
#pragma warning disable CA2201 // Inject exact native getter failures without starting Excel.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ChartDetectionDoesNotMisclassifyUnexpectedNativeFailure(bool pivot)
    {
        var failure = new COMException("Excel rejected the read.", unchecked((int)0x800A03EC));
        var chart = new FailingChart(failure);
        IChartStrategy strategy = pivot ? new PivotChartStrategy() : new RegularChartStrategy();

        Assert.Same(failure, Assert.Throws<COMException>(() => strategy.CanHandle(chart)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ChartDetectionToleratesOnlyKnownMissingMember(bool pivot)
    {
        var chart = new FailingChart(new COMException("No member.", unchecked((int)0x80020003)));
        IChartStrategy strategy = pivot ? new PivotChartStrategy() : new RegularChartStrategy();

        Assert.Equal(!pivot, strategy.CanHandle(chart));
    }

    public sealed class FailingChart(COMException failure)
    {
        public object PivotLayout => throw failure;
    }
#pragma warning restore CA2201
}
