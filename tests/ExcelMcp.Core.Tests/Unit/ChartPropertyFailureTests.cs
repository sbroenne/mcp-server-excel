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
    [Fact]
    public void ChartDetectionDoesNotMisclassifyUnexpectedNativeFailure()
    {
        var failure = new COMException("Excel rejected the read.", unchecked((int)0x800A03EC));

        Assert.Same(failure, Assert.Throws<COMException>(() =>
            PivotChartStrategy.HasPivotLayout(() => throw failure)));
    }

    [Fact]
    public void ChartDetectionToleratesOnlyKnownMissingMember()
    {
        var failure = new COMException("No member.", unchecked((int)0x80020003));

        Assert.False(PivotChartStrategy.HasPivotLayout(() => throw failure));
    }

    [Fact]
    public void ChartDetectionWithoutPivotLayoutReturnsFalse()
    {
        Assert.False(PivotChartStrategy.HasPivotLayout(() => null));
    }
#pragma warning restore CA2201
}
