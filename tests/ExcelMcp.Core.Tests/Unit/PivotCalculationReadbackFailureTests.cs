using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "PivotCalculation")]
[Trait("Layer", "Core")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PivotCalculationReadbackFailureTests
{
    [Fact]
    public void OlapMismatch_DoesNotClaimWriteWasIgnoredOrStateWasPreserved()
    {
        var error = PivotTableCommands.DescribeCalculationReadbackFailure(
            "Revenue", PivotFieldCalculation.PercentOfTotal, PivotFieldCalculation.Normal, true);

        Assert.Contains("Excel accepted the Show Values As write for 'Revenue'", error.Message);
        Assert.Contains("native readback did not verify", error.Message);
        Assert.Contains("Requested calculation: PercentOfTotal; native calculation: Normal", error.Message);
        Assert.Contains("Displayed values may already have changed", error.Message);
        Assert.Contains("list-fields reports the previous calculation", error.Message);
        Assert.Contains("Inspect get-data", error.Message);
        Assert.Contains("no changes have been rolled back", error.Message);
        Assert.DoesNotContain("did not apply", error.Message);
    }
}
