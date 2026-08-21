using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class OwnedProcessGuardTests
{
    [Theory]
    [InlineData(0, true)]
    [InlineData(1, false)]
    [InlineData(2, true)]
    public void IsAlive_ProbeResult_FailsOpenUnlessExitIsConfirmed(
        int probeValue,
        bool expected)
    {
        var probe = (OwnedProcessGuard.ProcessIdentityProbe)probeValue;
        Assert.Equal(expected, OwnedProcessGuard.IsAlive(probe));
    }
}
