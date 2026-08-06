using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class OwnedProcessIdentityTests
{
    [Fact]
    public void Matches_RequiresPidStartTimeNameAndExecutablePath()
    {
        var identity = new OwnedProcessIdentity(42, 100, "EXCEL", @"C:\Program Files\Microsoft Office\EXCEL.EXE");
        var exact = new ProcessIdentitySnapshot(42, 100, "EXCEL", @"C:\Program Files\Microsoft Office\EXCEL.EXE");

        Assert.True(OwnedProcessIdentityGuard.Matches(identity, exact));
        Assert.False(OwnedProcessIdentityGuard.Matches(identity, exact with { StartedAtUtcFileTime = 101 }));
        Assert.False(OwnedProcessIdentityGuard.Matches(identity, exact with { ProcessName = "notepad" }));
        Assert.False(OwnedProcessIdentityGuard.Matches(identity, exact with { ExecutablePath = @"C:\Other\EXCEL.EXE" }));
    }

    [Fact]
    public void TryKill_WhenStartTimeDoesNotMatch_FailsClosedAndLeavesProcessAlive()
    {
        using var current = Process.GetCurrentProcess();
        Assert.True(OwnedProcessIdentityGuard.TryCapture(current.Id, out var captured));
        var staleIdentity = captured with { StartedAtUtcFileTime = captured.StartedAtUtcFileTime - 1 };

        var killed = OwnedProcessIdentityGuard.TryKill(staleIdentity);

        Assert.False(killed);
        Assert.False(current.HasExited);
    }

    [Fact]
    public void IsAlive_WhenStartTimeDoesNotMatch_ReturnsFalse()
    {
        using var current = Process.GetCurrentProcess();
        Assert.True(OwnedProcessIdentityGuard.TryCapture(current.Id, out var captured));
        var reusedPid = captured with { StartedAtUtcFileTime = captured.StartedAtUtcFileTime + 1 };

        Assert.False(OwnedProcessIdentityGuard.IsAlive(reusedPid));
    }

    [Fact]
    public void CompatibilityPidTrackingOverloads_FailClosedForNonExcelProcess()
    {
        using var current = Process.GetCurrentProcess();

        // These public compatibility methods must remain callable, but must not
        // register or control an arbitrary non-Excel PID.
        SessionManager.TrackExcelProcess(current.Id);
        SessionManager.UntrackExcelProcess(current.Id);

        Assert.False(current.HasExited);
    }
}
