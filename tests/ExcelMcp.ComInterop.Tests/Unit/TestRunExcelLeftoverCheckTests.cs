using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Tests.Shared;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class TestRunExcelLeftoverCheckTests
{
    [Fact]
    public void FindRemaining_RecordedProcessExited_IsNotReported()
    {
        using var standIn = StartStandIn(seconds: 2);
        using var check = new TestRunExcelLeftoverCheck(_ => true);
        check.Record(IdentityOf(standIn));

        Assert.False(standIn.HasExited);
        Assert.Empty(check.FindRemaining(TimeSpan.FromSeconds(15)));
    }

    [Fact]
    public void FindRemaining_RecordedProcessRunning_IsReportedAndKeptRunningAtHostExit()
    {
        // Ends by itself after the test; the check must never stop it.
        using var standIn = StartStandIn(seconds: 30);
        var identity = Assert.IsType<ExcelProcessIdentity>(SessionManager.TrackExcelProcessIdentity(standIn.Id));
        try
        {
            using var check = new TestRunExcelLeftoverCheck(_ => true);
            check.Record(identity);

            Assert.Equal([identity], check.FindRemaining(TimeSpan.FromSeconds(1)));
            Assert.True(SessionManager.IsExcelProcessPreserved(identity));
            Assert.Contains(identity, SessionManager.GetTrackedExcelProcesses());
            Assert.False(standIn.HasExited);
        }
        finally
        {
            SessionManager.UntrackExcelProcess(identity);
        }
    }

    [Fact]
    public void Record_StartTimeDoesNotMatch_IsNotReported()
    {
        using var standIn = StartStandIn(seconds: 30);
        using var check = new TestRunExcelLeftoverCheck(_ => true);
        var identity = IdentityOf(standIn);
        check.Record(identity with { StartedAtUtcFileTime = identity.StartedAtUtcFileTime - 1 });

        Assert.Empty(check.FindRemaining(TimeSpan.Zero));
    }

    [Fact]
    public void Record_NotExcel_IsNotReported()
    {
        using var standIn = StartStandIn(seconds: 30);
        using var check = new TestRunExcelLeftoverCheck();
        check.Record(IdentityOf(standIn));

        Assert.Empty(check.FindRemaining(TimeSpan.Zero));
    }

    private static Process StartStandIn(int seconds) =>
        Process.Start(new ProcessStartInfo("pwsh")
        {
            ArgumentList = { "-NoProfile", "-Command", $"Start-Sleep -Seconds {seconds}" },
            UseShellExecute = false,
            CreateNoWindow = true
        })!;

    private static ExcelProcessIdentity IdentityOf(Process process) =>
        new(process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc());
}
