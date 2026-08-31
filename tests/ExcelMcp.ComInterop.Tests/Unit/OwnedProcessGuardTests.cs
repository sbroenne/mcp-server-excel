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

    [Fact]
    public async Task TerminationUnavailable_ProcessExitsDuringFinalObservation_Succeeds()
    {
        var waits = new Queue<ProcessTerminationPolicy.ProcessWaitOutcome>(
        [
            ProcessTerminationPolicy.ProcessWaitOutcome.TimedOut,
            ProcessTerminationPolicy.ProcessWaitOutcome.Exited
        ]);
        var terminated = false;

        var result = await ProcessTerminationPolicy.TryCompleteAsync(
            TimeSpan.FromSeconds(5),
            TimeSpan.FromSeconds(3),
            (_, _) => Task.FromResult(waits.Dequeue()),
            () => ProcessTerminationPolicy.ProcessTerminationOutcome.Unavailable,
            CancellationToken.None,
            value => terminated = value);

        Assert.True(result);
        Assert.False(terminated);
        Assert.Empty(waits);
    }

    [Fact]
    public async Task TerminationUnavailable_ProcessRemainsLive_Fails()
    {
        var waits = new Queue<ProcessTerminationPolicy.ProcessWaitOutcome>(
        [
            ProcessTerminationPolicy.ProcessWaitOutcome.TimedOut,
            ProcessTerminationPolicy.ProcessWaitOutcome.TimedOut
        ]);

        var result = await ProcessTerminationPolicy.TryCompleteAsync(
            TimeSpan.FromSeconds(5),
            TimeSpan.FromSeconds(3),
            (_, _) => Task.FromResult(waits.Dequeue()),
            () => ProcessTerminationPolicy.ProcessTerminationOutcome.Unavailable,
            CancellationToken.None,
            _ => { });

        Assert.False(result);
        Assert.Empty(waits);
    }

    [Fact]
    public void ProcessExitTimeout_MatchesPipeAndSessionTeardownBudget()
    {
        Assert.Equal(
            TimeSpan.FromSeconds(10),
            ProcessTerminationPolicy.ProcessExitTimeout);
    }
}
