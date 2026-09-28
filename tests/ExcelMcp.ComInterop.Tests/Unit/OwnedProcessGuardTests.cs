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

    [Theory]
    [InlineData(0, 5)]
    [InlineData(2, 3)]
    [InlineData(6, 0)]
    public async Task NormalShutdown_RequestTimeReducesFinalWaitWithinOriginalBudget(
        int requestSeconds, int expectedFinalWaitSeconds)
    {
        var clock = new ControlledClock();
        var waits = new List<TimeSpan>();
        Assert.Equal(TimeSpan.FromSeconds(10), ProcessTerminationPolicy.NormalGraceTimeout);
        Assert.Equal(TimeSpan.FromSeconds(5), ProcessTerminationPolicy.NormalForcedExitTimeout);
        Assert.Equal(TimeSpan.FromSeconds(15), ProcessTerminationPolicy.NormalShutdownBudget);
        Assert.Equal(ProcessTerminationPolicy.NormalShutdownBudget,
            ProcessTerminationPolicy.NormalGraceTimeout + ProcessTerminationPolicy.NormalForcedExitTimeout);

        var result = await ProcessTerminationPolicy.TryCompleteAsync(
            ProcessTerminationPolicy.NormalGraceTimeout,
            ProcessTerminationPolicy.NormalForcedExitTimeout,
            (timeout, _) =>
            {
                waits.Add(timeout);
                clock.Advance(timeout);
                return Task.FromResult(ProcessTerminationPolicy.ProcessWaitOutcome.TimedOut);
            },
            () =>
            {
                clock.Advance(TimeSpan.FromSeconds(requestSeconds));
                return ProcessTerminationPolicy.ProcessTerminationOutcome.Requested;
            },
            CancellationToken.None,
            _ => { },
            overallTimeout: ProcessTerminationPolicy.NormalShutdownBudget,
            timeProvider: clock);

        Assert.False(result);
        Assert.Equal(TimeSpan.FromSeconds(10), waits[0]);
        Assert.Equal(TimeSpan.FromSeconds(expectedFinalWaitSeconds), waits[1]);
        Assert.Equal(TimeSpan.FromSeconds(Math.Max(15, 10 + requestSeconds)), clock.Elapsed);
    }

    private sealed class ControlledClock : TimeProvider
    {
        internal TimeSpan Elapsed { get; private set; }
        public override long TimestampFrequency => TimeSpan.TicksPerSecond;
        public override long GetTimestamp() => Elapsed.Ticks;
        internal void Advance(TimeSpan duration) => Elapsed += duration;
    }
}
