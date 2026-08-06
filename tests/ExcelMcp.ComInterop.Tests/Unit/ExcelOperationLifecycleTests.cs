using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "ExcelBatch")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ExcelOperationLifecycleTests
{
    [Fact]
    public void InterruptWhileQueued_IsKnownNotStartedAndWorkerCannotStartLater()
    {
        var lifecycle = new ExcelOperationLifecycle();

        var outcome = lifecycle.Interrupt();

        Assert.Equal(ExcelOperationInterruption.NotStarted, outcome);
        Assert.Equal(ExcelOperationState.AbandonedBeforeStart, lifecycle.State);
        Assert.False(lifecycle.TryStart());
    }

    [Fact]
    public void InterruptAfterStart_IsUnknownAndPoisonsOutcome()
    {
        var lifecycle = new ExcelOperationLifecycle();
        Assert.True(lifecycle.TryStart());

        var outcome = lifecycle.Interrupt();

        Assert.Equal(ExcelOperationInterruption.StartedOutcomeUnknown, outcome);
        Assert.Equal(ExcelOperationState.OutcomeUnknown, lifecycle.State);
    }

    [Fact]
    public void CompletedBeforeInterrupt_IsKnownComplete()
    {
        var lifecycle = new ExcelOperationLifecycle();
        Assert.True(lifecycle.TryStart());
        lifecycle.MarkCompleted();

        var outcome = lifecycle.Interrupt();

        Assert.Equal(ExcelOperationInterruption.Completed, outcome);
        Assert.Equal(ExcelOperationState.Completed, lifecycle.State);
    }

    [Fact]
    public void LateWorkerCompletion_DoesNotEraseUnknownOutcome()
    {
        var lifecycle = new ExcelOperationLifecycle();
        Assert.True(lifecycle.TryStart());
        Assert.Equal(ExcelOperationInterruption.StartedOutcomeUnknown, lifecycle.Interrupt());

        lifecycle.MarkCompleted();

        Assert.Equal(ExcelOperationState.OutcomeUnknown, lifecycle.State);
    }
}
