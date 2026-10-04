using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
public sealed class PersistentServiceCleanupFailureTests
{
    [Fact]
    public void Run_PreservesPrimaryAndEveryCleanupFailureAndContinuesCleanup()
    {
        var primary = new InvalidOperationException("test");
        var first = new IOException("first cleanup");
        var second = new UnauthorizedAccessException("second cleanup");
        bool completed = false;

        var error = Assert.Throws<AggregateException>(() => PersistentServiceCleanupFailures.Run(
            () => throw primary,
            () => throw first,
            () => throw second,
            () => completed = true));

        var failures = error.Flatten().InnerExceptions;
        Assert.Equal(3, failures.Count);
        Assert.Contains(primary, failures);
        Assert.Contains(first, failures);
        Assert.Contains(second, failures);
        Assert.True(completed);
    }

    [Fact]
    public void Run_CleanupFailureCannotTurnPassingTestIntoSuccess()
    {
        var failure = new IOException("cleanup");
        var error = Assert.Throws<IOException>(() => PersistentServiceCleanupFailures.Run(
            () => { },
            () => throw failure));
        Assert.Same(failure, error);
    }

    [Fact]
    public void Run_PassingTestAndCleanupCompleteNormally()
    {
        var order = new List<string>();
        PersistentServiceCleanupFailures.Run(
            () => order.Add("test"),
            () => order.Add("first"),
            () => order.Add("second"));
        Assert.Equal(["test", "first", "second"], order);
    }

    [Fact]
    public void Combine_PreservesEveryCleanupFailure()
    {
        var closeFailure = new InvalidOperationException("close");
        var disposeFailure = new IOException("dispose");
        var directoryFailure = new UnauthorizedAccessException("directory");

        Exception? combined = null;
        combined = PersistentServiceCleanupFailures.Combine(
            combined,
            closeFailure);
        combined = PersistentServiceCleanupFailures.Combine(
            combined,
            disposeFailure);
        combined = PersistentServiceCleanupFailures.Combine(
            combined,
            directoryFailure);

        var aggregate = Assert.IsType<AggregateException>(combined);
        var failures = aggregate.Flatten().InnerExceptions;
        Assert.Equal(3, failures.Count);
        Assert.Contains(closeFailure, failures);
        Assert.Contains(disposeFailure, failures);
        Assert.Contains(directoryFailure, failures);
    }
}
