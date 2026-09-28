using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
public sealed class PersistentServiceCleanupFailureTests
{
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
