using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Workbook")]
[Trait("RequiresExcel", "false")]
public sealed class OperationCleanupTests
{
    [Fact]
    public void SuccessfulCleanupDoesNotChangePrimaryFailure()
    {
        var calls = 0;
        OperationCleanup.Run(() => calls++, new InvalidOperationException("Primary failure."));
        Assert.Equal(1, calls);
    }

    [Fact]
    public void CleanupFailureWithoutPrimaryPropagatesUnchanged()
    {
        var cleanup = new InvalidOperationException("Restoring state failed.");
        var failure = Assert.Throws<InvalidOperationException>(() => OperationCleanup.Run(() => throw cleanup, null));
        Assert.Same(cleanup, failure);
    }

    [Fact]
    public void PrimaryAndCleanupFailuresAreBothRetained()
    {
        var primary = new InvalidOperationException("Operation failed.");
        var cleanup = new InvalidOperationException("Restoring state failed.");
        var failures = Assert.Throws<AggregateException>(() => OperationCleanup.Run(() => throw cleanup, primary));
        Assert.Equal([primary, cleanup], failures.InnerExceptions);
    }

    [Fact]
    public void CancellationMessageTokenAndCleanupFailureArePreserved()
    {
        var token = new CancellationToken(canceled: true);
        var primary = new OperationCanceledException("Canceled operation.", token);
        var cleanup = new InvalidOperationException("Restoring state failed.");
        var canceled = Assert.Throws<OperationCanceledException>(() => OperationCleanup.Run(() => throw cleanup, primary));
        Assert.Equal(primary.Message, canceled.Message);
        Assert.Equal(token, canceled.CancellationToken);
        var failures = Assert.IsType<AggregateException>(canceled.InnerException);
        Assert.Equal([primary, cleanup], failures.InnerExceptions);
    }

}
