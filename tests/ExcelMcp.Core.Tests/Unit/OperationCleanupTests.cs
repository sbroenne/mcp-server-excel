using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands;
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

    [Theory]
    [InlineData(unchecked((int)0x80010108))]
    [InlineData(unchecked((int)0x800706BE))]
    [InlineData(unchecked((int)0x800706BA))]
    public void MutationFailureAndDisconnectedCleanup_CannotBeReportedAsRecoverable(int hresult)
    {
        var primary = new InvalidOperationException("Filter mutation failed.");
#pragma warning disable CA2201 // Synthetic native restoration error for the reporting guard.
        var cleanup = new COMException("Restoring ManualUpdate failed.", hresult);
#pragma warning restore CA2201
        var failure = Assert.Throws<AggregateException>(() => OperationCleanup.Run(() => throw cleanup, primary));

        Assert.Equal([primary, cleanup], failure.InnerExceptions);
        Assert.False(CreatedObjectFailure.CanReport(failure));
        Assert.False(CreatedObjectFailure.CanReport(new InvalidOperationException("Context", failure)));
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
