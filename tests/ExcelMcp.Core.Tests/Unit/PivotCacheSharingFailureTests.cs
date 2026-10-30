using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "PivotTables")]
[Trait("Layer", "Core")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PivotCacheSharingFailureTests
{
    [Theory]
    [InlineData(true, 1)]
    [InlineData(false, 2)]
    public void Failure_ReportsObservedStateWithoutClaimingRollback(bool changed, int actualIndex)
    {
        var failure = Marshal.GetExceptionForHR(unchecked((int)0x800A03EC))!;
        var error = PivotTableCommands.DescribeCacheSharingFailure(failure, "reading the actual cache", () => new()
        {
            SheetName = "Reports",
            PivotTableName = "Moved",
            PreviousCacheIndex = 2,
            CacheIndex = actualIndex,
            TargetCacheIndex = 1,
            Changed = changed,
            SharedPivotTables = [new() { SheetName = "Reports", PivotTableName = "Moved" }]
        });

        Assert.Same(failure, error.InnerException);
        Assert.Contains("reading the actual cache", error.Message, StringComparison.Ordinal);
        Assert.Contains(failure.Message, error.Message, StringComparison.Ordinal);
        Assert.Contains("No changes were undone", error.Message, StringComparison.Ordinal);
        Assert.Contains($"\"changed\":{changed.ToString().ToLowerInvariant()}", error.Message, StringComparison.Ordinal);
        Assert.Contains($"\"cacheIndex\":{actualIndex}", error.Message, StringComparison.Ordinal);
        Assert.Contains("\"success\":false", error.Message, StringComparison.Ordinal);
        Assert.Contains("\"sheetName\":\"Reports\"", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FailedInspection_PreservesBothCausesAndReportsUnknownState()
    {
        var failure = new InvalidOperationException("write failed");
        var readFailure = new InvalidOperationException("Excel disconnected");

        var error = PivotTableCommands.DescribeCacheSharingFailure(failure, "changing cache", () => throw readFailure);

        Assert.Equal(new Exception[] { failure, readFailure }, Assert.IsType<AggregateException>(error.InnerException).InnerExceptions);
        Assert.Contains("Actual cache state could not be read: Excel disconnected", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("\"cacheIndex\"", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CancelledInspection_PropagatesCancellation()
    {
        var cancellation = new OperationCanceledException();
        Assert.Same(cancellation, Assert.Throws<OperationCanceledException>(() =>
            PivotTableCommands.DescribeCacheSharingFailure(new InvalidOperationException(), "changing cache", () => throw cancellation)));
    }
}
