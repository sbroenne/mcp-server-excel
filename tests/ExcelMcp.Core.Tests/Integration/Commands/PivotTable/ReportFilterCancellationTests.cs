using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.PivotTable;

// Service cannot inject a cancelled token into an already-running STA callback.
[Collection("Sequential")]
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PivotTables")]
[Trait("Speed", "Medium")]
public sealed class ReportFilterCancellationTests(TempDirectoryFixture fixture) :
    IClassFixture<TempDirectoryFixture>
{
    [Fact]
    public void SetReportFilter_CancelledCallback_PropagatesBeforeWorkbookAccess()
    {
        using var innerBatch = ExcelSession.BeginBatch(fixture.CreateTestFile());
        var savedBefore = innerBatch.Execute((context, _) => context.Book.Saved);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var batch = new InjectedCancellationBatch(innerBatch, cancellation.Token);
        var commands = new PivotTableCommands();

        var error = Assert.ThrowsAny<OperationCanceledException>(() =>
            commands.SetReportFilter(batch, "MissingSheet", "MissingPivot",
                "[Sales].[Region]", []));

        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.False(innerBatch.HasTimedOutOperation);
        Assert.Equal(savedBefore, innerBatch.Execute((context, _) => context.Book.Saved));
    }
}
