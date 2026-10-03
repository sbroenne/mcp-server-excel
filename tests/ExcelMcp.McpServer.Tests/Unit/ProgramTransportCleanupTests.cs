using System.IO.Pipelines;
using Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
public sealed class ProgramTransportCleanupTests(ITestOutputHelper output)
{
    [Fact]
    public async Task StopAsync_FaultedServer_ReportsOriginalFailure()
    {
        var failure = new InvalidOperationException("unexpected test host fault");
        var rejected = await Assert.ThrowsAsync<AggregateException>(() =>
            ProgramTransportTestHost.StopAsync(
                null, new Pipe(), new Pipe(), Task.FromException(failure), output));

        Assert.Same(failure, Assert.Single(rejected.InnerExceptions));
    }

    [Fact]
    public async Task StopAsync_CancelledServer_AcceptsExpectedShutdownCancellation()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await ProgramTransportTestHost.StopAsync(
            null, new Pipe(), new Pipe(), Task.FromCanceled(cancellation.Token), output);
    }
}
