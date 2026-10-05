using System.IO.Pipelines;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Collection("ProgramTransport")]
[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "ProgramTransport")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ProgramTestTransportLifecycleTests(ITestOutputHelper output)
{
    [Theory]
    [InlineData("exception")]
    [InlineData("invalid-json")]
    [InlineData("missing-result")]
    public async Task UnexpectedBackendFailure_UsesSdkErrorWithoutPrivateDetails(string failureMode)
    {
        using var cancellation = new CancellationTokenSource();
        var input = new Pipe();
        var outputPipe = new Pipe();
        var backend = new ListBackend { FailureMode = failureMode };
        var host = await ProgramTransportTestHost.StartAsync(
            input, outputPipe, cancellation.Token, "ErrorPrivacyClient", () => backend);
        try
        {
            var result = await host.Client.CallToolAsync("file_read", new Dictionary<string, object?>
            {
                ["action"] = "list"
            });
            Assert.True(result.IsError);
            var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
            Assert.Equal("An error occurred invoking 'file_read'.", text);
            Assert.DoesNotContain("synthetic-private-detail", text, StringComparison.Ordinal);
        }
        finally
        {
            await ProgramTransportTestHost.StopAsync(host.Client, input, outputPipe, host.ServerTask, output, cancellation);
        }
    }

    [Fact]
    public async Task IndependentHosts_ShutdownDoesNotResetAnotherHost()
    {
        using var firstCancellation = new CancellationTokenSource();
        using var secondCancellation = new CancellationTokenSource();
        var firstInput = new Pipe();
        var firstOutput = new Pipe();
        var secondInput = new Pipe();
        var secondOutput = new Pipe();
        var firstBackend = new ListBackend();
        var secondBackend = new ListBackend();
        var first = await ProgramTransportTestHost.StartAsync(
            firstInput, firstOutput, firstCancellation.Token, "FirstHost", () => firstBackend);
        var second = await ProgramTransportTestHost.StartAsync(
            secondInput, secondOutput, secondCancellation.Token, "SecondHost", () => secondBackend);
        var firstStopped = false;
        try
        {
            var args = new Dictionary<string, object?> { ["action"] = "list" };
            Assert.False((await first.Client.CallToolAsync("file_read", args)).IsError is true);
            Assert.False((await second.Client.CallToolAsync("file_read", args)).IsError is true);
            await ProgramTransportTestHost.StopAsync(first.Client, firstInput, firstOutput, first.ServerTask, output, firstCancellation);
            firstStopped = true;
            Assert.True(firstBackend.Disposed);
            Assert.False(secondBackend.Disposed);
            Assert.False((await second.Client.CallToolAsync("file_read", args)).IsError is true);
        }
        finally
        {
            if (!firstStopped)
                await ProgramTransportTestHost.StopAsync(first.Client, firstInput, firstOutput, first.ServerTask, output, firstCancellation);
            await ProgramTransportTestHost.StopAsync(second.Client, secondInput, secondOutput, second.ServerTask, output, secondCancellation);
        }
        Assert.True(secondBackend.Disposed);
    }

    private sealed class ListBackend : IServiceBridgeBackend
    {
        internal bool Disposed { get; private set; }
        internal string? FailureMode { get; init; }
        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            ObjectDisposedException.ThrowIf(Disposed, this);
            if (FailureMode == "exception")
                throw new InvalidOperationException("synthetic-private-detail");
            if (FailureMode is "invalid-json" or "missing-result")
                return Task.FromResult(new ServiceResponse { Success = true, Result = FailureMode == "invalid-json" ? "synthetic-private-detail" : null });
            return Task.FromResult(new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[],"count":0}""" });
        }
        public bool ForceCloseSession(string sessionId) => throw new InvalidOperationException("Unexpected close.");
        public void Dispose() => Disposed = true;
    }
}
