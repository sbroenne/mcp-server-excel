using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Bridge = Sbroenne.ExcelMcp.McpServer.ServiceBridge.ServiceBridge;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
[Trait("RequiresExcel", "false")]
public sealed class BatchToolExecutionTests
{
    [Fact]
    public async Task Cancellation_StopsDispatchAndRetainsPendingWork()
    {
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var response = new TaskCompletionSource<ServiceResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        var backend = new Backend(_ =>
        {
            started.SetResult();
            return response.Task;
        });
        using var bridge = new Bridge(() => backend);
        using var cancellation = new CancellationTokenSource();
        var pending = ExcelBatchTool.ExcelBatch(BatchAction.Run, bridge,
            [new() { Command = "sheet.list" }, new() { Command = "sheet.list" }],
            "session-1", cancellationToken: cancellation.Token);
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => pending);
        Assert.Single(backend.Requests);
        Assert.True(backend.Requests[0].CallerCancellation.IsCancellationRequested);
        Assert.Equal(["session-1"], backend.Revealed);
        Assert.False(backend.Disposed);
        Assert.False(response.Task.IsCompleted);
        response.SetResult(new ServiceResponse { Success = true, Result = """{"success":true,"sheets":[]}""" });
    }

    [Fact]
    public async Task UnexpectedExecutionFailure_PropagatesInsteadOfBecomingInvalidInput()
    {
        var backend = new Backend(_ => throw new JsonException("Invalid Service response"));
        using var bridge = new Bridge(() => backend);
        var exception = await Assert.ThrowsAsync<JsonException>(() => ExcelBatchTool.ExcelBatch(
            BatchAction.Run, bridge, [new() { Command = "sheet.list" }, new() { Command = "sheet.list" }],
            "session-1"));
        Assert.Equal("Invalid Service response", exception.Message);
        Assert.Single(backend.Requests);
        Assert.False(backend.Disposed);
    }

    [Fact]
    public async Task AlreadyCancelled_DoesNotStartBackend()
    {
        var started = false;
        using var bridge = new Bridge(() =>
        {
            started = true;
            throw new InvalidOperationException("Should not start");
        });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelBatchTool.ExcelBatch(
            BatchAction.Run, bridge, [new() { Command = "sheet.list" }],
            cancellationToken: cancellation.Token));
        Assert.False(started);
    }

    private sealed class Backend(Func<ServiceRequest, Task<ServiceResponse>> process) : IServiceBridgeBackend
    {
        public List<ServiceRequest> Requests { get; } = [];
        public List<string> Revealed { get; } = [];
        public bool Disposed { get; private set; }

        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            Requests.Add(request);
            return process(request);
        }

        public void RevealSession(string sessionId) => Revealed.Add(sessionId);
        public void Dispose() => Disposed = true;
    }
}
