using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "ServiceBridge")]
[Trait("RequiresExcel", "false")]
public sealed class BridgeOwnershipRegressionTests
{
    [Fact]
    public async Task Dispose_PreventsStartingAnotherBackend()
    {
        var backend = new PendingCreationBackend();
        backend.Response.SetResult(new ServiceResponse { Success = true });
        using var bridge = new ServiceBridge.ServiceBridge(() => backend);
        bridge.Dispose();
        await Assert.ThrowsAsync<ObjectDisposedException>(() =>
            bridge.SendAsync("session.list", null, null, null, CancellationToken.None));
    }

    [Fact]
    public async Task CancelledCreation_PreservesBackendAndClosesLateSession()
    {
        var backend = new PendingCreationBackend();
        using var bridge = new ServiceBridge.ServiceBridge(() => backend);
        using var cancellation = new CancellationTokenSource();
        var request = bridge.SendAsync("session.create", null, null, null, cancellation.Token);
        await backend.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            request.WaitAsync(TimeSpan.FromSeconds(5)));

        try
        {
            Assert.False(backend.Disposed);
            backend.Response.SetResult(new ServiceResponse
            {
                Success = true,
                Result = """{"success":true,"sessionId":"late-session"}"""
            });
            Assert.Equal("late-session", await backend.Closed.Task.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.False(backend.Disposed);
        }
        finally
        {
            backend.Response.TrySetResult(new ServiceResponse { Success = false, ErrorMessage = "test cleanup" });
        }
    }

    private sealed class PendingCreationBackend : IServiceBridgeBackend
    {
        internal TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource<ServiceResponse> Response { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource<string> Closed { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal bool Disposed { get; private set; }

        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            Started.TrySetResult();
            return Response.Task;
        }

        public bool ForceCloseSession(string sessionId)
        {
            Closed.TrySetResult(sessionId);
            return true;
        }

        public void Dispose() => Disposed = true;
    }
}
