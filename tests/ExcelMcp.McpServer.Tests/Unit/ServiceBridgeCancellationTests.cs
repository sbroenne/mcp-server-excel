using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Bridge = Sbroenne.ExcelMcp.McpServer.ServiceBridge.ServiceBridge;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceBridge")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ServiceBridgeCancellationTests
{
    [Fact]
    public async Task SendAsync_WithSessionTimeout_ForceClosesSession()
    {
        var backend = new Backend();
        using var bridge = new Bridge(() => backend);
        var response = await bridge.SendAsync("sheet.list", "session-1", timeoutSeconds: 1);
        Assert.False(response.Success);
        Assert.Equal("Timeout", response.ErrorCategory);
        Assert.Equal(["session-1"], backend.ClosedSessions);
        Assert.False(backend.Disposed);
    }

    [Fact]
    public async Task ForwardToService_UsesExplicitCancellationToken()
    {
        var backend = new Backend();
        using var bridge = new Bridge(() => backend);
        using var cancellation = new CancellationTokenSource();
        var request = ExcelToolsBase.ForwardToServiceAsync(bridge, "sheet.list", "session-1", null, cancellation.Token);
        await backend.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => request);
        Assert.Equal(["session-1"], backend.ClosedSessions);
        Assert.False(backend.Disposed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FailedForcedClose_RetiresCapturedBackendAndAllowsRecovery(bool throws)
    {
        var first = new Backend { FailClose = true, ThrowOnClose = throws };
        var second = new Backend();
        second.Response.SetResult(new ServiceResponse { Success = true });
        var calls = 0;
        using var bridge = new Bridge(() => ++calls == 1 ? first : second);
        using var cancellation = new CancellationTokenSource();
        var request = bridge.SendAsync("sheet.list", "broken-session", cancellationToken: cancellation.Token);
        await first.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => request);
        Assert.True(first.Disposed);
        Assert.Equal(["broken-session"], first.ClosedSessions);
        Assert.True((await bridge.SendAsync("session.list")).Success);
        Assert.False(second.Disposed);
    }

    [Fact]
    public async Task LateFailedCleanup_DoesNotRetireReplacementBackend()
    {
        using var release = new ManualResetEventSlim();
        var closing = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var first = new Backend
        {
            CloseOverride = id =>
            {
                if (id == "old-session")
                {
                    closing.TrySetResult();
                    if (!release.Wait(TimeSpan.FromSeconds(10)))
                        throw new TimeoutException("Old cleanup was not released.");
                }
                return false;
            }
        };
        var replacement = new Backend();
        replacement.Response.SetResult(new ServiceResponse { Success = true });
        var factories = 0;
        using var bridge = new Bridge(() => ++factories == 1 ? first : replacement);
        using var oldCancellation = new CancellationTokenSource();
        using var otherCancellation = new CancellationTokenSource();
        var oldRequest = bridge.SendAsync("sheet.list", "old-session", cancellationToken: oldCancellation.Token);
        await first.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        var cancelling = Task.Run(oldCancellation.Cancel);
        try
        {
            await closing.Task.WaitAsync(TimeSpan.FromSeconds(5));
            var otherRequest = bridge.SendAsync("sheet.list", "other-session", cancellationToken: otherCancellation.Token);
            otherCancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => otherRequest);
            Assert.True(first.Disposed);
            Assert.True((await bridge.SendAsync("session.list")).Success);

            release.Set();
            await cancelling;
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => oldRequest);
            Assert.True((await bridge.SendAsync("session.list")).Success);
            Assert.Equal(2, factories);
            Assert.False(replacement.Disposed);
        }
        finally
        {
            release.Set();
            await cancelling;
        }
    }

    [Fact]
    public async Task CancellationRacingCompletion_DoesNotReportSuccess()
    {
        var backend = new Backend();
        using var bridge = new Bridge(() => backend);
        using var cancellation = new CancellationTokenSource();
        var request = bridge.SendAsync("sheet.list", "session-race", cancellationToken: cancellation.Token);
        await backend.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancellation.Cancel();
        backend.Response.TrySetResult(new ServiceResponse { Success = true });
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => request);
        Assert.Equal(["session-race"], backend.ClosedSessions);
    }

    [Fact]
    public async Task StartupFailure_ReportsCategoryWithoutLeakingPrivateDetails()
    {
        using var bridge = new Bridge(() => throw new FileNotFoundException("office runtime missing"));
        var response = await bridge.SendAsync("session.open");
        Assert.False(response.Success);
        Assert.Equal("ServiceStartup", response.ErrorCategory);
        Assert.Equal("FileNotFoundException", response.ExceptionType);
        Assert.DoesNotContain("office runtime missing", response.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Dispose_UnblocksActiveRequest()
    {
        var backend = new Backend();
        using var bridge = new Bridge(() => backend);
        var request = bridge.SendAsync("sheet.list");
        await backend.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        bridge.Dispose();
        var response = await request.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.True(backend.Disposed);
        Assert.False(response.Success);
        Assert.Equal("disposed", response.ErrorMessage);
    }

    [Fact]
    public async Task Dispose_DuringInitialization_DiscardsBackendWithoutRestart()
    {
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var backend = new Backend();
        var calls = 0;
        using var bridge = new Bridge(() =>
        {
            Interlocked.Increment(ref calls);
            entered.Set();
            if (!release.Wait(TimeSpan.FromSeconds(5)))
                throw new TimeoutException("Test factory was not released.");
            return backend;
        });
        var request = Task.Run(() => bridge.SendAsync("session.list"));
        try
        {
            Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
            bridge.Dispose();
        }
        finally
        {
            release.Set();
        }
        await Assert.ThrowsAsync<ObjectDisposedException>(() => request);
        Assert.True(backend.Disposed);
        Assert.Equal(1, calls);
    }

    [Fact]
    public async Task FailedDisposal_RemainsTerminalAndReportsFailure()
    {
        var backend = new Backend { ThrowOnDispose = true };
        backend.Response.SetResult(new ServiceResponse { Success = true });
        using var bridge = new Bridge(() => backend);
        Assert.True((await bridge.SendAsync("session.list")).Success);
        Assert.Throws<InvalidOperationException>(bridge.Dispose);
        await Assert.ThrowsAsync<ObjectDisposedException>(() => bridge.SendAsync("session.list"));
    }

    private sealed class Backend : IServiceBridgeBackend
    {
        internal TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource<ServiceResponse> Response { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal List<string> ClosedSessions { get; } = [];
        internal bool Disposed { get; private set; }
        internal bool FailClose { get; init; }
        internal bool ThrowOnClose { get; init; }
        internal bool ThrowOnDispose { get; init; }
        internal Func<string, bool>? CloseOverride { get; init; }

        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            Started.TrySetResult();
            return Response.Task;
        }

        public bool ForceCloseSession(string sessionId)
        {
            if (CloseOverride is not null)
                return CloseOverride(sessionId);
            ClosedSessions.Add(sessionId);
            if (ThrowOnClose)
                throw new InvalidOperationException("Synthetic close failure.");
            if (FailClose)
                return false;
            Response.TrySetResult(new ServiceResponse { Success = false, ErrorMessage = "closed" });
            return true;
        }

        public void Dispose()
        {
            Disposed = true;
            Response.TrySetResult(new ServiceResponse { Success = false, ErrorMessage = "disposed" });
            if (ThrowOnDispose)
                throw new InvalidOperationException("Synthetic disposal failure.");
        }
    }
}
