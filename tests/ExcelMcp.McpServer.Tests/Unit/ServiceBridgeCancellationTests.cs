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
    public async Task SendAsync_WithSessionTimeout_PreservesSession()
    {
        var backend = new Backend();
        using var bridge = new Bridge(() => backend);
        var response = await bridge.SendAsync("sheet.list", "session-1", timeoutSeconds: 1);
        Assert.False(response.Success);
        Assert.Equal("Timeout", response.ErrorCategory);
        Assert.Equal(["session-1"], backend.RevealedSessions);
        Assert.False(backend.Disposed);
        Assert.True(Assert.IsType<ServiceRequest>(backend.Request).CallerCancellation.IsCancellationRequested);
        Assert.False(backend.Response.Task.IsCompleted);
    }

    [Fact]
    public async Task SendAsync_QueuedDeadline_PreservesCancelledPartialStateAsTimeout()
    {
        var backend = new Backend { Response = new TaskCompletionSource<ServiceResponse>() };
        using var bridge = new Bridge(() => backend);
        var pending = bridge.SendAsync("powerquery.create", "session-1", timeoutSeconds: 2);
        await backend.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        // Register last and complete inline so the Service response wins the deadline race.
        using var queuedCancellation = Assert.IsType<ServiceRequest>(backend.Request).CallerCancellation.Register(() =>
            backend.Response.TrySetResult(new ServiceResponse
            {
                Success = false,
                Command = "powerquery.create",
                SessionId = "session-1",
                ErrorCategory = "Cancelled",
                ErrorMessage = "Queued operation was not executed.",
                InnerError = "Earlier edits remain.",
                PartialState = """{"queryCreated":true}"""
            }));
        var response = await pending;
        Assert.False(response.Success);
        Assert.Equal("Timeout", response.ErrorCategory);
        Assert.Equal(nameof(TimeoutException), response.ExceptionType);
        Assert.Equal("session-1", response.SessionId);
        Assert.Equal("""{"queryCreated":true}""", response.PartialState);
        Assert.Contains("Queued operation was not executed.", response.InnerError, StringComparison.Ordinal);
        Assert.Contains("Earlier edits remain.", response.InnerError, StringComparison.Ordinal);
        Assert.Equal(["session-1"], backend.RevealedSessions);
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
        Assert.Equal(["session-1"], backend.RevealedSessions);
        Assert.False(backend.Disposed);
        Assert.True(Assert.IsType<ServiceRequest>(backend.Request).CallerCancellation.IsCancellationRequested);
        Assert.False(backend.Response.Task.IsCompleted);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RevealFailure_PreservesBackendAndAllowsOtherRequests(bool throws)
    {
        var first = new Backend { ThrowOnReveal = throws };
        var calls = 0;
        using var bridge = new Bridge(() => { calls++; return first; });
        using var cancellation = new CancellationTokenSource();
        var request = bridge.SendAsync("sheet.list", "broken-session", cancellationToken: cancellation.Token);
        await first.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => request);
        Assert.False(first.Disposed);
        Assert.Equal(["broken-session"], first.RevealedSessions);
        first.Response.TrySetResult(new ServiceResponse { Success = true });
        Assert.True((await bridge.SendAsync("session.list")).Success);
        Assert.Equal(1, calls);
        Assert.False(first.Disposed);
    }

    [Fact]
    public async Task LateCompletion_DoesNotDisposeOrReplaceBackend()
    {
        var first = new Backend();
        var factories = 0;
        using var bridge = new Bridge(() => { factories++; return first; });
        using var oldCancellation = new CancellationTokenSource();
        var oldRequest = bridge.SendAsync("sheet.list", "old-session", cancellationToken: oldCancellation.Token);
        await first.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        oldCancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => oldRequest);
        first.Response.TrySetResult(new ServiceResponse { Success = true });
        Assert.True((await bridge.SendAsync("session.list")).Success);
        Assert.Equal(1, factories);
        Assert.False(first.Disposed);
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
        Assert.Equal(["session-race"], backend.RevealedSessions);
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
        backend.ThrowOnDispose = false;
        bridge.Dispose();
        Assert.True(backend.Disposed);
    }

    private sealed class Backend : IServiceBridgeBackend
    {
        internal TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource<ServiceResponse> Response { get; init; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal List<string> RevealedSessions { get; } = [];
        internal bool Disposed { get; private set; }
        internal bool ThrowOnReveal { get; init; }
        internal bool ThrowOnDispose { get; set; }
        internal ServiceRequest? Request { get; private set; }

        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            Request = request;
            Started.TrySetResult();
            return Response.Task;
        }

        public void RevealSession(string sessionId)
        {
            RevealedSessions.Add(sessionId);
            if (ThrowOnReveal)
                throw new InvalidOperationException("Synthetic reveal failure.");
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
