using Microsoft.Extensions.Logging;
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
    public async Task Dispose_FailureRetainsBackendForRetry()
    {
        var backend = new PendingCreationBackend();
        backend.Response.SetResult(new ServiceResponse { Success = true });
        var bridge = new ServiceBridge.ServiceBridge(() => backend);
        await bridge.SendAsync("session.list");
        backend.DisposeFailure = new InvalidOperationException("Excel work is still pending.");

        var error = Assert.Throws<InvalidOperationException>(bridge.Dispose);
        Assert.Equal("Excel work is still pending.", error.Message);
        Assert.Equal(1, backend.DisposeCalls);
        Assert.False(backend.Disposed);
        await Assert.ThrowsAsync<ObjectDisposedException>(() => bridge.SendAsync("session.list"));

        backend.DisposeFailure = null;
        bridge.Dispose();
        Assert.Equal(2, backend.DisposeCalls);
        Assert.True(backend.Disposed);
        bridge.Dispose();
        Assert.Equal(2, backend.DisposeCalls);
    }

    [Fact]
    public async Task Dispose_LogsShutdownReportBeforeDisposingBackend()
    {
        var backend = new PendingCreationBackend { ShutdownReport = "Excel process 42 was left as it is." };
        backend.Response.SetResult(new ServiceResponse { Success = true });
        var logger = new CapturingLogger();
        var bridge = new ServiceBridge.ServiceBridge(() => backend, logger);
        await bridge.SendAsync("session.list");

        bridge.Dispose();

        var entry = Assert.Single(logger.Entries);
        Assert.Equal(LogLevel.Warning, entry.Level);
        Assert.Contains("Excel process 42 was left as it is.", entry.Message, StringComparison.Ordinal);
        Assert.True(backend.Disposed);
    }

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
    public async Task CancelledCreation_PreservesBackendAndRevealsLateSession()
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
            backend.DisposeFailure = new InvalidOperationException("Startup is still pending.");
            Assert.Throws<InvalidOperationException>(bridge.Dispose);
            Assert.False(backend.Disposed);
            backend.Response.SetResult(new ServiceResponse
            {
                Success = true,
                Result = """{"success":true,"sessionId":"late-session"}"""
            });
            Assert.Equal("late-session", await backend.Revealed.Task.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.False(backend.Disposed);
            backend.DisposeFailure = null;
            bridge.Dispose();
            Assert.Equal(2, backend.DisposeCalls);
            Assert.True(backend.Disposed);
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
        internal TaskCompletionSource<string> Revealed { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal bool Disposed { get; private set; }
        internal Exception? DisposeFailure { get; set; }
        internal int DisposeCalls { get; private set; }
        internal string? ShutdownReport { get; init; }

        public string? RequestShutdown() => ShutdownReport;

        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            Started.TrySetResult();
            return Response.Task;
        }

        public void RevealSession(string sessionId)
        {
            Revealed.TrySetResult(sessionId);
        }

        public void Dispose()
        {
            DisposeCalls++;
            if (DisposeFailure is { } failure)
                throw failure;
            Disposed = true;
        }
    }

    private sealed class CapturingLogger : ILogger<ServiceBridge.ServiceBridge>
    {
        internal List<(LogLevel Level, string Message)> Entries { get; } = [];

        public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;
        public bool IsEnabled(LogLevel logLevel) => true;

        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception,
            Func<TState, Exception?, string> formatter) => Entries.Add((logLevel, formatter(state, exception)));
    }
}
