using Bridge = Sbroenne.ExcelMcp.McpServer.ServiceBridge.ServiceBridge;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Collection("ProgramTransport")]
[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceBridge")]
[Trait("Speed", "Fast")]
public sealed class ServiceBridgeCancellationTests : IDisposable
{
    public void Dispose()
    {
        Bridge.ResetForTests();
    }

    [Fact]
    public async Task SendAsync_WithSessionTimeout_ForceClosesSession()
    {
        var backend = new BlockingBackend();
        Bridge.SetServiceFactoryForTests(() => backend);

        var response = await Bridge.SendAsync(
            "sheet.list",
            sessionId: "session-1",
            timeoutSeconds: 1,
            cancellationToken: CancellationToken.None);

        Assert.False(response.Success);
        Assert.Contains("timed out", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Single(backend.ClosedSessions);
        Assert.Contains("session-1", backend.ClosedSessions);
        Assert.False(backend.Disposed);
    }

    [Fact]
    public async Task SendAsync_WithoutSessionCancellation_ResetsService()
    {
        var backend = new BlockingBackend();
        Bridge.SetServiceFactoryForTests(() => backend);

        using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(100));

        var response = await Bridge.SendAsync(
            "session.open",
            sessionId: null,
            cancellationToken: cts.Token);

        Assert.False(response.Success);
        Assert.Contains("cancelled", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.True(backend.Disposed);
    }

    [Fact]
    public async Task ForwardToService_UsesAmbientCancellationToken()
    {
        var backend = new BlockingBackend();
        Bridge.SetServiceFactoryForTests(() => backend);

        using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(100));
        using var cancellationScope = ExcelToolsBase.PushCancellationToken(cts.Token);

        var json = ExcelToolsBase.ForwardToService("sheet.list", "session-ambient");

        Assert.Contains("cancelled", json, StringComparison.OrdinalIgnoreCase);
        Assert.Single(backend.ClosedSessions);
        Assert.Contains("session-ambient", backend.ClosedSessions);
    }

    [Fact]
    public async Task SendAsync_WhenCancellationRacesWithCompletedResponse_ReturnsResponseWithoutCleanup()
    {
        var backend = new DelayedCompletionBackend();
        Bridge.SetServiceFactoryForTests(() => backend);

        using var cts = new CancellationTokenSource();
        var sendTask = Bridge.SendAsync(
            "sheet.list",
            sessionId: "session-race",
            cancellationToken: cts.Token);

        await backend.WaitForRequestAsync();

        cts.Cancel();
        backend.Complete(new ServiceResponse
        {
            Success = true,
            Result = """{"success":true}"""
        });

        var response = await sendTask;

        Assert.True(response.Success);
        Assert.Equal("""{"success":true}""", response.Result);
        Assert.Empty(backend.ClosedSessions);
        Assert.False(backend.Disposed);
    }

    [Fact]
    public async Task SendAsync_WhenServiceFactoryThrows_IncludesStartupFailureDetails()
    {
        Bridge.SetServiceFactoryForTests(static () => throw new FileNotFoundException("office runtime missing"));

        var response = await Bridge.SendAsync("session.open");

        Assert.False(response.Success);
        Assert.Contains("Failed to start ExcelMCP Service in-process", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("FileNotFoundException", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("office runtime missing", response.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public async Task DisposeIfOwnedBy_WithStaleOwner_DoesNotDisposeNewerService()
    {
        var backend = new BlockingBackend(completeImmediately: true);
        Bridge.SetTestOwnerToken(1);
        Bridge.SetServiceFactoryForTests(() => backend);

        var response = await Bridge.SendAsync("sheet.list");

        Assert.True(response.Success);
        Bridge.SetTestOwnerToken(2);

        Assert.False(Bridge.DisposeIfOwnedBy(2_147_483_647));
        Assert.False(backend.Disposed);
    }

    [Fact]
    public async Task DisposeIfOwnedBy_WithMatchingOwner_DisposesService()
    {
        var backend = new BlockingBackend(completeImmediately: true);
        Bridge.SetTestOwnerToken(42);
        Bridge.SetServiceFactoryForTests(() => backend);

        var response = await Bridge.SendAsync("sheet.list");

        Assert.True(response.Success);
        Assert.True(Bridge.DisposeIfOwnedBy(42));
        Assert.True(backend.Disposed);
    }

    private sealed class BlockingBackend : IServiceBridgeBackend
    {
        public List<string> ClosedSessions { get; } = [];

        public bool Disposed { get; private set; }

        private readonly TaskCompletionSource<ServiceResponse> _response =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly bool _completeImmediately;

        public BlockingBackend(bool completeImmediately = false)
        {
            _completeImmediately = completeImmediately;
        }

        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            if (_completeImmediately)
            {
                return Task.FromResult(new ServiceResponse
                {
                    Success = true
                });
            }

            return _response.Task;
        }

        public bool ForceCloseSession(string sessionId)
        {
            ClosedSessions.Add(sessionId);
            _response.TrySetResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = "closed"
            });
            return true;
        }

        public void Dispose()
        {
            Disposed = true;
            _response.TrySetResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = "disposed"
            });
        }
    }

    private sealed class DelayedCompletionBackend : IServiceBridgeBackend
    {
        private readonly TaskCompletionSource<bool> _requestStarted =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<ServiceResponse> _response =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        public List<string> ClosedSessions { get; } = [];

        public bool Disposed { get; private set; }

        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            _requestStarted.TrySetResult(true);
            return _response.Task;
        }

        public async Task WaitForRequestAsync()
        {
            await _requestStarted.Task;
        }

        public void Complete(ServiceResponse response)
        {
            _response.TrySetResult(response);
        }

        public bool ForceCloseSession(string sessionId)
        {
            ClosedSessions.Add(sessionId);
            _response.TrySetResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = "closed"
            });
            return true;
        }

        public void Dispose()
        {
            Disposed = true;
            _response.TrySetResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = "disposed"
            });
        }
    }
}
