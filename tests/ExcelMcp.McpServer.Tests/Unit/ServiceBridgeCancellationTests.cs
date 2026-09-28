using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Bridge = Sbroenne.ExcelMcp.McpServer.ServiceBridge.ServiceBridge;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Collection("ProgramTransport")]
[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceBridge")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
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
        Assert.Equal("Timeout", response.ErrorCategory);
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
        Assert.Equal("Cancelled", response.ErrorCategory);
        Assert.Contains("cancelled", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.True(backend.Disposed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SendAsync_WhenForcedCloseFails_ResetsCapturedService(bool forceCloseThrows)
    {
        var backend = new FailedForceCloseBackend(forceCloseThrows);
        Bridge.SetServiceFactoryForTests(() => backend);
        using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(100));

        var response = await Bridge.SendAsync(
            "sheet.list",
            sessionId: "session-failed-close",
            cancellationToken: cts.Token);

        Assert.False(response.Success);
        Assert.Equal("Cancelled", response.ErrorCategory);
        Assert.Equal(["session-failed-close"], backend.ClosedSessions);
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
        Assert.Equal("ServiceStartup", response.ErrorCategory);
        Assert.Contains("Failed to start ExcelMCP Service in-process", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("FileNotFoundException", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("office runtime missing", response.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public async Task DisposeIfOwnedBy_WithStaleOwner_DoesNotDisposeNewerService()
    {
        var firstBackend = new BlockingBackend(completeImmediately: true);
        Bridge.SetTestOwnerToken(1);
        Bridge.SetServiceFactoryForTests(() => firstBackend);

        Assert.True((await Bridge.SendAsync("sheet.list")).Success);

        Bridge.SetTestOwnerToken(2);
        var secondBackend = new BlockingBackend(completeImmediately: true);
        Bridge.SetServiceFactoryForTests(() => secondBackend);
        Assert.True((await Bridge.SendAsync("sheet.list")).Success);

        Assert.True(firstBackend.Disposed);
        Assert.False(Bridge.DisposeIfOwnedBy(1));
        Assert.False(secondBackend.Disposed);
    }

    [Fact]
    public async Task SendAsync_LateCancellationFromOldGeneration_DoesNotResetNewGeneration()
    {
        var firstBackend = new DetachedBlockingBackend();
        Bridge.SetServiceFactoryForTests(() => firstBackend);
        using var firstCancellation = new CancellationTokenSource();
        var firstRequest = Bridge.SendAsync(
            "session.open",
            cancellationToken: firstCancellation.Token);
        await firstBackend.WaitForRequestAsync();

        var secondBackend = new BlockingBackend(completeImmediately: true);
        Bridge.SetServiceFactoryForTests(() => secondBackend);
        Assert.True((await Bridge.SendAsync("sheet.list")).Success);

        firstCancellation.Cancel();
        var cancelledResponse = await firstRequest;

        Assert.Equal("Cancelled", cancelledResponse.ErrorCategory);
        Assert.True(firstBackend.Disposed);
        Assert.False(secondBackend.Disposed);
        Assert.True((await Bridge.SendAsync("sheet.list")).Success);
    }

    [Fact]
    public async Task SendAsync_CancelledRequestResetsBackendAndCompletesOtherBlockedRequest()
    {
        var backend = new MultiRequestBlockingBackend(expectedRequests: 2);
        using var lifetime = new ServiceBridgeLifetime(() => backend);
        using var cancellation = new CancellationTokenSource();

        var cancelledRequest = lifetime.SendAsync(
            "session.open",
            sessionId: null,
            args: null,
            timeoutSeconds: null,
            cancellation.Token);
        var otherRequest = lifetime.SendAsync(
            "session.open",
            sessionId: null,
            args: null,
            timeoutSeconds: null,
            CancellationToken.None);
        await backend.WaitForRequestsAsync();

        cancellation.Cancel();

        var cancelledResponse = await cancelledRequest.WaitAsync(TimeSpan.FromSeconds(5));
        var otherResponse = await otherRequest.WaitAsync(TimeSpan.FromSeconds(5));

        Assert.Equal("Cancelled", cancelledResponse.ErrorCategory);
        Assert.False(otherResponse.Success);
        Assert.Equal("disposed", otherResponse.ErrorMessage);
        Assert.True(backend.Disposed);
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

    [Fact]
    public async Task Dispose_DuringActiveRequest_DisposesBackendAndCompletesResponse()
    {
        var backend = new DelayedCompletionBackend();
        using var lifetime = new ServiceBridgeLifetime(() => backend);
        var sendTask = lifetime.SendAsync(
            "sheet.list",
            sessionId: null,
            args: null,
            timeoutSeconds: null,
            CancellationToken.None);
        await backend.WaitForRequestAsync();

        lifetime.Dispose();

        Assert.True(backend.Disposed);
        var response = await sendTask.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.False(response.Success);
        Assert.Equal("disposed", response.ErrorMessage);
    }

    [Fact]
    public async Task Dispose_DuringFactoryInitialization_DiscardsStaleBackend()
    {
        using var factoryEntered = new ManualResetEventSlim();
        using var releaseFactory = new ManualResetEventSlim();
        var staleBackend = new BlockingBackend(completeImmediately: true);
        var currentBackend = new BlockingBackend(completeImmediately: true);
        var factoryCalls = 0;
        using var lifetime = new ServiceBridgeLifetime(() =>
        {
            if (Interlocked.Increment(ref factoryCalls) == 1)
            {
                factoryEntered.Set();
                releaseFactory.Wait();
                return staleBackend;
            }

            return currentBackend;
        });

        var sendTask = Task.Run(() => lifetime.SendAsync(
            "sheet.list",
            sessionId: null,
            args: null,
            timeoutSeconds: null,
            CancellationToken.None));
        Assert.True(factoryEntered.Wait(TimeSpan.FromSeconds(5)));

        lifetime.Dispose();
        releaseFactory.Set();

        Assert.True((await sendTask).Success);
        Assert.True(staleBackend.Disposed);
        Assert.False(currentBackend.Disposed);
        Assert.Equal(2, factoryCalls);
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

    private sealed class MultiRequestBlockingBackend(int expectedRequests)
        : IServiceBridgeBackend
    {
        private readonly TaskCompletionSource<bool> _requestsStarted =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<ServiceResponse> _response =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _requestCount;

        internal bool Disposed { get; private set; }

        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            if (Interlocked.Increment(ref _requestCount) == expectedRequests)
            {
                _requestsStarted.TrySetResult(true);
            }

            return _response.Task;
        }

        internal Task<bool> WaitForRequestsAsync() => _requestsStarted.Task;

        public bool ForceCloseSession(string sessionId) => false;

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

    private sealed class FailedForceCloseBackend(bool forceCloseThrows)
        : IServiceBridgeBackend
    {
        private readonly TaskCompletionSource<ServiceResponse> _response =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        internal List<string> ClosedSessions { get; } = [];
        internal bool Disposed { get; private set; }

        public Task<ServiceResponse> ProcessAsync(ServiceRequest request) => _response.Task;

        public bool ForceCloseSession(string sessionId)
        {
            ClosedSessions.Add(sessionId);
            if (forceCloseThrows)
            {
                throw new InvalidOperationException("forced close failed");
            }

            return false;
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

    private sealed class DetachedBlockingBackend : IServiceBridgeBackend
    {
        private readonly TaskCompletionSource<bool> _requestStarted =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<ServiceResponse> _response =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        internal bool Disposed { get; private set; }

        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            _requestStarted.TrySetResult(true);
            return _response.Task;
        }

        internal Task<bool> WaitForRequestAsync() => _requestStarted.Task;

        public bool ForceCloseSession(string sessionId) => false;

        public void Dispose()
        {
            Disposed = true;
        }
    }
}
