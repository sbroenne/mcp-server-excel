using Bridge = Sbroenne.ExcelMcp.McpServer.ServiceBridge.ServiceBridge;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

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

    private sealed class BlockingBackend : IServiceBridgeBackend
    {
        private readonly TaskCompletionSource<ServiceResponse> _response =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        public List<string> ClosedSessions { get; } = [];

        public bool Disposed { get; private set; }

        public Task<ServiceResponse> ProcessAsync(ServiceRequest request) => _response.Task;

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
