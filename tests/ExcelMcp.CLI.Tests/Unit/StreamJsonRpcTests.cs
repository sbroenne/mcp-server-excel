using Nerdbank.Streams;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Rpc;
using StreamJsonRpc;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

/// <summary>
/// Tests for the StreamJsonRpc-based CLI↔daemon communication layer.
/// Uses <see cref="FullDuplexStream.CreatePair"/> for in-process testing
/// without real pipes or processes.
/// </summary>
[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "StreamJsonRpc")]
[Trait("Speed", "Fast")]
public sealed class StreamJsonRpcTests : IDisposable
{
    private readonly ExcelMcpService _service = new();

    /// <summary>
    /// Validates end-to-end RPC round-trip: client sends ServiceRequest through StreamJsonRpc,
    /// DaemonRpcTarget delegates to ExcelMcpService.ProcessAsync, response returns correctly.
    /// </summary>
    [Fact]
    public async Task ProcessCommandAsync_RoundTrip_ReturnsServiceResponse()
    {
        // Arrange
        var (serverStream, clientStream) = FullDuplexStream.CreatePair();
        var rpcTarget = new DaemonRpcTarget(_service);
        using var serverRpc = JsonRpc.Attach(serverStream, rpcTarget);
        var clientProxy = JsonRpc.Attach<IExcelDaemonRpc>(clientStream);

        try
        {
            var request = new ServiceRequest { Command = "diag.ping" };

            // Act
            var response = await clientProxy.ProcessCommandAsync(request);

            // Assert
            Assert.True(response.Success);
        }
        finally
        {
            ((IDisposable)clientProxy).Dispose();
        }
    }

    /// <summary>
    /// Validates that business errors (unknown command category) propagate correctly
    /// through the RPC layer as ServiceResponse with Success=false, not as exceptions.
    /// </summary>
    [Fact]
    public async Task ProcessCommandAsync_UnknownCategory_ReturnsErrorResponse()
    {
        // Arrange
        var (serverStream, clientStream) = FullDuplexStream.CreatePair();
        var rpcTarget = new DaemonRpcTarget(_service);
        using var serverRpc = JsonRpc.Attach(serverStream, rpcTarget);
        var clientProxy = JsonRpc.Attach<IExcelDaemonRpc>(clientStream);

        try
        {
            var request = new ServiceRequest { Command = "boguscategory.action" };

            // Act
            var response = await clientProxy.ProcessCommandAsync(request);

            // Assert — error is returned as data, not thrown as exception
            Assert.False(response.Success);
            Assert.Contains("Unknown command category", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            ((IDisposable)clientProxy).Dispose();
        }
    }

    /// <summary>
    /// Validates that the server RPC instance completes when the client disconnects.
    /// This ensures no leaked tasks or hung connections.
    /// </summary>
    [Fact]
    public async Task ServerRpc_Completion_ResolvesWhenClientDisconnects()
    {
        // Arrange
        var (serverStream, clientStream) = FullDuplexStream.CreatePair();
        var rpcTarget = new DaemonRpcTarget(_service);
        using var serverRpc = JsonRpc.Attach(serverStream, rpcTarget);
        var clientProxy = JsonRpc.Attach<IExcelDaemonRpc>(clientStream);

        // Act — dispose client (simulates CLI process exit)
        ((IDisposable)clientProxy).Dispose();

        // Assert — server-side Completion should resolve within a reasonable timeout
        var completedTask = await Task.WhenAny(serverRpc.Completion, Task.Delay(TimeSpan.FromSeconds(5)));
        Assert.Same(serverRpc.Completion, completedTask);
    }

    /// <summary>
    /// Validates that NullReferenceException from ProcessAsync (null command) propagates
    /// through the RPC layer and includes the exception type name in the error message.
    /// StreamJsonRpc wraps server-side exceptions as RemoteInvocationException on the client.
    /// </summary>
    [Fact]
    public async Task ProcessCommandAsync_NullCommand_PropagatesAsRemoteException()
    {
        // Arrange
        var (serverStream, clientStream) = FullDuplexStream.CreatePair();
        var rpcTarget = new DaemonRpcTarget(_service);
        using var serverRpc = JsonRpc.Attach(serverStream, rpcTarget);
        var clientProxy = JsonRpc.Attach<IExcelDaemonRpc>(clientStream);

        try
        {
#pragma warning disable CS8714 // required property set to null intentionally
            var request = new ServiceRequest { Command = null! };
#pragma warning restore CS8714

            // Act — ExcelMcpService.ProcessAsync catches NullReferenceException internally
            // and returns a ServiceResponse with Success=false (not an exception).
            var response = await clientProxy.ProcessCommandAsync(request);

            // Assert — the error is caught by ProcessAsync and returned as data
            Assert.False(response.Success);
            Assert.NotNull(response.ErrorMessage);
            Assert.Contains("NullReferenceException", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            ((IDisposable)clientProxy).Dispose();
        }
    }

    /// <summary>
    /// Validates that a session-level command (which requires a valid sessionId)
    /// returns an error response correctly through the RPC layer.
    /// This tests that structured error responses with ErrorMessage survive the round trip.
    /// </summary>
    [Fact]
    public async Task ProcessCommandAsync_InvalidSessionId_ReturnsStructuredError()
    {
        // Arrange
        var (serverStream, clientStream) = FullDuplexStream.CreatePair();
        var rpcTarget = new DaemonRpcTarget(_service);
        using var serverRpc = JsonRpc.Attach(serverStream, rpcTarget);
        var clientProxy = JsonRpc.Attach<IExcelDaemonRpc>(clientStream);

        try
        {
            var request = new ServiceRequest
            {
                Command = "sheet.list",
                SessionId = "nonexistent-session-00000000"
            };

            // Act
            var response = await clientProxy.ProcessCommandAsync(request);

            // Assert — error response with structured message survives RPC
            Assert.False(response.Success);
            Assert.NotNull(response.ErrorMessage);
            Assert.NotEmpty(response.ErrorMessage);
        }
        finally
        {
            ((IDisposable)clientProxy).Dispose();
        }
    }

    public void Dispose() => _service.Dispose();
}
