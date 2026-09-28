using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

public sealed class RecordingProgramTransportFixture :
    IAsyncLifetime,
    IAsyncDisposable
{
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private readonly SemaphoreSlim _callGate = new(1, 1);
    private readonly RecordingBackend _backend = new();
    private readonly FixtureOutputHelper _output = new();
    private McpClient? _client;
    private Task? _serverTask;
    private bool _disposed;

    public async Task InitializeAsync()
    {
        ServiceBridge.ServiceBridge.SetServiceFactoryForTests(() => _backend);
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            _cts.Token,
            "RecordingProgramTransportClient");
    }

    public async Task<CapturedToolCall> CallToolAsync(
        string toolName,
        Dictionary<string, object?> arguments,
        ServiceResponse response,
        string expectedCommand,
        string? expectedArgsJson)
    {
        return await CallToolAsync(
            toolName,
            arguments,
            response,
            expectedCommand,
            arguments["session_id"] as string
                ?? throw new InvalidOperationException(
                    "Session-scoped recording calls must supply a session_id."),
            expectedArgsJson);
    }

    public async Task<CapturedToolCall> CallToolAsync(
        string toolName,
        Dictionary<string, object?> arguments,
        ServiceResponse response,
        string expectedCommand,
        string? expectedSessionId,
        string? expectedArgsJson)
    {
        await _callGate.WaitAsync(_cts.Token);
        try
        {
            _backend.Prepare(response);
            var client = _client
                ?? throw new InvalidOperationException(
                    "The recording MCP transport fixture is not initialized.");
            var result = await client.CallToolAsync(
                toolName,
                arguments,
                cancellationToken: _cts.Token);
            var text = result.Content
                .OfType<TextContentBlock>()
                .FirstOrDefault()?.Text
                ?? throw new InvalidOperationException(
                    $"Unexpected response from MCP tool '{toolName}'.");
            AssertSerializedResult(text, response);
            var request = _backend.TakeRequest();
            RecordingToolTest.AssertRequest(
                request,
                expectedCommand,
                expectedSessionId,
                expectedArgsJson);
            return new CapturedToolCall(text, request);
        }
        finally
        {
            _callGate.Release();
        }
    }

    public async Task<string> CallToolWithoutDispatchAsync(
        string toolName,
        Dictionary<string, object?> arguments)
    {
        await _callGate.WaitAsync(_cts.Token);
        try
        {
            _backend.AssertIdle();
            var client = _client
                ?? throw new InvalidOperationException(
                    "The recording MCP transport fixture is not initialized.");
            var result = await client.CallToolAsync(
                toolName,
                arguments,
                cancellationToken: _cts.Token);
            _backend.AssertIdle();
            return result.Content
                .OfType<TextContentBlock>()
                .FirstOrDefault()?.Text
                ?? throw new InvalidOperationException(
                    $"Unexpected response from MCP tool '{toolName}'.");
        }
        finally
        {
            _callGate.Release();
        }
    }

    public async Task<IList<McpClientTool>> ListToolsAsync()
    {
        await _callGate.WaitAsync(_cts.Token);
        try
        {
            _backend.AssertIdle();
            var client = _client
                ?? throw new InvalidOperationException(
                    "The recording MCP transport fixture is not initialized.");
            var tools = await client.ListToolsAsync(cancellationToken: _cts.Token);
            _backend.AssertIdle();
            return tools;
        }
        finally
        {
            _callGate.Release();
        }
    }

    public async Task DisposeAsync()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        List<Exception>? failures = null;

        try
        {
            await ProgramTransportTestHost.StopAsync(
                _client,
                _clientToServerPipe,
                _serverToClientPipe,
                _serverTask,
                _output,
                _cts);
        }
        catch (Exception ex)
        {
            (failures ??= []).Add(ex);
        }

        try
        {
            ServiceBridge.ServiceBridge.ResetForTests();
        }
        catch (Exception ex)
        {
            (failures ??= []).Add(ex);
        }

        _client = null;
        _cts.Dispose();
        _callGate.Dispose();

        if (failures is not null)
        {
            throw new AggregateException(
                "Recording MCP transport fixture cleanup failed.",
                failures);
        }
    }

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsync();
        GC.SuppressFinalize(this);
    }

    public sealed record CapturedToolCall(
        string JsonResult,
        ServiceRequest Request);

    private static void AssertSerializedResult(
        string jsonResult,
        ServiceResponse response)
    {
        bool expectedSuccess = response.Success;
        if (response.Result is not null)
        {
            using var configuredResult = JsonDocument.Parse(response.Result);
            expectedSuccess = configuredResult.RootElement
                .GetProperty("success")
                .GetBoolean();
        }

        using var actual = JsonDocument.Parse(jsonResult);
        var root = actual.RootElement;
        Assert.Equal(expectedSuccess, root.GetProperty("success").GetBoolean());
        if (!expectedSuccess)
        {
            Assert.False(string.IsNullOrWhiteSpace(
                root.GetProperty("errorMessage").GetString()));
        }
    }

    private sealed class RecordingBackend : IServiceBridgeBackend
    {
        private ServiceResponse? _response;
        private ServiceRequest? _request;

        public void Prepare(ServiceResponse response)
        {
            if (_response is not null || _request is not null)
            {
                throw new InvalidOperationException(
                    "The previous recording call was not consumed.");
            }

            _response = response;
        }

        public void AssertIdle()
        {
            if (_response is not null || _request is not null)
            {
                throw new InvalidOperationException(
                    "The recording backend unexpectedly received a request.");
            }
        }

        public ServiceRequest TakeRequest()
        {
            var request = _request
                ?? throw new InvalidOperationException(
                    "The MCP tool did not dispatch a Service request.");
            _request = null;
            _response = null;
            return request;
        }

        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            if (_request is not null)
            {
                throw new InvalidOperationException(
                    "The MCP tool dispatched more than one Service request.");
            }

            _request = request;
            return Task.FromResult(
                _response
                ?? throw new InvalidOperationException(
                    "No recording response was configured."));
        }

        public bool ForceCloseSession(string sessionId) => false;

        public void Dispose()
        {
        }
    }

    private sealed class FixtureOutputHelper : ITestOutputHelper
    {
        public void WriteLine(string message)
        {
        }

        public void WriteLine(string format, params object[] args)
        {
        }
    }
}

[CollectionDefinition("RecordingProgramTransport", DisableParallelization = true)]
public sealed class RecordingProgramTransportGroup :
    ICollectionFixture<RecordingProgramTransportFixture>;
