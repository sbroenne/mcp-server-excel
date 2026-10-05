using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

public sealed class McpProgramTransportFixture :
    IAsyncLifetime,
    IAsyncDisposable
{
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private readonly SemaphoreSlim _sessionGate = new(1, 1);
    private readonly HashSet<string> _trackedSessionIds = new(StringComparer.Ordinal);
    private readonly McpOwnedExcelProcesses _ownedExcelProcesses = new();
    private readonly FixtureOutputHelper _output = new();
    private Task? _serverTask;
    private McpClient? _client;
    private string? _sharedSessionId;
    private string? _tempDirectory;
    private bool _openedSession;
    private bool _disposed;

    public async Task InitializeAsync()
    {
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            _cts.Token,
            "SharedProgramTransportClient",
            _ownedExcelProcesses.CreateBackend);
        _tempDirectory = Path.Join(
            Path.GetTempPath(),
            $"McpProgramTransport_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDirectory);
    }

    public async Task<string> GetSharedWorkbookSessionAsync()
    {
        await _sessionGate.WaitAsync(_cts.Token);
        try
        {
            if (!string.IsNullOrWhiteSpace(_sharedSessionId))
            {
                return _sharedSessionId;
            }

            var workbookPath = CreateTempPath("SharedWorkbook", ".xlsx");
            var result = await CallToolAsync(
                "file",
                new Dictionary<string, object?>
                {
                    ["action"] = "create",
                    ["path"] = workbookPath,
                    ["show"] = false
                },
                TimeSpan.FromSeconds(90));

            AssertSuccess(result, "file.create shared workbook");
            _sharedSessionId = GetJsonProperty(result, "workbook_session_id");
            Assert.False(string.IsNullOrWhiteSpace(_sharedSessionId));
            _openedSession = true;
            _trackedSessionIds.Add(_sharedSessionId!);
            return _sharedSessionId!;
        }
        finally
        {
            _sessionGate.Release();
        }
    }

    public async Task<string> CreateWorkbookSessionAsync(string workbookPath)
    {
        var result = await CallToolAsync(
            "file",
            new Dictionary<string, object?>
            {
                ["action"] = "create",
                ["path"] = workbookPath,
                ["show"] = false
            },
            TimeSpan.FromSeconds(90));

        AssertSuccess(result, "file.create");
        var sessionId = GetJsonProperty(result, "workbook_session_id");
        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        _openedSession = true;
        _trackedSessionIds.Add(sessionId!);
        return sessionId!;
    }

    public async Task CloseSessionAsync(string sessionId, bool save = false)
    {
        var result = await CallToolAsync(
            "file",
            new Dictionary<string, object?>
            {
                ["action"] = "close",
                ["workbook_session_id"] = sessionId,
                ["save"] = save
            },
            TimeSpan.FromSeconds(30));
        AssertSuccess(result, "file.close");
        _trackedSessionIds.Remove(sessionId);
    }

    public string CreateTempPath(string prefix, string extension)
    {
        if (string.IsNullOrWhiteSpace(_tempDirectory))
        {
            throw new InvalidOperationException(
                "The MCP transport fixture is not initialized.");
        }

        return Path.Join(
            _tempDirectory,
            $"{prefix}_{Guid.NewGuid():N}{extension}");
    }

    public async Task<string> CallToolAsync(
        string toolName,
        Dictionary<string, object?> arguments,
        TimeSpan? timeout = null)
    {
        var client = _client
            ?? throw new InvalidOperationException(
                "The MCP transport fixture is not initialized.");
        var callTask = client.CallToolAsync(
            toolName,
            arguments,
            cancellationToken: _cts.Token).AsTask();
        var result = timeout.HasValue
            ? await callTask.WaitAsync(timeout.Value, _cts.Token)
            : await callTask;
        var textBlock = result.Content.OfType<TextContentBlock>().FirstOrDefault();
        return textBlock?.Text
            ?? throw new InvalidOperationException(
                $"Unexpected response from MCP tool '{toolName}'.");
    }

    public async Task DisposeAsync()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        List<Exception>? failures = null;

        var closeFailures = await CloseTrackedSessionsAsync(
            _trackedSessionIds,
            sessionId =>
                CallToolAsync(
                    "file",
                    new Dictionary<string, object?>
                    {
                        ["action"] = "close",
                        ["workbook_session_id"] = sessionId,
                        ["save"] = false
                    },
                    TimeSpan.FromSeconds(30)));
        if (closeFailures.Count > 0)
        {
            (failures ??= []).AddRange(closeFailures);
        }

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
            await _ownedExcelProcesses.AssertExitedAsync(
                _output,
                _openedSession);
        }
        catch (Exception ex)
        {
            (failures ??= []).Add(ex);
        }

        try
        {
            if (!string.IsNullOrWhiteSpace(_tempDirectory)
                && Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, recursive: true);
            }
        }
        catch (Exception ex)
        {
            (failures ??= []).Add(ex);
        }

        _client = null;
        _cts.Dispose();
        _sessionGate.Dispose();
        _ownedExcelProcesses.Dispose();

        if (failures is not null)
        {
            throw new AggregateException(
                "Shared MCP transport fixture cleanup failed.",
                failures);
        }
    }

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsync();
        GC.SuppressFinalize(this);
    }

    internal static async Task<IReadOnlyList<Exception>> CloseTrackedSessionsAsync(
        ISet<string> trackedSessionIds,
        Func<string, Task<string>> closeSessionAsync)
    {
        List<Exception>? failures = null;

        foreach (var sessionId in trackedSessionIds.ToArray())
        {
            try
            {
                var result = await closeSessionAsync(sessionId);
                AssertSuccess(result, "file.close");
                trackedSessionIds.Remove(sessionId);
            }
            catch (Exception ex)
            {
                (failures ??= []).Add(ex);
            }
        }

        return failures ?? [];
    }

    private static void AssertSuccess(string jsonResult, string operationName)
    {
        using var json = JsonDocument.Parse(jsonResult);
        Assert.True(
            json.RootElement.GetProperty("success").GetBoolean(),
            $"{operationName} failed: {jsonResult}");
    }

    private static string? GetJsonProperty(
        string jsonResult,
        string propertyName)
    {
        using var json = JsonDocument.Parse(jsonResult);
        return json.RootElement.TryGetProperty(propertyName, out var property)
            ? property.GetString()
            : null;
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
