using System.Diagnostics;
using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Shared harness for MCP integration tests that exercise the real Program transport.
/// </summary>
[System.Diagnostics.CodeAnalysis.SuppressMessage("Design", "CA1001", Justification = "_cts is disposed in DisposeAsync.")]
public abstract class McpIntegrationTestBase : IAsyncLifetime
{
    private static readonly TimeSpan ExcelShutdownTimeout = TimeSpan.FromSeconds(15);
    private static readonly TimeSpan ExcelShutdownPollInterval = TimeSpan.FromMilliseconds(250);
    private readonly string _clientName;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private readonly HashSet<string> _trackedSessionIds = new(StringComparer.Ordinal);
    private readonly List<string> _tempDirectories = [];
    private readonly HashSet<int> _baselineExcelProcessIds = [];
    private Task? _serverTask;
    private bool _disposed;
    private bool _openedSession;

    protected McpIntegrationTestBase(ITestOutputHelper output, string clientName)
    {
        Output = output;
        _clientName = clientName;
    }

    protected ITestOutputHelper Output { get; }

    protected McpClient? Client { get; private set; }

    protected CancellationToken TestCancellationToken => _cts.Token;

    protected virtual TimeSpan ClientInitializationTimeout => TimeSpan.FromSeconds(30);

    protected virtual TimeSpan ServerShutdownTimeout => TimeSpan.FromSeconds(10);

    protected virtual TimeSpan ServerReadyTimeout => TimeSpan.FromSeconds(15);

    public async Task InitializeAsync()
    {
        CaptureBaselineExcelProcesses();

        try
        {
            (Client, _serverTask) = await ProgramTransportTestHost.StartAsync(
                _clientToServerPipe,
                _serverToClientPipe,
                TestCancellationToken,
                _clientName);

            await InitializeTestAsync();
        }
        catch
        {
            await DisposeAsync();
            throw;
        }
    }

    public async Task DisposeAsync()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        try
        {
            await CloseTrackedSessionsAsync();
            await BeforeServerShutdownAsync();
        }
        finally
        {
            await ProgramTransportTestHost.StopAsync(
                Client,
                _clientToServerPipe,
                _serverToClientPipe,
                _serverTask,
                Output,
                _cts);
            Client = null;

            if (!_cts.IsCancellationRequested)
            {
                await _cts.CancelAsync();
            }

            _cts.Dispose();

            await AfterServerShutdownAsync();
            await AssertNoLeakedExcelProcessesAsync();
            CleanupTempDirectories();
        }
    }

    protected virtual Task InitializeTestAsync() => Task.CompletedTask;

    protected virtual Task BeforeServerShutdownAsync() => Task.CompletedTask;

    protected virtual Task AfterServerShutdownAsync() => Task.CompletedTask;

    protected string CreateTempDirectory(string prefix)
    {
        var tempDirectory = Path.Join(Path.GetTempPath(), $"{prefix}_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDirectory);
        _tempDirectories.Add(tempDirectory);
        Output.WriteLine($"Test directory: {tempDirectory}");
        return tempDirectory;
    }

    protected async Task<string> CreateWorkbookSessionAsync(string workbookPath)
    {
        var createJson = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["path"] = workbookPath
        });

        AssertSetupSuccess(createJson, $"file.create ({Path.GetFileName(workbookPath)})");

        using var createDoc = JsonDocument.Parse(createJson);
        var sessionId = createDoc.RootElement.GetProperty("session_id").GetString();
        TrackSession(sessionId);
        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        return sessionId!;
    }

    protected async Task CreateWorksheetAsync(string sessionId, string sheetName)
    {
        var createSheetJson = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = sessionId,
            ["sheet_name"] = sheetName
        });

        AssertSetupSuccess(createSheetJson, $"worksheet.create ({sheetName})");
    }

    protected void TrackSession(string? sessionId)
    {
        if (!string.IsNullOrWhiteSpace(sessionId))
        {
            _openedSession = true;
            _trackedSessionIds.Add(sessionId);
        }
    }

    protected void UntrackSession(string? sessionId)
    {
        if (!string.IsNullOrWhiteSpace(sessionId))
        {
            _trackedSessionIds.Remove(sessionId);
        }
    }

    protected async Task CloseSessionAsync(string? sessionId, bool save)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return;
        }

        await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "close",
            ["session_id"] = sessionId,
            ["save"] = save
        });

        UntrackSession(sessionId);
    }

    protected async Task TryCloseSessionAsync(string? sessionId, bool save = false)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || Client == null)
        {
            return;
        }

        try
        {
            await CloseSessionAsync(sessionId, save);
        }
        catch (Exception ex)
        {
            Output.WriteLine($"Warning: Failed to close session {sessionId}: {ex.Message}");
        }
    }

    protected async Task<string> CallToolAsync(string toolName, Dictionary<string, object?> arguments, TimeSpan? timeout = null)
    {
        Assert.NotNull(Client);

        var callTask = Client!.CallToolAsync(toolName, arguments, cancellationToken: TestCancellationToken).AsTask();
        var result = timeout.HasValue
            ? await callTask.WaitAsync(timeout.Value, TestCancellationToken)
            : await callTask;

        Assert.NotNull(result);
        Assert.NotNull(result.Content);
        Assert.NotEmpty(result.Content);

        var textBlock = result.Content.OfType<TextContentBlock>().FirstOrDefault();
        Assert.NotNull(textBlock);

        return textBlock.Text;
    }

    protected static void AssertSuccess(string jsonResult, string operationName)
    {
        Assert.True(
            jsonResult.TrimStart().StartsWith('{'),
            $"{operationName} returned a non-JSON response. Response: {jsonResult}");

        using var json = JsonDocument.Parse(jsonResult);
        var root = json.RootElement;

        if (root.TryGetProperty("Success", out var successPascal))
        {
            Assert.True(successPascal.GetBoolean(), $"{operationName} failed: {jsonResult}");
            return;
        }

        Assert.True(root.GetProperty("success").GetBoolean(), $"{operationName} failed: {jsonResult}");
    }

    protected static void AssertSetupSuccess(string jsonResult, string operationName)
    {
        AssertSuccess(jsonResult, $"{operationName} setup");
    }

    protected static JsonDocument ParseJsonResult(string jsonResult, string operationName)
    {
        Assert.True(
            jsonResult.TrimStart().StartsWith('{'),
            $"{operationName} returned a non-JSON response. Response: {jsonResult}");

        return JsonDocument.Parse(jsonResult);
    }

    protected static void AssertFailureEnvelope(
        JsonElement root,
        string operationName,
        string expectedExceptionType,
        string? expectedErrorCategory = null,
        string? expectedHResult = null,
        string? expectedInnerError = null,
        bool allowOptionalNonEmptyInnerError = false)
    {
        Assert.False(root.GetProperty("success").GetBoolean(), $"{operationName} unexpectedly succeeded.");
        Assert.True(root.GetProperty("isError").GetBoolean(), $"{operationName} should return isError=true.");
        Assert.Equal(expectedExceptionType, root.GetProperty("exceptionType").GetString());

        var error = root.GetProperty("error").GetString();
        var errorMessage = root.GetProperty("errorMessage").GetString();

        Assert.False(string.IsNullOrWhiteSpace(errorMessage), $"{operationName} should return errorMessage.");
        Assert.Equal(errorMessage, error);

        AssertOptionalStringProperty(root, "errorCategory", expectedErrorCategory, operationName);
        AssertOptionalStringProperty(root, "hresult", expectedHResult, operationName);
        if (allowOptionalNonEmptyInnerError)
        {
            AssertOptionalNonEmptyStringProperty(root, "innerError", operationName);
        }
        else
        {
            AssertOptionalStringProperty(root, "innerError", expectedInnerError, operationName);
        }
    }

    protected static void AssertHasNonEmptyStringProperty(
        JsonElement root,
        string propertyName,
        string operationName)
    {
        Assert.True(root.TryGetProperty(propertyName, out var property), $"{operationName} should include {propertyName}.");
        Assert.False(string.IsNullOrWhiteSpace(property.GetString()), $"{operationName} should include a non-empty {propertyName}.");
    }

    protected static void AssertOptionalNonEmptyStringProperty(
        JsonElement root,
        string propertyName,
        string operationName)
    {
        if (!root.TryGetProperty(propertyName, out var property))
        {
            return;
        }

        Assert.False(string.IsNullOrWhiteSpace(property.GetString()), $"{operationName} should include a non-empty {propertyName} when present.");
    }

    protected static string? GetJsonProperty(string jsonResult, string propertyName)
    {
        using var json = JsonDocument.Parse(jsonResult);
        return json.RootElement.TryGetProperty(propertyName, out var prop) ? prop.GetString() : null;
    }

    private static void AssertOptionalStringProperty(
        JsonElement root,
        string propertyName,
        string? expectedValue,
        string operationName)
    {
        var hasProperty = root.TryGetProperty(propertyName, out var property);
        if (expectedValue is null)
        {
            Assert.False(hasProperty, $"{operationName} should not include {propertyName} when no value is expected.");
            return;
        }

        if (expectedValue.Length == 0)
        {
            Assert.True(hasProperty, $"{operationName} should include {propertyName}.");
            Assert.False(string.IsNullOrWhiteSpace(property.GetString()), $"{operationName} should include a non-empty {propertyName}.");
            return;
        }

        Assert.True(hasProperty, $"{operationName} should include {propertyName}.");
        Assert.Equal(expectedValue, property.GetString());
    }

    private async Task CloseTrackedSessionsAsync()
    {
        if (Client == null || _trackedSessionIds.Count == 0)
        {
            return;
        }

        foreach (var sessionId in _trackedSessionIds.ToArray())
        {
            await TryCloseSessionAsync(sessionId, save: false);
        }
    }

    private void CleanupTempDirectories()
    {
        foreach (var tempDirectory in _tempDirectories)
        {
            for (var attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    if (Directory.Exists(tempDirectory))
                    {
                        Directory.Delete(tempDirectory, recursive: true);
                    }

                    break;
                }
                catch (IOException) when (attempt < 3)
                {
                    Thread.Sleep(250);
                }
                catch (UnauthorizedAccessException) when (attempt < 3)
                {
                    Thread.Sleep(250);
                }
                catch (Exception ex)
                {
                    Output.WriteLine($"Warning: Failed to delete temp directory {tempDirectory}: {ex.Message}");
                    break;
                }
            }
        }
    }

    private void CaptureBaselineExcelProcesses()
    {
        _baselineExcelProcessIds.Clear();
        foreach (var processId in GetCurrentExcelProcessIds())
        {
            _baselineExcelProcessIds.Add(processId);
        }
    }

    private async Task AssertNoLeakedExcelProcessesAsync()
    {
        if (!_openedSession)
        {
            return;
        }

        var deadline = DateTime.UtcNow + ExcelShutdownTimeout;
        List<int> leakedExcelProcessIds;
        do
        {
            leakedExcelProcessIds = GetCurrentExcelProcessIds()
                .Where(processId => !_baselineExcelProcessIds.Contains(processId))
                .ToList();

            if (leakedExcelProcessIds.Count == 0)
            {
                return;
            }

            await Task.Delay(ExcelShutdownPollInterval);
        }
        while (DateTime.UtcNow < deadline);

        ForceKillProcesses(leakedExcelProcessIds);

        throw new XunitException(
            $"Excel processes started during the MCP integration test did not exit after shutdown. " +
            $"Baseline PIDs: [{string.Join(", ", _baselineExcelProcessIds.OrderBy(id => id))}]. " +
            $"Leaked PIDs: [{string.Join(", ", leakedExcelProcessIds.OrderBy(id => id))}].");
    }

    private static int[] GetCurrentExcelProcessIds()
    {
        return Process.GetProcessesByName("EXCEL")
            .Select(process => process.Id)
            .ToArray();
    }

    private void ForceKillProcesses(IEnumerable<int> processIds)
    {
        foreach (var processId in processIds.Distinct())
        {
            try
            {
                using var process = Process.GetProcessById(processId);
                Output.WriteLine($"Warning: Force-killing leaked Excel process {processId} after MCP integration test teardown.");
                process.Kill(entireProcessTree: true);
                process.WaitForExit((int)ExcelShutdownTimeout.TotalMilliseconds);
            }
            catch (ArgumentException)
            {
            }
            catch (Exception ex)
            {
                Output.WriteLine($"Warning: Failed to force-kill leaked Excel process {processId}: {ex.Message}");
            }
        }
    }
}
