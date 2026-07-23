using System.Collections.Concurrent;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Real-daemon regression coverage for lifecycle races that in-process service tests cannot see.
/// </summary>
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Slow")]
public sealed class CliDaemonLifecycleRegressionTests : IAsyncLifetime, IClassFixture<TempDirectoryFixture>
{
    private const int ReopenIterations = 3;
    private const int ConcurrentCreateCount = 4;

    private readonly ITestOutputHelper _output;
    private readonly TempDirectoryFixture _fixture;
    private readonly string _uniquePipeName = $"excelmcp-daemon-lifecycle-{Guid.NewGuid():N}";
    private readonly ConcurrentDictionary<string, string> _activeSessions = new(StringComparer.Ordinal);

    public CliDaemonLifecycleRegressionTests(TempDirectoryFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private Dictionary<string, string> UniquePipeEnv => new() { ["EXCELMCP_CLI_PIPE"] = _uniquePipeName };

    public async Task InitializeAsync()
    {
        await StopServiceAsync(environmentVariables: null, "initialize-default-stop");
        await StopServiceAsync(UniquePipeEnv, "initialize-unique-stop");
    }

    public async Task DisposeAsync()
    {
        foreach (var (sessionId, pipeName) in _activeSessions)
        {
            await CloseSessionBestEffortAsync(sessionId, save: false, EnvironmentForPipe(pipeName), $"cleanup-close-{sessionId}");
        }

        await StopServiceAsync(environmentVariables: null, "cleanup-default-stop");
        await StopServiceAsync(UniquePipeEnv, "cleanup-unique-stop");
    }

    [Fact(Timeout = 300000)]
    public async Task DefaultPipe_RepeatedCloseSaveStatusReopen_KeepsDaemonResponsiveAndSessionless()
    {
        for (var iteration = 0; iteration < ReopenIterations; iteration++)
        {
            var workbookPath = CreateNewWorkbookPath(nameof(DefaultPipe_RepeatedCloseSaveStatusReopen_KeepsDaemonResponsiveAndSessionless), iteration);
            var sheetName = $"Data{iteration}";
            var marker = $"default-pipe-reopen-{iteration}";

            var sessionId = await OpenSessionAsync(["session", "create", workbookPath], environmentVariables: null, $"default-{iteration}-create");
            try
            {
                await AssertSuccessAsync(["sheet", "create", "--session", sessionId, "--sheet-name", sheetName], environmentVariables: null, $"default-{iteration}-sheet-create", timeoutMs: 30000);
                await WriteMarkerAsync(sessionId, sheetName, marker, environmentVariables: null, $"default-{iteration}-write");
                await CloseSessionAsync(sessionId, save: true, environmentVariables: null, $"default-{iteration}-close-save");
            }
            catch
            {
                await CloseSessionBestEffortAsync(sessionId, save: false, environmentVariables: null, $"default-{iteration}-best-effort-close-created");
                throw;
            }

            await AssertServiceHealthyAsync(environmentVariables: null, expectedSessionCount: 0, $"default-{iteration}-status-after-close");

            var reopenedSessionId = await OpenSessionAsync(["session", "open", workbookPath], environmentVariables: null, $"default-{iteration}-reopen");
            try
            {
                var persisted = await ReadMarkerAsync(reopenedSessionId, sheetName, environmentVariables: null, $"default-{iteration}-read-after-reopen");
                Assert.Equal(marker, persisted);
                await CloseSessionAsync(reopenedSessionId, save: false, environmentVariables: null, $"default-{iteration}-close-reopened");
            }
            catch
            {
                await CloseSessionBestEffortAsync(reopenedSessionId, save: false, environmentVariables: null, $"default-{iteration}-best-effort-close-reopened");
                throw;
            }

            Assert.True(File.Exists(workbookPath), $"Workbook should remain on disk for diagnostics and reopen validation: {workbookPath}");
            await AssertServiceHealthyAsync(environmentVariables: null, expectedSessionCount: 0, $"default-{iteration}-status-after-reopen-close");
        }
    }

    [Fact(Timeout = 240000)]
    public async Task ConcurrentSessionCreate_OnColdUniquePipe_LeavesSingleResponsiveDaemon()
    {
        var environmentVariables = UniquePipeEnv;
        var createTasks = Enumerable.Range(0, ConcurrentCreateCount)
            .Select(async index =>
            {
                var workbookPath = CreateNewWorkbookPath(nameof(ConcurrentSessionCreate_OnColdUniquePipe_LeavesSingleResponsiveDaemon), index);
                var sessionId = await OpenSessionAsync(["session", "create", workbookPath], environmentVariables, $"concurrent-create-{index}");
                return new CreatedSession(sessionId, workbookPath);
            })
            .ToArray();

        var createdSessions = await Task.WhenAll(createTasks);
        Assert.Equal(ConcurrentCreateCount, createdSessions.Select(session => session.SessionId).Distinct(StringComparer.Ordinal).Count());

        await AssertServiceHealthyAsync(environmentVariables, expectedSessionCount: ConcurrentCreateCount, "concurrent-status-with-sessions");

        foreach (var session in createdSessions)
        {
            await CloseSessionAsync(session.SessionId, save: true, environmentVariables, $"concurrent-close-{session.SessionId}");
            Assert.True(File.Exists(session.WorkbookPath), $"Workbook should exist after close/save: {session.WorkbookPath}");
        }

        await AssertServiceHealthyAsync(environmentVariables, expectedSessionCount: 0, "concurrent-status-after-close");
    }

    private string CreateNewWorkbookPath(string testName, int iteration)
    {
        return Path.Combine(_fixture.TempDir, $"{testName}_{iteration}_{Guid.NewGuid():N}.xlsx");
    }

    private async Task<string> OpenSessionAsync(
        IReadOnlyList<string> args,
        Dictionary<string, string>? environmentVariables,
        string diagnosticLabel)
    {
        var (result, json) = await RunJsonSuccessAsync(args, environmentVariables, diagnosticLabel, timeoutMs: 45000);
        using (json)
        {
            var sessionId = json.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId), $"{diagnosticLabel} did not return a sessionId. Stdout: {result.Stdout}");
            _activeSessions.TryAdd(sessionId!, PipeNameFor(environmentVariables));
            return sessionId!;
        }
    }

    private async Task WriteMarkerAsync(
        string sessionId,
        string sheetName,
        string marker,
        Dictionary<string, string>? environmentVariables,
        string diagnosticLabel)
    {
        var valuesJson = JsonSerializer.Serialize(new[] { new[] { marker } });
        await AssertSuccessAsync(
            ["range", "set-values", "--session", sessionId, "--sheet-name", sheetName, "--range-address", "A1", "--values", valuesJson],
            environmentVariables,
            diagnosticLabel,
            timeoutMs: 30000);
    }

    private async Task<string> ReadMarkerAsync(
        string sessionId,
        string sheetName,
        Dictionary<string, string>? environmentVariables,
        string diagnosticLabel)
    {
        var (_, json) = await RunJsonSuccessAsync(
            ["range", "get-values", "--session", sessionId, "--sheet-name", sheetName, "--range-address", "A1"],
            environmentVariables,
            diagnosticLabel,
            timeoutMs: 30000);

        using (json)
        {
            return json.RootElement.GetProperty("values")[0][0].GetString() ?? string.Empty;
        }
    }

    private async Task CloseSessionAsync(
        string sessionId,
        bool save,
        Dictionary<string, string>? environmentVariables,
        string diagnosticLabel)
    {
        var args = save
            ? ["session", "close", "--session", sessionId, "--save"]
            : new[] { "session", "close", "--session", sessionId };

        await AssertSuccessAsync(args, environmentVariables, diagnosticLabel, timeoutMs: 45000);
        _activeSessions.TryRemove(sessionId, out _);
    }

    private async Task CloseSessionBestEffortAsync(
        string sessionId,
        bool save,
        Dictionary<string, string>? environmentVariables,
        string diagnosticLabel)
    {
        try
        {
            await CloseSessionAsync(sessionId, save, environmentVariables, diagnosticLabel);
        }
        catch (Exception ex)
        {
            _output.WriteLine($"[{diagnosticLabel}] best-effort close failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private async Task AssertServiceHealthyAsync(
        Dictionary<string, string>? environmentVariables,
        int expectedSessionCount,
        string diagnosticLabel)
    {
        var (_, statusJson) = await RunJsonSuccessAsync(["service", "status"], environmentVariables, diagnosticLabel, timeoutMs: 10000);
        using (statusJson)
        {
            Assert.True(statusJson.RootElement.GetProperty("running").GetBoolean(), $"{diagnosticLabel} daemon should be running.");
            Assert.Equal(expectedSessionCount, statusJson.RootElement.GetProperty("sessionCount").GetInt32());
        }
    }

    private async Task AssertSuccessAsync(
        IReadOnlyList<string> args,
        Dictionary<string, string>? environmentVariables,
        string diagnosticLabel,
        int timeoutMs)
    {
        var (_, json) = await RunJsonSuccessAsync(args, environmentVariables, diagnosticLabel, timeoutMs);
        json.Dispose();
    }

    private async Task<(CliResult Result, JsonDocument Json)> RunJsonSuccessAsync(
        IReadOnlyList<string> args,
        Dictionary<string, string>? environmentVariables,
        string diagnosticLabel,
        int timeoutMs)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(args, timeoutMs, environmentVariables, diagnosticLabel);
        _output.WriteLine($"[{diagnosticLabel}] Exit: {result.ExitCode}");
        _output.WriteLine($"[{diagnosticLabel}] Stdout: {result.Stdout}");
        _output.WriteLine($"[{diagnosticLabel}] Stderr: {result.Stderr}");

        Assert.Equal(0, result.ExitCode);
        Assert.DoesNotContain("ConnectionLostException", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("timed out", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
        Assert.True(
            json.RootElement.TryGetProperty("success", out var success) && success.GetBoolean(),
            $"{diagnosticLabel} returned success=false.{Environment.NewLine}{CliProcessHelper.DescribeDaemonState(environmentVariables)}{Environment.NewLine}Stdout: {result.Stdout}{Environment.NewLine}Stderr: {result.Stderr}");

        return (result, json);
    }

    private async Task StopServiceAsync(Dictionary<string, string>? environmentVariables, string diagnosticLabel)
    {
        var result = await CliProcessHelper.RunAsync(["service", "stop"], timeoutMs: 20000, environmentVariables: environmentVariables, diagnosticLabel: diagnosticLabel);
        _output.WriteLine($"[{diagnosticLabel}] Exit: {result.ExitCode}");
        _output.WriteLine($"[{diagnosticLabel}] Stdout: {result.Stdout}");
        _output.WriteLine($"[{diagnosticLabel}] Stderr: {result.Stderr}");
    }

    private static string PipeNameFor(Dictionary<string, string>? environmentVariables)
    {
        return environmentVariables != null && environmentVariables.TryGetValue("EXCELMCP_CLI_PIPE", out var pipeName)
            ? pipeName
            : string.Empty;
    }

    private static Dictionary<string, string>? EnvironmentForPipe(string pipeName)
    {
        return string.IsNullOrEmpty(pipeName)
            ? null
            : new Dictionary<string, string> { ["EXCELMCP_CLI_PIPE"] = pipeName };
    }

    private sealed record CreatedSession(string SessionId, string WorkbookPath);
}
