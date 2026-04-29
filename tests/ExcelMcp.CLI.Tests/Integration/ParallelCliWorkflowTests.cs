using System.Collections.Concurrent;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// End-to-end CLI stress coverage for multiple independent workbook workflows
/// running through the same excelcli daemon at the same time.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Slow")]
public sealed class ParallelCliWorkflowTests : IAsyncLifetime, IClassFixture<TempDirectoryFixture>
{
    private const int WorkflowCount = 4;

    private readonly ITestOutputHelper _output;
    private readonly string _pipeName = $"excelmcp-parallel-e2e-{Guid.NewGuid():N}";
    private readonly string _tempDir;
    private readonly ConcurrentDictionary<string, byte> _activeSessions = new(StringComparer.Ordinal);

    public ParallelCliWorkflowTests(TempDirectoryFixture fixture, ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(fixture.TempDir, $"ParallelCliWorkflow_{Guid.NewGuid():N}");
    }

    private Dictionary<string, string> TestEnv => new() { ["EXCELMCP_CLI_PIPE"] = _pipeName };

    public async Task InitializeAsync()
    {
        Directory.CreateDirectory(_tempDir);
        await StopServiceAsync("initialize-stop");
    }

    public async Task DisposeAsync()
    {
        foreach (var sessionId in _activeSessions.Keys)
        {
            await CloseSessionBestEffortAsync(sessionId, "cleanup-close");
        }

        await StopServiceAsync("cleanup-stop");

        try
        {
            if (Directory.Exists(_tempDir))
                Directory.Delete(_tempDir, recursive: true);
        }
        catch
        {
            // Best-effort cleanup for integration-test temp files.
        }
    }

    [Fact(Timeout = 180000)]
    public async Task ParallelMultiFileWorkflows_StayIsolatedAndLeaveNoSessions()
    {
        await AssertSuccessAsync(["service", "start"], "service-start", timeoutMs: 20000);

        var workflows = Enumerable.Range(0, WorkflowCount)
            .Select(RunWorkbookWorkflowAsync)
            .ToArray();

        var results = await Task.WhenAll(workflows);

        Assert.Equal(WorkflowCount, results.Length);
        Assert.Equal(WorkflowCount, results.Select(r => r.FilePath).Distinct(StringComparer.OrdinalIgnoreCase).Count());

        foreach (var result in results)
        {
            Assert.True(File.Exists(result.FilePath), $"Expected workbook to exist: {result.FilePath}");
            Assert.Equal($"workflow-{result.Index}", result.PersistedValue);
        }

        var (_, listJson) = await RunJsonSuccessAsync(["session", "list"], "final-session-list", timeoutMs: 10000);
        using (listJson)
        {
            Assert.Equal(0, listJson.RootElement.GetProperty("sessions").GetArrayLength());
        }

        var (_, statusJson) = await RunJsonSuccessAsync(["service", "status"], "final-service-status", timeoutMs: 10000);
        using (statusJson)
        {
            Assert.True(statusJson.RootElement.GetProperty("running").GetBoolean());
            Assert.Equal(0, statusJson.RootElement.GetProperty("sessionCount").GetInt32());
        }
    }

    private async Task<WorkflowResult> RunWorkbookWorkflowAsync(int index)
    {
        var filePath = Path.Combine(_tempDir, $"parallel-{index}.xlsx");
        var sheetName = $"Data{index}";
        var marker = $"workflow-{index}";
        var valuesJson = JsonSerializer.Serialize(new[] { new[] { marker }, new[] { $"file-{index}" } });

        var sessionId = await OpenSessionAsync(["session", "create", filePath], $"workflow-{index}-create");
        try
        {
            await AssertSuccessAsync(["sheet", "create", "--session", sessionId, "--sheet-name", sheetName], $"workflow-{index}-sheet-create", timeoutMs: 30000);
            await AssertSuccessAsync(
                ["range", "set-values", "--session", sessionId, "--sheet-name", sheetName, "--range-address", "A1:A2", "--values", valuesJson],
                $"workflow-{index}-write",
                timeoutMs: 30000);
            await AssertCellValueAsync(sessionId, sheetName, "A1", marker, $"workflow-{index}-read-before-close");
            await AssertSuccessAsync(
                ["rangeformat", "format-range", "--session", sessionId, "--sheet-name", sheetName, "--range-address", "A1:A2", "--bold", "true"],
                $"workflow-{index}-format",
                timeoutMs: 30000);
            await CloseSessionAsync(sessionId, save: true, $"workflow-{index}-close-save");
        }
        catch
        {
            await CloseSessionBestEffortAsync(sessionId, $"workflow-{index}-best-effort-close");
            throw;
        }

        var reopenedSessionId = await OpenSessionAsync(["session", "open", filePath], $"workflow-{index}-reopen");
        try
        {
            var persisted = await ReadCellValueAsync(reopenedSessionId, sheetName, "A1", $"workflow-{index}-read-after-reopen");
            await CloseSessionAsync(reopenedSessionId, save: false, $"workflow-{index}-close-reopened");
            return new WorkflowResult(index, filePath, persisted);
        }
        catch
        {
            await CloseSessionBestEffortAsync(reopenedSessionId, $"workflow-{index}-best-effort-close-reopened");
            throw;
        }
    }

    private async Task<string> OpenSessionAsync(IReadOnlyList<string> args, string diagnosticLabel)
    {
        var (result, json) = await RunJsonSuccessAsync(args, diagnosticLabel, timeoutMs: 30000);
        using (json)
        {
            var sessionId = json.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId), $"{diagnosticLabel} did not return a sessionId. Stdout: {result.Stdout}");
            _activeSessions.TryAdd(sessionId!, 0);
            return sessionId!;
        }
    }

    private async Task CloseSessionAsync(string sessionId, bool save, string diagnosticLabel)
    {
        var args = save
            ? ["session", "close", "--session", sessionId, "--save"]
            : new[] { "session", "close", "--session", sessionId };

        await AssertSuccessAsync(args, diagnosticLabel, timeoutMs: 30000);
        _activeSessions.TryRemove(sessionId, out _);
    }

    private async Task CloseSessionBestEffortAsync(string sessionId, string diagnosticLabel)
    {
        try
        {
            await CloseSessionAsync(sessionId, save: false, diagnosticLabel);
        }
        catch (Exception ex)
        {
            _output.WriteLine($"[{diagnosticLabel}] best-effort close failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private async Task AssertCellValueAsync(string sessionId, string sheetName, string rangeAddress, string expected, string diagnosticLabel)
    {
        var actual = await ReadCellValueAsync(sessionId, sheetName, rangeAddress, diagnosticLabel);
        Assert.Equal(expected, actual);
    }

    private async Task<string> ReadCellValueAsync(string sessionId, string sheetName, string rangeAddress, string diagnosticLabel)
    {
        var (_, json) = await RunJsonSuccessAsync(
            ["range", "get-values", "--session", sessionId, "--sheet-name", sheetName, "--range-address", rangeAddress],
            diagnosticLabel,
            timeoutMs: 30000);

        using (json)
        {
            return json.RootElement.GetProperty("values")[0][0].GetString() ?? string.Empty;
        }
    }

    private async Task AssertSuccessAsync(IReadOnlyList<string> args, string diagnosticLabel, int timeoutMs)
    {
        var (_, json) = await RunJsonSuccessAsync(args, diagnosticLabel, timeoutMs);
        json.Dispose();
    }

    private async Task<(CliResult Result, JsonDocument Json)> RunJsonSuccessAsync(
        IReadOnlyList<string> args,
        string diagnosticLabel,
        int timeoutMs)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(args, timeoutMs, TestEnv, diagnosticLabel);
        _output.WriteLine($"[{diagnosticLabel}] Exit: {result.ExitCode}");
        _output.WriteLine($"[{diagnosticLabel}] Stdout: {result.Stdout}");
        _output.WriteLine($"[{diagnosticLabel}] Stderr: {result.Stderr}");

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), $"{diagnosticLabel} returned success=false. Stdout: {result.Stdout}");

        return (result, json);
    }

    private async Task StopServiceAsync(string diagnosticLabel)
    {
        var result = await CliProcessHelper.RunAsync(["service", "stop"], timeoutMs: 20000, environmentVariables: TestEnv, diagnosticLabel: diagnosticLabel);
        _output.WriteLine($"[{diagnosticLabel}] Exit: {result.ExitCode}");
        _output.WriteLine($"[{diagnosticLabel}] Stdout: {result.Stdout}");
        _output.WriteLine($"[{diagnosticLabel}] Stderr: {result.Stderr}");
    }

    private sealed record WorkflowResult(int Index, string FilePath, string PersistedValue);
}
