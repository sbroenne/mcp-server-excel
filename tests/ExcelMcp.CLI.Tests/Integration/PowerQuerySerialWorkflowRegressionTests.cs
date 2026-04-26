using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Anonymized CLI regression for the copied-workbook workflow:
/// copy workbook -> write config -> run serial refresh chain -> close/save -> reopen -> verify.
/// 
/// The tracked test code is generic. Workbook-specific names stay in the local-only JSON scenario
/// consumed by <see cref="CliPowerQueryWorkflowFixture" />.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "VerySlow")]
[Trait("RunType", "OnDemand")]
public sealed class PowerQuerySerialWorkflowRegressionTests : IAsyncLifetime, IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly CliPowerQueryWorkflowFixture _workflowFixture;
    private string? _activeSessionId;

    public PowerQuerySerialWorkflowRegressionTests(ITestOutputHelper output)
    {
        _output = output;
        _workflowFixture = new CliPowerQueryWorkflowFixture();
    }

    public async Task InitializeAsync()
    {
        await CleanupSessionsAndExcelAsync();
        _workflowFixture.ResetWorkingCopy();
        _activeSessionId = null;
    }

    public async Task DisposeAsync()
    {
        if (!string.IsNullOrWhiteSpace(_activeSessionId))
        {
#pragma warning disable CA1031 // Best-effort cleanup for failed integration runs
            try
            {
                await CloseSessionAsync(_activeSessionId, save: false, "workflow-cleanup-close");
            }
            catch (Exception ex)
            {
                _output.WriteLine($"Cleanup session close failed: {ex.GetType().Name}");
            }
#pragma warning restore CA1031

            _activeSessionId = null;
        }

        await CleanupSessionsAndExcelAsync();
    }

    [Fact]
    public async Task CopiedWorkbook_TwoSerialWorkflowPasses_StayUsable()
    {
        // Ensure clean state before start
        EnsureNoExcelProcesses();

        try
        {
            await RunWorkflowPassAsync("workflow-pass-1");

            _workflowFixture.ResetWorkingCopy();
            await RunWorkflowPassAsync("workflow-pass-2");
        }
        finally
        {
            await CleanupSessionsAndExcelAsync();
            var excelProcesses = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            if (excelProcesses.Length > 0)
            {
                foreach (var p in excelProcesses)
                {
                    try { p.Kill(); } catch { }
                }
                Assert.Fail($"Leaked {excelProcesses.Length} EXCEL.EXE processes after workflow completion.");
            }
        }
    }

    private static void EnsureNoExcelProcesses()
    {
        var processes = System.Diagnostics.Process.GetProcessesByName("EXCEL");
        foreach (var p in processes)
        {
            try { p.Kill(); } catch { }
        }
    }

    private async Task RunWorkflowPassAsync(string diagnosticPrefix)
    {
        var startingSessionCount = await GetSessionCountAsync();

        _activeSessionId = await OpenSessionAsync(_workflowFixture.WorkingCopyPath, $"{diagnosticPrefix}-open");
        Assert.Equal(startingSessionCount + 1, await GetSessionCountAsync());

        try
        {
            await ApplySetupWriteAsync(_activeSessionId);
            await CloseSessionAsync(_activeSessionId, save: true, $"{diagnosticPrefix}-close-save");
            _activeSessionId = null;
            Assert.Equal(startingSessionCount, await GetSessionCountAsync());
            await WaitForNoExcelProcessesAsync();

            _activeSessionId = await OpenSessionAsync(_workflowFixture.WorkingCopyPath, $"{diagnosticPrefix}-verify-write");
            Assert.Equal(startingSessionCount + 1, await GetSessionCountAsync());

            await AssertPersistedSetupValueAsync(_activeSessionId);
            await RefreshWorkflowAsync(_activeSessionId);

            if (await IsSessionAliveAsync(_activeSessionId))
            {
                await CloseSessionAsync(_activeSessionId, save: false, $"{diagnosticPrefix}-post-refresh-close");
                _activeSessionId = null;
                Assert.Equal(startingSessionCount, await GetSessionCountAsync());
                await WaitForNoExcelProcessesAsync();
            }
            else
            {
                _activeSessionId = null;
            }

            _activeSessionId = await OpenSessionAsync(_workflowFixture.WorkingCopyPath, $"{diagnosticPrefix}-reopen");
            Assert.Equal(startingSessionCount + 1, await GetSessionCountAsync());
            await AssertWorkflowQueriesRemainAvailableAsync(_activeSessionId);

            await CloseSessionAsync(_activeSessionId, save: false, $"{diagnosticPrefix}-reopen-close");
            _activeSessionId = null;
            Assert.Equal(startingSessionCount, await GetSessionCountAsync());
            await WaitForNoExcelProcessesAsync();
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(_activeSessionId))
            {
#pragma warning disable CA1031 // Best-effort cleanup for partially completed pass
                try
                {
                    await CloseSessionAsync(_activeSessionId, save: false, $"{diagnosticPrefix}-best-effort-close");
                }
                catch
                {
                }
#pragma warning restore CA1031

                _activeSessionId = null;
            }

            await WaitForNoExcelProcessesAsync();
        }
    }

    private async Task ApplySetupWriteAsync(string sessionId)
    {
        var setup = _workflowFixture.Definition.ConfigUpdate;
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            [
                "range",
                "set-values",
                "--session",
                sessionId,
                "--sheet-name",
                setup.SheetName,
                "--range-address",
                setup.RangeAddress,
                "--values",
                setup.ValuesJson
            ],
            timeoutMs: setup.TimeoutMs,
            diagnosticLabel: "workflow-setup-write");

        Assert.True(result.ExitCode == 0, $"workflow-setup-write failed. Stdout: {result.Stdout} Stderr: {result.Stderr}");
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), $"workflow-setup-write returned success=false. Stdout: {result.Stdout}");
    }

    private async Task RefreshWorkflowAsync(string sessionId)
    {
        foreach (var (step, index) in _workflowFixture.Definition.RefreshSequence.Select((value, idx) => (value, idx)))
        {
            var result = await CliProcessHelper.RunAsync(
                ["powerquery", "refresh", "--session", sessionId, "--query-name", step.QueryName],
                timeoutMs: step.TimeoutMs,
                diagnosticLabel: $"workflow-refresh-step-{index + 1}");

            JsonDocument? json = null;
            if (!string.IsNullOrWhiteSpace(result.Stdout))
            {
                json = JsonDocument.Parse(result.Stdout);
            }

            bool reportedSuccess = json != null &&
                json.RootElement.TryGetProperty("success", out var successElement) &&
                successElement.GetBoolean();

            _output.WriteLine($"[Step {index + 1}] Stdout: {result.Stdout}");
            _output.WriteLine($"[Step {index + 1}] Stderr: {result.Stderr}");

            if (step.ExpectedSuccess == false)
            {
                Assert.True(
                    result.ExitCode != 0 || !reportedSuccess,
                    $"workflow-refresh-step-{index + 1} was expected to fail but succeeded. Stdout: {result.Stdout}");
                json?.Dispose();
                continue;
            }

            Assert.True(
                result.ExitCode == 0,
                $"workflow-refresh-step-{index + 1} failed. Stdout: {result.Stdout} Stderr: {result.Stderr}");
            Assert.True(
                reportedSuccess,
                $"workflow-refresh-step-{index + 1} returned success=false. Stdout: {result.Stdout}");
            json?.Dispose();
        }
    }

    private async Task AssertPersistedSetupValueAsync(string sessionId)
    {
        var setup = _workflowFixture.Definition.ConfigUpdate;
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            [
                "range",
                "get-values",
                "--session",
                sessionId,
                "--sheet-name",
                setup.SheetName,
                "--range-address",
                setup.RangeAddress
            ],
            timeoutMs: setup.TimeoutMs,
            diagnosticLabel: "workflow-read-setup");

        Assert.True(result.ExitCode == 0, $"workflow-read-setup failed. Stdout: {result.Stdout} Stderr: {result.Stderr}");
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), $"workflow-read-setup returned success=false. Stdout: {result.Stdout}");

        var actualValue = json.RootElement.GetProperty("values")[0][0];
        var expectedValue = CliPowerQueryWorkflowFixture.ExtractFirstScalarValue(setup.ValuesJson);
        AssertJsonScalarEquals(actualValue, expectedValue);
    }

    private async Task AssertWorkflowQueriesRemainAvailableAsync(string sessionId)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["powerquery", "list", "--session", sessionId],
            timeoutMs: 30000,
            diagnosticLabel: "workflow-list-queries");

        _output.WriteLine($"[workflow-list-queries] Stdout: {result.Stdout}");
        _output.WriteLine($"[workflow-list-queries] Stderr: {result.Stderr}");

        Assert.True(result.ExitCode == 0, $"workflow-list-queries failed. Stdout: {result.Stdout} Stderr: {result.Stderr}");
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), $"workflow-list-queries returned success=false. Stdout: {result.Stdout}");

        var queryNames = json.RootElement
            .GetProperty("queries")
            .EnumerateArray()
            .Select(query => query.GetProperty("name").GetString())
            .Where(name => !string.IsNullOrEmpty(name))
            .Cast<string>()
            .ToHashSet(StringComparer.Ordinal);

        foreach (var step in _workflowFixture.Definition.RefreshSequence)
        {
            Assert.Contains(step.QueryName, queryNames);
        }
    }

    private async Task<string> OpenSessionAsync(string workbookPath, string diagnosticLabel)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["session", "open", workbookPath],
            timeoutMs: 30000,
            diagnosticLabel: diagnosticLabel);

        _output.WriteLine($"[{diagnosticLabel}] Stdout: {result.Stdout}");
        _output.WriteLine($"[{diagnosticLabel}] Stderr: {result.Stderr}");

        Assert.True(result.ExitCode == 0, $"{diagnosticLabel} failed. Stdout: {result.Stdout} Stderr: {result.Stderr}");
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), $"{diagnosticLabel} returned success=false. Stdout: {result.Stdout}");

        return json.RootElement.GetProperty("sessionId").GetString()
            ?? throw new InvalidOperationException("session open did not return a sessionId.");
    }

    private async Task CloseSessionAsync(string sessionId, bool save, string diagnosticLabel)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["session", "close", "--session", sessionId, "--save", save ? "true" : "false"],
            timeoutMs: 60000,
            diagnosticLabel: diagnosticLabel);

        _output.WriteLine($"[{diagnosticLabel}] Stdout: {result.Stdout}");
        _output.WriteLine($"[{diagnosticLabel}] Stderr: {result.Stderr}");

        Assert.True(result.ExitCode == 0, $"{diagnosticLabel} failed. Stdout: {result.Stdout} Stderr: {result.Stderr}");
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), $"{diagnosticLabel} returned success=false. Stdout: {result.Stdout}");
    }

    private async Task<int> GetSessionCountAsync()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["session", "list"],
            timeoutMs: 10000,
            diagnosticLabel: "workflow-session-list");

        _output.WriteLine($"[workflow-session-list] Stdout: {result.Stdout}");
        _output.WriteLine($"[workflow-session-list] Stderr: {result.Stderr}");

        Assert.True(result.ExitCode == 0, $"workflow-session-list failed. Stdout: {result.Stdout} Stderr: {result.Stderr}");
        return json.RootElement.GetProperty("sessions").GetArrayLength();
    }

    private async Task<bool> IsSessionAliveAsync(string sessionId)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["session", "list"],
            timeoutMs: 10000,
            diagnosticLabel: "workflow-session-check");

        _output.WriteLine($"[workflow-session-check] Stdout: {result.Stdout}");
        _output.WriteLine($"[workflow-session-check] Stderr: {result.Stderr}");

        Assert.True(result.ExitCode == 0, $"workflow-session-check failed. Stdout: {result.Stdout} Stderr: {result.Stderr}");
        return json.RootElement
            .GetProperty("sessions")
            .EnumerateArray()
            .Any(session => string.Equals(session.GetProperty("sessionId").GetString(), sessionId, StringComparison.Ordinal));
    }

    private async Task CleanupSessionsAndExcelAsync()
    {
        try
        {
            var (result, json) = await CliProcessHelper.RunJsonAsync(
                ["session", "list"],
                timeoutMs: 10000,
                diagnosticLabel: "workflow-cleanup-list");

            if (result.ExitCode == 0 && json.RootElement.TryGetProperty("sessions", out var sessions))
            {
                foreach (var session in sessions.EnumerateArray())
                {
                    if (!session.TryGetProperty("sessionId", out var sessionIdElement))
                    {
                        continue;
                    }

                    var sessionId = sessionIdElement.GetString();
                    if (!string.IsNullOrWhiteSpace(sessionId))
                    {
                        await CloseSessionAsync(sessionId, save: false, $"workflow-cleanup-close-{sessionId}");
                    }
                }
            }
        }
        catch
        {
        }

        EnsureNoExcelProcesses();
        await WaitForNoExcelProcessesAsync();
    }

    private static async Task WaitForNoExcelProcessesAsync(int timeoutMs = 15000)
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        while (stopwatch.ElapsedMilliseconds < timeoutMs)
        {
            if (System.Diagnostics.Process.GetProcessesByName("EXCEL").Length == 0)
            {
                return;
            }

            await Task.Delay(500);
        }
    }

    private static void AssertJsonScalarEquals(JsonElement actual, JsonElement expected)
    {
        Assert.Equal(expected.ValueKind, actual.ValueKind);

        switch (expected.ValueKind)
        {
            case JsonValueKind.String:
                Assert.Equal(expected.GetString(), actual.GetString());
                break;
            case JsonValueKind.Number:
                Assert.Equal(expected.GetDecimal(), actual.GetDecimal());
                break;
            case JsonValueKind.True:
            case JsonValueKind.False:
                Assert.Equal(expected.GetBoolean(), actual.GetBoolean());
                break;
            case JsonValueKind.Null:
                Assert.Equal(JsonValueKind.Null, actual.ValueKind);
                break;
            default:
                throw new NotSupportedException($"Unsupported scalar comparison for {expected.ValueKind}.");
        }
    }

    public void Dispose()
    {
        _workflowFixture.Dispose();
        GC.SuppressFinalize(this);
    }
}
