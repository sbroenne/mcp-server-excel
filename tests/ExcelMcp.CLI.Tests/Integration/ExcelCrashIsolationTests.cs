using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Isolation experiments to determine root cause of Excel crash during
/// Power Query serial workflow. These tests isolate individual variables:
/// <list type="bullet">
/// <item>Single-pass: no prior MashupContainer crashes</item>
/// <item>Two-pass with MashupContainer cleanup between passes</item>
/// <item>Two-pass with connection property diagnostics</item>
/// </list>
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "VerySlow")]
[Trait("RunType", "OnDemand")]
public sealed class ExcelCrashIsolationTests : IAsyncLifetime, IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly CliPowerQueryWorkflowFixture _workflowFixture;
    private string? _activeSessionId;

    public ExcelCrashIsolationTests(ITestOutputHelper output)
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
#pragma warning disable CA1031 // Best-effort cleanup
            try
            {
                await CloseSessionAsync(_activeSessionId, save: false, "cleanup-close");
            }
            catch
            {
            }
#pragma warning restore CA1031

            _activeSessionId = null;
        }

        await CleanupSessionsAndExcelAsync();
    }

    /// <summary>
    /// Experiment 1: Single pass only — no prior MashupContainer crashes.
    /// If Excel still crashes, auto-refresh-on-open alone is the cause.
    /// If Excel survives, MashupContainer residue from a prior pass is needed.
    /// </summary>
    [Fact]
    public async Task Experiment1_SinglePass_NoExcelCrash()
    {
        EnsureNoExcelProcesses();
        KillMashupContainerProcesses("pre-test");

        _output.WriteLine("=== EXPERIMENT 1: Single pass, no prior crashes ===");

        var sessionId = await OpenSessionAsync(_workflowFixture.WorkingCopyPath, "single-pass-open");
        _activeSessionId = sessionId;

        // Log connection properties to diagnose RefreshOnFileOpen/BackgroundQuery state
        await LogConnectionPropertiesAsync(sessionId, "after-open");

        // Wait a moment for any auto-refresh to potentially kill Excel
        await Task.Delay(3000);
        _output.WriteLine("Waited 3s after open — checking if Excel is still alive...");

        // Verify session is still alive by listing sessions
        var (listResult, listJson) = await CliProcessHelper.RunJsonAsync(
            ["session", "list"],
            timeoutMs: 10000,
            diagnosticLabel: "single-pass-check-alive");

        _output.WriteLine($"[session-list] Stdout: {listResult.Stdout}");

        var sessions = listJson.RootElement.GetProperty("sessions");
        Assert.True(sessions.GetArrayLength() > 0, "Session should still be alive after 3s wait");

        // Now attempt a simple operation (range set-values on a safe sheet)
        var setup = _workflowFixture.Definition.ConfigUpdate;
        var (writeResult, _) = await CliProcessHelper.RunJsonAsync(
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
            diagnosticLabel: "single-pass-write");

        _output.WriteLine($"[single-pass-write] Stdout: {writeResult.Stdout}");
        _output.WriteLine($"[single-pass-write] Stderr: {writeResult.Stderr}");
        Assert.True(writeResult.ExitCode == 0, $"Write failed: {writeResult.Stdout}");

        // Now attempt the PQ refreshes
        await RefreshWorkflowAsync(sessionId, "single-pass");

        // Close session
        await CloseSessionAsync(sessionId, save: false, "single-pass-close");
        _activeSessionId = null;
        await WaitForNoExcelProcessesAsync();

        _output.WriteLine("=== EXPERIMENT 1 RESULT: Excel survived single pass ===");
    }

    /// <summary>
    /// Experiment 2: Two passes, killing MashupContainer processes between passes.
    /// If pass 2 survives, MashupContainer residue confirmed as the trigger.
    /// </summary>
    [Fact]
    public async Task Experiment2_TwoPasses_KillMashupContainerBetween_NoExcelCrash()
    {
        EnsureNoExcelProcesses();
        KillMashupContainerProcesses("pre-test");

        _output.WriteLine("=== EXPERIMENT 2: Two passes, kill MashupContainer between ===");

        // Pass 1
        _output.WriteLine("--- Pass 1 ---");
        await RunMinimalPassAsync("pass-1");

        // Kill MashupContainer between passes
        _output.WriteLine("--- Between passes: killing MashupContainer ---");
        KillMashupContainerProcesses("between-passes");
        await Task.Delay(2000); // Allow processes to fully exit

        // Reset workbook
        _workflowFixture.ResetWorkingCopy();

        // Pass 2
        _output.WriteLine("--- Pass 2 ---");
        await RunMinimalPassAsync("pass-2");

        _output.WriteLine("=== EXPERIMENT 2 RESULT: Both passes survived ===");
    }

    /// <summary>
    /// Experiment 3: Two passes WITHOUT killing MashupContainer between passes.
    /// This is the original failing scenario — should now pass with the auto-refresh fix.
    /// </summary>
    [Fact]
    public async Task Experiment3_TwoPasses_NoMashupCleanup_FixApplied()
    {
        EnsureNoExcelProcesses();
        KillMashupContainerProcesses("pre-test");

        _output.WriteLine("=== EXPERIMENT 3: Two passes, NO MashupContainer cleanup (fix applied) ===");

        // Pass 1
        _output.WriteLine("--- Pass 1 ---");
        await RunMinimalPassAsync("pass-1");

        // NO MashupContainer cleanup — this is the original failing scenario
        LogMashupContainerProcesses("between-passes-no-cleanup");

        // Reset workbook
        _workflowFixture.ResetWorkingCopy();

        // Pass 2 — should survive now with auto-refresh suppression fix
        _output.WriteLine("--- Pass 2 (should survive with fix) ---");
        await RunMinimalPassAsync("pass-2");

        _output.WriteLine("=== EXPERIMENT 3 RESULT: Both passes survived with fix ===");
    }

    /// <summary>
    /// Experiment 4: Verify connection properties on real workbook.
    /// Diagnoses whether connections have RefreshOnFileOpen=true and BackgroundQuery=true.
    /// </summary>
    [Fact]
    public async Task Experiment4_DiagnoseConnectionProperties()
    {
        EnsureNoExcelProcesses();
        KillMashupContainerProcesses("pre-test");

        _output.WriteLine("=== EXPERIMENT 4: Connection property diagnosis ===");

        var sessionId = await OpenSessionAsync(_workflowFixture.WorkingCopyPath, "diag-open");
        _activeSessionId = sessionId;

        await LogConnectionPropertiesAsync(sessionId, "after-open-with-fix");

        // Also check MashupContainer processes after open
        LogMashupContainerProcesses("after-open");

        await CloseSessionAsync(sessionId, save: false, "diag-close");
        _activeSessionId = null;

        _output.WriteLine("=== EXPERIMENT 4 COMPLETE ===");
    }

    private async Task RunMinimalPassAsync(string prefix)
    {
        var sessionId = await OpenSessionAsync(_workflowFixture.WorkingCopyPath, $"{prefix}-open");
        _activeSessionId = sessionId;

        try
        {
            // Log connection state
            await LogConnectionPropertiesAsync(sessionId, $"{prefix}-after-open");

            // Wait for potential auto-refresh to settle
            await Task.Delay(2000);

            // Verify Excel is alive
            var (listResult, listJson) = await CliProcessHelper.RunJsonAsync(
                ["session", "list"],
                timeoutMs: 10000,
                diagnosticLabel: $"{prefix}-check-alive");

            _output.WriteLine($"[{prefix}-check-alive] Stdout: {listResult.Stdout}");

            bool sessionAlive = listJson.RootElement
                .GetProperty("sessions")
                .EnumerateArray()
                .Any(s => s.GetProperty("sessionId").GetString() == sessionId);

            Assert.True(sessionAlive,
                $"Excel process died after opening workbook in {prefix}. " +
                "Check MashupContainer crash logs in Windows Event Log.");

            // Apply config write
            var setup = _workflowFixture.Definition.ConfigUpdate;
            var (writeResult, _) = await CliProcessHelper.RunJsonAsync(
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
                diagnosticLabel: $"{prefix}-write");

            _output.WriteLine($"[{prefix}-write] ExitCode: {writeResult.ExitCode}");
            Assert.True(writeResult.ExitCode == 0, $"{prefix}-write failed: {writeResult.Stdout}");

            // Run PQ refreshes (expected to fail - missing data sources)
            await RefreshWorkflowAsync(sessionId, prefix);

            // Close session
            await CloseSessionAsync(sessionId, save: false, $"{prefix}-close");
            _activeSessionId = null;
            await WaitForNoExcelProcessesAsync();

            _output.WriteLine($"--- {prefix} completed successfully ---");
        }
        catch
        {
            // Best-effort cleanup
            if (!string.IsNullOrWhiteSpace(_activeSessionId))
            {
#pragma warning disable CA1031
                try
                {
                    await CloseSessionAsync(_activeSessionId, save: false, $"{prefix}-error-close");
                }
                catch
                {
                }
#pragma warning restore CA1031

                _activeSessionId = null;
            }

            await WaitForNoExcelProcessesAsync();

            throw;
        }
    }

    private async Task RefreshWorkflowAsync(string sessionId, string prefix)
    {
        foreach (var (step, index) in _workflowFixture.Definition.RefreshSequence.Select((value, idx) => (value, idx)))
        {
            var result = await CliProcessHelper.RunAsync(
                ["powerquery", "refresh", "--session", sessionId, "--query-name", step.QueryName],
                timeoutMs: step.TimeoutMs,
                diagnosticLabel: $"{prefix}-refresh-{index + 1}");

            _output.WriteLine($"[{prefix}-refresh-{index + 1}] Exit: {result.ExitCode}, Stdout: {result.Stdout}");

            if (step.ExpectedSuccess == false)
            {
                // Expected failure — just log and continue
                continue;
            }

            Assert.True(result.ExitCode == 0,
                $"{prefix}-refresh-{index + 1} failed unexpectedly: {result.Stdout}");
        }
    }

    private async Task LogConnectionPropertiesAsync(string sessionId, string label)
    {
        var result = await CliProcessHelper.RunAsync(
            ["connection", "list", "--session", sessionId],
            timeoutMs: 30000,
            diagnosticLabel: $"connection-list-{label}");

        _output.WriteLine($"[connection-list-{label}] Stdout: {result.Stdout}");

        if (!string.IsNullOrWhiteSpace(result.Stdout))
        {
            try
            {
                using var doc = JsonDocument.Parse(result.Stdout);
                if (doc.RootElement.TryGetProperty("connections", out var connections))
                {
                    foreach (var conn in connections.EnumerateArray())
                    {
                        string name = conn.TryGetProperty("name", out var n) ? n.GetString() ?? "?" : "?";
                        string type = conn.TryGetProperty("type", out var t) ? t.GetString() ?? "?" : "?";
                        bool bgQuery = conn.TryGetProperty("backgroundQuery", out var bg) && bg.GetBoolean();
                        bool rfoOpen = conn.TryGetProperty("refreshOnFileOpen", out var rfo) && rfo.GetBoolean();
                        bool isPQ = conn.TryGetProperty("isPowerQuery", out var pq) && pq.GetBoolean();

                        _output.WriteLine(
                            $"  [{label}] Connection: {name}, Type: {type}, " +
                            $"BackgroundQuery: {bgQuery}, RefreshOnFileOpen: {rfoOpen}, IsPowerQuery: {isPQ}");
                    }
                }
            }
            catch (JsonException)
            {
                _output.WriteLine($"  [{label}] Could not parse connection list JSON");
            }
        }
    }

    private void KillMashupContainerProcesses(string label)
    {
        var mashupProcesses = Process.GetProcessesByName("Microsoft.Mashup.Container.Loader");
        var mashupNetHostProcesses = Process.GetProcessesByName("Microsoft.Mashup.Container.NetHost");

        _output.WriteLine($"[{label}] MashupContainer.Loader processes: {mashupProcesses.Length}");
        _output.WriteLine($"[{label}] MashupContainer.NetHost processes: {mashupNetHostProcesses.Length}");

        foreach (var p in mashupProcesses.Concat(mashupNetHostProcesses))
        {
            try
            {
                _output.WriteLine($"[{label}] Killing MashupContainer PID {p.Id}");
                p.Kill();
            }
            catch
            {
            }
        }
    }

    private void LogMashupContainerProcesses(string label)
    {
        var mashupProcesses = Process.GetProcessesByName("Microsoft.Mashup.Container.Loader");
        var mashupNetHostProcesses = Process.GetProcessesByName("Microsoft.Mashup.Container.NetHost");

        _output.WriteLine($"[{label}] MashupContainer.Loader processes: {mashupProcesses.Length}");
        _output.WriteLine($"[{label}] MashupContainer.NetHost processes: {mashupNetHostProcesses.Length}");

        foreach (var p in mashupProcesses.Concat(mashupNetHostProcesses))
        {
            _output.WriteLine($"  PID {p.Id}, Started: {p.StartTime:HH:mm:ss}");
        }
    }

    private async Task<string> OpenSessionAsync(string workbookPath, string label)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["session", "open", workbookPath],
            timeoutMs: 30000,
            diagnosticLabel: label);

        _output.WriteLine($"[{label}] Stdout: {result.Stdout}");
        _output.WriteLine($"[{label}] Stderr: {result.Stderr}");

        Assert.True(result.ExitCode == 0, $"{label} failed: {result.Stdout} {result.Stderr}");
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), $"{label} returned success=false");

        return json.RootElement.GetProperty("sessionId").GetString()
            ?? throw new InvalidOperationException($"{label} did not return sessionId");
    }

    private async Task CloseSessionAsync(string sessionId, bool save, string label)
    {
        var (result, _) = await CliProcessHelper.RunJsonAsync(
            ["session", "close", "--session", sessionId, "--save", save ? "true" : "false"],
            timeoutMs: 60000,
            diagnosticLabel: label);

        _output.WriteLine($"[{label}] Stdout: {result.Stdout}");
        _output.WriteLine($"[{label}] Stderr: {result.Stderr}");

        Assert.True(result.ExitCode == 0, $"{label} failed: {result.Stdout} {result.Stderr}");
    }

    private async Task CleanupSessionsAndExcelAsync()
    {
        try
        {
            var (result, json) = await CliProcessHelper.RunJsonAsync(
                ["session", "list"],
                timeoutMs: 10000,
                diagnosticLabel: "crash-cleanup-list");

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
                        await CloseSessionAsync(sessionId, save: false, $"crash-cleanup-close-{sessionId}");
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
        var stopwatch = Stopwatch.StartNew();
        while (stopwatch.ElapsedMilliseconds < timeoutMs)
        {
            if (Process.GetProcessesByName("EXCEL").Length == 0)
            {
                return;
            }

            await Task.Delay(500);
        }
    }

    private static void EnsureNoExcelProcesses()
    {
        var processes = Process.GetProcessesByName("EXCEL");
        foreach (var p in processes)
        {
            try { p.Kill(); } catch { }
        }
    }

    public void Dispose()
    {
        _workflowFixture.Dispose();
        GC.SuppressFinalize(this);
    }
}
