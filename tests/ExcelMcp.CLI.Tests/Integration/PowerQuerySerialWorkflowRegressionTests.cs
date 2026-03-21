using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// REGRESSION TESTS for Bug Report (v1.8.32): Serial Power Query workflow via CLI daemon.
///
/// BUG PATTERN FROM REPORT (CLI-specific):
/// - Multiple excelcli commands using --session flag (daemon mode)
/// - Serial Power Query refresh operations on one session
/// - One refresh times out or runs long
/// - Later operations on same session fail or hang
/// - Session close hangs
/// - Excel.exe and excelcli daemon remain running
///
/// DIFFERENCE FROM ComInterop control tests:
/// - ComInterop tests: Direct ExcelBatch/SessionManager API (all PASSED - controls)
/// - These tests: CLI daemon with pipe communication (where bug likely reproduces)
/// - Focus: Daemon-level session reuse, pipe communication, process cleanup
///
/// HYPOTHESIS: The bug is in the daemon/pipe layer, not in ComInterop layer.
/// If these tests FAIL (RED), we've found the real repro layer.
/// </summary>
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "VerySlow")]
[Trait("RunType", "OnDemand")]
[Collection("Service")]
public class PowerQuerySerialWorkflowRegressionTests : IClassFixture<ServiceFixture>, IAsyncLifetime
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private string? _testFile;
    private string? _sessionId;

    public PowerQuerySerialWorkflowRegressionTests(ServiceFixture serviceFixture, ITestOutputHelper output)
    {
        _ = serviceFixture; // Force fixture initialization
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"pq-cli-serial-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public async Task InitializeAsync()
    {
        // ServiceFixture auto-starts the daemon

        // Create test workbook manually - no CLI file create command
        _testFile = Path.Combine(_tempDir, $"pq-serial-test-{Guid.NewGuid():N}.xlsx");

        // Create empty workbook using Core API directly for test setup
        Sbroenne.ExcelMcp.ComInterop.Session.ExcelSession.CreateNew<int>(
            _testFile,
            isMacroEnabled: false,
            (ctx, ct) =>
            {
                // No-op - just need the file created
                return 0;
            });

        _output.WriteLine($"=== Created test file: {_testFile} ===");

        // Now use CLI to open session on existing file
        _output.WriteLine($"=== Opening session on: {_testFile} ===");
        var openResult = await CliProcessHelper.RunAsync($"session open \"{_testFile}\"", timeoutMs: 30000);
        _output.WriteLine($"Open stdout: {openResult.Stdout}");
        _output.WriteLine($"Open stderr: {openResult.Stderr}");
        _output.WriteLine($"Open exit code: {openResult.ExitCode}");

        if (openResult.ExitCode != 0)
        {
            throw new InvalidOperationException($"Failed to open session. Exit code: {openResult.ExitCode}, Output: {openResult.Stdout}, Error: {openResult.Stderr}");
        }

        var openJson = System.Text.Json.JsonDocument.Parse(openResult.Stdout);
        _sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
        Assert.NotNull(_sessionId);

        // Add a simple Power Query (List.Generate with 5000 rows to make refresh take a few seconds)
        string mCode = """
            let
                Source = List.Generate(
                    () => [i = 0],
                    each [i] < 5000,
                    each [i = [i] + 1],
                    each [ID = [i], Name = "Item_" & Text.From([i])]
                ),
                AsTable = Table.FromRecords(Source)
            in
                AsTable
            """;

        var createQueryResult = await CliProcessHelper.RunAsync(
            $"powerquery create --session {_sessionId} --query-name TestQuery --load-to Sheet1!A1 --m-code \"{mCode.Replace("\"", "\"\"")}\"",
            timeoutMs: 60000);

        _output.WriteLine($"Create query stdout: {createQueryResult.Stdout}");
        _output.WriteLine($"Create query stderr: {createQueryResult.Stderr}");
        _output.WriteLine($"Create query exit code: {createQueryResult.ExitCode}");

        if (createQueryResult.ExitCode != 0)
        {
            throw new InvalidOperationException($"Failed to create Power Query. Exit code: {createQueryResult.ExitCode}, Output: {createQueryResult.Stdout}, Error: {createQueryResult.Stderr}");
        }
    }

    public async Task DisposeAsync()
    {
        // Best-effort cleanup: close session if it still exists
        if (!string.IsNullOrEmpty(_sessionId))
        {
#pragma warning disable CA1031 // Intentional: best-effort cleanup
            try
            {
                await CliProcessHelper.RunAsync($"session close --session {_sessionId}", timeoutMs: 30000);
            }
            catch (Exception ex)
            {
                _output.WriteLine($"Best-effort session close failed: {ex.Message}");
            }
#pragma warning restore CA1031
        }

        // Delete test file and temp directory
        if (_testFile != null && File.Exists(_testFile))
        {
#pragma warning disable CA1031
            try { File.Delete(_testFile); } catch (Exception) { /* best effort */ }
#pragma warning restore CA1031
        }

        if (Directory.Exists(_tempDir))
        {
#pragma warning disable CA1031
            try { Directory.Delete(_tempDir, recursive: true); } catch (Exception) { /* best effort */ }
#pragma warning restore CA1031
        }
    }

    /// <summary>
    /// REGRESSION TEST: Serial Power Query refreshes via CLI daemon.
    ///
    /// WORKFLOW (models bug report exactly):
    /// 1. session open (done in InitializeAsync)
    /// 2. powerquery refresh --session (operation A - should succeed)
    /// 3. powerquery refresh --session (operation B - should succeed)
    /// 4. powerquery refresh --session (operation C - should succeed)
    /// 5. session close --session (should succeed)
    ///
    /// EXPECTED (if bug exists): Test FAILS because:
    /// - One of the refreshes hangs/times out
    /// - Later refreshes fail or hang
    /// - Session close hangs
    /// - Excel process doesn't get cleaned up
    ///
    /// EXPECTED (if bug fixed): Test PASSES because:
    /// - All refreshes complete successfully
    /// - Session closes cleanly
    /// - Excel process is cleaned up
    /// </summary>
    [Fact]
    public async Task SerialRefreshes_ViaCliDaemon_AllSucceed()
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();

        // Operation A: First refresh
        _output.WriteLine("=== Operation A: First refresh ===");
        var (refresh1Result, refresh1Json) = await CliProcessHelper.RunJsonAsync(
            $"powerquery refresh --session {_sessionId} --query-name TestQuery",
            timeoutMs: 90000);

        _output.WriteLine($"Refresh 1 stdout: {refresh1Result.Stdout}");
        _output.WriteLine($"Refresh 1 exit code: {refresh1Result.ExitCode}");
        _output.WriteLine($"Refresh 1 elapsed: {sw.Elapsed.TotalSeconds:F1}s");

        Assert.Equal(0, refresh1Result.ExitCode);
        var success1 = refresh1Json.RootElement.GetProperty("success").GetBoolean();
        Assert.True(success1, "First refresh should succeed");

        sw.Restart();

        // Operation B: Second refresh (this is where bug report shows poisoning)
        _output.WriteLine("=== Operation B: Second refresh ===");
        var (refresh2Result, refresh2Json) = await CliProcessHelper.RunJsonAsync(
            $"powerquery refresh --session {_sessionId} --query-name TestQuery",
            timeoutMs: 90000);

        _output.WriteLine($"Refresh 2 stdout: {refresh2Result.Stdout}");
        _output.WriteLine($"Refresh 2 exit code: {refresh2Result.ExitCode}");
        _output.WriteLine($"Refresh 2 elapsed: {sw.Elapsed.TotalSeconds:F1}s");

        Assert.Equal(0, refresh2Result.ExitCode);
        var success2 = refresh2Json.RootElement.GetProperty("success").GetBoolean();
        Assert.True(success2, "Second refresh should succeed (bug report suggests this might fail)");

        sw.Restart();

        // Operation C: Third refresh
        _output.WriteLine("=== Operation C: Third refresh ===");
        var (refresh3Result, refresh3Json) = await CliProcessHelper.RunJsonAsync(
            $"powerquery refresh --session {_sessionId} --query-name TestQuery",
            timeoutMs: 90000);

        _output.WriteLine($"Refresh 3 stdout: {refresh3Result.Stdout}");
        _output.WriteLine($"Refresh 3 exit code: {refresh3Result.ExitCode}");
        _output.WriteLine($"Refresh 3 elapsed: {sw.Elapsed.TotalSeconds:F1}s");

        Assert.Equal(0, refresh3Result.ExitCode);
        var success3 = refresh3Json.RootElement.GetProperty("success").GetBoolean();
        Assert.True(success3, "Third refresh should succeed");

        sw.Restart();

        // Close session
        _output.WriteLine("=== Session close ===");
        var (closeResult, closeJson) = await CliProcessHelper.RunJsonAsync(
            $"session close --session {_sessionId}",
            timeoutMs: 60000);

        _output.WriteLine($"Close stdout: {closeResult.Stdout}");
        _output.WriteLine($"Close exit code: {closeResult.ExitCode}");
        _output.WriteLine($"Close elapsed: {sw.Elapsed.TotalSeconds:F1}s");

        Assert.Equal(0, closeResult.ExitCode);
        var closeSuccess = closeJson.RootElement.GetProperty("success").GetBoolean();
        Assert.True(closeSuccess, "Session close should succeed (bug report suggests this might hang)");

        // Verify session is gone
        var (listResult, listJson) = await CliProcessHelper.RunJsonAsync("session list", timeoutMs: 10000);
        var sessions = listJson.RootElement.GetProperty("sessions").EnumerateArray().ToList();
        Assert.DoesNotContain(sessions, s => s.GetProperty("sessionId").GetString() == _sessionId);
    }

    /// <summary>
    /// REGRESSION TEST: Reopen same file immediately after serial workflow.
    ///
    /// Bug report mentions: "New session creation should not be blocked by stale dead sessions"
    ///
    /// WORKFLOW:
    /// 1. Serial refreshes (like test above)
    /// 2. Close session
    /// 3. Immediately reopen same file
    ///
    /// EXPECTED (if bug exists): Test FAILS because:
    /// - Reopen hangs or fails
    /// - Stale session metadata blocks new session
    ///
    /// EXPECTED (if bug fixed): Test PASSES because:
    /// - Reopen succeeds immediately
    /// </summary>
    [Fact]
    public async Task ReopenSameFile_AfterSerialRefreshes_Succeeds()
    {
        // Do serial refreshes first
        for (int i = 0; i < 3; i++)
        {
            var (refreshResult, _) = await CliProcessHelper.RunJsonAsync(
                $"powerquery refresh --session {_sessionId} --query-name TestQuery",
                timeoutMs: 90000);
            Assert.Equal(0, refreshResult.ExitCode);
        }

        // Close session
        var (closeResult, _) = await CliProcessHelper.RunJsonAsync(
            $"session close --session {_sessionId}",
            timeoutMs: 60000);
        Assert.Equal(0, closeResult.ExitCode);

        // Immediately reopen same file
        _output.WriteLine("=== Reopen same file ===");
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var (reopenResult, reopenJson) = await CliProcessHelper.RunJsonAsync(
            $"session open \"{_testFile}\"",
            timeoutMs: 60000);

        _output.WriteLine($"Reopen stdout: {reopenResult.Stdout}");
        _output.WriteLine($"Reopen exit code: {reopenResult.ExitCode}");
        _output.WriteLine($"Reopen elapsed: {sw.Elapsed.TotalSeconds:F1}s");

        Assert.Equal(0, reopenResult.ExitCode);
        var newSessionId = reopenJson.RootElement.GetProperty("sessionId").GetString();
        Assert.NotNull(newSessionId);
        Assert.NotEqual(_sessionId, newSessionId);

        // Verify new session works
        var (listResult, _) = await CliProcessHelper.RunJsonAsync(
            $"powerquery list --session {newSessionId}",
            timeoutMs: 30000);
        Assert.Equal(0, listResult.ExitCode);

        // Close new session
        await CliProcessHelper.RunAsync($"session close --session {newSessionId}", timeoutMs: 60000);
    }

    /// <summary>
    /// REGRESSION TEST: With simulated long-running refresh (large dataset).
    ///
    /// This test uses a 20K row Power Query to make refresh take longer,
    /// increasing the chance of exposing timing-related bugs.
    /// </summary>
    [Fact]
    public async Task LargerDataset_SerialRefreshes_StillSucceeds()
    {
        // Create a larger Power Query (20K rows)
        string largeMCode = """
            let
                Source = List.Generate(
                    () => [i = 0],
                    each [i] < 20000,
                    each [i = [i] + 1],
                    each [
                        ID = [i],
                        Name = "Item_" & Text.From([i]),
                        Value = Number.Round(Number.RandomBetween(1, 10000), 2),
                        Category = if Number.Mod([i], 3) = 0 then "A" else if Number.Mod([i], 3) = 1 then "B" else "C"
                    ]
                ),
                AsTable = Table.FromRecords(Source)
            in
                AsTable
            """;

        var (createLargeQueryResult, _) = await CliProcessHelper.RunJsonAsync(
            $"powerquery create --session {_sessionId} --query-name LargeQuery --load-to Sheet2!A1 --m-code \"{largeMCode.Replace("\"", "\"\"")}\"",
            timeoutMs: 120000);

        Assert.Equal(0, createLargeQueryResult.ExitCode);

        // Do multiple refreshes
        var sw = System.Diagnostics.Stopwatch.StartNew();
        for (int i = 0; i < 2; i++)
        {
            _output.WriteLine($"=== Large dataset refresh {i + 1} ===");
            sw.Restart();

            var (refreshResult, refreshJson) = await CliProcessHelper.RunJsonAsync(
                $"powerquery refresh --session {_sessionId} --query-name LargeQuery",
                timeoutMs: 180000); // 3 minutes for large dataset

            _output.WriteLine($"Refresh {i + 1} elapsed: {sw.Elapsed.TotalSeconds:F1}s");
            _output.WriteLine($"Refresh {i + 1} exit code: {refreshResult.ExitCode}");

            Assert.Equal(0, refreshResult.ExitCode);
            var success = refreshJson.RootElement.GetProperty("success").GetBoolean();
            Assert.True(success, $"Large dataset refresh {i + 1} should succeed");
        }
    }
}
