using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// REGRESSION TESTS for Bug Report (v1.8.32): Serial Power Query workflow leaves sessions wedged.
///
/// BUG PATTERN FROM REPORT:
/// - Multiple Power Query refreshes in sequence on one session
/// - One refresh times out or is cancelled in the middle
/// - Later operations on the same session fail or hang
/// - Session close/reopen on same workbook fails or hangs
/// - EXCEL.EXE and excelcli remain running until manually killed
///
/// These tests model the ACTUAL reported workflow, not just isolated timeout smoke tests.
/// The existing timeout tests prove timeout detection and cleanup work in isolation.
/// These tests prove whether the SERIAL RECOVERY PATH works across realistic multi-operation workflows.
///
/// TEST STRATEGY:
/// - Operation A (success) → Operation B (timeout) → Operation C (attempt)
/// - Verify C's behavior: does it fail fast? does it hang? is error message useful?
/// - Verify cleanup: does session close cleanly? can we reopen the same file?
/// - Verify process cleanup: does Excel process get killed? any leaks?
///
/// Follow-up callbacks must not execute after timeout. Disposal must release
/// the owned process, and reopening must return the same usable workbook.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelBatch")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")]
[Trait("RequiresExcel", "true")]
public class ExcelBatchSerialWorkflowTests : IAsyncLifetime, IDisposable
{
    private readonly ITestOutputHelper _output;
    private static string? _staticTestFile;
    private string? _testFileCopy;
    private readonly OwnedExcelProcessScope _owned = new();

    public ExcelBatchSerialWorkflowTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public Task InitializeAsync()
    {
        if (_staticTestFile == null)
        {
            var testFolder = Path.Join(AppContext.BaseDirectory, "Integration", "Session", "TestFiles");
            _staticTestFile = Path.Join(testFolder, "batch-test-static.xlsx");

            if (!File.Exists(_staticTestFile))
            {
                throw new FileNotFoundException($"Static test file not found at {_staticTestFile}.");
            }
        }

        _testFileCopy = Path.Join(Path.GetTempPath(), $"batch-serial-{Guid.NewGuid():N}.xlsx");
        File.Copy(_staticTestFile, _testFileCopy, overwrite: true);

        return Task.Delay(500);
    }

    public Task DisposeAsync()
    {
        Dispose();
        return Task.CompletedTask;
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);
        SessionTestCleanup.AssertExitedAndDelete(_owned,
            _testFileCopy is null ? [] : [_testFileCopy]);
    }

    /// <summary>
    /// REGRESSION TEST: Models the core bug pattern.
    /// 
    /// WORKFLOW:
    /// 1. Operation A: Quick read (success)
    /// 2. Operation B: Long operation that times out
    /// 3. Operation C: Another quick read (should fail fast with useful message)
    /// 
    /// EXPECTED (if bug exists): Test FAILS because:
    /// - Operation C hangs or takes too long (session poisoning)
    /// - Error message is not helpful for recovery
    /// - Dispose hangs or leaks Excel process
    /// 
    /// EXPECTED (if bug fixed): Test PASSES because:
    /// - Operation C fails FAST (under 1s) with "previous operation timed out" message
    /// - Dispose completes quickly (under 30s)
    /// - Excel process is killed
    /// </summary>
    [Fact]
    public void SerialWorkflow_TimeoutInMiddle_LaterOperationsFailFast()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatchWithTimeouts(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(3),
            startupTimeout: ComInteropConstants.DefaultOperationTimeout,
            _testFileCopy!);

        int? excelPid = batch.ExcelProcessId;
        _output.WriteLine($"Started batch, Excel PID: {excelPid}");

        // Operation A: Success
        _output.WriteLine("Operation A: Quick read (should succeed)");
        var operationASw = Stopwatch.StartNew();
        SessionWorkbookAssertions.AssertIdentity(batch, _testFileCopy!);
        operationASw.Stop();
        _output.WriteLine($"Operation A verified the workbook in {operationASw.Elapsed.TotalMilliseconds:F0}ms.");

        // Operation B: Timeout
        _output.WriteLine("Operation B: Long operation (should timeout)");
        var operationBSw = Stopwatch.StartNew();
        var timeoutException = Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });
        operationBSw.Stop();
        _output.WriteLine($"  ✓ Operation B timed out in {operationBSw.Elapsed.TotalSeconds:F1}s: {timeoutException.Message}");

        // Operation C: Should fail FAST, not hang
        _output.WriteLine("Operation C: Another quick read (should fail FAST)");
        var operationCSw = Stopwatch.StartNew();
        var rejectedCallbackRan = false;
        var operationCException = Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                rejectedCallbackRan = true;
                return 42;
            });
        });
        operationCSw.Stop();

        _output.WriteLine($"  Operation C threw in {operationCSw.Elapsed.TotalMilliseconds:F0}ms: {operationCException.Message}");

        // REGRESSION ASSERTION: Operation C must fail fast (< 1 second)
        // BUG SYMPTOM: If this fails, operation C is hanging/poisoning the session
        Assert.True(operationCSw.Elapsed < TimeSpan.FromSeconds(1),
            $"REGRESSION: Operation C took {operationCSw.Elapsed.TotalSeconds:F1}s after earlier timeout. " +
            "Expected < 1s fail-fast. This indicates session poisoning — later operations are not failing fast.");

        // Error message should guide recovery
        Assert.Contains("previous operation", operationCException.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(rejectedCallbackRan);
        Assert.True(batch.HasTimedOutOperation);

        // Dispose should complete quickly
        _output.WriteLine("Disposing batch...");
        var disposeSw = Stopwatch.StartNew();
        batch.Dispose();
        disposeSw.Stop();
        _output.WriteLine($"  Dispose completed in {disposeSw.Elapsed.TotalSeconds:F1}s");

        Assert.True(disposeSw.Elapsed < TimeSpan.FromSeconds(30),
            $"REGRESSION: Dispose took {disposeSw.Elapsed.TotalSeconds:F1}s — expected < 30s");

        Assert.NotNull(excelPid);
        _owned.AssertAllExited();

        _output.WriteLine("✓ Serial workflow test passed: later operations failed fast, dispose cleaned up");
    }

    /// <summary>
    /// REGRESSION TEST: After timeout, can we immediately reopen the same workbook in a new session?
    ///
    /// WORKFLOW:
    /// 1. Session A: Open workbook → timeout
    /// 2. Dispose Session A
    /// 3. Session B: Reopen SAME workbook immediately
    /// 4. Verify Session B works
    ///
    /// EXPECTED (if bug exists): Test FAILS because:
    /// - Session B creation hangs or times out
    /// - Reopen fails with "file in use" or COM errors
    /// - Stale Session A state blocks Session B
    ///
    /// EXPECTED (if bug fixed): Test PASSES because:
    /// - Session B opens quickly (under 10s)
    /// - Session B operations succeed
    /// - No file locking or stale session interference
    /// </summary>
    [Fact]
    public void SerialWorkflow_ReopenAfterTimeout_NewSessionWorks()
    {
        int? sessionAPid = null;
        int? sessionBPid = null;
        using var owned = new OwnedExcelProcessScope();

        try
        {
            // Session A: Timeout
            _output.WriteLine("Session A: Opening workbook (will timeout)");
            using var sessionA = ExcelSession.BeginBatchWithTimeouts(
                show: false,
                operationTimeout: TimeSpan.FromSeconds(3),
                startupTimeout: ComInteropConstants.DefaultOperationTimeout,
                _testFileCopy!);
            sessionAPid = sessionA.ExcelProcessId;
            _output.WriteLine($"  Session A started, Excel PID: {sessionAPid}");

            // Quick warmup
            SessionWorkbookAssertions.AssertIdentity(sessionA, _testFileCopy!);

            // Trigger timeout
            _output.WriteLine("  Triggering timeout in Session A...");
            Assert.Throws<TimeoutException>(() =>
            {
                sessionA.Execute((ctx, ct) =>
                {
                    Thread.Sleep(TimeSpan.FromSeconds(30));
                    return 0;
                });
            });

            // Dispose Session A
            _output.WriteLine("  Disposing Session A...");
            var disposeASw = Stopwatch.StartNew();
            sessionA.Dispose();
            disposeASw.Stop();
            _output.WriteLine($"  ✓ Session A disposed in {disposeASw.Elapsed.TotalSeconds:F1}s");

            owned.AssertAllExited();

            // Session B: Reopen same file
            _output.WriteLine("Session B: Reopening SAME workbook immediately");
            var sessionBSw = Stopwatch.StartNew();
            using var sessionB = ExcelSession.BeginBatch(
                show: false,
                operationTimeout: TimeSpan.FromSeconds(30),
                _testFileCopy!);
            sessionBSw.Stop();
            sessionBPid = sessionB.ExcelProcessId;
            _output.WriteLine($"  Session B opened in {sessionBSw.Elapsed.TotalSeconds:F1}s, Excel PID: {sessionBPid}");

            // REGRESSION ASSERTION: Reopen should be fast (< 10 seconds)
            // BUG SYMPTOM: If this fails, stale Session A is blocking reopen
            Assert.True(sessionBSw.Elapsed < TimeSpan.FromSeconds(10),
                $"REGRESSION: Reopening same file after timeout took {sessionBSw.Elapsed.TotalSeconds:F1}s. " +
                "Expected < 10s. This indicates stale session state or file locking from Session A.");

            // Verify Session B actually works
            _output.WriteLine("  Verifying Session B operations work...");
            SessionWorkbookAssertions.AssertIdentity(sessionB, _testFileCopy!);
            SessionWorkbookAssertions.WriteMarker(sessionB, "reopened-batch");
            Assert.Equal("reopened-batch", SessionWorkbookAssertions.ReadMarker(sessionB));

            // Clean up Session B
            sessionB.Dispose();
            owned.AssertAllExited();

            _output.WriteLine("✓ Reopen test passed: new session opened quickly and worked after timeout cleanup");
        }
        finally
        {
            owned.AssertAllExited();
        }
    }

    /// <summary>
    /// REGRESSION TEST: Multiple timeouts in sequence (A timeout → B timeout → C attempt).
    ///
    /// WORKFLOW:
    /// 1. Operation A: Timeout
    /// 2. Operation B: Timeout (again on same session)
    /// 3. Operation C: Quick read (should still fail fast)
    ///
    /// EXPECTED (if bug exists): Test FAILS because:
    /// - Multiple timeouts poison session worse
    /// - Operation C hangs or shows degraded behavior
    ///
    /// EXPECTED (if bug fixed): Test PASSES because:
    /// - All operations after first timeout fail fast
    /// - Error messages stay consistent
    /// </summary>
    [Fact]
    public void SerialWorkflow_MultipleTimeouts_AllFollowUpOperationsFailFast()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatchWithTimeouts(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(3),
            startupTimeout: ComInteropConstants.DefaultOperationTimeout,
            _testFileCopy!);

        _output.WriteLine("Warming up session...");
        SessionWorkbookAssertions.AssertIdentity(batch, _testFileCopy!);

        // Operation A: First timeout
        _output.WriteLine("Operation A: First timeout");
        Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });

        // Operation B: Second timeout attempt (should fail fast)
        _output.WriteLine("Operation B: Second timeout attempt (should fail fast)");
        var operationBSw = Stopwatch.StartNew();
        var rejectedB = false;
        var exceptionB = Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                rejectedB = true;
                return 0;
            });
        });
        operationBSw.Stop();
        _output.WriteLine($"  Operation B threw in {operationBSw.Elapsed.TotalMilliseconds:F0}ms");

        Assert.True(operationBSw.Elapsed < TimeSpan.FromSeconds(1),
            $"Operation B (second timeout) took {operationBSw.Elapsed.TotalSeconds:F1}s — expected < 1s");
        Assert.False(rejectedB);
        Assert.Contains("previous operation", exceptionB.Message, StringComparison.OrdinalIgnoreCase);

        // Operation C: Quick read (should also fail fast)
        _output.WriteLine("Operation C: Quick read (should fail fast)");
        var operationCSw = Stopwatch.StartNew();
        var rejectedC = false;
        var exceptionC = Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                rejectedC = true;
                return 42;
            });
        });
        operationCSw.Stop();
        _output.WriteLine($"  Operation C threw in {operationCSw.Elapsed.TotalMilliseconds:F0}ms");

        // REGRESSION ASSERTION: Even after multiple timeouts, later operations fail fast
        Assert.True(operationCSw.Elapsed < TimeSpan.FromSeconds(1),
            $"REGRESSION: Operation C after multiple timeouts took {operationCSw.Elapsed.TotalSeconds:F1}s. " +
            "Expected < 1s. Session poisoning may be cumulative.");
        Assert.False(rejectedC);
        Assert.Contains("previous operation", exceptionC.Message, StringComparison.OrdinalIgnoreCase);

        // Dispose should still work
        var disposeSw = Stopwatch.StartNew();
        batch.Dispose();
        disposeSw.Stop();
        _output.WriteLine($"Dispose completed in {disposeSw.Elapsed.TotalSeconds:F1}s");

        Assert.True(disposeSw.Elapsed < TimeSpan.FromSeconds(30),
            $"Dispose after multiple timeouts took {disposeSw.Elapsed.TotalSeconds:F1}s");
        _owned.AssertAllExited();

        _output.WriteLine("✓ Multiple timeouts test passed: all follow-ups failed fast");
    }
}
