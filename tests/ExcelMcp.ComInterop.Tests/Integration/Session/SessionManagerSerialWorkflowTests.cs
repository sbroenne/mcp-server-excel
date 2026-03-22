using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// REGRESSION TESTS for Bug Report (v1.8.32): SessionManager serial workflow recovery.
///
/// BUG PATTERN FROM REPORT:
/// - Serial operations through SessionManager (like CLI daemon uses)
/// - Timeout/cancel in middle of workflow
/// - Later session operations fail or hang
/// - New session creation on same file fails or hangs
/// - Dead session cleanup doesn't happen fast enough
///
/// These tests validate SessionManager-level recovery paths that the ExcelBatch tests don't cover:
/// - Does GetSession return the poisoned session or clean it up?
/// - Does CreateSession work immediately after a timeout on the same file?
/// - Does ActiveSessionCount reflect reality after timeout cleanup?
/// - Does stale session metadata block new session creation?
///
/// DIFFERENCE FROM ExcelBatchSerialWorkflowTests:
/// - ExcelBatch tests: Direct batch API (no session manager)
/// - These tests: SessionManager API (like CLI daemon uses)
/// - Focus: Session lifecycle, reopen, metadata cleanup
///
/// EXPECTED OUTCOME (if bug still exists):
/// - These tests should FAIL (RED) because SessionManager recovery path has gaps
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SessionManager")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")]
public class SessionManagerSerialWorkflowTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly List<string> _testFiles = [];

    private static readonly string TemplateFilePath = Path.Combine(
        Path.GetDirectoryName(typeof(SessionManagerSerialWorkflowTests).Assembly.Location)!,
        "Integration", "Session", "TestFiles", "batch-test-static.xlsx");

    public SessionManagerSerialWorkflowTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"SessionMgrSerialTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);

        foreach (var file in _testFiles.Where(File.Exists))
        {
#pragma warning disable CA1031 // Intentional: best-effort test cleanup
            try { File.Delete(file); } catch (Exception) { /* best effort */ }
#pragma warning restore CA1031
        }

        if (Directory.Exists(_tempDir))
        {
#pragma warning disable CA1031
            try { Directory.Delete(_tempDir, recursive: true); } catch (Exception) { /* best effort */ }
#pragma warning restore CA1031
        }
    }

    private string CreateTestFile(string testName)
    {
        var filePath = Path.Combine(_tempDir, $"{testName}_{Guid.NewGuid():N}.xlsx");
        File.Copy(TemplateFilePath, filePath);
        _testFiles.Add(filePath);
        return filePath;
    }

    /// <summary>
    /// REGRESSION TEST: After timeout, subsequent GetSession on same session ID should handle poisoned state.
    ///
    /// WORKFLOW:
    /// 1. CreateSession → sessionId
    /// 2. GetSession(sessionId) → batch
    /// 3. batch.Execute → timeout
    /// 4. GetSession(sessionId) again → what happens?
    ///
    /// EXPECTED (if bug exists): Test FAILS because:
    /// - GetSession returns the poisoned batch that keeps failing
    /// - OR GetSession returns null but session isn't cleaned from ActiveSessionCount
    /// - Caller has no clear recovery path
    ///
    /// EXPECTED (if bug fixed): Test PASSES because:
    /// - GetSession returns null after timeout (session auto-cleaned)
    /// - OR GetSession returns batch but operations fail fast with useful errors
    /// - ActiveSessionCount is accurate
    /// </summary>
    [Fact]
    public void SerialWorkflow_GetSessionAfterTimeout_ReturnsNullOrFailsFast()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(SerialWorkflow_GetSessionAfterTimeout_ReturnsNullOrFailsFast));
        using var manager = new SessionManager();

        var sessionId = manager.CreateSession(testFile, operationTimeout: TimeSpan.FromSeconds(3));
        _output.WriteLine($"Created session: {sessionId}");

        var batch = manager.GetSession(sessionId);
        Assert.NotNull(batch);

        // Warm up
        batch.Execute((ctx, ct) => { _ = ctx.Book.Worksheets[1]; return 0; });

        // Trigger timeout
        _output.WriteLine("Triggering timeout...");
        Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });

        // Act: Get session again after timeout
        _output.WriteLine("Getting session after timeout...");
        var batchAfterTimeout = manager.GetSession(sessionId);

        // Assert: Either null (cleaned) or still present but operations fail fast
        if (batchAfterTimeout == null)
        {
            _output.WriteLine("  Session returned null — session was auto-cleaned");
            Assert.Equal(0, manager.ActiveSessionCount);
        }
        else
        {
            _output.WriteLine("  Session still exists — testing if operations fail fast");

            var sw = Stopwatch.StartNew();
            var ex = Assert.Throws<TimeoutException>(() =>
            {
                batchAfterTimeout.Execute((ctx, ct) =>
                {
                    dynamic sheet = ctx.Book.Worksheets[1];
                    return sheet.Name?.ToString() ?? "unknown";
                });
            });
            sw.Stop();

            _output.WriteLine($"  Operation threw in {sw.Elapsed.TotalMilliseconds:F0}ms: {ex.Message}");

            // REGRESSION ASSERTION: If session still exists, operations must fail FAST
            Assert.True(sw.Elapsed < TimeSpan.FromSeconds(1),
                $"REGRESSION: Operation after timeout took {sw.Elapsed.TotalSeconds:F1}s. " +
                "Expected < 1s. Session is poisoned but not failing fast.");
        }

        _output.WriteLine("✓ GetSession after timeout test passed");
    }

    /// <summary>
    /// REGRESSION TEST: After timeout + CloseSession, can we immediately CreateSession on same file?
    ///
    /// WORKFLOW:
    /// 1. Session A: CreateSession(file) → timeout → CloseSession(force:true)
    /// 2. Session B: CreateSession(SAME file) immediately
    /// 3. Verify Session B works
    ///
    /// EXPECTED (if bug exists): Test FAILS because:
    /// - CreateSession on same file hangs or times out
    /// - Stale metadata from Session A blocks Session B
    /// - File locking from dead Excel blocks reopen
    ///
    /// EXPECTED (if bug fixed): Test PASSES because:
    /// - CreateSession succeeds quickly (under 10s)
    /// - Session B operations work
    /// - No interference from Session A
    /// </summary>
    [Fact]
    public void SerialWorkflow_CreateSessionAfterTimeoutCleanup_SucceedsOnSameFile()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(SerialWorkflow_CreateSessionAfterTimeoutCleanup_SucceedsOnSameFile));
        using var manager = new SessionManager();

        // Session A: Timeout and force close
        _output.WriteLine("Session A: Creating and timing out...");
        var sessionAId = manager.CreateSession(testFile, operationTimeout: TimeSpan.FromSeconds(3));
        var sessionA = manager.GetSession(sessionAId)!;

        sessionA.Execute((ctx, ct) => { _ = ctx.Book.Worksheets[1]; return 0; });

        _output.WriteLine("  Triggering timeout in Session A...");
        Assert.Throws<TimeoutException>(() =>
        {
            sessionA.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });

        _output.WriteLine("  Force closing Session A...");
        var closeResult = manager.CloseSession(sessionAId, save: false, force: true);
        Assert.True(closeResult, "CloseSession should succeed");
        Assert.Equal(0, manager.ActiveSessionCount);

        // Wait for cleanup
        Thread.Sleep(2000);

        // Session B: Create on same file immediately
        _output.WriteLine("Session B: Creating on SAME file immediately...");
        var sessionBSw = Stopwatch.StartNew();
        var sessionBId = manager.CreateSession(testFile, operationTimeout: TimeSpan.FromSeconds(30));
        sessionBSw.Stop();

        _output.WriteLine($"  Session B created in {sessionBSw.Elapsed.TotalSeconds:F1}s");

        // REGRESSION ASSERTION: CreateSession should be fast (< 10s)
        // BUG SYMPTOM: If this fails, stale Session A state is blocking Session B creation
        Assert.True(sessionBSw.Elapsed < TimeSpan.FromSeconds(10),
            $"REGRESSION: CreateSession on same file after timeout cleanup took {sessionBSw.Elapsed.TotalSeconds:F1}s. " +
            "Expected < 10s. Stale session metadata or file locking may be interfering.");

        // Verify Session B works
        _output.WriteLine("  Verifying Session B operations...");
        var sessionB = manager.GetSession(sessionBId);
        Assert.NotNull(sessionB);

        var sheetName = sessionB.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            return sheet.Name?.ToString() ?? "unknown";
        });

        _output.WriteLine($"  ✓ Session B read succeeded: {sheetName}");
        Assert.Equal(1, manager.ActiveSessionCount);

        // Cleanup
        manager.CloseSession(sessionBId, save: false, force: false);

        _output.WriteLine("✓ CreateSession after timeout cleanup test passed");
    }

    /// <summary>
    /// REGRESSION TEST: ActiveSessionCount stays accurate through timeout workflow.
    ///
    /// WORKFLOW:
    /// 1. CreateSession → ActiveSessionCount = 1
    /// 2. Timeout
    /// 3. CloseSession(force:true) → ActiveSessionCount should be 0
    /// 4. GetSession should return null
    ///
    /// EXPECTED (if bug exists): Test FAILS because:
    /// - ActiveSessionCount stays at 1 after CloseSession
    /// - Stale session metadata leaks
    ///
    /// EXPECTED (if bug fixed): Test PASSES because:
    /// - ActiveSessionCount reflects reality at each step
    /// - Session is truly removed after force close
    /// </summary>
    [Fact]
    public void SerialWorkflow_ActiveSessionCount_RemainsAccurateThroughTimeout()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(SerialWorkflow_ActiveSessionCount_RemainsAccurateThroughTimeout));
        using var manager = new SessionManager();

        Assert.Equal(0, manager.ActiveSessionCount);

        // Create session
        var sessionId = manager.CreateSession(testFile, operationTimeout: TimeSpan.FromSeconds(3));
        _output.WriteLine($"Session created: {sessionId}");
        Assert.Equal(1, manager.ActiveSessionCount);

        var batch = manager.GetSession(sessionId);
        Assert.NotNull(batch);

        // Warm up
        batch.Execute((ctx, ct) => { _ = ctx.Book.Worksheets[1]; return 0; });

        // Timeout
        _output.WriteLine("Triggering timeout...");
        Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });

        // ASSERTION: ActiveSessionCount should still be 1 (session not auto-removed yet)
        // This is expected — timeout doesn't auto-close the session
        Assert.Equal(1, manager.ActiveSessionCount);

        // Force close
        _output.WriteLine("Force closing session...");
        var closed = manager.CloseSession(sessionId, save: false, force: true);
        Assert.True(closed);

        // REGRESSION ASSERTION: ActiveSessionCount should now be 0
        Assert.Equal(0, manager.ActiveSessionCount);

        // REGRESSION ASSERTION: GetSession should return null
        Assert.Null(manager.GetSession(sessionId));

        _output.WriteLine("✓ ActiveSessionCount remained accurate through timeout workflow");
    }

    /// <summary>
    /// REGRESSION TEST: Multiple sessions, timeout one, others still work.
    ///
    /// WORKFLOW:
    /// 1. Session A (file1) → success
    /// 2. Session B (file2) → timeout
    /// 3. Session C (file3) → success
    /// 4. Session A again → should still work (not poisoned by B's timeout)
    ///
    /// EXPECTED (if bug exists): Test FAILS because:
    /// - Session B timeout poisons SessionManager state
    /// - Sessions A and C fail after B times out
    ///
    /// EXPECTED (if bug fixed): Test PASSES because:
    /// - Session isolation works
    /// - Only Session B is poisoned
    /// - Sessions A and C continue working
    /// </summary>
    [Fact]
    public void SerialWorkflow_MultipleSessionsOneTimeout_OthersUnaffected()
    {
        // Arrange
        var fileA = CreateTestFile($"{nameof(SerialWorkflow_MultipleSessionsOneTimeout_OthersUnaffected)}_A");
        var fileB = CreateTestFile($"{nameof(SerialWorkflow_MultipleSessionsOneTimeout_OthersUnaffected)}_B");
        var fileC = CreateTestFile($"{nameof(SerialWorkflow_MultipleSessionsOneTimeout_OthersUnaffected)}_C");

        using var manager = new SessionManager();

        // Session A: Normal timeout
        _output.WriteLine("Session A: Creating (file A, normal timeout)");
        var sessionAId = manager.CreateSession(fileA, operationTimeout: TimeSpan.FromSeconds(30));
        var sessionA = manager.GetSession(sessionAId)!;
        var sheetA1 = sessionA.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            return sheet.Name?.ToString() ?? "unknown";
        });
        _output.WriteLine($"  ✓ Session A read: {sheetA1}");

        // Session B: Short timeout (will timeout)
        _output.WriteLine("Session B: Creating (file B, short timeout)");
        var sessionBId = manager.CreateSession(fileB, operationTimeout: TimeSpan.FromSeconds(3));
        var sessionB = manager.GetSession(sessionBId)!;
        sessionB.Execute((ctx, ct) => { _ = ctx.Book.Worksheets[1]; return 0; }); // warmup

        _output.WriteLine("  Triggering timeout in Session B...");
        Assert.Throws<TimeoutException>(() =>
        {
            sessionB.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });
        _output.WriteLine("  ✓ Session B timed out");

        // Session C: Normal timeout (created after B timeout)
        _output.WriteLine("Session C: Creating (file C, normal timeout, AFTER B timeout)");
        var sessionCId = manager.CreateSession(fileC, operationTimeout: TimeSpan.FromSeconds(30));
        var sessionC = manager.GetSession(sessionCId)!;
        var sheetC1 = sessionC.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            return sheet.Name?.ToString() ?? "unknown";
        });
        _output.WriteLine($"  ✓ Session C read: {sheetC1}");

        // Session A again: Should still work
        _output.WriteLine("Session A: Reading again (AFTER B timeout)");
        var sheetA2 = sessionA.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            return sheet.Name?.ToString() ?? "unknown";
        });
        _output.WriteLine($"  ✓ Session A read again: {sheetA2}");

        // REGRESSION ASSERTION: Sessions A and C should have worked normally
        Assert.NotEmpty(sheetA1);
        Assert.NotEmpty(sheetA2);
        Assert.NotEmpty(sheetC1);
        Assert.Equal(3, manager.ActiveSessionCount);

        // Cleanup
        manager.CloseSession(sessionAId, save: false, force: false);
        manager.CloseSession(sessionBId, save: false, force: true);
        manager.CloseSession(sessionCId, save: false, force: false);

        _output.WriteLine("✓ Multiple sessions test passed: Session B timeout didn't affect A or C");
    }
}
