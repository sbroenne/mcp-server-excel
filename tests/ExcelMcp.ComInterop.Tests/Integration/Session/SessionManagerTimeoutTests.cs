using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Tests for SessionManager behavior when operations timeout.
///
/// These tests validate the integration between ExcelBatch timeout detection
/// and SessionManager session cleanup — the complete recovery path that was
/// missing before Bug 8 (Feb 2026).
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test that SessionManager.CloseSession(force:true) works after timeout
/// - ✅ Test that session is removed after timeout + force close
/// - ✅ Test that subsequent GetSession returns null after timeout cleanup
/// - ✅ Test that Excel process is cleaned up end-to-end through SessionManager
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SessionManager")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")]
[Trait("RequiresExcel", "true")]
public class SessionManagerTimeoutTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly List<string> _testFiles = [];
    private readonly OwnedExcelProcessScope _owned = new();

    private static readonly string TemplateFilePath = Path.Combine(
        Path.GetDirectoryName(typeof(SessionManagerTimeoutTests).Assembly.Location)!,
        "Integration", "Session", "TestFiles", "batch-test-static.xlsx");

    public SessionManagerTimeoutTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"SessionMgrTimeoutTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);

        SessionTestCleanup.AssertExitedAndDelete(_owned, _testFiles, _tempDir);
    }

    private string CreateTestFile(string testName)
    {
        var filePath = Path.Combine(_tempDir, $"{testName}_{Guid.NewGuid():N}.xlsx");
        File.Copy(TemplateFilePath, filePath);
        _testFiles.Add(filePath);
        return filePath;
    }

    /// <summary>
    /// REGRESSION TEST: After a timeout, CloseSession(force:true) must succeed and remove the session.
    /// This simulates what WithSessionAsync does when it catches TimeoutException.
    /// Before Bug 8 fix, there was no TimeoutException handler — the session leaked.
    /// </summary>
    [Fact]
    public void CloseSession_AfterTimeout_RemovesSessionAndCleansUp()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(CloseSession_AfterTimeout_RemovesSessionAndCleansUp));
        using var manager = new SessionManager();

        // Create session with very short timeout
        var sessionId = manager.CreateSessionWithTimeouts(
            testFile,
            show: false,
            operationTimeout: TimeSpan.FromSeconds(3),
            startupTimeout: ComInteropConstants.DefaultOperationTimeout);
        _output.WriteLine($"Session created: {sessionId}");

        var batch = manager.GetSession(sessionId);
        Assert.NotNull(batch);

        // Warm up
        SessionWorkbookAssertions.AssertIdentity(batch, testFile);

        // Trigger timeout
        var ex = Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });
        _output.WriteLine($"Timeout triggered: {ex.Message}");
        Assert.Contains("timed out", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(batch.HasTimedOutOperation);

        // Act — simulate what WithSessionAsync does: force-close the session
        var closed = manager.CloseSession(sessionId, save: false, force: true);

        // Assert
        Assert.True(closed, "CloseSession should succeed after timeout");
        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.Null(manager.GetSession(sessionId));
        Assert.Empty(manager.ActiveSessionIds);
        _owned.AssertAllExited();

        _output.WriteLine("✓ Session cleaned up after timeout");
    }

    /// <summary>
    /// REGRESSION TEST: After timeout + force close, the Excel process must be terminated.
    /// This is the end-to-end test for the complete Bug 8 recovery chain:
    /// timeout → force close → pre-emptive kill → process cleanup.
    /// </summary>
    [Fact]
    public void CloseSession_AfterTimeout_ExcelProcessIsTerminated()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(CloseSession_AfterTimeout_ExcelProcessIsTerminated));
        using var manager = new SessionManager();

        var sessionId = manager.CreateSessionWithTimeouts(
            testFile,
            show: false,
            operationTimeout: TimeSpan.FromSeconds(3),
            startupTimeout: ComInteropConstants.DefaultOperationTimeout);
        var batch = Assert.IsAssignableFrom<IExcelBatch>(manager.GetSession(sessionId));
        int? excelPid = batch.ExcelProcessId;
        Assert.NotNull(excelPid);
        _output.WriteLine($"Session {sessionId}, Excel PID: {excelPid}");

        // Warm up
        SessionWorkbookAssertions.AssertIdentity(batch, testFile);

        // Trigger timeout
        Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });

        // Act — force close (this triggers Dispose → pre-emptive kill)
        var sw = Stopwatch.StartNew();
        Assert.True(manager.CloseSession(sessionId, save: false, force: true));
        sw.Stop();
        _output.WriteLine($"CloseSession took {sw.Elapsed.TotalSeconds:F1}s");

        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.Empty(manager.ActiveSessionIds);
        Assert.Null(manager.GetSession(sessionId));
        _owned.AssertAllExited();
    }

    /// <summary>
    /// Test that normal operations still work with custom timeout — only long operations fail.
    /// </summary>
    [Fact]
    public void Execute_WithinTimeout_SucceedsNormally()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(Execute_WithinTimeout_SucceedsNormally));
        using var manager = new SessionManager();

        var sessionId = manager.CreateSession(testFile, operationTimeout: TimeSpan.FromSeconds(30));
        var batch = Assert.IsAssignableFrom<IExcelBatch>(manager.GetSession(sessionId));

        // Act — quick operation should succeed
        SessionWorkbookAssertions.AssertIdentity(batch, testFile);
        SessionWorkbookAssertions.WriteMarker(batch, "within-timeout");
        Assert.Equal("within-timeout", SessionWorkbookAssertions.ReadMarker(batch));
        Assert.False(batch.HasTimedOutOperation);
        Assert.True(manager.CloseSession(sessionId));
        Assert.Empty(manager.ActiveSessionIds);
        _owned.AssertAllExited();
    }
}
