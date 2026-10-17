using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Expired waits retain the session and exact process until execution completes.
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
    /// Even force-close must refuse while an expired operation is still running.
    /// </summary>
    [Fact]
    public void CloseSession_AfterTimeout_RefusesUntilExecutionCompletes()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(CloseSession_AfterTimeout_RefusesUntilExecutionCompletes));
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
        SessionWorkbookAssertions.WriteMarker(batch, "Retained unsaved contents");

        // Trigger timeout
        var ex = Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(5));
                return 0;
            });
        });
        _output.WriteLine($"Timeout triggered: {ex.Message}");
        Assert.Contains("timed out", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(batch.HasTimedOutOperation);

        Assert.Throws<ExcelBusyException>(() => manager.CloseSession(sessionId, save: false, force: true));
        Assert.Same(batch, manager.GetSession(sessionId));
        Assert.True(batch.IsExcelProcessAlive());
        WaitUntilReady(batch);
        Assert.False(batch.HasTimedOutOperation);
        Assert.Equal("Retained unsaved contents", SessionWorkbookAssertions.ReadMarker(batch));
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
    /// The exact process survives expiry and exits only after a later explicit close.
    /// </summary>
    [Fact]
    public void CloseSession_AfterTimeout_PreservesExactProcessUntilExplicitClose()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(CloseSession_AfterTimeout_PreservesExactProcessUntilExplicitClose));
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
                Thread.Sleep(TimeSpan.FromSeconds(5));
                return 0;
            });
        });

        Assert.Throws<ExcelBusyException>(() => manager.CloseSession(sessionId, save: false, force: true));
        Assert.Equal(excelPid, batch.ExcelProcessId);
        Assert.True(batch.IsExcelProcessAlive());
        Assert.Equal(sessionId, Assert.Single(manager.ActiveSessionIds));
        WaitUntilReady(batch);
        var sw = Stopwatch.StartNew();
        Assert.True(manager.CloseSession(sessionId, save: false, force: true));
        sw.Stop();
        _output.WriteLine($"CloseSession took {sw.Elapsed.TotalSeconds:F1}s");

        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.Empty(manager.ActiveSessionIds);
        Assert.Null(manager.GetSession(sessionId));
        _owned.AssertAllExited();
    }

    private static void WaitUntilReady(IExcelBatch batch)
    {
        var state = Assert.IsAssignableFrom<IExcelBatchRefreshState>(batch);
        Assert.True(SpinWait.SpinUntil(
            () => state.GetRefreshState() == WorkbookRefreshState.Ready,
            TimeSpan.FromSeconds(20)), "Excel did not become ready after execution completed.");
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
