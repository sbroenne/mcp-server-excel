using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Tests for SessionManager operation tracking functionality.
/// Verifies that BeginOperation/EndOperation tracking works correctly
/// and that CloseSession is blocked when operations are running.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "true")]
[Collection("Sequential")]
public class SessionManagerOperationTrackingTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly List<string> _testFiles = [];

    public SessionManagerOperationTrackingTests(ITestOutputHelper output)
    {
        _output = output;
#pragma warning disable CA3003 // Path.Combine is safe here - test code with controlled inputs
        _tempDir = Path.Combine(Path.GetTempPath(), $"SessionManagerOpTrackingTests_{Guid.NewGuid():N}");
#pragma warning restore CA3003
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);

        foreach (var file in _testFiles.Where(File.Exists))
        {
#pragma warning disable CA1031 // Catch general exception - best effort cleanup in test disposal
            try { File.Delete(file); } catch { /* best effort */ }
#pragma warning restore CA1031
        }

        if (Directory.Exists(_tempDir))
        {
#pragma warning disable CA1031 // Catch general exception - best effort cleanup in test disposal
            try { Directory.Delete(_tempDir, recursive: true); } catch { /* best effort */ }
#pragma warning restore CA1031
        }

        Thread.Sleep(500);
    }

    private string CreateTestFile(string testName)
    {
        var fileName = $"{testName}_{Guid.NewGuid():N}.xlsx";
#pragma warning disable CA3003 // Path.Combine is safe here - test code with controlled inputs
        var filePath = Path.Combine(_tempDir, fileName);
#pragma warning restore CA3003

        ExcelSession.CreateNew(
            filePath,
            isMacroEnabled: false,
            (ctx, ct) => 0);

        _testFiles.Add(filePath);
        return filePath;
    }

    #region BeginOperation / EndOperation

    [Fact]
    public void BeginOperation_IncrementsCounter()
    {
        var testFile = CreateTestFile(nameof(BeginOperation_IncrementsCounter));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        Assert.Equal(0, manager.GetActiveOperationCount(sessionId));

        manager.BeginOperation(sessionId);
        Assert.Equal(1, manager.GetActiveOperationCount(sessionId));

        manager.BeginOperation(sessionId);
        Assert.Equal(2, manager.GetActiveOperationCount(sessionId));

        manager.EndOperation(sessionId);
        manager.EndOperation(sessionId);
        manager.CloseSession(sessionId);
    }

    [Fact]
    public void EndOperation_DecrementsCounter()
    {
        var testFile = CreateTestFile(nameof(EndOperation_DecrementsCounter));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        manager.BeginOperation(sessionId);
        manager.BeginOperation(sessionId);
        Assert.Equal(2, manager.GetActiveOperationCount(sessionId));

        manager.EndOperation(sessionId);
        Assert.Equal(1, manager.GetActiveOperationCount(sessionId));

        manager.EndOperation(sessionId);
        Assert.Equal(0, manager.GetActiveOperationCount(sessionId));

        manager.CloseSession(sessionId);
    }

    [Fact]
    public void EndOperation_DoesNotGoBelowZero()
    {
        var testFile = CreateTestFile(nameof(EndOperation_DoesNotGoBelowZero));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        // End without begin
        manager.EndOperation(sessionId);
        manager.EndOperation(sessionId);

        Assert.Equal(0, manager.GetActiveOperationCount(sessionId));

        manager.CloseSession(sessionId);
    }

    [Fact]
    public void GetActiveOperationCount_NonExistentSession_ReturnsZero()
    {
        using var manager = new SessionManager();

        Assert.Equal(0, manager.GetActiveOperationCount("nonexistent"));
        Assert.Equal(0, manager.GetActiveOperationCount(null!));
        Assert.Equal(0, manager.GetActiveOperationCount(""));
    }

    [Fact]
    public void BeginEndOperation_NullSessionId_DoesNotThrow()
    {
        using var manager = new SessionManager();

        // Should not throw
        manager.BeginOperation(null!);
        manager.BeginOperation("");
        manager.EndOperation(null!);
        manager.EndOperation("");
    }

    #endregion

    #region IsExcelVisible

    [Fact]
    public void IsExcelVisible_SessionWithShowExcelFalse_ReturnsFalse()
    {
        var testFile = CreateTestFile(nameof(IsExcelVisible_SessionWithShowExcelFalse_ReturnsFalse));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, showExcel: false);

        Assert.False(manager.IsExcelVisible(sessionId));

        manager.CloseSession(sessionId);
    }

    [Fact]
    public void IsExcelVisible_SessionWithShowExcelTrue_ReturnsTrue()
    {
        var testFile = CreateTestFile(nameof(IsExcelVisible_SessionWithShowExcelTrue_ReturnsTrue));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, showExcel: true);

        Assert.True(manager.IsExcelVisible(sessionId));

        manager.CloseSession(sessionId);
    }

    [Fact]
    public void IsExcelVisible_NonExistentSession_ReturnsFalse()
    {
        using var manager = new SessionManager();

        Assert.False(manager.IsExcelVisible("nonexistent"));
        Assert.False(manager.IsExcelVisible(null!));
    }

    #endregion

    #region ValidateClose

    [Fact]
    public void ValidateClose_NoOperationsRunning_CanCloseTrue()
    {
        var testFile = CreateTestFile(nameof(ValidateClose_NoOperationsRunning_CanCloseTrue));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        var result = manager.ValidateClose(sessionId);

        Assert.True(result.SessionExists);
        Assert.True(result.CanClose);
        Assert.Equal(0, result.ActiveOperationCount);
        Assert.Null(result.BlockingReason);

        manager.CloseSession(sessionId);
    }

    [Fact]
    public void ValidateClose_OperationsRunning_CanCloseFalse()
    {
        var testFile = CreateTestFile(nameof(ValidateClose_OperationsRunning_CanCloseFalse));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        manager.BeginOperation(sessionId);
        manager.BeginOperation(sessionId);

        var result = manager.ValidateClose(sessionId);

        Assert.True(result.SessionExists);
        Assert.False(result.CanClose);
        Assert.Equal(2, result.ActiveOperationCount);
        Assert.NotNull(result.BlockingReason);
        Assert.Contains("2 operation(s) still running", result.BlockingReason);

        manager.EndOperation(sessionId);
        manager.EndOperation(sessionId);
        manager.CloseSession(sessionId);
    }

    [Fact]
    public void ValidateClose_NonExistentSession_SessionExistsFalse()
    {
        using var manager = new SessionManager();

        var result = manager.ValidateClose("nonexistent");

        Assert.False(result.SessionExists);
        Assert.False(result.CanClose);
        Assert.NotNull(result.BlockingReason);
        Assert.Contains("not found", result.BlockingReason);
    }

    [Fact]
    public void ValidateClose_NullSessionId_SessionExistsFalse()
    {
        using var manager = new SessionManager();

        var result = manager.ValidateClose(null!);

        Assert.False(result.SessionExists);
        Assert.Contains("required", result.BlockingReason);
    }

    [Fact]
    public void ValidateClose_IncludesVisibilityInfo()
    {
        var testFile = CreateTestFile(nameof(ValidateClose_IncludesVisibilityInfo));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, showExcel: true);

        var result = manager.ValidateClose(sessionId);

        Assert.True(result.IsExcelVisible);

        manager.CloseSession(sessionId);
    }

    #endregion

    #region CloseSession with Operation Tracking

    [Fact]
    public void CloseSession_OperationsRunning_ThrowsInvalidOperationException()
    {
        var testFile = CreateTestFile(nameof(CloseSession_OperationsRunning_ThrowsInvalidOperationException));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        manager.BeginOperation(sessionId);

        var ex = Assert.Throws<InvalidOperationException>(
            () => manager.CloseSession(sessionId));

        Assert.Contains("1 operation(s) still running", ex.Message);
        Assert.Contains("Wait for all operations to complete", ex.Message);

        // Session should still be open
        Assert.Equal(1, manager.ActiveSessionCount);

        // Clean up
        manager.EndOperation(sessionId);
        manager.CloseSession(sessionId);
    }

    [Fact]
    public void CloseSession_OperationsComplete_ClosesSuccessfully()
    {
        var testFile = CreateTestFile(nameof(CloseSession_OperationsComplete_ClosesSuccessfully));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        // Simulate operation lifecycle
        manager.BeginOperation(sessionId);
        manager.EndOperation(sessionId);

        // Should now be able to close
        var closed = manager.CloseSession(sessionId);

        Assert.True(closed);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void CloseSession_ForceTrue_ClosesEvenWithRunningOperations()
    {
        var testFile = CreateTestFile(nameof(CloseSession_ForceTrue_ClosesEvenWithRunningOperations));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        manager.BeginOperation(sessionId);
        manager.BeginOperation(sessionId);

        // Force close should work even with operations running
        var closed = manager.CloseSession(sessionId, save: false, force: true);

        Assert.True(closed);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    #endregion

    #region Cleanup on Close

    [Fact]
    public void CloseSession_CleansUpOperationTracking()
    {
        var testFile = CreateTestFile(nameof(CloseSession_CleansUpOperationTracking));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, showExcel: true);

        // Set some state
        manager.BeginOperation(sessionId);
        manager.EndOperation(sessionId);

        manager.CloseSession(sessionId);

        // After close, these should return defaults
        Assert.Equal(0, manager.GetActiveOperationCount(sessionId));
        Assert.False(manager.IsExcelVisible(sessionId));
    }

    [Fact]
    public void Dispose_CleansUpAllTracking()
    {
        var testFile1 = CreateTestFile($"{nameof(Dispose_CleansUpAllTracking)}_1");
        var testFile2 = CreateTestFile($"{nameof(Dispose_CleansUpAllTracking)}_2");
        using var manager = new SessionManager();

        var session1 = manager.CreateSession(testFile1, showExcel: true);
        var session2 = manager.CreateSession(testFile2, showExcel: false);

        manager.BeginOperation(session1);
        manager.BeginOperation(session2);

        manager.Dispose();

        // All tracking should be cleared
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    #endregion
}
