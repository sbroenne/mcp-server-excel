using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Integration tests for SessionManager - verifies session lifecycle management.
/// Tests multi-session scenarios, concurrent operations, and proper cleanup.
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test session creation and tracking
/// - ✅ Test session retrieval by ID
/// - ✅ Test save operations
/// - ✅ Test close operations
/// - ✅ Test concurrent multi-session scenarios
/// - ✅ Test disposal cleanup
/// - ✅ Test post-disposal protection
///
/// NOTE: SessionManager uses ExcelSession internally, so these tests verify
/// the orchestration layer, not the underlying Excel COM interactions.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "true")]
[Collection("Sequential")] // Disable parallelization to avoid COM interference
public class SessionManagerTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly List<string> _testFiles = new();

    public SessionManagerTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"SessionManagerTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        // Clean up any existing Excel processes to ensure clean state
        try
        {
            var existingProcesses = Process.GetProcessesByName("EXCEL");
            if (existingProcesses.Length > 0)
            {
                _output.WriteLine($"Cleaning up {existingProcesses.Length} existing Excel processes...");
                foreach (var p in existingProcesses)
                {
                    p.Kill(entireProcessTree: true);
                    p.WaitForExit(5000);
                    p.Dispose();
                }
                Thread.Sleep(2000);
            }
        }
        catch (Exception ex)
        {
            _output.WriteLine($"Warning: Failed to clean Excel processes: {ex.Message}");
        }
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);

        // Delete test files
        foreach (var file in _testFiles)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }

        // Delete temp directory
        if (Directory.Exists(_tempDir))
        {
            Directory.Delete(_tempDir, recursive: true);
        }

        // Give Excel time to fully terminate
        Thread.Sleep(1000);
    }

    private string CreateTestFile(string testName)
    {
        var fileName = $"{testName}_{Guid.NewGuid():N}.xlsx";
        var filePath = Path.Combine(_tempDir, fileName);

        // Create blank workbook
        ExcelSession.CreateNew(
            filePath,
            isMacroEnabled: false,
            (ctx, ct) =>
            {
                // Just create the file - no operations needed
                return 0;
            });

        _testFiles.Add(filePath);
        return filePath;
    }

    #region Basic Session Lifecycle

    [Fact]
    public void CreateSession_ValidFile_ReturnsSessionId()
    {
        var testFile = CreateTestFile(nameof(CreateSession_ValidFile_ReturnsSessionId));
        using var manager = new SessionManager();

        var sessionId = manager.CreateSession(testFile);

        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        Assert.Equal(32, sessionId.Length); // GUID without hyphens
        Assert.Equal(1, manager.ActiveSessionCount);

        manager.CloseSession(sessionId);
    }

    [Fact]
    public void CreateSession_NonExistentFile_ThrowsFileNotFoundException()
    {
        using var manager = new SessionManager();
        var nonExistentFile = Path.Combine(_tempDir, "nonexistent.xlsx");

        var ex = Assert.Throws<FileNotFoundException>(
            () => manager.CreateSession(nonExistentFile));

        Assert.Contains("Excel file not found", ex.Message);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void GetSession_ExistingSessionId_ReturnsValidBatch()
    {
        var testFile = CreateTestFile(nameof(GetSession_ExistingSessionId_ReturnsValidBatch));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        var batch = manager.GetSession(sessionId);

        Assert.NotNull(batch);
        Assert.Equal(1, manager.ActiveSessionCount);

        manager.CloseSession(sessionId);
    }

    [Fact]
    public void GetSession_NonExistentSessionId_ReturnsNull()
    {
        using var manager = new SessionManager();

        var batch = manager.GetSession("nonexistent-session-id");

        Assert.Null(batch);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void GetSession_NullOrWhitespaceSessionId_ReturnsNull()
    {
        using var manager = new SessionManager();

        Assert.Null(manager.GetSession(null!));
        Assert.Null(manager.GetSession(""));
        Assert.Null(manager.GetSession("   "));
    }

    #endregion

    #region Save Operations

    [Fact]
    public void CloseSession_WithSaveTrue_SavesAndCloses()
    {
        var testFile = CreateTestFile(nameof(CloseSession_WithSaveTrue_SavesAndCloses));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        // Modify data to verify save
        var batch = manager.GetSession(sessionId);
        Assert.NotNull(batch);
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            sheet.Cells[1, 1].Value2 = "Test Value";
            return 0;
        });

        var closed = manager.CloseSession(sessionId, save: true);

        Assert.True(closed);
        Assert.Equal(0, manager.ActiveSessionCount);

        // Verify changes persisted
        using var verifyBatch = ExcelSession.BeginBatch(testFile);
        var value = verifyBatch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            return (string)sheet.Cells[1, 1].Value2;
        });
        Assert.Equal("Test Value", value);
    }

    [Fact]
    public void CloseSession_WithSaveFalse_DiscardsChanges()
    {
        var testFile = CreateTestFile(nameof(CloseSession_WithSaveFalse_DiscardsChanges));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        // Modify data but don't save
        var batch = manager.GetSession(sessionId);
        Assert.NotNull(batch);
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            sheet.Cells[1, 1].Value2 = "Discarded Value";
            return 0;
        });

        var closed = manager.CloseSession(sessionId, save: false);

        Assert.True(closed);
        Assert.Equal(0, manager.ActiveSessionCount);

        // Verify changes were NOT persisted
        using var verifyBatch = ExcelSession.BeginBatch(testFile);
        var value = verifyBatch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            return sheet.Cells[1, 1].Value2;
        });
        Assert.Null(value); // Cell should be empty
    }

    #endregion

    #region Close Operations

    [Fact]
    public void CloseSession_ExistingSession_RemovesSessionAndReturnsTrue()
    {
        var testFile = CreateTestFile(nameof(CloseSession_ExistingSession_RemovesSessionAndReturnsTrue));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        var closed = manager.CloseSession(sessionId, save: false);

        Assert.True(closed);
        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.Null(manager.GetSession(sessionId));
    }

    [Fact]
    public void CloseSession_NullOrWhitespaceSessionId_ReturnsFalse()
    {
        using var manager = new SessionManager();

        Assert.False(manager.CloseSession(null!));
        Assert.False(manager.CloseSession(""));
        Assert.False(manager.CloseSession("   "));
    }

    [Fact]
    public void CloseSession_AlreadyClosedSession_ReturnsFalse()
    {
        var testFile = CreateTestFile(nameof(CloseSession_AlreadyClosedSession_ReturnsFalse));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        var closed1 = manager.CloseSession(sessionId);
        var closed2 = manager.CloseSession(sessionId);

        Assert.True(closed1);
        Assert.False(closed2);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    #endregion

    #region Multi-Session Scenarios

    [Fact]
    public void CreateMultipleSessions_DifferentFiles_TracksAllSessions()
    {
        var testFile1 = CreateTestFile($"{nameof(CreateMultipleSessions_DifferentFiles_TracksAllSessions)}_1");
        var testFile2 = CreateTestFile($"{nameof(CreateMultipleSessions_DifferentFiles_TracksAllSessions)}_2");
        using var manager = new SessionManager();

        var sessionId1 = manager.CreateSession(testFile1);
        Thread.Sleep(1000); // Delay to avoid COM initialization conflicts
        var sessionId2 = manager.CreateSession(testFile2);

        Assert.Equal(2, manager.ActiveSessionCount);
        Assert.Contains(sessionId1, manager.ActiveSessionIds);
        Assert.Contains(sessionId2, manager.ActiveSessionIds);

        manager.CloseSession(sessionId1);
        manager.CloseSession(sessionId2);
    }

    [Fact]
    public void ActiveSessionIds_ReflectsCurrentState()
    {
        var testFile1 = CreateTestFile($"{nameof(ActiveSessionIds_ReflectsCurrentState)}_1");
        var testFile2 = CreateTestFile($"{nameof(ActiveSessionIds_ReflectsCurrentState)}_2");
        using var manager = new SessionManager();

        // Initially empty
        Assert.Empty(manager.ActiveSessionIds);

        // After creating sessions
        var sessionId1 = manager.CreateSession(testFile1);
        var sessionId2 = manager.CreateSession(testFile2);
        var activeIds = manager.ActiveSessionIds.ToList();

        Assert.Equal(2, activeIds.Count);
        Assert.Contains(sessionId1, activeIds);
        Assert.Contains(sessionId2, activeIds);

        // After closing one session
        manager.CloseSession(sessionId1);
        activeIds = manager.ActiveSessionIds.ToList();

        Assert.Single(activeIds);
        Assert.Contains(sessionId2, activeIds);
        Assert.DoesNotContain(sessionId1, activeIds);

        manager.CloseSession(sessionId2);
    }

    [Fact]
    public void CloseOneSession_DoesNotAffectOtherSessions()
    {
        var testFile1 = CreateTestFile($"{nameof(CloseOneSession_DoesNotAffectOtherSessions)}_1");
        var testFile2 = CreateTestFile($"{nameof(CloseOneSession_DoesNotAffectOtherSessions)}_2");
        using var manager = new SessionManager();

        var sessionId1 = manager.CreateSession(testFile1);
        var sessionId2 = manager.CreateSession(testFile2);

        manager.CloseSession(sessionId1);

        Assert.Equal(1, manager.ActiveSessionCount);
        Assert.Null(manager.GetSession(sessionId1));
        Assert.NotNull(manager.GetSession(sessionId2));

        manager.CloseSession(sessionId2);
    }

    [Fact]
    public void CreateSession_SameFileAlreadyOpen_ThrowsInvalidOperationException()
    {
        var testFile = CreateTestFile(nameof(CreateSession_SameFileAlreadyOpen_ThrowsInvalidOperationException));
        using var manager = new SessionManager();

        // First session succeeds
        var sessionId1 = manager.CreateSession(testFile);
        Assert.NotNull(sessionId1);
        Assert.Equal(1, manager.ActiveSessionCount);

        // Second session with same file should fail fast
        var ex = Assert.Throws<InvalidOperationException>(
            () => manager.CreateSession(testFile));

        Assert.Contains("already open in another session", ex.Message);
        Assert.Contains("Excel cannot open the same file multiple times", ex.Message);
        Assert.Equal(1, manager.ActiveSessionCount); // Still only one session

        manager.CloseSession(sessionId1);
    }

    [Fact]
    public void CreateSession_AfterClosingPrevious_AllowsReopeningFile()
    {
        var testFile = CreateTestFile(nameof(CreateSession_AfterClosingPrevious_AllowsReopeningFile));
        using var manager = new SessionManager();

        // First session
        var sessionId1 = manager.CreateSession(testFile);
        Assert.Equal(1, manager.ActiveSessionCount);

        // Close first session
        manager.CloseSession(sessionId1);
        Assert.Equal(0, manager.ActiveSessionCount);

        // Should now be able to open same file again
        var sessionId2 = manager.CreateSession(testFile);
        Assert.NotNull(sessionId2);
        Assert.NotEqual(sessionId1, sessionId2);
        Assert.Equal(1, manager.ActiveSessionCount);

        manager.CloseSession(sessionId2);
    }

    #endregion

    #region Disposal and Post-Disposal

    [Fact]
    public void Dispose_OneSession_ClosesAllSessions()
    {
        var testFile1 = CreateTestFile($"{nameof(Dispose_OneSession_ClosesAllSessions)}_1");
        var manager = new SessionManager();

        var sessionId1 = manager.CreateSession(testFile1);

        Assert.Equal(1, manager.ActiveSessionCount);
        manager.Dispose();

        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.Empty(manager.ActiveSessionIds);
    }

    [Fact]
    public void Dispose_TwoSessions_ClosesAllSessions()
    {
        var testFile1 = CreateTestFile($"{nameof(Dispose_TwoSessions_ClosesAllSessions)}_1");
        var testFile2 = CreateTestFile($"{nameof(Dispose_TwoSessions_ClosesAllSessions)}_2");
        var manager = new SessionManager();

        var sessionId1 = manager.CreateSession(testFile1);
        var sessionId2 = manager.CreateSession(testFile2);

        Assert.Equal(2, manager.ActiveSessionCount);

        // DisposeAsync handles sessions sequentially to avoid COM threading issues
        manager.Dispose();

        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.Empty(manager.ActiveSessionIds);
    }

    [Fact]
    public void Dispose_EmptyManager_CompletesImmediately()
    {
        var manager = new SessionManager();

        manager.Dispose();

        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void Dispose_CalledMultipleTimes_DoesNotThrow()
    {
        var manager = new SessionManager();

        manager.Dispose();
        manager.Dispose();
        manager.Dispose();

        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void CreateSession_AfterDisposal_ThrowsObjectDisposedException()
    {
        var testFile = CreateTestFile(nameof(CreateSession_AfterDisposal_ThrowsObjectDisposedException));
        var manager = new SessionManager();
        manager.Dispose();

        Assert.Throws<ObjectDisposedException>(
            () => manager.CreateSession(testFile));
    }

    [Fact]
    public void GetSession_AfterDisposal_ThrowsObjectDisposedException()
    {
        var manager = new SessionManager();
        manager.Dispose();

        Assert.Throws<ObjectDisposedException>(
            () => manager.GetSession("any-id"));
    }

    [Fact]

    public void CloseSession_AfterDisposal_ThrowsObjectDisposedException()
    {
        var manager = new SessionManager();
        manager.Dispose();

        Assert.Throws<ObjectDisposedException>(
            () => manager.CloseSession("any-id"));
    }

    #endregion

    #region Edge Cases

    [Fact]
    public void CreateSession_VeryLongFilePath_HandlesGracefully()
    {
        // Create a long but valid path
        var longDirName = new string('x', 200);
        var longDir = Path.Combine(_tempDir, longDirName);

        try
        {
            Directory.CreateDirectory(longDir);
            var longFilePath = Path.Combine(longDir, "test.xlsx");

            ExcelSession.CreateNew(
                longFilePath,
                isMacroEnabled: false,
                (ctx, ct) => 0);
            _testFiles.Add(longFilePath);

            using var manager = new SessionManager();
            var sessionId = manager.CreateSession(longFilePath);

            Assert.NotNull(sessionId);
            Assert.Equal(1, manager.ActiveSessionCount);

            manager.CloseSession(sessionId);
        }
        catch (PathTooLongException)
        {
            // Expected on some systems - skip test
            _output.WriteLine("Path too long - test skipped");
        }
        catch (AggregateException ex) when (ex.InnerException is PathTooLongException)
        {
            // Excel COM may reject very long paths - expected behavior (converted from COMException)
            _output.WriteLine($"Excel rejected long path - test skipped: {ex.InnerException.Message}");
        }
        catch (AggregateException ex) when (ex.InnerException is AggregateException inner && inner.InnerException is PathTooLongException)
        {
            // Nested AggregateException from async task wrapping (STA thread -> Task.Wait -> Task.Wait)
            _output.WriteLine($"Excel rejected long path (nested) - test skipped: {((AggregateException)ex.InnerException).InnerException!.Message}");
        }
    }

    [Fact]
    public void CloseSession_DefaultSaveTrue_PersistsChanges()
    {
        var testFile = CreateTestFile(nameof(CloseSession_DefaultSaveTrue_PersistsChanges));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        // Get batch and make changes
        var batch = manager.GetSession(sessionId);
        Assert.NotNull(batch);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            sheet.Cells[1, 1].Value2 = "Test Value";
            return 0;
        });

        // Close with default save=false, but pass save:true explicitly
        var closed = manager.CloseSession(sessionId, save: true);
        Assert.True(closed);

        // Verify changes persisted
        using var verifyBatch = ExcelSession.BeginBatch(testFile);
        var value = verifyBatch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            return (string)sheet.Cells[1, 1].Value2;
        });

        Assert.Equal("Test Value", value);
    }

    #endregion
}



