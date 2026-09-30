using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration;

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
    }

    /// <summary>
    /// Path to the template xlsx file used for fast test file creation.
    /// Copying a template is ~1000x faster than spawning Excel to create a new workbook.
    /// </summary>
    private static readonly string TemplateFilePath = Path.Combine(
        Path.GetDirectoryName(typeof(SessionManagerTests).Assembly.Location)!,
        "Integration", "Session", "TestFiles", "batch-test-static.xlsx");

    private string CreateTestFile(string testName)
    {
        var fileName = $"{testName}_{Guid.NewGuid():N}.xlsx";
        var filePath = Path.Combine(_tempDir, fileName);

        // PERFORMANCE OPTIMIZATION: Copy from template instead of spawning Excel.
        // This reduces test file creation from ~7-14 seconds to <10ms.
        // Copying the saved fixture avoids spawning a full Excel process
        // for each test file, causing 30+ second test execution times.
        File.Copy(TemplateFilePath, filePath);

        _testFiles.Add(filePath);
        return filePath;
    }

    private string CreateTestFileWithPathLength(int length)
    {
        const string fileName = "test.xlsx";
        var directoryNameLength = length - _tempDir.Length - fileName.Length - 2;
        Assert.InRange(directoryNameLength, 1, 255);
        var directory = Path.Combine(_tempDir, new string('x', directoryNameLength));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, fileName);
        Assert.Equal(length, path.Length);
        File.Copy(TemplateFilePath, path);
        _testFiles.Add(path);
        return path;
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

    #region Save As Path Reservation

    [Fact]
    public void ReserveSessionFilePath_ConcurrentSessions_AllowsOnlyOneOwner()
    {
        var firstFile = CreateTestFile(nameof(ReserveSessionFilePath_ConcurrentSessions_AllowsOnlyOneOwner) + "_First");
        var secondFile = CreateTestFile(nameof(ReserveSessionFilePath_ConcurrentSessions_AllowsOnlyOneOwner) + "_Second");
        var targetPath = Path.Combine(_tempDir, "SharedSaveAsTarget.xlsx");
        using var manager = new SessionManager();
        var firstSessionId = manager.CreateSession(firstFile);
        var secondSessionId = manager.CreateSession(secondFile);
        var reservations = new ConcurrentBag<(string SessionId, string Path)>();
        var failures = new ConcurrentBag<Exception>();
        using var start = new Barrier(2);

        void Reserve(string sessionId, string path)
        {
            start.SignalAndWait();
            try
            {
                reservations.Add((sessionId, manager.ReserveSessionFilePath(sessionId, path)));
            }
            catch (Exception ex)
            {
                failures.Add(ex);
            }
        }

        Parallel.Invoke(
            () => Reserve(firstSessionId, targetPath),
            () => Reserve(secondSessionId, targetPath.ToUpperInvariant()));

        var reservation = Assert.Single(reservations);
        Assert.IsType<InvalidOperationException>(Assert.Single(failures));
        manager.ReleaseSessionFilePathReservation(reservation.SessionId, reservation.Path);
        manager.CloseSession(firstSessionId);
        manager.CloseSession(secondSessionId);
    }

    [Fact]
    public void CloseSession_ReleasesOutstandingPathReservation()
    {
        var firstFile = CreateTestFile(nameof(CloseSession_ReleasesOutstandingPathReservation) + "_First");
        var secondFile = CreateTestFile(nameof(CloseSession_ReleasesOutstandingPathReservation) + "_Second");
        var targetPath = Path.Combine(_tempDir, "ReleasedSaveAsTarget.xlsx");
        using var manager = new SessionManager();
        var firstSessionId = manager.CreateSession(firstFile);
        var secondSessionId = manager.CreateSession(secondFile);

        manager.ReserveSessionFilePath(firstSessionId, targetPath);
        manager.CloseSession(firstSessionId);

        var reservation = manager.ReserveSessionFilePath(secondSessionId, targetPath);
        Assert.Equal(Path.GetFullPath(targetPath), reservation, ignoreCase: true);
        manager.ReleaseSessionFilePathReservation(secondSessionId, reservation);
        manager.CloseSession(secondSessionId);
    }

    [Fact]
    public void ReserveSessionFilePath_SameSessionRejectsConcurrentTarget()
    {
        var sourceFile = CreateTestFile(nameof(ReserveSessionFilePath_SameSessionRejectsConcurrentTarget));
        var firstTargetPath = Path.Combine(_tempDir, "FirstSaveAsTarget.xlsx");
        var secondTargetPath = Path.Combine(_tempDir, "SecondSaveAsTarget.xlsx");
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(sourceFile);

        var firstReservation = manager.ReserveSessionFilePath(sessionId, firstTargetPath);
        var exception = Assert.Throws<InvalidOperationException>(
            () => manager.ReserveSessionFilePath(sessionId, secondTargetPath));
        Assert.Contains("already has a Save As operation in progress", exception.Message, StringComparison.Ordinal);

        manager.ReleaseSessionFilePathReservation(sessionId, firstReservation);
        var secondReservation = manager.ReserveSessionFilePath(sessionId, secondTargetPath);
        Assert.Equal(Path.GetFullPath(secondTargetPath), secondReservation, ignoreCase: true);
        manager.ReleaseSessionFilePathReservation(sessionId, secondReservation);
        manager.CloseSession(sessionId);
    }

    [Fact]
    public void CreateSessionForNewFile_ReservedPath_RejectsBeforeStartingExcel()
    {
        var sourceFile = CreateTestFile(nameof(CreateSessionForNewFile_ReservedPath_RejectsBeforeStartingExcel));
        var targetPath = Path.Combine(_tempDir, "ReservedNewWorkbook.xlsx");
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(sourceFile);
        manager.ReserveSessionFilePath(sessionId, targetPath);

        var exception = Assert.Throws<InvalidOperationException>(
            () => manager.CreateSessionForNewFile(targetPath));

        Assert.Contains("already open or reserved", exception.Message, StringComparison.Ordinal);
        manager.ReleaseSessionFilePathReservation(sessionId, targetPath);
        manager.CloseSession(sessionId);
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
    public void CreateSession_FileLockedByAnotherProcess_DoesNotLeakExcelProcess()
    {
        var testFile = CreateTestFile(nameof(CreateSession_FileLockedByAnotherProcess_DoesNotLeakExcelProcess));
        using var manager = new SessionManager();

        using var owned = new OwnedExcelProcessScope();

        using (var fileLock = new FileStream(testFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            var ex = Assert.Throws<InvalidOperationException>(() => manager.CreateSession(testFile));
            Assert.Contains("already open", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        owned.AssertAllExited(expectProcess: false);
        Assert.Equal(0, manager.ActiveSessionCount);
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

        manager.CreateSession(testFile1);
        manager.CreateSession(testFile2);

        Assert.Equal(2, manager.ActiveSessionCount);

        // DisposeAsync handles sessions sequentially to avoid COM threading issues
        manager.Dispose();

        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.Empty(manager.ActiveSessionIds);
    }

    [Fact]
    public void Dispose_EmptyManager_CompletesImmediately()
    {
        using var manager = new SessionManager();

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
    public void CreateSession_AtDocumentedPathLimit_OpensAndCloses()
    {
        // Excel documents a 218-character limit including the full path.
        var path = CreateTestFileWithPathLength(218);
        using var owned = new OwnedExcelProcessScope();
        var manager = new SessionManager();
        var operationFailure = Record.Exception(() =>
        {
            var sessionId = manager.CreateSession(path);

            Assert.False(string.IsNullOrWhiteSpace(sessionId));
            Assert.Equal(1, manager.ActiveSessionCount);
            Assert.Equal(path, Assert.IsAssignableFrom<IExcelBatch>(manager.GetSession(sessionId)).WorkbookPath);
            Assert.True(manager.CloseSession(sessionId, save: false));
            Assert.Equal(0, manager.ActiveSessionCount);
            Assert.Empty(manager.ActiveSessionIds);
        });
        var disposalFailure = Record.Exception(manager.Dispose);
        var processFailure = Record.Exception(() => owned.AssertAllExited());
        Assert.All(new[] { operationFailure, disposalFailure, processFailure }, failure => Assert.Null(failure));
    }

    [Fact]
    public void CreateSession_OverlongFilePath_RejectsAndReleasesResources()
    {
        // Some Excel versions open paths beyond 218 characters; exceed the legacy Windows limit too.
        var path = CreateTestFileWithPathLength(300);
        using var owned = new OwnedExcelProcessScope();
        var manager = new SessionManager();
        var operationFailure = Record.Exception(() =>
        {
            // Retry the same path to prove failed startup released its reservation.
            for (var attempt = 0; attempt < 2; attempt++)
            {
                var failure = Assert.Throws<InvalidOperationException>(() => manager.CreateSession(path));
                var excelFailure = Assert.IsType<COMException>(failure.InnerException);
                Assert.Equal(unchecked((int)0x800A03EC), excelFailure.HResult);
                Assert.False(string.IsNullOrWhiteSpace(excelFailure.Message));
                Assert.Equal(0, manager.ActiveSessionCount);
                Assert.Empty(manager.ActiveSessionIds);

                using var file = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                Assert.True(file.Length > 0);
            }

            var shorterPath = Path.Combine(_tempDir, "recovered.xlsx");
            File.Move(path, shorterPath);
            _testFiles.Add(shorterPath);
            var sessionId = manager.CreateSession(shorterPath);
            Assert.Equal(1, manager.ActiveSessionCount);
            Assert.True(manager.CloseSession(sessionId, save: false));
            Assert.Equal(0, manager.ActiveSessionCount);
            Assert.Empty(manager.ActiveSessionIds);
        });
        var disposalFailure = Record.Exception(manager.Dispose);
        var processFailure = Record.Exception(() => owned.AssertAllExited());
        Assert.All(new[] { operationFailure, disposalFailure, processFailure }, failure => Assert.Null(failure));
    }

    [Fact]
    public void CloseSession_SaveTrue_PersistsChanges()
    {
        var testFile = CreateTestFile(nameof(CloseSession_SaveTrue_PersistsChanges));
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
