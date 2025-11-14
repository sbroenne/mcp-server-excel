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
[Trait("RunType", "OnDemand")]
[Trait("RequiresExcel", "true")]
[Collection("Sequential")] // Disable parallelization to avoid COM interference
public class SessionManagerTests : IAsyncLifetime
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

    public async Task InitializeAsync()
    {
        // Kill any existing Excel processes to ensure clean state
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
                await Task.Delay(2000);
            }
        }
        catch (Exception ex)
        {
            _output.WriteLine($"Warning: Failed to clean Excel processes: {ex.Message}");
        }
    }

    public async Task DisposeAsync()
    {
        // Delete test files
        foreach (var file in _testFiles)
        {
            try
            {
                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
            catch
            {
                // Best effort
            }
        }

        // Delete temp directory
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
            // Best effort
        }

        // Give Excel time to fully terminate
        await Task.Delay(1000);
    }

    private async Task<string> CreateTestFileAsync(string testName)
    {
        var fileName = $"{testName}_{Guid.NewGuid():N}.xlsx";
        var filePath = Path.Combine(_tempDir, fileName);

        // Create blank workbook
        await ExcelSession.CreateNewAsync(
            filePath,
            isMacroEnabled: false,
            async (ctx, ct) =>
            {
                // Just create the file - no operations needed
                return await Task.FromResult(0);
            });

        _testFiles.Add(filePath);
        return filePath;
    }

    #region Basic Session Lifecycle

    [Fact]
    public async Task CreateSession_ValidFile_ReturnsUniqueSessionId()
    {
        var testFile = await CreateTestFileAsync(nameof(CreateSession_ValidFile_ReturnsUniqueSessionId));
        await using var manager = new SessionManager();

        var sessionId = await manager.CreateSessionAsync(testFile);

        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        Assert.Equal(32, sessionId.Length); // GUID without hyphens
        Assert.Equal(1, manager.ActiveSessionCount);

        await manager.CloseSessionAsync(sessionId);
    }

    [Fact]
    public async Task CreateSession_NonExistentFile_ThrowsFileNotFoundException()
    {
        await using var manager = new SessionManager();
        var nonExistentFile = Path.Combine(_tempDir, "nonexistent.xlsx");

        var ex = await Assert.ThrowsAsync<FileNotFoundException>(
            async () => await manager.CreateSessionAsync(nonExistentFile));

        Assert.Contains("Excel file not found", ex.Message);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public async Task GetSession_ExistingSessionId_ReturnsValidBatch()
    {
        var testFile = await CreateTestFileAsync(nameof(GetSession_ExistingSessionId_ReturnsValidBatch));
        await using var manager = new SessionManager();
        var sessionId = await manager.CreateSessionAsync(testFile);

        var batch = manager.GetSession(sessionId);

        Assert.NotNull(batch);
        Assert.Equal(1, manager.ActiveSessionCount);

        await manager.CloseSessionAsync(sessionId);
    }

    [Fact]
    public async Task GetSession_NonExistentSessionId_ReturnsNull()
    {
        await using var manager = new SessionManager();

        var batch = manager.GetSession("nonexistent-session-id");

        Assert.Null(batch);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public async Task GetSession_NullOrWhitespaceSessionId_ReturnsNull()
    {
        await using var manager = new SessionManager();

        Assert.Null(manager.GetSession(null!));
        Assert.Null(manager.GetSession(""));
        Assert.Null(manager.GetSession("   "));
    }

    #endregion

    #region Save Operations

    [Fact]
    public async Task SaveSession_ExistingSession_ReturnsTrueAndKeepsSessionActive()
    {
        var testFile = await CreateTestFileAsync(nameof(SaveSession_ExistingSession_ReturnsTrueAndKeepsSessionActive));
        await using var manager = new SessionManager();
        var sessionId = await manager.CreateSessionAsync(testFile);

        var saved = await manager.SaveSessionAsync(sessionId);

        Assert.True(saved);
        Assert.Equal(1, manager.ActiveSessionCount);
        Assert.NotNull(manager.GetSession(sessionId));

        await manager.CloseSessionAsync(sessionId);
    }

    [Fact]
    public async Task SaveSession_NonExistentSession_ReturnsFalse()
    {
        await using var manager = new SessionManager();

        var saved = await manager.SaveSessionAsync("nonexistent-session-id");

        Assert.False(saved);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public async Task SaveSession_MultipleConsecutiveSaves_AllSucceed()
    {
        var testFile = await CreateTestFileAsync(nameof(SaveSession_MultipleConsecutiveSaves_AllSucceed));
        await using var manager = new SessionManager();
        var sessionId = await manager.CreateSessionAsync(testFile);

        var saved1 = await manager.SaveSessionAsync(sessionId);
        var saved2 = await manager.SaveSessionAsync(sessionId);
        var saved3 = await manager.SaveSessionAsync(sessionId);

        Assert.True(saved1);
        Assert.True(saved2);
        Assert.True(saved3);
        Assert.Equal(1, manager.ActiveSessionCount);

        await manager.CloseSessionAsync(sessionId);
    }

    #endregion

    #region Close Operations

    [Fact]
    public async Task CloseSession_ExistingSession_RemovesSessionAndReturnsTrue()
    {
        var testFile = await CreateTestFileAsync(nameof(CloseSession_ExistingSession_RemovesSessionAndReturnsTrue));
        await using var manager = new SessionManager();
        var sessionId = await manager.CreateSessionAsync(testFile);

        var closed = await manager.CloseSessionAsync(sessionId);

        Assert.True(closed);
        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.Null(manager.GetSession(sessionId));
    }

    [Fact]
    public async Task CloseSession_NonExistentSession_ReturnsFalse()
    {
        await using var manager = new SessionManager();

        var closed = await manager.CloseSessionAsync("nonexistent-session-id");

        Assert.False(closed);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public async Task CloseSession_NullOrWhitespaceSessionId_ReturnsFalse()
    {
        await using var manager = new SessionManager();

        Assert.False(await manager.CloseSessionAsync(null!));
        Assert.False(await manager.CloseSessionAsync(""));
        Assert.False(await manager.CloseSessionAsync("   "));
    }

    [Fact]
    public async Task CloseSession_AlreadyClosedSession_ReturnsFalse()
    {
        var testFile = await CreateTestFileAsync(nameof(CloseSession_AlreadyClosedSession_ReturnsFalse));
        await using var manager = new SessionManager();
        var sessionId = await manager.CreateSessionAsync(testFile);

        var closed1 = await manager.CloseSessionAsync(sessionId);
        var closed2 = await manager.CloseSessionAsync(sessionId);

        Assert.True(closed1);
        Assert.False(closed2);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    #endregion

    #region Multi-Session Scenarios

    [Fact(Timeout = 20000)]
    public async Task CreateMultipleSessions_DifferentFiles_TracksAllSessions()
    {
        var testFile1 = await CreateTestFileAsync($"{nameof(CreateMultipleSessions_DifferentFiles_TracksAllSessions)}_1");
        var testFile2 = await CreateTestFileAsync($"{nameof(CreateMultipleSessions_DifferentFiles_TracksAllSessions)}_2");
        await using var manager = new SessionManager();

        var sessionId1 = await manager.CreateSessionAsync(testFile1);
        await Task.Delay(1000); // Delay to avoid COM initialization conflicts
        var sessionId2 = await manager.CreateSessionAsync(testFile2);

        Assert.Equal(2, manager.ActiveSessionCount);
        Assert.Contains(sessionId1, manager.ActiveSessionIds);
        Assert.Contains(sessionId2, manager.ActiveSessionIds);

        await manager.CloseSessionAsync(sessionId1);
        await manager.CloseSessionAsync(sessionId2);
    }

    [Fact]
    public async Task ActiveSessionIds_ReflectsCurrentState()
    {
        var testFile1 = await CreateTestFileAsync($"{nameof(ActiveSessionIds_ReflectsCurrentState)}_1");
        var testFile2 = await CreateTestFileAsync($"{nameof(ActiveSessionIds_ReflectsCurrentState)}_2");
        await using var manager = new SessionManager();

        // Initially empty
        Assert.Empty(manager.ActiveSessionIds);

        // After creating sessions
        var sessionId1 = await manager.CreateSessionAsync(testFile1);
        var sessionId2 = await manager.CreateSessionAsync(testFile2);
        var activeIds = manager.ActiveSessionIds.ToList();

        Assert.Equal(2, activeIds.Count);
        Assert.Contains(sessionId1, activeIds);
        Assert.Contains(sessionId2, activeIds);

        // After closing one session
        await manager.CloseSessionAsync(sessionId1);
        activeIds = manager.ActiveSessionIds.ToList();

        Assert.Single(activeIds);
        Assert.Contains(sessionId2, activeIds);
        Assert.DoesNotContain(sessionId1, activeIds);

        await manager.CloseSessionAsync(sessionId2);
    }

    [Fact]
    public async Task CloseOneSession_DoesNotAffectOtherSessions()
    {
        var testFile1 = await CreateTestFileAsync($"{nameof(CloseOneSession_DoesNotAffectOtherSessions)}_1");
        var testFile2 = await CreateTestFileAsync($"{nameof(CloseOneSession_DoesNotAffectOtherSessions)}_2");
        await using var manager = new SessionManager();

        var sessionId1 = await manager.CreateSessionAsync(testFile1);
        var sessionId2 = await manager.CreateSessionAsync(testFile2);

        await manager.CloseSessionAsync(sessionId1);

        Assert.Equal(1, manager.ActiveSessionCount);
        Assert.Null(manager.GetSession(sessionId1));
        Assert.NotNull(manager.GetSession(sessionId2));

        await manager.CloseSessionAsync(sessionId2);
    }

    [Fact]
    public async Task CreateSession_SameFileAlreadyOpen_ThrowsInvalidOperationException()
    {
        var testFile = await CreateTestFileAsync(nameof(CreateSession_SameFileAlreadyOpen_ThrowsInvalidOperationException));
        await using var manager = new SessionManager();

        // First session succeeds
        var sessionId1 = await manager.CreateSessionAsync(testFile);
        Assert.NotNull(sessionId1);
        Assert.Equal(1, manager.ActiveSessionCount);

        // Second session with same file should fail fast
        var ex = await Assert.ThrowsAsync<InvalidOperationException>(
            async () => await manager.CreateSessionAsync(testFile));

        Assert.Contains("already open in another session", ex.Message);
        Assert.Contains("Excel cannot open the same file multiple times", ex.Message);
        Assert.Equal(1, manager.ActiveSessionCount); // Still only one session

        await manager.CloseSessionAsync(sessionId1);
    }

    [Fact]
    public async Task CreateSession_AfterClosingPrevious_AllowsReopeningFile()
    {
        var testFile = await CreateTestFileAsync(nameof(CreateSession_AfterClosingPrevious_AllowsReopeningFile));
        await using var manager = new SessionManager();

        // First session
        var sessionId1 = await manager.CreateSessionAsync(testFile);
        Assert.Equal(1, manager.ActiveSessionCount);

        // Close first session
        await manager.CloseSessionAsync(sessionId1);
        Assert.Equal(0, manager.ActiveSessionCount);

        // Should now be able to open same file again
        var sessionId2 = await manager.CreateSessionAsync(testFile);
        Assert.NotNull(sessionId2);
        Assert.NotEqual(sessionId1, sessionId2);
        Assert.Equal(1, manager.ActiveSessionCount);

        await manager.CloseSessionAsync(sessionId2);
    }

    #endregion

    #region Disposal and Post-Disposal

    [Fact]
    public async Task DisposeAsync_OneSession_ClosesAllSessions()
    {
        var testFile1 = await CreateTestFileAsync($"{nameof(DisposeAsync_OneSession_ClosesAllSessions)}_1");
        var manager = new SessionManager();

        var sessionId1 = await manager.CreateSessionAsync(testFile1);

        Assert.Equal(1, manager.ActiveSessionCount);
        await manager.DisposeAsync();

        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.Empty(manager.ActiveSessionIds);
    }

    [Fact(Timeout = 30000)]
    [Trait("RunType", "OnDemand")]
    public async Task DisposeAsync_TwoSessions_ClosesAllSessions()
    {
        var testFile1 = await CreateTestFileAsync($"{nameof(DisposeAsync_TwoSessions_ClosesAllSessions)}_1");
        var testFile2 = await CreateTestFileAsync($"{nameof(DisposeAsync_TwoSessions_ClosesAllSessions)}_2");
        var manager = new SessionManager();

        var sessionId1 = await manager.CreateSessionAsync(testFile1);
        var sessionId2 = await manager.CreateSessionAsync(testFile2);

        Assert.Equal(2, manager.ActiveSessionCount);

        // DisposeAsync handles sessions sequentially to avoid COM threading issues
        await manager.DisposeAsync();

        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.Empty(manager.ActiveSessionIds);
    }

    [Fact]
    public async Task DisposeAsync_EmptyManager_CompletesImmediately()
    {
        var manager = new SessionManager();

        await manager.DisposeAsync();

        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public async Task DisposeAsync_CalledMultipleTimes_DoesNotThrow()
    {
        var manager = new SessionManager();

        await manager.DisposeAsync();
        await manager.DisposeAsync();
        await manager.DisposeAsync();

        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public async Task CreateSession_AfterDisposal_ThrowsObjectDisposedException()
    {
        var testFile = await CreateTestFileAsync(nameof(CreateSession_AfterDisposal_ThrowsObjectDisposedException));
        var manager = new SessionManager();
        await manager.DisposeAsync();

        await Assert.ThrowsAsync<ObjectDisposedException>(
            async () => await manager.CreateSessionAsync(testFile));
    }

    [Fact]
    public async Task GetSession_AfterDisposal_ThrowsObjectDisposedException()
    {
        var manager = new SessionManager();
        await manager.DisposeAsync();

        Assert.Throws<ObjectDisposedException>(
            () => manager.GetSession("any-id"));
    }

    [Fact]
    public async Task SaveSession_AfterDisposal_ThrowsObjectDisposedException()
    {
        var manager = new SessionManager();
        await manager.DisposeAsync();

        await Assert.ThrowsAsync<ObjectDisposedException>(
            async () => await manager.SaveSessionAsync("any-id"));
    }

    [Fact]
    public async Task CloseSession_AfterDisposal_ThrowsObjectDisposedException()
    {
        var manager = new SessionManager();
        await manager.DisposeAsync();

        await Assert.ThrowsAsync<ObjectDisposedException>(
            async () => await manager.CloseSessionAsync("any-id"));
    }

    #endregion

    #region Edge Cases

    [Fact]
    public async Task CreateSession_VeryLongFilePath_HandlesGracefully()
    {
        // Create a long but valid path
        var longDirName = new string('x', 200);
        var longDir = Path.Combine(_tempDir, longDirName);

        try
        {
            Directory.CreateDirectory(longDir);
            var longFilePath = Path.Combine(longDir, "test.xlsx");

            await ExcelSession.CreateNewAsync(
                longFilePath,
                isMacroEnabled: false,
                async (ctx, ct) => await Task.FromResult(0));
            _testFiles.Add(longFilePath);

            await using var manager = new SessionManager();
            var sessionId = await manager.CreateSessionAsync(longFilePath);

            Assert.NotNull(sessionId);
            Assert.Equal(1, manager.ActiveSessionCount);

            await manager.CloseSessionAsync(sessionId);
        }
        catch (PathTooLongException)
        {
            // Expected on some systems - skip test
            _output.WriteLine("Path too long - test skipped");
        }
    }

    [Fact]
    public async Task SaveSession_AfterDataModification_PersistsChanges()
    {
        var testFile = await CreateTestFileAsync(nameof(SaveSession_AfterDataModification_PersistsChanges));
        await using var manager = new SessionManager();
        var sessionId = await manager.CreateSessionAsync(testFile);

        // Get batch and make changes
        var batch = manager.GetSession(sessionId);
        Assert.NotNull(batch);

        await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            sheet.Cells[1, 1].Value2 = "Test Value";
            return await Task.FromResult(0);
        });

        // Save changes
        var saved = await manager.SaveSessionAsync(sessionId);
        Assert.True(saved);

        await manager.CloseSessionAsync(sessionId);

        // Verify changes persisted
        await using var verifyBatch = await ExcelSession.BeginBatchAsync(testFile);
        var value = await verifyBatch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            return await Task.FromResult((string)sheet.Cells[1, 1].Value2);
        });

        Assert.Equal("Test Value", value);
    }

    #endregion
}
