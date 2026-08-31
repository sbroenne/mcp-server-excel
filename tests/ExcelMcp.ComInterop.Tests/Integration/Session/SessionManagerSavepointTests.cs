using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "WorkbookSavepoints")]
[Trait("RequiresExcel", "true")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")]
public sealed class SessionManagerSavepointTests : IDisposable
{
    private static readonly string TemplateFilePath = Path.Combine(
        Path.GetDirectoryName(typeof(SessionManagerSavepointTests).Assembly.Location)!,
        "Integration", "Session", "TestFiles", "batch-test-static.xlsx");

    private readonly string _tempDirectory = Path.Combine(
        Path.GetTempPath(),
        $"SessionManagerSavepointTests_{Guid.NewGuid():N}");

    public SessionManagerSavepointTests()
    {
        Directory.CreateDirectory(_tempDirectory);
    }

    [Fact]
    public void RollbackSavepoint_RestoresUnsavedStateAndPreservesSessionIdentity()
    {
        var workbookPath = CreateWorkbook();
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(workbookPath);
        var originalBatch = manager.GetSession(sessionId)!;

        SetCell(originalBatch, "before");
        var savepoint = manager.CreateSavepoint(sessionId, "before-change");
        SetCell(originalBatch, "after");

        var rollback = manager.RollbackSavepoint(sessionId, savepoint.Name);
        var restoredBatch = manager.GetSession(sessionId);

        Assert.NotNull(restoredBatch);
        Assert.Same(originalBatch, restoredBatch);
        Assert.Equal(sessionId, rollback.SessionId);
        Assert.Equal(Path.GetFullPath(workbookPath), rollback.WorkbookPath, ignoreCase: true);
        Assert.Equal("before", GetCell(restoredBatch));
        Assert.Single(manager.GetSavepoints(sessionId));

        manager.CloseSession(sessionId, save: false);
        using var persistedBatch = ExcelSession.BeginBatch(workbookPath);
        Assert.Equal("before", GetCell(persistedBatch));
    }

    [Fact]
    public void ReleaseSavepoint_RemovesOnlyTheRequestedSnapshot()
    {
        var workbookPath = CreateWorkbook();
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(workbookPath);

        manager.CreateSavepoint(sessionId, "first");
        manager.CreateSavepoint(sessionId, "second");

        Assert.True(manager.ReleaseSavepoint(sessionId, "first"));
        var remaining = Assert.Single(manager.GetSavepoints(sessionId));
        Assert.Equal("second", remaining.Name);
        Assert.False(manager.ReleaseSavepoint(sessionId, "missing"));

        manager.CloseSession(sessionId);
    }

    [Fact]
    public void CreateSavepoint_DuplicateNameIsRejectedWithoutReplacingSnapshot()
    {
        var workbookPath = CreateWorkbook();
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(workbookPath);

        SetCell(manager.GetSession(sessionId)!, "original");
        manager.CreateSavepoint(sessionId, "stable");
        SetCell(manager.GetSession(sessionId)!, "changed");

        var exception = Assert.Throws<InvalidOperationException>(
            () => manager.CreateSavepoint(sessionId, "stable"));
        Assert.Contains("already exists", exception.Message, StringComparison.Ordinal);

        manager.RollbackSavepoint(sessionId, "stable");
        Assert.Equal("original", GetCell(manager.GetSession(sessionId)!));
        manager.CloseSession(sessionId);
    }

    [Theory]
    [InlineData("")]
    [InlineData("contains space")]
    [InlineData("../escape")]
    [InlineData("slash/name")]
    public void CreateSavepoint_InvalidNameIsRejected(string name)
    {
        var workbookPath = CreateWorkbook();
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(workbookPath);

        Assert.Throws<ArgumentException>(() => manager.CreateSavepoint(sessionId, name));
        Assert.Empty(manager.GetSavepoints(sessionId));
        manager.CloseSession(sessionId);
    }

    [Fact]
    public void CreateSavepoint_NinthSnapshotIsRejectedBySessionLimit()
    {
        var workbookPath = CreateWorkbook();
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(workbookPath);

        for (var index = 0; index < SessionManager.SavepointLimits.MaxSavepointsPerSession; index++)
        {
            manager.CreateSavepoint(sessionId, $"point-{index}");
        }

        var exception = Assert.Throws<WorkbookSavepointStorageLimitException>(
            () => manager.CreateSavepoint(sessionId, "over-limit"));
        Assert.Contains("maximum", exception.Message, StringComparison.Ordinal);
        Assert.Equal(
            SessionManager.SavepointLimits.MaxSavepointsPerSession,
            manager.GetSavepoints(sessionId).Count);
        manager.CloseSession(sessionId);
    }

    [Fact]
    public void CreateSavepoint_ReadOnlyWorkbookIsRejected()
    {
        var workbookPath = CreateWorkbook();
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(workbookPath);
        manager.GetSession(sessionId)!.Execute((context, _) =>
            context.Book.ChangeFileAccess(Excel.XlFileAccess.xlReadOnly));

        var exception = Assert.Throws<InvalidOperationException>(
            () => manager.CreateSavepoint(sessionId, "read-only"));
        Assert.Contains("read-only", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(manager.GetSavepoints(sessionId));
        manager.CloseSession(sessionId);
    }

    [Fact]
    public void RollbackSavepoint_ReopenFailureRecoversPreRollbackState()
    {
        var workbookPath = CreateWorkbook();
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(workbookPath);
        var batch = manager.GetSession(sessionId)!;
        SetCell(batch, "checkpoint");
        manager.CreateSavepoint(sessionId, "stable");
        SetCell(batch, "current");

        var reopenAttempts = 0;
        ExcelBatch.BeforeWorkbookRestoreOpenHook = _ =>
        {
            if (Interlocked.Increment(ref reopenAttempts) == 1)
            {
                throw new IOException("Injected rollback reopen failure.");
            }
        };

        try
        {
            var exception = Assert.Throws<WorkbookSavepointRollbackException>(
                () => manager.RollbackSavepoint(sessionId, "stable"));
            Assert.True(exception.SessionRecovered);
            Assert.False(exception.SessionClosed);
            Assert.Same(batch, manager.GetSession(sessionId));
            Assert.Equal("current", GetCell(batch));
            Assert.Single(manager.GetSavepoints(sessionId));
        }
        finally
        {
            ExcelBatch.BeforeWorkbookRestoreOpenHook = null;
            manager.CloseSession(sessionId);
        }
    }

    [Fact]
    public void RollbackSavepoint_RecoveryFailureClosesSessionAndPreservesRecoveryFile()
    {
        var workbookPath = CreateWorkbook();
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(workbookPath);
        var batch = manager.GetSession(sessionId)!;
        SetCell(batch, "checkpoint");
        manager.CreateSavepoint(sessionId, "stable");
        SetCell(batch, "current");
        ExcelBatch.BeforeWorkbookRestoreOpenHook = _ =>
            throw new IOException("Injected reopen failure.");

        string? recoveryFilePath = null;
        try
        {
            var exception = Assert.Throws<WorkbookSavepointRollbackException>(
                () => manager.RollbackSavepoint(sessionId, "stable"));
            recoveryFilePath = exception.RecoveryFilePath;

            Assert.False(exception.SessionRecovered);
            Assert.True(exception.SessionClosed);
            Assert.Null(manager.GetSession(sessionId));
            Assert.False(string.IsNullOrWhiteSpace(recoveryFilePath));
            Assert.True(File.Exists(recoveryFilePath));
        }
        finally
        {
            ExcelBatch.BeforeWorkbookRestoreOpenHook = null;
            if (recoveryFilePath != null && File.Exists(recoveryFilePath))
            {
                File.Delete(recoveryFilePath);
            }
        }
    }

    [Fact]
    public void EndOperation_DoesNotDisposeSavepointsOwnedByOtherSessions()
    {
        var firstWorkbook = CreateWorkbook();
        var secondWorkbook = CreateWorkbook();
        using var manager = new SessionManager();
        var firstSession = manager.CreateSession(firstWorkbook);
        var secondSession = manager.CreateSession(secondWorkbook);

        manager.CreateSavepoint(secondSession, "second-session");
        manager.BeginOperation(firstSession);
        manager.EndOperation(firstSession);

        var savepoint = Assert.Single(manager.GetSavepoints(secondSession));
        Assert.Equal("second-session", savepoint.Name);
        manager.CloseSession(firstSession);
        manager.CloseSession(secondSession);
    }

    [Fact]
    public void SavepointMutation_IsRejectedWhileAnotherOperationIsActive()
    {
        var workbookPath = CreateWorkbook();
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(workbookPath);
        manager.CreateSavepoint(sessionId, "stable");
        manager.BeginOperation(sessionId);

        try
        {
            Assert.Throws<InvalidOperationException>(
                () => manager.CreateSavepoint(sessionId, "concurrent"));
            Assert.Throws<InvalidOperationException>(
                () => manager.RollbackSavepoint(sessionId, "stable"));
        }
        finally
        {
            manager.EndOperation(sessionId);
            manager.CloseSession(sessionId);
        }
    }

    [Fact]
    public void RollbackSavepoint_PreCancelledRequestLeavesSessionAndUnsavedStateIntact()
    {
        var workbookPath = CreateWorkbook();
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(workbookPath);
        var batch = manager.GetSession(sessionId)!;
        SetCell(batch, "checkpoint");
        manager.CreateSavepoint(sessionId, "stable");
        SetCell(batch, "current");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            manager.RollbackSavepoint(
                sessionId,
                "stable",
                cancellationToken: cancellation.Token));

        Assert.Same(batch, manager.GetSession(sessionId));
        Assert.Equal("current", GetCell(batch));
        Assert.Single(manager.GetSavepoints(sessionId));
        manager.CloseSession(sessionId);
    }

    private string CreateWorkbook()
    {
        var path = Path.Combine(_tempDirectory, $"{Guid.NewGuid():N}.xlsx");
        File.Copy(TemplateFilePath, path);
        return path;
    }

    private static void SetCell(IExcelBatch batch, string value)
    {
        batch.Execute((context, _) =>
        {
            context.Book.Worksheets[1].Cells[1, 1].Value2 = value;
        });
    }

    private static string? GetCell(IExcelBatch batch)
    {
        return batch.Execute((context, _) =>
            Convert.ToString(context.Book.Worksheets[1].Cells[1, 1].Value2));
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDirectory))
        {
            Directory.Delete(_tempDirectory, recursive: true);
        }
    }
}
