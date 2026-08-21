using System.Collections.Concurrent;
using System.Reflection;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class SessionManagerTeardownFailureTests
{
    [Fact]
    public void CloseSession_OneShotDisposeFailure_RemainsQuarantinedUntilCleanupConfirmed()
    {
        using var manager = new SessionManager();
        var batch = new OneShotFailingBatch();
        const string sessionId = "one-shot-dispose-failure";
        var filePath = Path.GetFullPath("one-shot-dispose-failure.xlsx");
        RegisterSession(manager, sessionId, filePath, batch);

        var firstFailure = Assert.Throws<InvalidOperationException>(
            () => manager.CloseSession(sessionId));
        var retryFailure = Assert.Throws<InvalidOperationException>(
            () => manager.CloseSession(sessionId));

        Assert.Same(firstFailure, retryFailure);
        Assert.Equal(1, batch.DisposeCallCount);
        Assert.Equal(1, manager.ActiveSessionCount);
        Assert.Contains(sessionId, manager.ActiveSessionIds);
        Assert.True(manager.TryGetFilePath(sessionId, out var retainedPath));
        Assert.Equal(filePath, retainedPath);
        Assert.False(manager.TryBeginOperation(sessionId, out _, out var errorMessage));
        Assert.Contains("quarantined", errorMessage, StringComparison.OrdinalIgnoreCase);

        batch.ConfirmOwnedCleanup();

        Assert.True(manager.CloseSession(sessionId));
        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.False(manager.TryGetFilePath(sessionId, out _));
    }

    private static void RegisterSession(
        SessionManager manager,
        string sessionId,
        string filePath,
        IExcelBatch batch)
    {
        GetField<ConcurrentDictionary<string, IExcelBatch>>(manager, "_activeSessions")
            [sessionId] = batch;
        GetField<ConcurrentDictionary<string, string>>(manager, "_activeFilePaths")
            [filePath] = sessionId;
        GetField<ConcurrentDictionary<string, string>>(manager, "_sessionFilePaths")
            [sessionId] = filePath;
        GetField<ConcurrentDictionary<string, int>>(manager, "_activeOperationCounts")
            [sessionId] = 0;
    }

    private static T GetField<T>(SessionManager manager, string name) where T : class =>
        Assert.IsType<T>(
            typeof(SessionManager)
                .GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!
                .GetValue(manager));

    private sealed class OneShotFailingBatch : IExcelBatch, IExcelBatchTeardownState
    {
        private bool _cleanupConfirmed;

        public int DisposeCallCount { get; private set; }

        public string WorkbookPath => Path.GetFullPath("one-shot-dispose-failure.xlsx");

        public Microsoft.Extensions.Logging.ILogger Logger => NullLogger.Instance;

        public IReadOnlyDictionary<string, Excel.Workbook> Workbooks =>
            new Dictionary<string, Excel.Workbook>();

        public bool HasTimedOutOperation => false;

        public int? ExcelProcessId => Environment.ProcessId;

        public TimeSpan OperationTimeout => TimeSpan.FromSeconds(1);

        public bool IsExcelVisible => false;

        public void Dispose()
        {
            DisposeCallCount++;
            if (DisposeCallCount == 1)
            {
                throw new InvalidOperationException("synthetic teardown failure");
            }
        }

        public void ConfirmOwnedCleanup() => _cleanupConfirmed = true;

        public bool TryConfirmOwnedProcessTeardown() => _cleanupConfirmed;

        public bool IsExcelProcessAlive() => !_cleanupConfirmed;

        public void UpdateWorkbookPath(string workbookPath) =>
            throw new NotSupportedException();

        public Excel.Workbook GetWorkbook(string filePath) =>
            throw new NotSupportedException();

        public void Execute(
            Action<ExcelContext, CancellationToken> operation,
            CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public T Execute<T>(
            Func<ExcelContext, CancellationToken, T> operation,
            CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public void Save(CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();
    }
}
