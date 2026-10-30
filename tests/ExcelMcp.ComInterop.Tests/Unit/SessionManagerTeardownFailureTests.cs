using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;
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
[Collection("Sequential")]
public sealed class SessionManagerTeardownFailureTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SuccessfulRelease_RecordsTheExactIdentityBeforeEndingOwnership(bool recordingFails)
    {
        var released = new ExcelProcessIdentity(int.MaxValue, 1);
        var other = new ExcelProcessIdentity(int.MaxValue, 2);
        var failure = new IOException("Test-only ownership record failure.");
        var observed = new List<ExcelProcessIdentity>();
        SessionManager.TrackExcelProcess(released);
        SessionManager.TrackExcelProcess(other);
        SessionManager.PreserveExcelProcess(other);
        void Record(ExcelProcessIdentity identity)
        {
            observed.Add(identity);
            Assert.True(SessionManager.IsExcelProcessTracked(identity));
            if (recordingFails)
                throw failure;
        }
        SessionManager.ExcelProcessIdentityReleased += Record;
        try
        {
            if (recordingFails)
                Assert.Same(failure, Assert.Throws<IOException>(() => SessionManager.UntrackExcelProcess(released)));
            else
                SessionManager.UntrackExcelProcess(released);
            Assert.Equal(released, Assert.Single(observed));
            Assert.Equal(recordingFails, SessionManager.IsExcelProcessTracked(released));
            Assert.True(SessionManager.IsExcelProcessTracked(other));
            Assert.True(SessionManager.IsExcelProcessPreserved(other));
        }
        finally
        {
            SessionManager.ExcelProcessIdentityReleased -= Record;
            SessionManager.UntrackExcelProcess(released);
            SessionManager.UntrackExcelProcess(other);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Close_Timings_RecordStagesWithoutChangingSaveOrFailure(bool saveFails)
    {
        using var manager = new SessionManager();
        var batch = new ConfigurableFailingBatch(saveFails, false)
        {
            BeforeSave = () => Thread.Sleep(30)
        };
        RegisterSession(manager, "timed-close", batch.WorkbookPath, batch);
        using var timings = SessionDiagnostics.BeginCloseTimings();
        try
        {
            if (saveFails)
            {
                var failure = Assert.Throws<InvalidOperationException>(
                    () => manager.CloseSession("timed-close", save: true));
                Assert.Same(batch.SaveException, failure.InnerException);
                Assert.Equal(1, manager.ActiveSessionCount);
                Assert.Equal(0, batch.CloseCallCount);
            }
            else
            {
                Assert.True(manager.CloseSession("timed-close", save: true));
                Assert.Equal(0, manager.ActiveSessionCount);
                Assert.Equal(1, batch.CloseCallCount);
            }
            Assert.Equal(1, batch.SaveCallCount);

            var measured = timings.Snapshot();
            Assert.True(measured["save"] >= 20);
            Assert.Contains("readiness", measured.Keys);
            Assert.Contains("discovery-release", measured.Keys);
            Assert.Equal(!saveFails, measured.ContainsKey("session-teardown"));
            var json = JsonSerializer.Serialize(measured);
            Assert.DoesNotContain("timed-close", json, StringComparison.Ordinal);
            Assert.DoesNotContain(batch.WorkbookPath, json, StringComparison.Ordinal);
        }
        finally
        {
            batch.SaveFails = false;
        }
    }

    [Fact]
    public async Task Close_Timings_PreserveOriginalFailureAndCancellation()
    {
        using var timings = SessionDiagnostics.BeginCloseTimings();
        var failure = new InvalidOperationException("Original close failure.");
        var calls = 0;
        Action fail = () => { calls++; throw failure; };
        Assert.Same(failure, Assert.Throws<InvalidOperationException>(
            () => SessionDiagnostics.MeasureCloseStage("workbook-close", fail)));
        Assert.Equal(1, calls);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var cancelled = new OperationCanceledException(cancellation.Token);
        async Task FailAsync()
        {
            using var timing = SessionDiagnostics.BeginCloseStage("worker-completion");
            await Task.FromException(cancelled);
        }
        Assert.Same(cancelled, await Assert.ThrowsAsync<OperationCanceledException>(FailAsync));
        Assert.Contains("workbook-close", timings.Snapshot().Keys);
        Assert.Contains("worker-completion", timings.Snapshot().Keys);
    }

    [Fact]
    public async Task Close_Timings_IsolateConcurrentRequestsAndRestoreOuterScope()
    {
        using var outer = SessionDiagnostics.BeginCloseTimings();
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var arrivals = 0;
        async Task<Dictionary<string, double>> Measure(string stage)
        {
            using var timings = SessionDiagnostics.BeginCloseTimings();
            using (SessionDiagnostics.BeginCloseStage(stage))
            {
                if (Interlocked.Increment(ref arrivals) == 2)
                    started.SetResult();
                await started.Task;
                await Task.Delay(20);
            }
            return timings.Snapshot();
        }
        var measured = await Task.WhenAll(Task.Run(() => Measure("first")), Task.Run(() => Measure("second")));
        Assert.Equal("first", Assert.Single(measured[0]).Key);
        Assert.Equal("second", Assert.Single(measured[1]).Key);
        Assert.All(measured, result => Assert.True(Assert.Single(result).Value >= 10));
        Assert.Same(outer, SessionDiagnostics.CurrentCloseTimings);
        Assert.Empty(outer.Snapshot());
        outer.Record("workbook-release", 1.234);
        outer.Record("workbook-release", 1.234);
        Assert.Equal(2.47, outer.Snapshot()["workbook-release"]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Dispose_ExitedExcel_BypassesReadinessButRequiresSuccessfulTeardown(bool disposeFails)
    {
        using var manager = new SessionManager();
        var batch = new ConfigurableFailingBatch(false, disposeFails)
        {
            ProcessAlive = false,
            RefreshState = WorkbookRefreshState.Unknown
        };
        var path = Path.GetFullPath("exited-before-shutdown.xlsx");
        RegisterSession(manager, "exited", path, batch);
        try
        {
            if (disposeFails)
            {
                var failure = Assert.Throws<AggregateException>(manager.Dispose);
                Assert.Contains("synthetic teardown failure", failure.ToString());
                Assert.Same(batch, GetField<ConcurrentDictionary<string, IExcelBatch>>(manager, "_activeSessions")["exited"]);
                Assert.True(manager.TryGetFilePath("exited", out var retainedPath));
                Assert.Equal(path, retainedPath);
                Assert.Equal(1, batch.DisposeCallCount);
                batch.DisposeFails = false;
            }
            manager.Dispose();
            Assert.Equal(disposeFails ? 2 : 1, batch.DisposeCallCount);
            Assert.Equal(0, batch.CloseCallCount);
            Assert.Equal(0, batch.SaveCallCount);
            Assert.Equal(0, manager.ActiveSessionCount);
            Assert.False(manager.HasActiveConnections);
            Assert.Empty(GetField<ConcurrentDictionary<string, string>>(manager, "_sessionFilePaths"));
            Assert.Empty(GetField<ConcurrentDictionary<string, string>>(manager, "_activeFilePaths"));
        }
        finally
        {
            batch.DisposeFails = false;
            batch.RefreshState = WorkbookRefreshState.Ready;
        }
    }

    [Fact]
    public void Dispose_DoesNotRequestAnImplicitSave()
    {
        using var manager = new SessionManager();
        var batch = new ConfigurableFailingBatch(false, false);
        RegisterSession(manager, "no-implicit-save", Path.GetFullPath("no-implicit-save.xlsx"), batch);

        manager.Dispose();

        Assert.Equal(0, batch.SaveCallCount);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void ValidateClose_ReadinessRevealsExcel_ReportsUpdatedVisibility()
    {
        using var manager = new SessionManager();
        var batch = new ConfigurableFailingBatch(false, false) { RevealOnRefreshRead = true };
        RegisterSession(manager, "revealed", Path.GetFullPath("revealed.xlsx"), batch);
        Assert.False(batch.IsExcelVisible);
        var validation = manager.ValidateClose("revealed");
        Assert.True(batch.IsExcelVisible);
        Assert.True(validation.CanClose);
        Assert.Equal(WorkbookRefreshState.Ready, validation.RefreshState);
        Assert.True(validation.IsExcelVisible);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Close_UncommittedFailure_OrdinaryShutdownDoesNotSave(bool cancelled)
    {
        using var manager = new SessionManager();
        Exception failure = cancelled
            ? new OperationCanceledException("Cancelled before close.")
            : new InvalidOperationException("Readiness failed before close.");
        var batch = new ConfigurableFailingBatch(false, false) { BeforeCloseFailure = failure };
        RegisterSession(manager, "uncommitted", Path.GetFullPath("uncommitted-close.xlsx"), batch);
        Assert.Same(failure, Record.Exception(() => manager.CloseSession("uncommitted")));
        Assert.False(batch.ShutdownCommitted);
        Assert.Equal(0, batch.DisposeCallCount);
        Assert.Same(batch, manager.GetSession("uncommitted"));
        batch.BeforeCloseFailure = null;
        manager.Dispose();
        Assert.Equal(0, batch.SaveCallCount);
        Assert.Equal(2, batch.CloseCallCount);
        Assert.Equal(1, batch.DisposeCallCount);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void Dispose_UncommittedCloseFailure_AllowsEditsWithoutSavingOnRetry()
    {
        using var manager = new SessionManager();
        var batch = new ConfigurableFailingBatch(false, false)
        {
            BeforeCloseFailure = new InvalidOperationException("Readiness failed before close.")
        };
        RegisterSession(manager, "shutdown-uncommitted", Path.GetFullPath("uncommitted-shutdown.xlsx"), batch);
        Assert.Throws<AggregateException>(manager.Dispose);
        Assert.False(batch.ShutdownCommitted);
        Assert.Equal(0, batch.SaveCallCount);
        Assert.True(manager.TryBeginOperation("shutdown-uncommitted", out var retained, out _));
        Assert.Same(batch, retained);
        manager.EndOperation("shutdown-uncommitted");
        batch.BeforeCloseFailure = null;
        manager.Dispose();
        Assert.Equal(0, batch.SaveCallCount);
        Assert.Equal(1, batch.DisposeCallCount);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void Dispose_PreviousTeardownFailure_ResumesDisposalWithoutClosingWorkbookAgain()
    {
        using var manager = new SessionManager();
        var batch = new ConfigurableFailingBatch(false, true);
        var path = Path.GetFullPath("pending-teardown.xlsx");
        RegisterSession(manager, "pending-teardown", path, batch);
        Assert.Throws<InvalidOperationException>(() => manager.CloseSession("pending-teardown"));
        Assert.Equal(1, batch.CloseCallCount);
        Assert.Equal(1, batch.DisposeCallCount);
        batch.DisposeFails = false;
        batch.CloseBusy = true;
        try
        {
            manager.Dispose();
            Assert.Equal(1, batch.CloseCallCount);
            Assert.Equal(2, batch.DisposeCallCount);
            Assert.Equal(0, batch.SaveCallCount);
            Assert.Equal(0, manager.ActiveSessionCount);
            Assert.Empty(GetField<ConcurrentDictionary<string, string>>(manager, "_activeFilePaths"));
            Assert.Empty(GetField<ConcurrentDictionary<string, string>>(manager, "_sessionFilePaths"));
        }
        finally
        {
            batch.CloseBusy = false;
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Close_LateBusyRefusal_RetainsPathAndAllowsRetry(bool save)
    {
        using var manager = new SessionManager();
        var batch = new ConfigurableFailingBatch(false, false) { CloseBusy = true };
        var path = Path.GetFullPath("late-close-busy.xlsx");
        RegisterSession(manager, "late-busy", path, batch);
        try
        {
            Assert.True(manager.ValidateClose("late-busy").CanClose);
            Assert.Throws<ExcelBusyException>(() => manager.CloseSession("late-busy", save));
            Assert.Equal(0, batch.DisposeCallCount);
            Assert.Same(batch, manager.GetSession("late-busy"));
            Assert.True(manager.TryGetFilePath("late-busy", out var retainedPath));
            Assert.Equal(path, retainedPath);
        }
        finally
        {
            batch.CloseBusy = false;
        }
        Assert.True(manager.CloseSession("late-busy", save));
        Assert.Equal(1, batch.DisposeCallCount);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void Dispose_LateCloseRefusal_RetainsBlockedSessionAndAllowsRetry()
    {
        using var manager = new SessionManager();
        var blocked = new ConfigurableFailingBatch(false, false) { CloseBusy = true };
        var healthy = new ConfigurableFailingBatch(false, false);
        var path = Path.GetFullPath("late-shutdown-busy.xlsx");
        RegisterSession(manager, "blocked", path, blocked);
        RegisterSession(manager, "healthy", Path.GetFullPath("late-shutdown-ready.xlsx"), healthy);
        try
        {
            var failure = Assert.Throws<AggregateException>(manager.Dispose);
            Assert.Contains(failure.InnerExceptions, ex => ex.InnerException is ExcelBusyException);
            Assert.Equal(0, blocked.DisposeCallCount);
            Assert.Same(blocked, manager.GetSession("blocked"));
            Assert.True(manager.TryGetFilePath("blocked", out var retainedPath));
            Assert.Equal(path, retainedPath);
            Assert.Equal(1, healthy.DisposeCallCount);
        }
        finally
        {
            blocked.CloseBusy = false;
        }
        manager.Dispose();
        Assert.Equal(1, blocked.DisposeCallCount);
        Assert.Equal(1, healthy.DisposeCallCount);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void Close_ExplicitForce_PreservesExistingTeardownPolicy()
    {
        using var manager = new SessionManager();
        var batch = new ConfigurableFailingBatch(false, false) { CloseBusy = true };
        RegisterSession(manager, "forced", Path.GetFullPath("forced-close.xlsx"), batch);
        Assert.True(manager.CloseSession("forced", force: true));
        Assert.Equal(0, batch.CloseCallCount);
        Assert.Equal(1, batch.DisposeCallCount);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void Close_SaveReadinessRace_RetainsBlockedSessionAndAllowsRetry()
    {
        using var manager = new SessionManager();
        var blocked = new ConfigurableFailingBatch(saveFails: false, disposeFails: false) { SaveBusy = true };
        var healthy = new ConfigurableFailingBatch(saveFails: false, disposeFails: false);
        var path = Path.GetFullPath("shutdown-busy.xlsx");
        RegisterSession(manager, "blocked", path, blocked);
        RegisterSession(manager, "healthy", Path.GetFullPath("shutdown-ready.xlsx"), healthy);
        try
        {
            var failure = Assert.Throws<InvalidOperationException>(() => manager.CloseSession("blocked", save: true));
            Assert.IsType<ExcelBusyException>(failure.InnerException);
            Assert.Equal(0, blocked.DisposeCallCount);
            Assert.Same(blocked, manager.GetSession("blocked"));
            Assert.True(manager.TryGetFilePath("blocked", out var retainedPath));
            Assert.Equal(path, retainedPath);
            Assert.Equal(2, manager.ActiveSessionCount);
            Assert.Equal(0, healthy.DisposeCallCount);
        }
        finally
        {
            blocked.SaveBusy = false;
        }
        manager.Dispose();
        Assert.Equal(1, blocked.SaveCallCount);
        Assert.Equal(1, blocked.DisposeCallCount);
        Assert.Equal(1, healthy.DisposeCallCount);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void Dispose_TeardownFailures_AreReportedAfterEverySessionIsAttempted()
    {
        using var manager = new SessionManager();
        var firstBatch = new OneShotFailingBatch();
        var secondBatch = new OneShotFailingBatch();
        RegisterSession(manager, "first-session", Path.GetFullPath("first-session.xlsx"), firstBatch);
        RegisterSession(manager, "second-session", Path.GetFullPath("second-session.xlsx"), secondBatch);

        var failure = Assert.Throws<AggregateException>(() => manager.Dispose());

        Assert.Equal(1, firstBatch.DisposeCallCount);
        Assert.Equal(1, secondBatch.DisposeCallCount);
        Assert.Equal(2, failure.InnerExceptions.Count);
        Assert.All(failure.InnerExceptions, exception =>
        {
            Assert.IsType<InvalidOperationException>(exception);
            Assert.Contains("synthetic teardown failure", exception.Message, StringComparison.Ordinal);
        });
        Assert.Contains(failure.InnerExceptions,
            exception => exception.Message.Contains("first-session", StringComparison.Ordinal));
        Assert.Contains(failure.InnerExceptions,
            exception => exception.Message.Contains("second-session", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Close_SaveFailure_RetainsSessionWithoutAttemptingTeardown(bool disposeFails)
    {
        using var manager = new SessionManager();
        var batch = new ConfigurableFailingBatch(saveFails: true, disposeFails);
        var path = Path.GetFullPath("save-failure.xlsx");
        RegisterSession(manager, "save-failure", path, batch);
        try
        {
            var failure = Assert.Throws<InvalidOperationException>(() => manager.CloseSession("save-failure", save: true));

            Assert.Equal(1, batch.SaveCallCount);
            Assert.Equal(0, batch.DisposeCallCount);
            Assert.Same(batch.SaveException, failure.InnerException);
            Assert.Equal(1, manager.ActiveSessionCount);
            Assert.Same(batch, manager.GetSession("save-failure"));
            Assert.True(manager.TryGetFilePath("save-failure", out var retainedPath));
            Assert.Equal(path, retainedPath);
        }
        finally
        {
            batch.SaveFails = false;
            batch.DisposeFails = false;
        }
        manager.Dispose();
        Assert.Equal(1, batch.SaveCallCount);
        Assert.Equal(1, batch.DisposeCallCount);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void Dispose_DoesNotInvokeSaveEvenWhenSaveWouldFail()
    {
        using var manager = new SessionManager();
        var failingBatch = new ConfigurableFailingBatch(saveFails: true, disposeFails: false);
        var healthyBatch = new ConfigurableFailingBatch(saveFails: false, disposeFails: false);
        RegisterSession(manager, "failing-save", Path.GetFullPath("failing-save.xlsx"), failingBatch);
        RegisterSession(manager, "healthy-save", Path.GetFullPath("healthy-save.xlsx"), healthyBatch);

        try
        {
            manager.Dispose();
            Assert.Equal(0, failingBatch.SaveCallCount);
            Assert.Equal(1, failingBatch.DisposeCallCount);
            Assert.Equal(0, healthyBatch.SaveCallCount);
            Assert.Equal(1, healthyBatch.DisposeCallCount);
            Assert.Equal(0, manager.ActiveSessionCount);
        }
        finally
        {
            failingBatch.SaveFails = false;
        }
        manager.Dispose();
        Assert.Equal(1, failingBatch.DisposeCallCount);
        Assert.Equal(1, healthyBatch.DisposeCallCount);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

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

    private sealed class OneShotFailingBatch : IExcelBatch, IExcelBatchTeardownState, IExcelBatchRefreshState
    {
        public WorkbookRefreshState GetRefreshState() => WorkbookRefreshState.Ready;
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

        public void Save(CancellationToken cancellationToken = default)
        {
        }
    }

    private sealed class ConfigurableFailingBatch(bool saveFails, bool disposeFails)
        : IExcelBatch, IExcelBatchRefreshState, IExcelBatchCloseState
    {
        public bool SaveFails { get; set; } = saveFails;
        public bool DisposeFails { get; set; } = disposeFails;
        public bool SaveBusy { get; set; }
        public Action? BeforeSave { get; init; }
        public bool CloseBusy { get; set; }
        public bool ProcessAlive { get; set; } = true;
        public WorkbookRefreshState RefreshState { get; set; } = WorkbookRefreshState.Ready;
        public Exception? BeforeCloseFailure { get; set; }
        public bool RevealOnRefreshRead { get; set; }
        public bool ShutdownCommitted { get; private set; }
        public int CloseCallCount { get; private set; }
        public WorkbookRefreshState GetRefreshState()
        {
            if (RevealOnRefreshRead)
                IsExcelVisible = true;
            return RefreshState;
        }
        internal InvalidOperationException SaveException { get; } =
            new("synthetic save failure");

        public int SaveCallCount { get; private set; }

        public int DisposeCallCount { get; private set; }

        public string WorkbookPath => Path.GetFullPath("configurable-failure.xlsx");

        public Microsoft.Extensions.Logging.ILogger Logger => NullLogger.Instance;

        public IReadOnlyDictionary<string, Excel.Workbook> Workbooks =>
            new Dictionary<string, Excel.Workbook>();

        public bool HasTimedOutOperation => false;

        public int? ExcelProcessId => Environment.ProcessId;

        public TimeSpan OperationTimeout => TimeSpan.FromSeconds(1);

        public bool IsExcelVisible { get; private set; }

        public void Close()
        {
            CloseCallCount++;
            if (CloseBusy)
                ExcelBusyException.ThrowIfNotReady(WorkbookRefreshState.Refreshing, "close");
            if (BeforeCloseFailure is { } failure)
                throw failure;
            Dispose();
        }

        public void Dispose()
        {
            ShutdownCommitted = true;
            DisposeCallCount++;
            if (DisposeFails)
            {
                throw new InvalidOperationException("synthetic teardown failure");
            }
        }

        public bool IsExcelProcessAlive() => ProcessAlive;

        public void Save(CancellationToken cancellationToken = default)
        {
            BeforeSave?.Invoke();
            SaveCallCount++;
            if (SaveBusy)
                ExcelBusyException.ThrowIfNotReady(WorkbookRefreshState.Refreshing, "save");
            if (SaveFails)
            {
                throw SaveException;
            }
        }

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
    }
}
