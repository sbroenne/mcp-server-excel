using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration;

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
    private readonly string _tempDir;
    private readonly List<string> _testFiles = [];

    public SessionManagerOperationTrackingTests(ITestOutputHelper _ /* injected by xUnit */)
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"SessionManagerOpTrackingTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);

        var failures = new List<Exception>();
        foreach (var file in _testFiles.Where(File.Exists))
        {
            try { File.Delete(file); }
            catch (IOException exception) { failures.Add(exception); }
            catch (UnauthorizedAccessException exception) { failures.Add(exception); }
        }

        if (Directory.Exists(_tempDir))
        {
            try { Directory.Delete(_tempDir, recursive: true); }
            catch (IOException exception) { failures.Add(exception); }
            catch (UnauthorizedAccessException exception) { failures.Add(exception); }
        }
        if (failures.Count != 0)
            throw new AggregateException("Operation-tracking test cleanup failed.", failures);
    }

    /// <summary>
    /// Path to the template xlsx file used for fast test file creation.
    /// Copying a template is ~1000x faster than spawning Excel to create a new workbook.
    /// </summary>
    private static readonly string TemplateFilePath = Path.Combine(
        Path.GetDirectoryName(typeof(SessionManagerOperationTrackingTests).Assembly.Location)!,
        "Integration", "Session", "TestFiles", "batch-test-static.xlsx");

    private string CreateTestFile(string testName)
    {
        var fileName = $"{testName}_{Guid.NewGuid():N}.xlsx";
#pragma warning disable CA3003 // Path.Combine is safe here - test code with controlled inputs
        var filePath = Path.Combine(_tempDir, fileName);
#pragma warning restore CA3003

        // PERFORMANCE OPTIMIZATION: Copy from template instead of spawning Excel.
        // This reduces test file creation from ~7-14 seconds to <10ms.
        File.Copy(TemplateFilePath, filePath);

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
        Assert.Equal(1, manager.GetActiveOperationCount(sessionId));
        manager.EndOperation(sessionId);
        Assert.Equal(0, manager.GetActiveOperationCount(sessionId));
        AssertSessionClosed(manager, sessionId);
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

        AssertSessionClosed(manager, sessionId);
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
    public void BeginOperation_NullOrEmptySessionId_Throws_AndEndOperationDoesNotThrow()
    {
        using var manager = new SessionManager();

        var nullException = Assert.Throws<InvalidOperationException>(() => manager.BeginOperation(null!));
        var emptyException = Assert.Throws<InvalidOperationException>(() => manager.BeginOperation(""));

        Assert.Equal("sessionId is required", nullException.Message);
        Assert.Equal("sessionId is required", emptyException.Message);

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
        var sessionId = manager.CreateSession(testFile, show: false);

        Assert.False(manager.IsExcelVisible(sessionId));

        AssertSessionClosed(manager, sessionId);
    }

    [Fact]
    public void IsExcelVisible_SessionWithShowExcelTrue_ReturnsTrue()
    {
        var testFile = CreateTestFile(nameof(IsExcelVisible_SessionWithShowExcelTrue_ReturnsTrue));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        Assert.True(manager.IsExcelVisible(sessionId));

        AssertSessionClosed(manager, sessionId);
    }

    [Fact]
    public void IsExcelVisible_WhenApplicationVisibilityChanges_ReturnsLiveComState()
    {
        var testFile = CreateTestFile(nameof(IsExcelVisible_WhenApplicationVisibilityChanges_ReturnsLiveComState));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        Assert.True(manager.TryBeginOperation(sessionId, out var batch, out var errorMessage), errorMessage);
        try
        {
            batch.Execute((ctx, ct) =>
            {
                ctx.App.Visible = false;
                return true;
            });
        }
        finally
        {
            manager.EndOperation(sessionId);
        }

        Assert.False(manager.IsExcelVisible(sessionId));

        AssertSessionClosed(manager, sessionId);
    }

    [Fact]
    public void IsExcelVisible_NonExistentSession_ReturnsFalse()
    {
        using var manager = new SessionManager();

        Assert.False(manager.IsExcelVisible("nonexistent"));
        Assert.False(manager.IsExcelVisible(null!));
    }

    [Fact]
    public void TrackExcelProcess_SynchronousPersistenceRunsBeforeReturn()
    {
        var persistenceObserved = false;
        void PersistIdentity(ExcelProcessIdentity identity)
        {
            if (identity.ProcessId == Environment.ProcessId)
            {
                persistenceObserved = true;
            }
        }

        ExcelProcessIdentity? trackedIdentity = null;
        SessionManager.ExcelProcessIdentityTracked += PersistIdentity;
        try
        {
            trackedIdentity = SessionManager.TrackExcelProcessIdentity(Environment.ProcessId);

            Assert.True(persistenceObserved);
        }
        finally
        {
            SessionManager.ExcelProcessIdentityTracked -= PersistIdentity;
            if (trackedIdentity is { } identity)
            {
                SessionManager.UntrackExcelProcess(identity);
            }
        }
    }

    [Fact]
    public void TrackExcelProcess_SynchronousPersistenceFailurePropagates()
    {
        using var process = Process.GetCurrentProcess();
        var identity = new ExcelProcessIdentity(
            process.Id,
            process.StartTime.ToUniversalTime().ToFileTimeUtc());
        void FailPersistence(ExcelProcessIdentity _) =>
            throw new IOException("synthetic durable persistence failure");

        SessionManager.ExcelProcessIdentityTracked += FailPersistence;
        try
        {
            var exception = Assert.ThrowsAny<InvalidOperationException>(() =>
                SessionManager.TrackExcelProcess(identity));

            Assert.Contains("persist", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.IsType<IOException>(exception.InnerException);
            var persistenceException = Assert.IsType<ExcelProcessPersistenceException>(exception);
            Assert.Equal(identity, persistenceException.Identity);
        }
        finally
        {
            SessionManager.ExcelProcessIdentityTracked -= FailPersistence;
            SessionManager.UntrackExcelProcess(identity);
        }
    }

    [Fact]
    public void LegacyProcessTrackingMembers_RemainUsable()
    {
        using var notificationReceived = new ManualResetEventSlim();
        void LegacySubscriber(IReadOnlyCollection<int> processIds)
        {
            if (processIds.Contains(Environment.ProcessId))
            {
                notificationReceived.Set();
            }
        }

        SessionManager.TrackedExcelProcessesChanged += LegacySubscriber;
        try
        {
            SessionManager.TrackExcelProcess(Environment.ProcessId);

            Assert.True(notificationReceived.Wait(TimeSpan.FromSeconds(5)));
            Assert.Contains(Environment.ProcessId, SessionManager.GetTrackedExcelProcessIds());
        }
        finally
        {
            SessionManager.TrackedExcelProcessesChanged -= LegacySubscriber;
            SessionManager.UntrackExcelProcess(Environment.ProcessId);
        }
    }

    [Fact]
    public void TryTerminateOwnedProcess_MismatchedStartTime_DoesNotTerminateReusedPid()
    {
        using var currentProcess = System.Diagnostics.Process.GetCurrentProcess();
        var staleIdentity = new ExcelProcessIdentity(
            currentProcess.Id,
            currentProcess.StartTime.ToUniversalTime().ToFileTimeUtc() + 1);

#pragma warning disable RS0030 // Checks the product refuses a wrong identity; the target is this test host, never Excel.
        var terminated = ExcelBatch.TryTerminateOwnedProcess(
            staleIdentity,
            waitBeforeTermination: TimeSpan.Zero,
            waitAfterTermination: TimeSpan.FromSeconds(1));
#pragma warning restore RS0030

        Assert.True(terminated);
        Assert.False(currentProcess.HasExited);
    }

    [Fact]
    public void FinalizeOwnedProcessTeardown_TerminationFailureRetainsExactIdentity()
    {
        using var currentProcess = Process.GetCurrentProcess();
        var identity = new ExcelProcessIdentity(
            currentProcess.Id,
            currentProcess.StartTime.ToUniversalTime().ToFileTimeUtc());
        var unrelatedIdentity = identity with
        {
            StartedAtUtcFileTime = identity.StartedAtUtcFileTime - 1
        };
        SessionManager.TrackExcelProcess(identity);
        SessionManager.TrackExcelProcess(unrelatedIdentity);
        ExcelProcessIdentity? observedIdentity = null;

        try
        {
            var exception = Assert.Throws<InvalidOperationException>(() =>
                ExcelBatch.FinalizeOwnedProcessTeardown(
                    identity,
                    candidate =>
                    {
                        observedIdentity = candidate;
                        return false;
                    }));

            Assert.Equal(identity, observedIdentity);
            Assert.Contains(identity, SessionManager.GetTrackedExcelProcesses());
            Assert.Contains(unrelatedIdentity, SessionManager.GetTrackedExcelProcesses());
            Assert.Contains(
                identity.ProcessId.ToString(System.Globalization.CultureInfo.InvariantCulture),
                exception.Message,
                StringComparison.Ordinal);
        }
        finally
        {
            SessionManager.UntrackExcelProcess(identity);
            SessionManager.UntrackExcelProcess(unrelatedIdentity);
        }
    }

    [Fact]
    public void FinalizeOwnedProcessTeardown_ConfirmedExitUntracksOnlyExactIdentity()
    {
        using var currentProcess = Process.GetCurrentProcess();
        var identity = new ExcelProcessIdentity(
            currentProcess.Id,
            currentProcess.StartTime.ToUniversalTime().ToFileTimeUtc());
        var unrelatedIdentity = identity with
        {
            StartedAtUtcFileTime = identity.StartedAtUtcFileTime - 1
        };
        SessionManager.TrackExcelProcess(identity);
        SessionManager.TrackExcelProcess(unrelatedIdentity);

        try
        {
            ExcelBatch.FinalizeOwnedProcessTeardown(identity, candidate => candidate == identity);

            Assert.DoesNotContain(identity, SessionManager.GetTrackedExcelProcesses());
            Assert.Contains(unrelatedIdentity, SessionManager.GetTrackedExcelProcesses());
            Assert.False(currentProcess.HasExited);
        }
        finally
        {
            SessionManager.UntrackExcelProcess(identity);
            SessionManager.UntrackExcelProcess(unrelatedIdentity);
        }
    }

    [Fact]
    public void StartupTeardown_UnconfirmedExitRetainsIdentityWithoutTermination()
    {
        using var currentProcess = Process.GetCurrentProcess();
        var identity = new ExcelProcessIdentity(
            currentProcess.Id,
            currentProcess.StartTime.ToUniversalTime().ToFileTimeUtc());
        SessionManager.TrackExcelProcess(identity);
        ExcelProcessIdentity? confirmationCandidate = null;

        try
        {
            var exception = Assert.Throws<InvalidOperationException>(() =>
                ExcelBatch.FinalizeOwnedProcessTeardown(
                    identity,
                    candidate =>
                    {
                        confirmationCandidate = candidate;
                        return false;
                    }));

            Assert.Equal(identity, confirmationCandidate);
            Assert.Contains(identity, SessionManager.GetTrackedExcelProcesses());
            Assert.Contains("remain recorded", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(currentProcess.HasExited);
        }
        finally
        {
            SessionManager.UntrackExcelProcess(identity);
        }
    }

    [Fact]
    public void UntrackExcelProcess_OldIdentityDoesNotRemoveReusedPidIdentity()
    {
        using var currentProcess = System.Diagnostics.Process.GetCurrentProcess();
        var replacementIdentity = new ExcelProcessIdentity(
            currentProcess.Id,
            currentProcess.StartTime.ToUniversalTime().ToFileTimeUtc());
        var oldIdentity = replacementIdentity with
        {
            StartedAtUtcFileTime = replacementIdentity.StartedAtUtcFileTime - 1
        };

        SessionManager.TrackExcelProcess(oldIdentity);
        SessionManager.TrackExcelProcess(replacementIdentity);
        try
        {
            SessionManager.UntrackExcelProcess(oldIdentity);

            Assert.DoesNotContain(oldIdentity, SessionManager.GetTrackedExcelProcesses());
            Assert.Contains(replacementIdentity, SessionManager.GetTrackedExcelProcesses());
        }
        finally
        {
            SessionManager.UntrackExcelProcess(oldIdentity);
            SessionManager.UntrackExcelProcess(replacementIdentity);
        }
    }

    [Fact]
    public void TrackExcelProcess_NotificationIncludesCapturedStartTime()
    {
        using var notificationReceived = new ManualResetEventSlim();
        ExcelProcessIdentity? observedIdentity = null;
        void CaptureSubscriber(IReadOnlyCollection<ExcelProcessIdentity> processes)
        {
            observedIdentity = processes.Single(process => process.ProcessId == Environment.ProcessId);
            notificationReceived.Set();
        }

        using var currentProcess = System.Diagnostics.Process.GetCurrentProcess();
        var expectedStartTime = currentProcess.StartTime.ToUniversalTime().ToFileTimeUtc();
        ExcelProcessIdentity? trackedIdentity = null;
        SessionManager.TrackedExcelProcessIdentitiesChanged += CaptureSubscriber;
        try
        {
            trackedIdentity = SessionManager.TrackExcelProcessIdentity(Environment.ProcessId);

            Assert.True(notificationReceived.Wait(TimeSpan.FromSeconds(5)));
            Assert.Equal(
                new ExcelProcessIdentity(Environment.ProcessId, expectedStartTime),
                observedIdentity);
        }
        finally
        {
            SessionManager.TrackedExcelProcessIdentitiesChanged -= CaptureSubscriber;
            if (trackedIdentity is { } identity)
            {
                SessionManager.UntrackExcelProcess(identity);
            }
        }
    }

    [Fact]
    public void TrackExcelProcess_SubscriberFailure_DoesNotBreakSessionLifecycle()
    {
        using var notificationReceived = new ManualResetEventSlim();
        void ThrowingSubscriber(IReadOnlyCollection<ExcelProcessIdentity> _)
        {
            notificationReceived.Set();
            throw new InvalidOperationException("Simulated tracker failure.");
        }

        SessionManager.TrackedExcelProcessIdentitiesChanged += ThrowingSubscriber;
        ExcelProcessIdentity? trackedIdentity = null;
        try
        {
            var exception = Record.Exception(() =>
                trackedIdentity = SessionManager.TrackExcelProcessIdentity(Environment.ProcessId));

            Assert.Null(exception);
            Assert.True(notificationReceived.Wait(TimeSpan.FromSeconds(5)));
        }
        finally
        {
            SessionManager.TrackedExcelProcessIdentitiesChanged -= ThrowingSubscriber;
            if (trackedIdentity is { } identity)
            {
                SessionManager.UntrackExcelProcess(identity);
            }
        }
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

        AssertSessionClosed(manager, sessionId);
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
        var remaining = manager.ValidateClose(sessionId);
        Assert.False(remaining.CanClose);
        Assert.Equal(1, remaining.ActiveOperationCount);
        manager.EndOperation(sessionId);
        var ready = manager.ValidateClose(sessionId);
        Assert.True(ready.CanClose);
        Assert.Null(ready.BlockingReason);
        AssertSessionClosed(manager, sessionId);
    }

    [Fact]
    public void ValidateClose_NonExistentSession_SessionExistsFalse()
    {
        using var manager = new SessionManager();

        var result = manager.ValidateClose("nonexistent");

        Assert.False(result.SessionExists);
        Assert.False(result.CanClose);
        Assert.Equal(0, result.ActiveOperationCount);
        Assert.NotNull(result.BlockingReason);
        Assert.Contains("not found", result.BlockingReason);
    }

    [Fact]
    public void ValidateClose_NullSessionId_SessionExistsFalse()
    {
        using var manager = new SessionManager();

        var result = manager.ValidateClose(null!);

        Assert.False(result.SessionExists);
        Assert.False(result.CanClose);
        Assert.Equal(0, result.ActiveOperationCount);
        Assert.Contains("required", result.BlockingReason);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ValidateClose_IncludesVisibilityInfo(bool slowInspection)
    {
        var testFile = CreateTestFile(nameof(ValidateClose_IncludesVisibilityInfo));
        using var owned = new OwnedExcelProcessScope();
        using var release = new ManualResetEventSlim();
        using var inspectionStarted = new ManualResetEventSlim();
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);
        var batch = Assert.IsType<ExcelBatch>(manager.GetSession(sessionId));
        if (slowInspection)
        {
            batch.BeforeRefreshStateReadHookForTests = () =>
            {
                inspectionStarted.Set();
                Assert.True(release.Wait(TimeSpan.FromSeconds(10)), "Readiness inspection was not released.");
            };
        }

        var failure = Record.Exception(() =>
        {
            if (slowInspection)
            {
                var pending = manager.ValidateClose(sessionId);
                var started = SpinWait.SpinUntil(() =>
                {
                    if (inspectionStarted.IsSet)
                        return true;
                    pending = manager.ValidateClose(sessionId);
                    return inspectionStarted.IsSet;
                }, TimeSpan.FromSeconds(5));
                Assert.True(started, $"The readiness inspection did not start. {DescribeCloseValidation(pending)}");
                Assert.Equal("unknown", pending.ExcelState);
                Assert.False(pending.CanClose, DescribeCloseValidation(pending));
                Assert.NotNull(pending.BlockingReason);
                Assert.True(pending.IsExcelVisible, DescribeCloseValidation(pending));
                Assert.True(pending.SessionExists, DescribeCloseValidation(pending));
                Assert.Equal(0, pending.ActiveOperationCount);
                batch.BeforeRefreshStateReadHookForTests = null;
                release.Set();
            }

            var result = WaitForCloseReady(manager, sessionId);

            Assert.True(result.IsExcelVisible, DescribeCloseValidation(result));
            Assert.True(result.SessionExists, DescribeCloseValidation(result));
            Assert.True(result.CanClose, DescribeCloseValidation(result));
            Assert.Equal("ready", result.ExcelState);
            Assert.Equal(0, result.ActiveOperationCount);
            Assert.Null(result.BlockingReason);
        });

        batch.BeforeRefreshStateReadHookForTests = null;
        release.Set();
        var closeFailure = Record.Exception(() =>
        {
            WaitForCloseReady(manager, sessionId);
            AssertSessionClosed(manager, sessionId);
        });
        var processFailure = Record.Exception(() => owned.AssertAllExited());
        var failures = new[] { failure, closeFailure, processFailure }.OfType<Exception>().ToArray();
        if (failures.Length > 0)
            throw new AggregateException("Visibility validation or normal session close failed.", failures);
    }

    private static CloseValidationResult WaitForCloseReady(SessionManager manager, string sessionId)
    {
        var result = manager.ValidateClose(sessionId);
        var ready = SpinWait.SpinUntil(() =>
        {
            if (result.CanClose)
                return true;
            result = manager.ValidateClose(sessionId);
            return result.CanClose;
        }, TimeSpan.FromSeconds(10));
        Assert.True(ready, $"Excel did not become ready to close. {DescribeCloseValidation(result)}");
        return result;
    }

    private static string DescribeCloseValidation(CloseValidationResult result) =>
        $"ExcelState={result.ExcelState}; BlockingReason={result.BlockingReason ?? "<null>"}; " +
        $"IsExcelVisible={result.IsExcelVisible}; ActiveOperationCount={result.ActiveOperationCount}; " +
        $"SessionExists={result.SessionExists}; CanClose={result.CanClose}";

    #endregion

    #region CloseSession with Operation Tracking

    [Fact]
    public void CloseSession_OperationsRunning_ThrowsInvalidOperationException()
    {
        var testFile = CreateTestFile(nameof(CloseSession_OperationsRunning_ThrowsInvalidOperationException));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);
        var batch = Assert.IsAssignableFrom<IExcelBatch>(manager.GetSession(sessionId));
        WriteGuard(batch, 21);
        AssertGuard(batch, 21);

        manager.BeginOperation(sessionId);

        var ex = Assert.Throws<InvalidOperationException>(
            () => manager.CloseSession(sessionId));

        Assert.Contains("1 operation(s) still running", ex.Message);
        Assert.Contains("Wait for all operations to complete", ex.Message);

        // Session should still be open
        Assert.Equal(1, manager.ActiveSessionCount);
        Assert.Same(batch, manager.GetSession(sessionId));
        Assert.Equal(1, manager.GetActiveOperationCount(sessionId));
        Assert.False(manager.ValidateClose(sessionId).CanClose);
        AssertGuard(batch, 21);

        // Clean up
        manager.EndOperation(sessionId);
        Assert.True(manager.ValidateClose(sessionId).CanClose);
        AssertSessionClosed(manager, sessionId);
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
        AssertClosedTracking(manager, sessionId);
        manager.EndOperation(sessionId);
        AssertClosedTracking(manager, sessionId);
    }

    #endregion

    #region Cleanup on Close

    [Fact]
    public void CloseSession_CleansUpOperationTracking()
    {
        var testFile = CreateTestFile(nameof(CloseSession_CleansUpOperationTracking));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        // Set some state
        manager.BeginOperation(sessionId);
        manager.EndOperation(sessionId);

        AssertSessionClosed(manager, sessionId);

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

        var session1 = manager.CreateSession(testFile1, show: true);
        var session2 = manager.CreateSession(testFile2, show: false);

        manager.BeginOperation(session1);
        manager.BeginOperation(session2);

        // Shutdown refuses sessions with active operations, leaves them open, and admits no new sessions.
        var refused = Assert.Throws<AggregateException>(manager.Dispose);
        Assert.Equal(2, refused.InnerExceptions.Count);
        Assert.All(refused.InnerExceptions, failure => Assert.IsType<ExcelBusyException>(failure.InnerException));
        Assert.Equal(2, manager.ActiveSessionCount);
        Assert.Throws<ObjectDisposedException>(() =>
            manager.CreateSession(CreateTestFile($"{nameof(Dispose_CleansUpAllTracking)}_3")));

        manager.EndOperation(session1);
        manager.EndOperation(session2);
        manager.Dispose();

        // All tracking should be cleared
        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.Equal(0, manager.GetActiveOperationCount(session1));
        Assert.Equal(0, manager.GetActiveOperationCount(session2));
        Assert.False(manager.IsExcelVisible(session1));
        Assert.False(manager.IsExcelVisible(session2));
        Assert.Throws<ObjectDisposedException>(() => manager.TryBeginOperation(session1, out _, out _));
        Assert.Throws<ObjectDisposedException>(() => manager.GetSession(session2));
    }

    [Theory]
    [InlineData(null, SessionOperationError.MissingSessionId, "sessionId is required")]
    [InlineData("", SessionOperationError.MissingSessionId, "sessionId is required")]
    [InlineData("   ", SessionOperationError.MissingSessionId, "sessionId is required")]
    [InlineData("missing-session", SessionOperationError.NotFound, "not found")]
    public void TryBeginOperation_InvalidIdentity_PreservesExistingSession(
        string? rejectedId, SessionOperationError expectedError, string expectedMessage)
    {
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(CreateTestFile(nameof(TryBeginOperation_InvalidIdentity_PreservesExistingSession)));
        var original = Assert.IsAssignableFrom<IExcelBatch>(manager.GetSession(sessionId));
        WriteGuard(original, 17);
        manager.BeginOperation(sessionId);
        try
        {
            Assert.False(manager.TryBeginOperation(rejectedId!, out var rejectedBatch, out var message, out var error));
            Assert.Null(rejectedBatch);
            Assert.Equal(expectedError, error);
            Assert.Contains(expectedMessage, message, StringComparison.Ordinal);
            Assert.Equal(1, manager.ActiveSessionCount);
            Assert.Equal(1, manager.GetActiveOperationCount(sessionId));
            Assert.Same(original, manager.GetSession(sessionId));
            AssertGuard(original, 17);
        }
        finally
        {
            manager.EndOperation(sessionId);
        }
        AssertSessionClosed(manager, sessionId);
    }

    [Fact]
    public void TryBeginOperation_FailedOperation_EndReleasesTrackingAndSessionRemainsUsable()
    {
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(CreateTestFile(nameof(TryBeginOperation_FailedOperation_EndReleasesTrackingAndSessionRemainsUsable)));
        var original = Assert.IsAssignableFrom<IExcelBatch>(manager.GetSession(sessionId));
        WriteGuard(original, 19);
        Assert.True(manager.TryBeginOperation(sessionId, out var batch, out var message, out var error), message);
        Assert.Null(message);
        Assert.Equal(SessionOperationError.None, error);
        Assert.Same(original, batch);
        Assert.Equal(1, manager.GetActiveOperationCount(sessionId));
        try
        {
            var exception = Assert.Throws<InvalidOperationException>(() =>
                batch.Execute((_, _) => throw new InvalidOperationException("synthetic operation failure")));
            Assert.Contains("synthetic operation failure", exception.Message);
            Assert.Equal(1, manager.GetActiveOperationCount(sessionId));
            Assert.False(manager.ValidateClose(sessionId).CanClose);
        }
        finally
        {
            manager.EndOperation(sessionId);
        }
        Assert.Equal(0, manager.GetActiveOperationCount(sessionId));
        Assert.True(manager.ValidateClose(sessionId).CanClose);
        AssertGuard(original, 19);
        WriteGuard(original, 23);
        AssertGuard(original, 23);
        AssertSessionClosed(manager, sessionId);
    }

    private static void AssertSessionClosed(SessionManager manager, string sessionId)
    {
        Assert.True(manager.CloseSession(sessionId));
        AssertClosedTracking(manager, sessionId);
    }

    private static void AssertClosedTracking(SessionManager manager, string sessionId)
    {
        Assert.Null(manager.GetSession(sessionId));
        Assert.Equal(0, manager.GetActiveOperationCount(sessionId));
        Assert.False(manager.IsExcelVisible(sessionId));
        var validation = manager.ValidateClose(sessionId);
        Assert.False(validation.SessionExists);
        Assert.False(validation.CanClose);
        Assert.Equal(0, validation.ActiveOperationCount);
        Assert.False(manager.TryBeginOperation(sessionId, out var batch, out var message, out var error));
        Assert.Null(batch);
        Assert.Equal(SessionOperationError.NotFound, error);
        Assert.Contains("not found", message, StringComparison.Ordinal);
    }

    private static void WriteGuard(IExcelBatch batch, int value) =>
        batch.Execute((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? input = null;
            Excel.Range? formula = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                input = sheet.Range["A1"];
                formula = sheet.Range["B1"];
                input.Value2 = value;
                formula.Formula = "=A1*2";
                formula.Calculate();
            }
            finally
            {
                ComUtilities.Release(ref formula);
                ComUtilities.Release(ref input);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    private static void AssertGuard(IExcelBatch batch, int expected) =>
        batch.Execute((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? input = null;
            Excel.Range? formula = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                input = sheet.Range["A1"];
                formula = sheet.Range["B1"];
                Assert.Equal(expected, Convert.ToInt32((object?)input.Value2, System.Globalization.CultureInfo.InvariantCulture));
                Assert.Equal("=A1*2", (object?)formula.Formula);
                Assert.Equal(expected * 2, Convert.ToInt32((object?)formula.Value2, System.Globalization.CultureInfo.InvariantCulture));
            }
            finally
            {
                ComUtilities.Release(ref formula);
                ComUtilities.Release(ref input);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    #endregion
}
