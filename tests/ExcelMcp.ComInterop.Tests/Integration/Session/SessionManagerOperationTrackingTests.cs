using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

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

        foreach (var file in _testFiles.Where(File.Exists))
        {
#pragma warning disable CA1031 // Catch general exception - best effort cleanup in test disposal
            try { File.Delete(file); } catch (Exception) { /* best effort */ }
#pragma warning restore CA1031
        }

        if (Directory.Exists(_tempDir))
        {
#pragma warning disable CA1031 // Catch general exception - best effort cleanup in test disposal
            try { Directory.Delete(_tempDir, recursive: true); } catch (Exception) { /* best effort */ }
#pragma warning restore CA1031
        }
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

        manager.CloseSession(sessionId);
    }

    [Fact]
    public void IsExcelVisible_SessionWithShowExcelTrue_ReturnsTrue()
    {
        var testFile = CreateTestFile(nameof(IsExcelVisible_SessionWithShowExcelTrue_ReturnsTrue));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        Assert.True(manager.IsExcelVisible(sessionId));

        manager.CloseSession(sessionId);
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

        manager.CloseSession(sessionId);
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

        var terminated = ExcelBatch.TryTerminateOwnedProcess(
            staleIdentity,
            waitBeforeTermination: TimeSpan.Zero,
            waitAfterTermination: TimeSpan.FromSeconds(1));

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
    public void FinalizeFailedStartupOwnedProcess_ConfirmedExitUntracksExactIdentity()
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
            ExcelBatch.FinalizeFailedStartupOwnedProcess(
                identity,
                _ => false,
                candidate => candidate == identity);

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
    public void FinalizeFailedStartupOwnedProcess_LiveUnkillableIdentityRemainsTracked()
    {
        using var currentProcess = Process.GetCurrentProcess();
        var identity = new ExcelProcessIdentity(
            currentProcess.Id,
            currentProcess.StartTime.ToUniversalTime().ToFileTimeUtc());
        SessionManager.TrackExcelProcess(identity);
        ExcelProcessIdentity? terminationCandidate = null;
        ExcelProcessIdentity? confirmationCandidate = null;

        try
        {
            var exception = Assert.Throws<InvalidOperationException>(() =>
                ExcelBatch.FinalizeFailedStartupOwnedProcess(
                    identity,
                    candidate =>
                    {
                        terminationCandidate = candidate;
                        return false;
                    },
                    candidate =>
                    {
                        confirmationCandidate = candidate;
                        return false;
                    }));

            Assert.Equal(identity, terminationCandidate);
            Assert.Equal(identity, confirmationCandidate);
            Assert.Contains(identity, SessionManager.GetTrackedExcelProcesses());
            Assert.Contains("remains tracked", exception.Message, StringComparison.OrdinalIgnoreCase);
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
        var sessionId = manager.CreateSession(testFile, show: true);

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
        var sessionId = manager.CreateSession(testFile, show: true);

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

        var session1 = manager.CreateSession(testFile1, show: true);
        var session2 = manager.CreateSession(testFile2, show: false);

        manager.BeginOperation(session1);
        manager.BeginOperation(session2);

        manager.Dispose();

        // All tracking should be cleared
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    #endregion
}
