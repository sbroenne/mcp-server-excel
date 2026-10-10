using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.ExceptionServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration;

[Collection("Sequential")]
[Trait("Category", "Integration")]
[Trait("RunType", "OnDemand")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "true")]
public sealed class SessionManagerStartupShutdownTests
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Startup_RequiresProcessIdentityBeforeOpeningWorkbook(bool transientFailure, bool createNew)
    {
        var path = Path.Combine(Path.GetTempPath(), $"startup-identity-{Guid.NewGuid():N}.xlsx");
        if (!createNew)
            File.Copy(Path.Combine(AppContext.BaseDirectory,
                "Integration", "Session", "TestFiles", "batch-test-static.xlsx"), path);

        using var manager = new SessionManager();
        using var owned = new OwnedExcelProcessScope();
        var identities = new List<ExcelProcessIdentity>();
        var attempts = 0;
        var opened = false;
        string? session = null;
        Exception? failure = null;
        var cleanupFailures = new List<Exception>();
        ExcelBatch.TrackProcessIdentityHookForTests = processId =>
        {
            var identity = SessionManager.TrackExcelProcessIdentity(processId);
            Assert.NotNull(identity);
            identities.Add(identity.Value);
            attempts++;
            return transientFailure && attempts > 1 ? identity : null;
        };
        ExcelBatch.AfterWorkbookOpenHookForTests = (_, _) => opened = true;
        try
        {
            var startupFailure = Record.Exception(() => session = createNew
                ? manager.CreateSessionForNewFile(path)
                : manager.CreateSession(path));
            if (transientFailure)
            {
                Assert.Null(startupFailure);
                Assert.Equal(2, attempts);
                Assert.True(opened);
                var batch = Assert.IsAssignableFrom<IExcelBatch>(manager.GetSession(session!));
                Assert.Equal(WorkbookRefreshState.Ready,
                    Assert.IsAssignableFrom<IExcelBatchRefreshState>(batch).GetRefreshState());
                Assert.True(manager.ValidateClose(session!).CanClose);
                SessionWorkbookAssertions.WriteMarker(batch, "Identity capture recovered");
                batch.Save();
                Assert.True(manager.CloseSession(session!, save: false));
                session = null;
                ExcelBatch.TrackProcessIdentityHookForTests = null;
                session = manager.CreateSession(path);
                Assert.Equal("Identity capture recovered", SessionWorkbookAssertions.ReadMarker(
                    Assert.IsAssignableFrom<IExcelBatch>(manager.GetSession(session))));
                Assert.True(manager.ValidateClose(session).CanClose);
            }
            else
            {
                var error = Assert.IsType<InvalidOperationException>(startupFailure);
                Assert.Contains("process identity", error.ToString(), StringComparison.OrdinalIgnoreCase);
                Assert.Equal(3, attempts);
                Assert.False(opened);
                Assert.Null(session);
                Assert.Equal(0, manager.ActiveSessionCount);
                Assert.Empty(GetField<ConcurrentDictionary<string, string>>(manager, "_activeFilePaths"));
                if (createNew) Assert.False(File.Exists(path));
                owned.AssertAllExited();
            }
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            ExcelBatch.TrackProcessIdentityHookForTests = null;
            ExcelBatch.AfterWorkbookOpenHookForTests = null;
            if (session != null)
                CaptureCleanup(() => manager.CloseSession(session, save: false, force: true));
            CaptureCleanup(manager.Dispose);
            CaptureCleanup(() => owned.AssertAllExited());
            foreach (var identity in identities.Where(OwnedProcessGuard.TryConfirmExited).Distinct())
                SessionManager.UntrackExcelProcess(identity);
            CaptureCleanup(() => File.Delete(path));
        }
        if (cleanupFailures.Count > 0)
        {
            if (failure != null) cleanupFailures.Insert(0, failure);
            throw new AggregateException("Process identity startup regression or cleanup failed.", cleanupFailures);
        }
        if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();

        void CaptureCleanup(Action cleanup)
        {
            try { cleanup(); }
            catch (Exception ex) { cleanupFailures.Add(ex); }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Dispose_DuringStartup_RejectsLateSessionAndReleasesOwnership(bool createNew)
    {
        var path = Path.Combine(Path.GetTempPath(), $"startup-shutdown-{Guid.NewGuid():N}.xlsx");
        if (!createNew)
        {
            File.Copy(Path.Combine(AppContext.BaseDirectory,
                "Integration", "Session", "TestFiles", "batch-test-static.xlsx"), path);
        }

        using var manager = new SessionManager();
        using var owned = new OwnedExcelProcessScope();
        using var startupReached = new ManualResetEventSlim();
        using var releaseStartup = new ManualResetEventSlim();
        var paths = GetField<ConcurrentDictionary<string, string>>(manager, "_activeFilePaths");
        var sessions = GetField<ConcurrentDictionary<string, IExcelBatch>>(manager, "_activeSessions");
        SessionManager.ExcelProcessIdentityTracked += HoldStartup;
        var creation = Task.Run(() => createNew
            ? manager.CreateSessionForNewFile(path)
            : manager.CreateSession(path));
        Exception? primaryFailure = null;
        var cleanupFailures = new List<Exception>();
        var creationObserved = false;
        try
        {
            Assert.True(startupReached.Wait(TimeSpan.FromSeconds(30)), "Excel startup never reached ownership tracking.");
            Assert.True(paths.ContainsKey(path), "Startup must hold its file-path reservation.");
            Assert.False(creation.IsCompleted, "Session creation must remain blocked before publication.");
            Assert.Empty(sessions);
            manager.Dispose();
            releaseStartup.Set();
            var failure = await Record.ExceptionAsync(
                async () => await creation.WaitAsync(TimeSpan.FromSeconds(150)));
            creationObserved = true;
            Assert.IsType<ObjectDisposedException>(Assert.IsType<InvalidOperationException>(failure).InnerException);
            Assert.Equal(0, manager.ActiveSessionCount);
            Assert.Empty(paths);
            foreach (var field in new[] { "_sessionFilePaths", "_activeOperationCounts", "_showExcelFlags", "_sessionOrigins", "_sessionCreatedAt" })
                Assert.Empty(GetField<System.Collections.IDictionary>(manager, field));
            owned.AssertAllExited();
        }
        catch (Exception ex)
        {
            primaryFailure = ex;
        }
        finally
        {
            releaseStartup.Set();
            if (!creationObserved)
            {
                var error = await Record.ExceptionAsync(async () => await creation.WaitAsync(TimeSpan.FromSeconds(150)));
                if (error is not null) cleanupFailures.Add(error);
            }
            SessionManager.ExcelProcessIdentityTracked -= HoldStartup;
            foreach (var batch in sessions.Values)
                CaptureCleanup(batch.Dispose);
            CaptureCleanup(manager.Dispose);
            CaptureCleanup(() => owned.AssertAllExited(expectProcess: startupReached.IsSet));
            CaptureCleanup(() => File.Delete(path));
        }
        if (cleanupFailures.Count > 0)
        {
            if (primaryFailure is not null) cleanupFailures.Insert(0, primaryFailure);
            throw new AggregateException("Startup/shutdown regression or cleanup failed.", cleanupFailures);
        }
        if (primaryFailure is not null) ExceptionDispatchInfo.Capture(primaryFailure).Throw();

        void HoldStartup(ExcelProcessIdentity _)
        {
            startupReached.Set();
            Assert.True(releaseStartup.Wait(TimeSpan.FromSeconds(45)), "Startup gate was not released.");
        }

        void CaptureCleanup(Action cleanup)
        {
            try { cleanup(); }
            catch (Exception ex) { cleanupFailures.Add(ex); }
        }
    }

    private static T GetField<T>(SessionManager manager, string name) where T : class =>
        Assert.IsAssignableFrom<T>(typeof(SessionManager)
            .GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(manager));
}
