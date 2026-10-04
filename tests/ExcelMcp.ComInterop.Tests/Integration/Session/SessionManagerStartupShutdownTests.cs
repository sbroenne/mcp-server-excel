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
