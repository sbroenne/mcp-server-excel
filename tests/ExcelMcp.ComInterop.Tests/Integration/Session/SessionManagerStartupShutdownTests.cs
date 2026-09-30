using System.Collections.Concurrent;
using System.Reflection;
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
        var paths = GetField<ConcurrentDictionary<string, string>>(manager, "_activeFilePaths");
        var sessions = GetField<ConcurrentDictionary<string, IExcelBatch>>(manager, "_activeSessions");
        var creation = Task.Run(() => createNew
            ? manager.CreateSessionForNewFile(path)
            : manager.CreateSession(path));

        try
        {
            var reserved = SpinWait.SpinUntil(
                () => paths.ContainsKey(path) || creation.IsCompleted,
                TimeSpan.FromSeconds(15));
            manager.Dispose();
            var failure = await Record.ExceptionAsync(
                async () => await creation.WaitAsync(TimeSpan.FromSeconds(150)));

            Assert.True(reserved, "Startup never acquired its file path.");
            Assert.IsType<ObjectDisposedException>(Assert.IsType<InvalidOperationException>(failure).InnerException);
            Assert.Equal(0, manager.ActiveSessionCount);
            Assert.Empty(paths);
            foreach (var field in new[] { "_sessionFilePaths", "_activeOperationCounts", "_showExcelFlags", "_sessionOrigins", "_sessionCreatedAt" })
                Assert.Empty(GetField<System.Collections.IDictionary>(manager, field));
        }
        finally
        {
            // Reclaim any late session if the regression fails against the old implementation.
            await Record.ExceptionAsync(async () => await creation.WaitAsync(TimeSpan.FromSeconds(150)));
            foreach (var batch in sessions.Values)
                batch.Dispose();
            File.Delete(path);
        }
    }

    private static T GetField<T>(SessionManager manager, string name) where T : class =>
        Assert.IsAssignableFrom<T>(typeof(SessionManager)
            .GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(manager));
}
