using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Collection("Mac backend state")]
public sealed class MacExcelSessionManagerTests
{
    [Fact]
    public async Task Execute_AfterUncertainOfficeMutation_RejectsFurtherOperations()
    {
        using var manager = CreateManager();
        var sessionId = await manager.OpenAsync(
            "/tmp/uncertain-office-mutation.xlsx",
            show: false,
            TimeSpan.FromSeconds(5));
        var session = Assert.Single(manager.Sessions);
        session.MarkUnsafe("The dispatched mutation may still complete.");

        var error = await Assert.ThrowsAsync<InvalidOperationException>(
            () => manager.ExecuteAsync(sessionId, _ => Task.FromResult(true)));

        Assert.Contains("unsafe", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("may still complete", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Execute_QueuedBeforeUncertainMutation_DoesNotRunAfterSessionBecomesUnsafe(
        bool requiresRecovery)
    {
        using var manager = CreateManager();
        var sessionId = await manager.OpenAsync(
            "/tmp/queued-after-uncertain-mutation.xlsx",
            show: false,
            TimeSpan.FromSeconds(5));
        var session = Assert.Single(manager.Sessions);
        var firstStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseFirst = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var secondRan = false;
        var first = manager.ExecuteAsync(sessionId, async _ =>
        {
            firstStarted.SetResult();
            await releaseFirst.Task;
            return true;
        });
        await firstStarted.Task;
        var second = manager.ExecuteAsync(sessionId, _ =>
        {
            secondRan = true;
            return Task.FromResult(true);
        });
        await WaitForAsync(() => session.PendingOperations == 2);

        if (requiresRecovery) manager.RequireRecovery(sessionId);
        else session.MarkUnsafe("The dispatched mutation may still complete.");
        releaseFirst.SetResult();

        Assert.True(await first);
        await Assert.ThrowsAsync<InvalidOperationException>(() => second);
        Assert.False(secondRan);
    }

    [Fact]
    public async Task Close_DrainsAdmittedOperationsBeforeDisposingSession()
    {
        using var manager = CreateManager();
        var sessionId = await manager.OpenAsync(
            "/tmp/session-race.xlsx",
            show: false,
            TimeSpan.FromSeconds(5));
        var session = Assert.Single(manager.Sessions);
        var releaseFirst = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var firstStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var secondRan = false;

        var first = manager.ExecuteAsync(sessionId, async _ =>
        {
            firstStarted.SetResult();
            await releaseFirst.Task;
            return true;
        });
        await firstStarted.Task;
        var second = manager.ExecuteAsync(sessionId, _ =>
        {
            secondRan = true;
            return Task.FromResult(true);
        });
        await WaitForAsync(() => session.PendingOperations == 2);

        var close = manager.CloseAsync(sessionId, save: false);
        Assert.False(close.IsCompleted);
        releaseFirst.SetResult();

        Assert.True(await first);
        Assert.True(await second);
        Assert.True(await close);
        Assert.True(secondRan);
        await Assert.ThrowsAsync<KeyNotFoundException>(
            () => manager.ExecuteAsync(sessionId, _ => Task.FromResult(true)));
    }

    [Fact]
    public async Task ConcurrentClose_UsesOneBackendClose()
    {
        var closeCalls = 0;
        var closeStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseClose = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var backend = new MacExcelBackend(async (start, input, cancellationToken) =>
        {
            if (start.FileName != "/usr/bin/open" && AutomationCommand(start) == "session.close")
            {
                Interlocked.Increment(ref closeCalls);
                closeStarted.SetResult();
                await releaseClose.Task.WaitAsync(cancellationToken);
            }

            return Success();
        });
        using var manager = new MacExcelSessionManager(backend);
        var sessionId = await manager.OpenAsync(
            "/tmp/concurrent-close.xlsx",
            show: false,
            TimeSpan.FromSeconds(5));

        var first = manager.CloseAsync(sessionId, save: false);
        await closeStarted.Task;
        var second = manager.CloseAsync(sessionId, save: false);
        releaseClose.SetResult();

        Assert.True(await first);
        Assert.True(await second);
        Assert.Equal(1, closeCalls);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Create_CopiesOpaqueTemplateBeforeOpening(bool macroEnabled)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-template-");
        var extension = macroEnabled ? ".xlsm" : ".xlsx";
        var path = Path.Combine(directory.FullName, $"created{extension}");
        var expected = macroEnabled ? "opaque-xlsm" : "opaque-xlsx";
        string? observed = null;
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            if (start.FileName == "/usr/bin/open")
            {
                observed = File.ReadAllText(path);
            }
            return Task.FromResult(Success());
        });
        using var manager = new MacExcelSessionManager(
            backend,
            (destination, isMacroEnabled) =>
                File.WriteAllText(destination, isMacroEnabled ? "opaque-xlsm" : "opaque-xlsx"));

        var sessionId = await manager.CreateAsync(
            path,
            macroEnabled,
            show: false,
            TimeSpan.FromSeconds(5));

        Assert.Equal(expected, observed);
        Assert.True(await manager.CloseAsync(sessionId, save: false));
        File.Delete(path);
        directory.Delete();
    }

    [Fact]
    public async Task Create_WhenDestinationAppears_DoesNotDeleteForeignFile()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-template-race-");
        var path = Path.Combine(directory.FullName, "created.xlsx");
        await File.WriteAllTextAsync(path, "foreign");
        using var manager = new MacExcelSessionManager(
            CreateBackend(),
            (destination, _) =>
            {
                using var stream = new FileStream(
                    destination,
                    FileMode.CreateNew,
                    FileAccess.Write,
                    FileShare.None);
                stream.WriteByte(1);
            });

        await Assert.ThrowsAsync<IOException>(
            () => manager.CreateAsync(path, macroEnabled: false, show: false, TimeSpan.FromSeconds(5)));

        Assert.Equal("foreign", await File.ReadAllTextAsync(path));
        File.Delete(path);
        directory.Delete();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Create_OnlyDeletesNewWorkbookWhenFailurePrecedesHandoff(bool afterHandoff)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-open-recovery-");
        var path = Path.Combine(directory.FullName, "created.xlsx");
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            var fail = start.FileName != "/usr/bin/open"
                && AutomationCommand(start) == (afterHandoff ? "session.open" : "session.prepare-open");
            return Task.FromResult(fail
                ? Success(
                    """{"success":false,"errorCategory":"ComInterop","errorMessage":"Open was not confirmed."}""")
                : Success());
        });
        try
        {
            using var manager = new MacExcelSessionManager(
                backend,
                (destination, _) => File.WriteAllText(destination, "opaque"));
            var error = await Assert.ThrowsAsync<MacExcelOperationException>(() =>
                manager.CreateAsync(path, macroEnabled: false, show: false, TimeSpan.FromSeconds(5)));

            Assert.Equal(afterHandoff, File.Exists(path));
            Assert.Equal(afterHandoff ? "RecoveryRequired" : "ComInterop", error.ErrorCategory);
            Assert.Empty(manager.Sessions);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
            directory.Delete();
        }
    }

    [Fact]
    public async Task Close_TimeoutWhileWorkbookRemainsOpenCancelsLogicalClose()
    {
        var backend = new MacExcelBackend(async (start, input, cancellationToken) =>
        {
            if (start.FileName != "/usr/bin/open")
            {
                switch (AutomationCommand(start))
                {
                    case "session.close":
                        await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
                        break;
                    case "session.is-open":
                        return Success("""{"success":true,"errorMessage":"","open":true}""");
                }
            }
            return Success();
        });
        using var manager = new MacExcelSessionManager(backend);
        var sessionId = await manager.OpenAsync(
            "/tmp/close-timeout-open.xlsx",
            show: false,
            TimeSpan.FromMilliseconds(50));

        await Assert.ThrowsAsync<TimeoutException>(
            () => manager.CloseAsync(sessionId, save: false));

        Assert.True(await manager.ExecuteAsync(sessionId, _ => Task.FromResult(true)));
    }

    [Fact]
    public async Task Close_IndeterminateTimeoutMarksSessionRecoveryRequired()
    {
        var backend = new MacExcelBackend(async (start, input, cancellationToken) =>
        {
            if (start.FileName != "/usr/bin/open"
                && AutomationCommand(start) is "session.close" or "session.is-open")
            {
                await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
            }
            return Success();
        });
        using var manager = new MacExcelSessionManager(backend);
        var sessionId = await manager.OpenAsync(
            "/tmp/close-timeout-indeterminate.xlsx",
            show: false,
            TimeSpan.FromMilliseconds(50));

        var error = await Assert.ThrowsAsync<InvalidOperationException>(
            () => manager.CloseAsync(sessionId, save: false));

        Assert.Contains("could not determine", error.Message, StringComparison.OrdinalIgnoreCase);
        await Assert.ThrowsAsync<InvalidOperationException>(
            () => manager.ExecuteAsync(sessionId, _ => Task.FromResult(true)));
        Assert.True(Assert.Single(manager.Sessions).RequiresRecovery);
    }

    private static MacExcelSessionManager CreateManager() =>
        new(CreateBackend(), (destination, _) => File.WriteAllText(destination, "opaque"));

    private static MacExcelBackend CreateBackend() =>
        new((start, input, cancellationToken) => Task.FromResult(Success()));

    private static MacProcessResult Success(
        string output = """{"success":true,"errorMessage":""}""") =>
        new(0, output, "");

    private static async Task WaitForAsync(Func<bool> predicate)
    {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        while (!predicate())
        {
            await Task.Delay(10, timeout.Token);
        }
    }

    private static string AutomationCommand(System.Diagnostics.ProcessStartInfo start)
    {
        var arguments = start.ArgumentList.ToArray();
        var marker = Array.IndexOf(arguments, MacAutomationHost.Marker);
        Assert.True(marker >= 0 && marker + 1 < arguments.Length);
        return arguments[marker + 1];
    }
}
