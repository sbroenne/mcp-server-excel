using Sbroenne.ExcelMcp.Service.Mac;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Collection("Mac backend state")]
public sealed class MacExcelSessionManagerTests
{
    private static string TestPath(string name) => Path.Combine(Path.GetTempPath(), name);

    [Fact]
    public async Task UnconfirmedOpen_RetainsRecoverySessionAndRejectsRepeatedOpenAndCreate()
    {
        var calls = 0;
        var backend = new MacExcelBackend((start, _, _) =>
        {
            calls++;
            return Task.FromResult(start.FileName != "/usr/bin/open" && AutomationCommand(start) == "session.open"
                ? Success("""{"success":false,"errorCategory":"ComInterop","errorMessage":"Attachment unconfirmed."}""")
                : Success());
        });
        using var manager = new MacExcelSessionManager(backend, (_, _) => throw new InvalidOperationException("Must not create."));
        var path = TestPath($"recovery-{Guid.NewGuid():N}.xlsx");
        var failure = await Assert.ThrowsAsync<MacExcelOperationException>(() =>
            manager.OpenAsync(path, false, TimeSpan.FromSeconds(5)));
        Assert.Equal("RecoveryRequired", failure.ErrorCategory);
        var recovery = Assert.Single(manager.Sessions);
        Assert.True(recovery.HasUnconfirmedOpen);
        Assert.True(recovery.RequiresRecovery);
        var callsBeforeRetries = calls;

        await Assert.ThrowsAsync<MacExcelOperationException>(() => manager.OpenAsync(path, false, TimeSpan.FromSeconds(5)));
        await Assert.ThrowsAsync<MacExcelOperationException>(() => manager.CreateAsync(path, false, false, TimeSpan.FromSeconds(5)));
        await Assert.ThrowsAsync<InvalidOperationException>(() => manager.ExecuteAsync(recovery.SessionId, _ => Task.FromResult(true)));
        await Assert.ThrowsAsync<MacExcelOperationException>(() => manager.CloseAsync(recovery.SessionId, false));
        Assert.Equal(callsBeforeRetries, calls);
        manager.Dispose();
        Assert.Equal(callsBeforeRetries, calls);
    }

    [Fact]
    public async Task Execute_AfterUncertainOfficeMutation_RejectsFurtherOperations()
    {
        using var manager = CreateManager();
        var sessionId = await manager.OpenAsync(
            TestPath("uncertain-office-mutation.xlsx"),
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
            TestPath("queued-after-uncertain-mutation.xlsx"),
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
            TestPath("session-race.xlsx"),
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
            TestPath("concurrent-close.xlsx"),
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
    public async Task ConcurrentClose_RejectsConflictingSaveIntent(bool save)
    {
        var closeCalls = 0;
        var closeStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseClose = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var backend = new MacExcelBackend(async (start, input, cancellationToken) =>
        {
            if (start.FileName != "/usr/bin/open" && AutomationCommand(start) == "session.close")
            {
                using var arguments = JsonDocument.Parse(input!);
                Assert.Equal(save, arguments.RootElement.GetProperty("save").GetBoolean());
                Interlocked.Increment(ref closeCalls);
                closeStarted.SetResult();
                await releaseClose.Task.WaitAsync(cancellationToken);
            }
            return Success();
        });
        using var manager = new MacExcelSessionManager(backend);
        var sessionId = await manager.OpenAsync(TestPath("conflicting-close.xlsx"), false, TimeSpan.FromSeconds(5));
        var first = manager.CloseAsync(sessionId, save);
        await closeStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        try
        {
            var error = await Assert.ThrowsAsync<MacExcelOperationException>(
                () => manager.CloseAsync(sessionId, !save).WaitAsync(TimeSpan.FromSeconds(1)));
            Assert.Equal("InvalidOperation", error.ErrorCategory);
            Assert.Contains("save", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, closeCalls);
        }
        finally
        {
            releaseClose.SetResult();
            Assert.True(await first);
        }
    }

    [Fact]
    public async Task Dispose_SavesConfirmedWorkbook()
    {
        bool? saved = null;
        var backend = new MacExcelBackend((start, input, _) =>
        {
            if (start.FileName != "/usr/bin/open" && AutomationCommand(start) == "session.close")
            {
                using var arguments = JsonDocument.Parse(input!);
                saved = arguments.RootElement.GetProperty("save").GetBoolean();
            }
            return Task.FromResult(Success());
        });
        using var manager = new MacExcelSessionManager(backend);
        await manager.OpenAsync(TestPath("shutdown-save.xlsx"), false, TimeSpan.FromSeconds(5));

        manager.Dispose();

        Assert.True(saved);
        Assert.Empty(manager.Sessions);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Dispose_LeavesRecoveryAndUnsafeWorkbooksUntouched(bool requiresRecovery)
    {
        var closeCalls = 0;
        var backend = new MacExcelBackend((start, input, _) =>
        {
            if (start.FileName != "/usr/bin/open" && AutomationCommand(start) == "session.close")
            {
                closeCalls++;
            }
            return Task.FromResult(Success());
        });
        var manager = new MacExcelSessionManager(backend);
        var sessionId = await manager.OpenAsync(
            TestPath($"shutdown-uncertain-{requiresRecovery}.xlsx"),
            false,
            TimeSpan.FromSeconds(5));
        var session = Assert.Single(manager.Sessions);
        if (requiresRecovery)
        {
            manager.RequireRecovery(sessionId);
        }
        else
        {
            session.MarkUnsafe("Mutation outcome is uncertain.");
        }

        manager.Dispose();

        Assert.Equal(0, closeCalls);
        Assert.Empty(manager.Sessions);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Close_RejectsRecoveryAndUnsafeSessionsWithoutTouchingWorkbook(bool requiresRecovery)
    {
        var closeCalls = 0;
        var backend = new MacExcelBackend((start, _, _) =>
        {
            if (start.FileName != "/usr/bin/open" && AutomationCommand(start) == "session.close")
            {
                closeCalls++;
            }
            return Task.FromResult(Success());
        });
        using var manager = new MacExcelSessionManager(backend);
        var sessionId = await manager.OpenAsync(
            TestPath($"close-uncertain-{requiresRecovery}.xlsx"),
            false,
            TimeSpan.FromSeconds(5));
        var session = Assert.Single(manager.Sessions);
        if (requiresRecovery)
        {
            manager.RequireRecovery(sessionId);
        }
        else
        {
            session.MarkUnsafe("Mutation outcome is uncertain.");
        }

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(
            () => manager.CloseAsync(sessionId, save: false));

        Assert.Equal("RecoveryRequired", error.ErrorCategory);
        Assert.Equal(0, closeCalls);
        Assert.Single(manager.Sessions);
    }

    [Fact]
    public async Task Open_UsesCanonicalPathIdentityForSymlinkAliases()
    {
        if (!OperatingSystem.IsMacOS())
        {
            return;
        }

        var directory = Directory.CreateTempSubdirectory("excelmcp-canonical-path-");
        var realDirectory = Directory.CreateDirectory(Path.Combine(directory.FullName, "real"));
        var aliasDirectory = Path.Combine(directory.FullName, "alias");
        Directory.CreateSymbolicLink(aliasDirectory, realDirectory.FullName);
        var realPath = Path.Combine(realDirectory.FullName, "workbook.xlsx");
        await File.WriteAllTextAsync(realPath, "opaque");
        using var manager = CreateManager();
        try
        {
            await manager.OpenAsync(realPath, false, TimeSpan.FromSeconds(5));

            var error = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                manager.OpenAsync(
                    Path.Combine(aliasDirectory, "workbook.xlsx"),
                    false,
                    TimeSpan.FromSeconds(5)));

            Assert.Contains("already open", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Single(manager.Sessions);
        }
        finally
        {
            manager.Dispose();
            File.Delete(aliasDirectory);
            directory.Delete(recursive: true);
        }
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
    public async Task Create_PreservesCreatedWorkbookWhenOpenFails(bool afterHandoff)
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

            Assert.True(File.Exists(path));
            Assert.Equal("opaque", await File.ReadAllTextAsync(path));
            Assert.Equal(afterHandoff ? "RecoveryRequired" : "ComInterop", error.ErrorCategory);
            if (afterHandoff)
            {
                var recovery = Assert.Single(manager.Sessions);
                Assert.True(recovery.HasUnconfirmedOpen);
                Assert.True(recovery.RequiresRecovery);
            }
            else
            {
                Assert.Empty(manager.Sessions);
            }
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
            TestPath("close-timeout-open.xlsx"),
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
            TestPath("close-timeout-indeterminate.xlsx"),
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
