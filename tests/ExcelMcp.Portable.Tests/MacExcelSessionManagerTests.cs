using System.IO.Compression;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacExcelSessionManagerTests
{
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

            return new MacProcessResult(0, """{"success":true,"errorMessage":""}""", "");
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
    [InlineData(false, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")]
    [InlineData(true, "application/vnd.ms-excel.sheet.macroEnabled.main+xml")]
    public async Task Create_WritesValidPackageBeforeOpening(
        bool macroEnabled,
        string expectedWorkbookContentType)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-package-");
        var extension = macroEnabled ? ".xlsm" : ".xlsx";
        var path = Path.Combine(directory.FullName, $"created{extension}");
        string? packageContentType = null;
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            if (start.FileName == "/usr/bin/open")
            {
                using var archive = ZipFile.OpenRead(path);
                using var reader = new StreamReader(archive.GetEntry("[Content_Types].xml")!.Open());
                packageContentType = reader.ReadToEnd();
            }
            return Task.FromResult(new MacProcessResult(0, """{"success":true,"errorMessage":""}""", ""));
        });
        using var manager = new MacExcelSessionManager(backend);

        var sessionId = await manager.CreateAsync(path, macroEnabled, show: false, TimeSpan.FromSeconds(5));
        Assert.True(File.Exists(path));
        Assert.Contains(expectedWorkbookContentType, packageContentType, StringComparison.Ordinal);
        Assert.True(await manager.CloseAsync(sessionId, save: false));

        File.Delete(path);
        directory.Delete();
    }

    [Fact]
    public async Task Create_WhenDestinationAppears_DoesNotDeleteForeignFile()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-package-race-");
        var path = Path.Combine(directory.FullName, "created.xlsx");
        await File.WriteAllTextAsync(path, "foreign");
        using var manager = CreateManager();

        await Assert.ThrowsAsync<IOException>(
            () => manager.CreateAsync(path, macroEnabled: false, show: false, TimeSpan.FromSeconds(5)));

        Assert.Equal("foreign", await File.ReadAllTextAsync(path));
        File.Delete(path);
        directory.Delete();
    }

    [Theory]
    [InlineData(false, "original")]
    [InlineData(true, "updated")]
    public async Task PackageMutation_CloseControlsPersistence(
        bool save,
        string expectedContent)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-package-transaction-");
        var path = Path.Combine(directory.FullName, "query.xlsx");
        await File.WriteAllTextAsync(path, "original");
        using var manager = CreateManager();
        var sessionId = await manager.OpenAsync(
            path,
            show: false,
            TimeSpan.FromSeconds(5));

        await manager.ExecuteAsync(sessionId, async session =>
        {
            await manager.MutatePackageAsync(
                session,
                workingPath => File.WriteAllText(workingPath, "updated"),
                static () => Task.CompletedTask);
            return true;
        });

        Assert.Equal("updated", await File.ReadAllTextAsync(path));
        Assert.True(await manager.CloseAsync(sessionId, save));
        Assert.Equal(expectedContent, await File.ReadAllTextAsync(path));

        File.Delete(path);
        directory.Delete();
    }

    [Fact]
    public async Task PackageMutation_PostReopenFailureRollsBackOperation()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-package-rollback-");
        var path = Path.Combine(directory.FullName, "query.xlsx");
        await File.WriteAllTextAsync(path, "original");
        using var manager = CreateManager();
        var sessionId = await manager.OpenAsync(
            path,
            show: true,
            TimeSpan.FromSeconds(5));

        var error = await Assert.ThrowsAsync<InvalidOperationException>(
            () => manager.ExecuteAsync(sessionId, async session =>
            {
                await manager.MutatePackageAsync(
                    session,
                    workingPath => File.WriteAllText(workingPath, "updated"),
                    () => throw new InvalidOperationException("refresh failed"));
                return true;
            }));

        Assert.Contains("refresh failed", error.Message, StringComparison.Ordinal);
        Assert.Equal("original", await File.ReadAllTextAsync(path));
        Assert.True(await manager.CloseAsync(sessionId, save: false));

        File.Delete(path);
        directory.Delete();
    }

    [Fact]
    public async Task PackageMutation_UserSaveImmediatelyBeforeCloseBecomesBaseline()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-package-save-race-");
        var path = Path.Combine(directory.FullName, "query.xlsx");
        await File.WriteAllTextAsync(path, "original");
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            if (start.FileName != "/usr/bin/open" &&
                AutomationCommand(start) == "session.close-if-saved")
            {
                File.WriteAllText(path, "latest-user-save");
            }

            return Task.FromResult(Success());
        });
        using var manager = new MacExcelSessionManager(backend);
        var sessionId = await manager.OpenAsync(
            path,
            show: false,
            TimeSpan.FromSeconds(5));

        await manager.ExecuteAsync(sessionId, async session =>
        {
            await manager.MutatePackageAsync(
                session,
                workingPath => File.WriteAllText(workingPath, "updated"),
                static () => Task.CompletedTask);
            return true;
        });

        Assert.True(await manager.CloseAsync(sessionId, save: false));
        Assert.Equal("latest-user-save", await File.ReadAllTextAsync(path));

        File.Delete(path);
        directory.Delete();
    }

    [Fact]
    public async Task PackageMutation_EditAfterStateCheckRejectsAtomicCloseWithoutDiscardingEdits()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-package-dirty-race-");
        var path = Path.Combine(directory.FullName, "query.xlsx");
        await File.WriteAllTextAsync(path, "original");
        var commands = new List<string>();
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            if (start.FileName != "/usr/bin/open")
            {
                var command = AutomationCommand(start);
                commands.Add(command);
                if (command == "session.close-if-saved")
                {
                    return Task.FromResult(new MacProcessResult(
                        0,
                        """{"success":false,"errorMessage":"Workbook has unsaved changes.","errorCategory":"InvalidOperation"}""",
                        ""));
                }
                if (command == "session.is-open")
                {
                    return Task.FromResult(new MacProcessResult(
                        0,
                        """{"success":true,"errorMessage":"","open":true}""",
                        ""));
                }
            }

            return Task.FromResult(new MacProcessResult(0, """{"success":true,"errorMessage":""}""", ""));
        });
        using var manager = new MacExcelSessionManager(backend);
        var sessionId = await manager.OpenAsync(
            path,
            show: false,
            TimeSpan.FromSeconds(5));
        var mutationRan = false;

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(
            () => manager.ExecuteAsync(sessionId, async session =>
            {
                await manager.MutatePackageAsync(
                    session,
                    workingPath =>
                    {
                        mutationRan = true;
                        File.WriteAllText(workingPath, "updated");
                    },
                    static () => Task.CompletedTask);
                return true;
            }));

        Assert.Contains("unsaved changes", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(mutationRan);
        Assert.Contains("session.close-if-saved", commands);
        Assert.DoesNotContain("session.close", commands);
        Assert.Equal("original", await File.ReadAllTextAsync(path));
        Assert.Empty(TransactionArtifacts(directory));
        Assert.True(await manager.ExecuteAsync(sessionId, _ => Task.FromResult(true)));
        Assert.True(await manager.CloseAsync(sessionId, save: false));

        File.Delete(path);
        directory.Delete();
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public async Task PackageMutation_SetupCopyFailureRemovesNewTransactionArtifacts(int failingCopy)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-package-setup-failure-");
        var path = Path.Combine(directory.FullName, "query.xlsx");
        await File.WriteAllTextAsync(path, "original");
        var copyCount = 0;
        using var manager = CreateManager((source, destination) =>
        {
            copyCount++;
            if (copyCount == failingCopy)
            {
                throw new IOException($"Copy {failingCopy} failed.");
            }
            File.Copy(source, destination, overwrite: false);
        });
        var sessionId = await manager.OpenAsync(
            path,
            show: false,
            TimeSpan.FromSeconds(5));

        var error = await Assert.ThrowsAsync<IOException>(
            () => manager.ExecuteAsync(sessionId, async session =>
            {
                await manager.MutatePackageAsync(
                    session,
                    workingPath => File.WriteAllText(workingPath, "updated"),
                    static () => Task.CompletedTask);
                return true;
            }));

        Assert.Contains($"Copy {failingCopy} failed", error.Message, StringComparison.Ordinal);
        var session = Assert.Single(manager.Sessions);
        Assert.Null(session.PackageBaselinePath);
        Assert.Null(session.PackageTransactionPath);
        Assert.Empty(TransactionArtifacts(directory));
        Assert.Equal("original", await File.ReadAllTextAsync(path));
        Assert.True(await manager.CloseAsync(sessionId, save: false));

        File.Delete(path);
        directory.Delete();
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public async Task PackageMutation_SetupCopyFailurePreservesExistingTransaction(int failingSetupCopy)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-package-existing-setup-failure-");
        var path = Path.Combine(directory.FullName, "query.xlsx");
        await File.WriteAllTextAsync(path, "original");
        var copyCount = 0;
        int? failAtCopy = null;
        using var manager = CreateManager((source, destination) =>
        {
            copyCount++;
            if (copyCount == failAtCopy)
            {
                throw new IOException($"Setup copy {failingSetupCopy} failed.");
            }
            File.Copy(source, destination, overwrite: false);
        });
        var sessionId = await manager.OpenAsync(
            path,
            show: false,
            TimeSpan.FromSeconds(5));
        await manager.ExecuteAsync(sessionId, async session =>
        {
            await manager.MutatePackageAsync(
                session,
                workingPath => File.WriteAllText(workingPath, "staged"),
                static () => Task.CompletedTask);
            return true;
        });
        var session = Assert.Single(manager.Sessions);
        var baselinePath = Assert.IsType<string>(session.PackageBaselinePath);
        var transactionPath = Assert.IsType<string>(session.PackageTransactionPath);
        failAtCopy = copyCount + failingSetupCopy;

        await Assert.ThrowsAsync<IOException>(
            () => manager.ExecuteAsync(sessionId, async currentSession =>
            {
                await manager.MutatePackageAsync(
                    currentSession,
                    workingPath => File.WriteAllText(workingPath, "second"),
                    static () => Task.CompletedTask);
                return true;
            }));

        Assert.Equal(baselinePath, session.PackageBaselinePath);
        Assert.Equal(transactionPath, session.PackageTransactionPath);
        Assert.True(File.Exists(baselinePath));
        Assert.True(File.Exists(transactionPath));
        Assert.Equal("staged", await File.ReadAllTextAsync(path));
        Assert.Equal(
            new[] { baselinePath, transactionPath }.Order(StringComparer.Ordinal).ToArray(),
            TransactionArtifacts(directory).Order(StringComparer.Ordinal).ToArray());
        Assert.True(await manager.CloseAsync(sessionId, save: false));
        Assert.Equal("original", await File.ReadAllTextAsync(path));

        File.Delete(path);
        directory.Delete();
    }

    [Fact]
    public async Task Open_StalePackageTransactionFailsBeforeExcelHandoff()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-package-stale-");
        var path = Path.Combine(directory.FullName, "query.xlsx");
        var transactionPath = Path.Combine(
            directory.FullName,
            ".query.xlsx.excelmcp-pq-transaction.json");
        await File.WriteAllTextAsync(path, "original");
        await File.WriteAllTextAsync(transactionPath, """{"baseline":"retained.tmp"}""");
        using var manager = CreateManager();

        var error = await Assert.ThrowsAsync<InvalidOperationException>(
            () => manager.OpenAsync(path, show: false, TimeSpan.FromSeconds(5)));

        Assert.Contains("interrupted Power Query package transaction", error.Message, StringComparison.OrdinalIgnoreCase);

        File.Delete(transactionPath);
        File.Delete(path);
        directory.Delete();
    }

    [Fact]
    public async Task Close_RestoreFailureRemovesClosedSessionAndRetainsJournal()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-package-close-failure-");
        var path = Path.Combine(directory.FullName, "query.xlsx");
        await File.WriteAllTextAsync(path, "original");
        using var manager = CreateManager();
        var sessionId = await manager.OpenAsync(
            path,
            show: false,
            TimeSpan.FromSeconds(5));
        await manager.ExecuteAsync(sessionId, async session =>
        {
            await manager.MutatePackageAsync(
                session,
                workingPath => File.WriteAllText(workingPath, "updated"),
                static () => Task.CompletedTask);
            return true;
        });
        var session = Assert.Single(manager.Sessions);
        var transactionPath = Assert.IsType<string>(session.PackageTransactionPath);
        File.Delete(Assert.IsType<string>(session.PackageBaselinePath));

        await Assert.ThrowsAsync<FileNotFoundException>(
            () => manager.CloseAsync(sessionId, save: false));

        await Assert.ThrowsAsync<KeyNotFoundException>(
            () => manager.ExecuteAsync(sessionId, _ => Task.FromResult(true)));
        Assert.True(File.Exists(transactionPath));

        File.Delete(transactionPath);
        File.Delete(path);
        directory.Delete();
    }

    [Fact]
    public async Task Close_TimeoutAfterWorkbookClosedReconcilesAndRestoresBaseline()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-package-close-timeout-");
        var path = Path.Combine(directory.FullName, "query.xlsx");
        await File.WriteAllTextAsync(path, "original");
        var open = true;
        var timeoutFinalClose = false;
        var backend = new MacExcelBackend(async (start, input, cancellationToken) =>
        {
            if (start.FileName == "/usr/bin/open")
            {
                return Success();
            }

            switch (AutomationCommand(start))
            {
                case "session.prepare-open":
                    return Success();
                case "session.open":
                    open = true;
                    return Success();
                case "session.close-if-saved":
                    open = false;
                    return Success();
                case "session.close" when timeoutFinalClose:
                    open = false;
                    await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
                    throw new InvalidOperationException("Unreachable.");
                case "session.close":
                    open = false;
                    return Success();
                case "session.is-open":
                    Assert.Equal(path, InputFilePath(input));
                    return Success($$"""{"success":true,"errorMessage":"","open":{{JsonSerializer.Serialize(open)}}}""");
                default:
                    return Success();
            }
        });
        using var manager = new MacExcelSessionManager(backend);
        var sessionId = await manager.OpenAsync(
            path,
            show: false,
            TimeSpan.FromMilliseconds(50));
        await manager.ExecuteAsync(sessionId, async session =>
        {
            await manager.MutatePackageAsync(
                session,
                workingPath => File.WriteAllText(workingPath, "updated"),
                static () => Task.CompletedTask);
            return true;
        });
        timeoutFinalClose = true;

        Assert.True(await manager.CloseAsync(sessionId, save: false));

        Assert.Equal("original", await File.ReadAllTextAsync(path));
        Assert.Empty(TransactionArtifacts(directory));
        Assert.Empty(manager.Sessions);

        File.Delete(path);
        directory.Delete();
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
            if (start.FileName != "/usr/bin/open" &&
                AutomationCommand(start) is "session.close" or "session.is-open")
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
        Assert.True(Assert.Single(manager.Sessions).RequiresPackageRecovery);
    }

    [Fact]
    public async Task Dispose_CloseFailureRetainsStagedWorkbookAndRecoveryArtifacts()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-package-dispose-failure-");
        var path = Path.Combine(directory.FullName, "query.xlsx");
        await File.WriteAllTextAsync(path, "original");
        var failFinalClose = false;
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            if (start.FileName != "/usr/bin/open" &&
                AutomationCommand(start) == "session.close" &&
                failFinalClose)
            {
                return Task.FromResult(Success(
                    """{"success":false,"errorMessage":"close failed","errorCategory":"ComInterop"}"""));
            }

            return Task.FromResult(Success());
        });
        var manager = new MacExcelSessionManager(backend);
        var sessionId = await manager.OpenAsync(
            path,
            show: false,
            TimeSpan.FromSeconds(5));
        await manager.ExecuteAsync(sessionId, async session =>
        {
            await manager.MutatePackageAsync(
                session,
                workingPath => File.WriteAllText(workingPath, "updated"),
                static () => Task.CompletedTask);
            return true;
        });
        var session = Assert.Single(manager.Sessions);
        var baselinePath = Assert.IsType<string>(session.PackageBaselinePath);
        var transactionPath = Assert.IsType<string>(session.PackageTransactionPath);
        failFinalClose = true;

        manager.Dispose();

        Assert.Equal("updated", await File.ReadAllTextAsync(path));
        Assert.True(File.Exists(baselinePath));
        Assert.True(File.Exists(transactionPath));

        File.Delete(baselinePath);
        File.Delete(transactionPath);
        File.Delete(path);
        directory.Delete();
    }

    [Fact]
    public async Task Dispose_CloseTimeoutAfterSideEffectRestoresBaseline()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-package-dispose-timeout-");
        var path = Path.Combine(directory.FullName, "query.xlsx");
        await File.WriteAllTextAsync(path, "original");
        var timeoutFinalClose = false;
        var backend = new MacExcelBackend(async (start, input, cancellationToken) =>
        {
            if (start.FileName != "/usr/bin/open")
            {
                switch (AutomationCommand(start))
                {
                    case "session.close" when timeoutFinalClose:
                        await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
                        break;
                    case "session.is-open":
                        Assert.Equal(path, InputFilePath(input));
                        return Success("""{"success":true,"errorMessage":"","open":false}""");
                }
            }
            return Success();
        });
        var manager = new MacExcelSessionManager(backend);
        var sessionId = await manager.OpenAsync(
            path,
            show: false,
            TimeSpan.FromMilliseconds(50));
        await manager.ExecuteAsync(sessionId, async session =>
        {
            await manager.MutatePackageAsync(
                session,
                workingPath => File.WriteAllText(workingPath, "updated"),
                static () => Task.CompletedTask);
            return true;
        });
        timeoutFinalClose = true;

        manager.Dispose();

        Assert.Equal("original", await File.ReadAllTextAsync(path));
        Assert.Empty(TransactionArtifacts(directory));

        File.Delete(path);
        directory.Delete();
    }

    private static MacExcelSessionManager CreateManager(Action<string, string>? copyFile = null)
    {
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
            Task.FromResult(new MacProcessResult(0, """{"success":true,"errorMessage":""}""", "")));
        return new MacExcelSessionManager(backend, copyFile);
    }

    private static string[] TransactionArtifacts(DirectoryInfo directory) =>
        Directory.GetFiles(directory.FullName)
            .Where(candidate =>
                Path.GetFileName(candidate).Contains("excelmcp-pq-", StringComparison.Ordinal))
            .ToArray();

    private static MacProcessResult Success(
        string output = """{"success":true,"errorMessage":""}""") =>
        new(0, output, "");

    private static string InputFilePath(string? input)
    {
        using var document = JsonDocument.Parse(Assert.IsType<string>(input));
        return Assert.IsType<string>(document.RootElement.GetProperty("filePath").GetString());
    }

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
