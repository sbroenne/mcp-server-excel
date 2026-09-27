using System.IO.Compression;
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

    private static MacExcelSessionManager CreateManager()
    {
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
            Task.FromResult(new MacProcessResult(0, """{"success":true,"errorMessage":""}""", "")));
        return new MacExcelSessionManager(backend);
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
