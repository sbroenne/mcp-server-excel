using System.Diagnostics;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacExcelBackendTests
{
    [Fact]
    public async Task Open_HandsOffExactPathBeforeAttachingWorkbook()
    {
        var calls = new List<ProcessStartInfo>();
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            calls.Add(start);
            return Task.FromResult(new MacProcessResult(0, """{"success":true,"errorMessage":""}""", ""));
        });
        const string path = "/tmp/workbook with 'quotes' and spaces.xlsx";

        await backend.InvokeAsync("session.open", new { filePath = path, show = false }, TimeSpan.FromSeconds(5));

        Assert.Equal(3, calls.Count);
        Assert.Equal("session.prepare-open", AutomationCommand(calls[0]));
        Assert.Equal("/usr/bin/open", calls[1].FileName);
        Assert.Equal(["-g", "-b", "com.microsoft.Excel", path], calls[1].ArgumentList.ToArray());
        Assert.Equal("session.open", AutomationCommand(calls[2]));
    }

    [Theory]
    [InlineData(-1743)]
    [InlineData(-1744)]
    [InlineData(-600)]
    public async Task PermissionFailure_DoesNotLaunchAnyProcess(int status)
    {
        var calls = new List<ProcessStartInfo>();
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            calls.Add(start);
            return Task.FromResult(new MacProcessResult(0,
                $$"""{"success":false,"errorMessage":"Automation unavailable (OSStatus {{status}}).","errorCategory":"ComInterop"}""",
                ""));
        });

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = "/tmp/test.xlsx" }, TimeSpan.FromSeconds(5)));

        Assert.Single(calls);
        Assert.Equal("session.prepare-open", AutomationCommand(calls[0]));
        Assert.Contains(status.ToString(System.Globalization.CultureInfo.InvariantCulture), error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task HandoffFailure_DoesNotAttachWorkbook()
    {
        var count = 0;
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            count++;
            return Task.FromResult(start.FileName == "/usr/bin/open"
                ? new MacProcessResult(1, "", "LaunchServices rejected the file")
                : new MacProcessResult(0, """{"success":true}""", ""));
        });

        var error = await Assert.ThrowsAsync<InvalidOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = "/tmp/test.xlsx" }, TimeSpan.FromSeconds(5)));

        Assert.Equal(2, count);
        Assert.Contains("LaunchServices", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task DuplicateWorkbookName_DoesNotLaunchHandoff()
    {
        var calls = new List<ProcessStartInfo>();
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            calls.Add(start);
            return Task.FromResult(new MacProcessResult(0,
                """{"success":false,"errorMessage":"A workbook with the same name is already open in shared Excel.","errorCategory":"ComInterop"}""",
                ""));
        });

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = "/tmp/duplicate.xlsx" }, TimeSpan.FromSeconds(5)));

        Assert.Single(calls);
        Assert.Equal("session.prepare-open", AutomationCommand(calls[0]));
        Assert.Contains("same name", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConcurrentOpens_DoNotInterleavePreflightAndHandoff()
    {
        var calls = new List<string>();
        var firstHandoffStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseFirstHandoff = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var backend = new MacExcelBackend(async (start, input, cancellationToken) =>
        {
            var command = start.FileName == "/usr/bin/open"
                ? $"open:{start.ArgumentList[^1]}"
                : AutomationCommand(start);
            lock (calls)
            {
                calls.Add(command);
            }

            if (command == "open:/tmp/first.xlsx")
            {
                firstHandoffStarted.SetResult();
                await releaseFirstHandoff.Task.WaitAsync(cancellationToken);
            }

            return new MacProcessResult(0, """{"success":true,"errorMessage":""}""", "");
        });

        var first = backend.InvokeAsync(
            "session.open", new { filePath = "/tmp/first.xlsx" }, TimeSpan.FromSeconds(5));
        await firstHandoffStarted.Task;
        var second = backend.InvokeAsync(
            "session.open", new { filePath = "/tmp/second.xlsx" }, TimeSpan.FromSeconds(5));
        await Task.Delay(50);

        lock (calls)
        {
            Assert.Equal(["session.prepare-open", "open:/tmp/first.xlsx"], calls);
        }

        releaseFirstHandoff.SetResult();
        await Task.WhenAll(first, second);
        lock (calls)
        {
            Assert.Equal(
                [
                    "session.prepare-open",
                    "open:/tmp/first.xlsx",
                    "session.open",
                    "session.prepare-open",
                    "open:/tmp/second.xlsx",
                    "session.open"
                ],
                calls);
        }
    }

    [Fact]
    public async Task Timeout_BoundsTheWholeOperation()
    {
        var backend = new MacExcelBackend(async (start, input, cancellationToken) =>
        {
            await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
            throw new InvalidOperationException("Unreachable.");
        });

        await Assert.ThrowsAsync<TimeoutException>(() => backend.InvokeAsync(
            "session.open", new { filePath = "/tmp/test.xlsx" }, TimeSpan.FromMilliseconds(50)));
    }

    private static string AutomationCommand(ProcessStartInfo start)
    {
        var arguments = start.ArgumentList.ToArray();
        var marker = Array.IndexOf(arguments, MacAutomationHost.Marker);
        Assert.True(marker >= 0 && marker + 1 < arguments.Length);
        return arguments[marker + 1];
    }
}
