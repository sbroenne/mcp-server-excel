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
        }, () => 0);
        const string path = "/tmp/workbook with 'quotes' and spaces.xlsx";

        await backend.InvokeAsync("session.open", new { filePath = path, show = false }, TimeSpan.FromSeconds(5));

        Assert.Equal(3, calls.Count);
        Assert.Equal("session.prepare-open", calls[0].ArgumentList[3]);
        Assert.Equal("/usr/bin/open", calls[1].FileName);
        Assert.Equal(["-g", "-b", "com.microsoft.Excel", path], calls[1].ArgumentList.ToArray());
        Assert.Equal("session.open", calls[2].ArgumentList[3]);
    }

    [Theory]
    [InlineData(-1743)]
    [InlineData(-1744)]
    [InlineData(-600)]
    public async Task PermissionFailure_DoesNotLaunchAnyProcess(int status)
    {
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
            throw new InvalidOperationException("A process must not be launched."), () => status);

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = "/tmp/test.xlsx" }, TimeSpan.FromSeconds(5)));

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
        }, () => 0);

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
        }, () => 0);

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = "/tmp/duplicate.xlsx" }, TimeSpan.FromSeconds(5)));

        Assert.Single(calls);
        Assert.Equal("session.prepare-open", calls[0].ArgumentList[3]);
        Assert.Contains("same name", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Timeout_BoundsTheWholeOperation()
    {
        var backend = new MacExcelBackend(async (start, input, cancellationToken) =>
        {
            await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
            throw new InvalidOperationException("Unreachable.");
        }, () => 0);

        await Assert.ThrowsAsync<TimeoutException>(() => backend.InvokeAsync(
            "session.open", new { filePath = "/tmp/test.xlsx" }, TimeSpan.FromMilliseconds(50)));
    }
}
