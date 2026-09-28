using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Collection("Mac backend state")]
public sealed class MacExcelBackendTests
{
    [Fact]
    public async Task StructuredCommandFailure_CanRemainAResultDto()
    {
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
            Task.FromResult(new MacProcessResult(
                0,
                """{"success":false,"filePath":"/tmp/test.xlsx","isPythonError":true,"errorMessage":"#PYTHON! - Python code raised an error"}""",
                "")));

        var result = await backend.InvokeAsync(
            "pythoninexcel.get-result",
            new { filePath = "/tmp/test.xlsx" },
            TimeSpan.FromSeconds(5),
            allowFailureResult: true);

        Assert.False(result.GetProperty("success").GetBoolean());
        Assert.True(result.GetProperty("isPythonError").GetBoolean());
    }

    [Fact]
    public async Task AutomationFailure_IsNotMistakenForStructuredCommandFailure()
    {
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
            Task.FromResult(new MacProcessResult(
                0,
                """{"success":false,"errorCategory":"ComInterop","errorMessage":"Worksheet does not exist."}""",
                "")));

        await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "pythoninexcel.get-result",
            new { filePath = "/tmp/test.xlsx" },
            TimeSpan.FromSeconds(5),
            allowFailureResult: true));
    }

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

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = "/tmp/test.xlsx" }, TimeSpan.FromSeconds(5)));

        Assert.Equal(2, count);
        Assert.Equal("RecoveryRequired", error.ErrorCategory);
        Assert.Contains("LaunchServices", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("""{"success":false,"errorMessage":"Workbook attachment timed out.","errorCategory":"ComInterop"}""")]
    [InlineData("invalid JSON")]
    [InlineData("{}")]
    [InlineData("""{"success":true,"errorMessage":"Excel rejected the workbook."}""")]
    public async Task Open_UnconfirmedAttachmentRequiresRecoveryInsteadOfAutomaticRetry(string response)
    {
        var calls = 0;
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            calls++;
            return Task.FromResult(new MacProcessResult(
                0, calls == 3 ? response : """{"success":true}""", ""));
        });

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = "/tmp/unconfirmed.xlsx" }, TimeSpan.FromSeconds(5)));

        Assert.Equal(3, calls);
        Assert.Equal("RecoveryRequired", error.ErrorCategory);
        Assert.NotNull(error.InnerException);
        Assert.Contains("Do not retry automatically", error.Message, StringComparison.Ordinal);
        Assert.Contains("desktop", error.Message, StringComparison.Ordinal);
        Assert.Contains("dialogs", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Open_TimeoutDuringOrAfterHandoffHasUncertainOutcome(bool duringHandoff)
    {
        var calls = 0;
        var backend = new MacExcelBackend(async (start, input, cancellationToken) =>
        {
            calls++;
            if (calls == (duringHandoff ? 2 : 3))
            {
                await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
            }
            return new MacProcessResult(0, """{"success":true}""", "");
        });

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = "/tmp/open-timeout.xlsx" }, TimeSpan.FromMilliseconds(100)));

        Assert.Equal(duringHandoff ? 2 : 3, calls);
        Assert.Equal("RecoveryRequired", error.ErrorCategory);
    }

    [Theory]
    [InlineData("open")]
    [InlineData("create")]
    public async Task SharedService_PreservesRecoveryCategoryAndFileAfterUnconfirmedOpen(string action)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-service-open-recovery-");
        var path = Path.Combine(directory.FullName, "workbook.xlsx");
        if (action == "open") MacWorkbookPackage.Create(path, macroEnabled: false);
        var calls = 0;
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            calls++;
            return Task.FromResult(new MacProcessResult(0, calls == 3
                ? """{"success":false,"errorCategory":"ComInterop","errorMessage":"Attachment not confirmed."}"""
                : """{"success":true}""", ""));
        });
        try
        {
            using var service = new ExcelMcpService(
                backend,
                (_, _) => throw new InvalidOperationException("Helper must not be probed."),
                (_, _, _, _) => throw new InvalidOperationException("Helper must not be invoked."));
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = $"session.{action}",
                Args = JsonSerializer.Serialize(new { filePath = path, timeoutSeconds = 10 })
            });

            Assert.False(response.Success);
            Assert.Equal("RecoveryRequired", response.ErrorCategory);
            Assert.Equal($"session.{action}", response.Command);
            Assert.Null(response.SessionId);
            Assert.Null(response.Result);
            Assert.Equal(0, service.SessionCount);
            Assert.Contains("Do not retry automatically", response.ErrorMessage, StringComparison.Ordinal);
            Assert.True(File.Exists(path));
            service.Dispose();
            Assert.Equal(3, calls);
        }
        finally
        {
            File.Delete(path);
            directory.Delete();
        }
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
