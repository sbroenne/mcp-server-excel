using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "Session")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ServiceWorkbookDiscoveryTests
{
    [Fact]
    public async Task Discover_RetriesFailedConnectionAfterOtherHostReleasesLease()
    {
        using var process = System.Diagnostics.Process.GetCurrentProcess();
        var identity = new ExcelProcessIdentity(
            process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc());
        using var service = new ExcelMcpService();
        using var releaseLease = new ManualResetEventSlim();
        var leaseReady = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var leaseTask = Task.Run(() =>
        {
            using var lease = new ExcelApplicationLease(identity);
            leaseReady.SetResult();
            Assert.True(releaseLease.Wait(TimeSpan.FromSeconds(30)));
        });
        RunningExcelWindows.FindHookForTests = () => new RunningExcelWindows.Discovery(
            [new RunningExcelWindows.ApplicationWindow(nint.Zero, identity)], []);
        try
        {
            await leaseReady.Task.WaitAsync(TimeSpan.FromSeconds(10));
            var blocked = await service.SessionManager.DiscoverOpenWorkbooksAsync();
            var error = Assert.Single(blocked.Errors);
            Assert.False(blocked.Complete);
            Assert.Equal(identity.ProcessId, error.ProcessId);
            Assert.Contains("another active ExcelMcp host", error.Message, StringComparison.Ordinal);

            releaseLease.Set();
            await leaseTask;
            var retried = await service.SessionManager.DiscoverOpenWorkbooksAsync();
            var retriedError = Assert.Single(retried.Errors);
            Assert.False(retried.Complete);
            Assert.Equal(identity.ProcessId, retriedError.ProcessId);
            Assert.Contains("window no longer belongs", retriedError.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("another active ExcelMcp host", retriedError.Message, StringComparison.Ordinal);
        }
        finally
        {
            releaseLease.Set();
            RunningExcelWindows.FindHookForTests = null;
            await leaseTask;
        }
    }

    [Fact]
    public async Task Discover_ReturnsPartialNativeErrorsAndMarksResultIncomplete()
    {
        using var service = new ExcelMcpService();
        RunningExcelWindows.FindHookForTests = () => new RunningExcelWindows.Discovery(
            [],
            [new WorkbookDiscoveryError(1234, "Window callback failed.")]);

        try
        {
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.discover"
            });

            Assert.True(response.Success);
            using var result = JsonDocument.Parse(response.Result!);
            var root = result.RootElement;
            Assert.True(root.GetProperty("success").GetBoolean());
            Assert.False(root.GetProperty("complete").GetBoolean());
            Assert.Empty(root.GetProperty("workbooks").EnumerateArray());
            var error = Assert.Single(root.GetProperty("errors").EnumerateArray());
            Assert.Equal(1234, error.GetProperty("processId").GetInt32());
            Assert.Equal("Window callback failed.", error.GetProperty("message").GetString());
        }
        finally
        {
            RunningExcelWindows.FindHookForTests = null;
        }
    }
}
