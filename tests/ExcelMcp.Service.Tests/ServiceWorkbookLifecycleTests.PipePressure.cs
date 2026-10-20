using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Fact]
    public async Task Pipe_BlockedSessionPreservesOtherWorkbookAndBoundedStatus()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var blocked = await CreateSessionAsync(service, Path.Join(directory, "blocked.xlsx"));
            sessions[blocked] = 0;
            var healthy = await CreateSessionAsync(service, Path.Join(directory, "healthy.xlsx"));
            sessions[healthy] = 0;
            await WriteMarkerAsync(service, blocked, "Original unsaved edits");
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(blocked));
            using var release = new ManualResetEventSlim();
            var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            batch.BeforeRefreshStateReadHookForTests = () =>
            {
                entered.TrySetResult();
                Assert.True(release.Wait(TimeSpan.FromSeconds(90)));
            };
            var pipeName = $"excelmcp-real-pressure-{Guid.NewGuid():N}";
            var host = service.RunAsync(pipeName);
            using var originalClient = new ServiceClient(pipeName);
            var original = originalClient.SendAsync(new ServiceRequest { Command = "session.save", SessionId = blocked });
            var check = await Record.ExceptionAsync(async () =>
            {
                await entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
                var repeated = Enumerable.Range(0, 20).Select(async _ =>
                {
                    using var client = new ServiceClient(pipeName);
                    var response = await client.SendAsync(new ServiceRequest
                    {
                        Command = "session.save",
                        SessionId = blocked
                    }, TimeSpan.FromSeconds(2));
                    Assert.False(response.Success);
                    Assert.Equal("Busy", response.ErrorCategory);
                });
                await Task.WhenAll(repeated);
                using var observer = new ServiceClient(pipeName);
                Assert.True((await observer.SendAsync(new ServiceRequest { Command = "service.ping" },
                    TimeSpan.FromSeconds(2))).Success);
                var list = await observer.SendAsync(new ServiceRequest
                {
                    Command = "session.list"
                }, TimeSpan.FromSeconds(2));
                RequireSuccess(list);
                using var status = JsonDocument.Parse(list.Result!);
                var retained = Assert.Single(status.RootElement.GetProperty("sessions").EnumerateArray(),
                    item => item.GetProperty("sessionId").GetString() == blocked);
                Assert.True(retained.GetProperty("operationPending").GetBoolean());
                Assert.False(retained.GetProperty("canClose").GetBoolean());
                var write = await observer.SendAsync(new ServiceRequest
                {
                    Command = "range.set-values",
                    SessionId = healthy,
                    Args = """{"sheetName":"Sheet1","rangeAddress":"A1","overwritePolicy":"allow","values":[["Independent workbook edits"]]}"""
                }, TimeSpan.FromSeconds(5));
                Assert.True(write.Success, write.ErrorMessage);
                Assert.Equal("Independent workbook edits", await ReadMarkerAsync(service, healthy));
                Assert.False(original.IsCompleted);
            });
            batch.BeforeRefreshStateReadHookForTests = null;
            release.Set();
            var completion = await Record.ExceptionAsync(async () =>
            {
                Assert.True((await original.WaitAsync(TimeSpan.FromSeconds(10))).Success);
                Assert.Equal("Original unsaved edits", await ReadMarkerAsync(service, blocked));
                Assert.Equal("Independent workbook edits", await ReadMarkerAsync(service, healthy));
                await CloseSessionAsync(service, blocked, save: false);
                sessions.TryRemove(blocked, out _);
                await CloseSessionAsync(service, healthy, save: false);
                sessions.TryRemove(healthy, out _);
            });
            var shutdown = await Record.ExceptionAsync(async () =>
            {
                service.RequestShutdown();
                await host.WaitAsync(TimeSpan.FromSeconds(10));
            });
            Assert.All(new[] { check, completion, shutdown }, failure => Assert.Null(failure));
        });
    }

}
