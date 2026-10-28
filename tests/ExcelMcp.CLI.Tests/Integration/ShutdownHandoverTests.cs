using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Slow")]
public sealed class ShutdownHandoverTests(Xunit.Abstractions.ITestOutputHelper output)
{
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(false, false, true)]
    public async Task ServiceStop_PreservesWorkbookAcrossDaemonExitAndRestart(
        bool retainedOwnership, bool blockedWorker, bool crashedController)
    {
        var path = Path.Join(Path.GetTempPath(), $"shutdown-handover-{Guid.NewGuid():N}.xlsx");
        var environment = new Dictionary<string, string>
        {
            ["EXCELMCP_CLI_PIPE"] = $"excelmcp-handover-{Guid.NewGuid():N}"
        };
        using var cleanupService = new ExcelMcpService();
        ExcelProcessIdentity? ownedIdentity = null;
        Process? ownedProcess = null;
        string? sessionId = null;
        string? cleanupSessionId = null;
        Exception? failure = null;
        try
        {
            var created = await RunAsync(["session", "create", path]);
            sessionId = created.GetProperty("sessionId").GetString()!;
            var initialDiscovery = await RunAsync(["session", "discover"]);
            var ownedBook = Assert.Single(initialDiscovery.GetProperty("workbooks").EnumerateArray(),
                book => book.GetProperty("location").GetString() == path);
            // The test created this process in its child daemon; retain ownership for orderly cleanup.
            ownedIdentity = SessionManager.TrackExcelProcessIdentity(ownedBook.GetProperty("processId").GetInt32());
            Assert.NotNull(ownedIdentity);
            ownedProcess = Process.GetProcessById(ownedIdentity.Value.ProcessId);
            Assert.Equal(ownedIdentity.Value.StartedAtUtcFileTime, ownedProcess.StartTime.ToUniversalTime().ToFileTimeUtc());
            _ = ownedProcess.Handle;
            await RunAsync(["range", "set-values", "--session", sessionId, "--sheet-name", "Sheet1",
                "--range-address", "A1", "--values", """[["left open"]]"""]);
            var pipeName = environment["EXCELMCP_CLI_PIPE"];
            var recorded = DaemonProcessTracker.ReadProcessSnapshot(pipeName).Snapshot!;
            var originalWorker = Assert.Single(new WorkerRegistry().Read(), record =>
                record.Controller.ProcessId == recorded.DaemonProcess.ProcessId && record.Excel == ownedIdentity);
            if (crashedController)
            {
                ownedProcess.Refresh();
                Assert.Equal(IntPtr.Zero, ownedProcess.MainWindowHandle);
                using var daemon = Process.GetProcessById(recorded.DaemonProcess.ProcessId);
                Assert.Equal(recorded.DaemonProcess.StartedAtUtcFileTime,
                    daemon.StartTime.ToUniversalTime().ToFileTimeUtc());
                _ = daemon.Handle;
                daemon.Kill();
                await daemon.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
                sessionId = null;
                var revealWait = Stopwatch.StartNew();
                while (revealWait.Elapsed < TimeSpan.FromSeconds(8))
                {
                    ownedProcess.Refresh();
                    if (ownedProcess.MainWindowHandle != IntPtr.Zero)
                        break;
                    await Task.Delay(50);
                }
                Assert.NotEqual(IntPtr.Zero, ownedProcess.MainWindowHandle);
            }
            if (blockedWorker)
            {
                using var worker = Process.GetProcessById(originalWorker.Worker.ProcessId);
                Assert.Equal(originalWorker.Worker, WorkerProcess.From(worker));
                Assert.Equal(0, NtSuspendProcess(worker.Handle));
                try
                {
                    var controlWait = Stopwatch.StartNew();
                    var responsive = await RunAsync(["service", "status"]);
                    Assert.True(responsive.GetProperty("running").GetBoolean());
                    Assert.Equal(1, responsive.GetProperty("sessionCount").GetInt32());
                    Assert.True(controlWait.Elapsed < TimeSpan.FromSeconds(5));
                    var list = await RunAsync(["session", "list"]);
                    Assert.Equal(sessionId, Assert.Single(list.GetProperty("sessions").EnumerateArray())
                        .GetProperty("sessionId").GetString());
                    var stopping = Stopwatch.StartNew();
                    var stop = await RunAsync(["service", "stop"]);
                    sessionId = null;
                    Assert.True(stopping.Elapsed < TimeSpan.FromSeconds(12));
                    Assert.Contains(originalWorker.Worker.ProcessId.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        stop.GetProperty("shutdownWarning").GetString());
                    Assert.False(worker.HasExited);
                    Assert.False(ownedProcess.HasExited);
                }
                finally
                {
                    Assert.Equal(0, NtResumeProcess(worker.Handle));
                }
            }
            else
            {
                await RunAsync(["service", "stop"]);
                sessionId = null;
            }
            if (retainedOwnership)
            {
                // Recreate the durable state of an earlier owner without killing Excel's automation host.
                DaemonProcessTracker.Clear(pipeName);
                var earlierService = DaemonProcessTracker.RegisterProcess(pipeName,
                    recorded.DaemonProcess.ProcessId, recorded.DaemonProcess.StartedAtUtcFileTime);
                DaemonProcessTracker.RecordExcelProcesses(pipeName, earlierService, recorded.ExcelProcesses.ToArray());
            }
            var stopped = await RunAsync(["service", "stop"]);
            var retained = retainedOwnership || blockedWorker || crashedController;
            Assert.Equal(retained ? 1 : 0, stopped.GetProperty("retainedExcelProcessCount").GetInt32());
            Assert.False(ownedProcess.HasExited);
            sessionId = null;

            await RunAsync(["service", "start"]);
            var status = await RunAsync(["service", "status"]);
            Assert.Equal(retained ? 1 : 0, status.GetProperty("retainedExcelProcessCount").GetInt32());
            if (retained)
            {
                var remaining = Assert.Single(status.GetProperty("retainedExcelProcesses").EnumerateArray());
                Assert.Equal(ownedIdentity.Value.ProcessId, remaining.GetProperty("processId").GetInt32());
                Assert.Equal("unknown", remaining.GetProperty("workbookState").GetString());
                Assert.Equal(DateTime.FromFileTimeUtc(ownedIdentity.Value.StartedAtUtcFileTime),
                    remaining.GetProperty("startTime").GetDateTime());
            }
            var empty = await RunAsync(["session", "list"]);
            Assert.Empty(empty.GetProperty("sessions").EnumerateArray());
            var discovery = await RunAsync(["session", "discover"]);
            Assert.True(discovery.GetProperty("complete").GetBoolean());
            Assert.Empty(discovery.GetProperty("errors").EnumerateArray());
            var book = Assert.Single(discovery.GetProperty("workbooks").EnumerateArray(),
                book => book.GetProperty("location").GetString() == path);
            Assert.False(book.GetProperty("saved").GetBoolean());
            Assert.Equal(ownedIdentity.Value.ProcessId, book.GetProperty("processId").GetInt32());
            var attached = await RunAsync(["session", "attach", "--workbook-id", book.GetProperty("workbookId").GetString()!]);
            sessionId = attached.GetProperty("sessionId").GetString()!;
            if (blockedWorker || crashedController)
            {
                var reclaimed = Assert.Single(new WorkerRegistry().Read(), record =>
                    record.Excel == ownedIdentity && record.Worker.IsAlive());
                Assert.Equal(originalWorker.Worker, reclaimed.Worker);
                Assert.NotEqual(originalWorker.Controller, reclaimed.Controller);
            }
            var info = await RunAsync(["workbook", "get-info", "--session", sessionId]);
            Assert.False(info.GetProperty("saved").GetBoolean());
            var values = await RunAsync(["range", "get-values", "--session", sessionId,
                "--sheet-name", "Sheet1", "--range-address", "A1"]);
            Assert.Equal("left open", values.GetProperty("values")[0][0].GetString());
            await RunAsync(["service", "stop"]);
            sessionId = null;
            var cleanupConnection = await cleanupService.ProcessAsync(new ServiceRequest { Command = "session.discover" });
            Assert.True(cleanupConnection.Success, cleanupConnection.ErrorMessage);
            using (var cleanupDiscovery = JsonDocument.Parse(cleanupConnection.Result!))
            {
                var candidate = Assert.Single(cleanupDiscovery.RootElement.GetProperty("workbooks").EnumerateArray(),
                    candidate => candidate.GetProperty("location").GetString() == path
                        && candidate.GetProperty("processId").GetInt32() == ownedIdentity.Value.ProcessId);
                var cleanupAttachment = await cleanupService.ProcessAsync(new ServiceRequest
                {
                    Command = "session.attach",
                    Args = JsonSerializer.Serialize(new { workbookId = candidate.GetProperty("workbookId").GetString() })
                });
                Assert.True(cleanupAttachment.Success, cleanupAttachment.ErrorMessage);
                cleanupSessionId = cleanupAttachment.SessionId!;
            }
            await CloseCleanupSessionAsync();
            cleanupService.Dispose();
            Assert.True(ownedProcess.WaitForExit(10000), $"The test-owned Excel process is still running: {ownedIdentity}.");
            var reopened = await RunAsync(["session", "open", path]);
            sessionId = reopened.GetProperty("sessionId").GetString()!;
            values = await RunAsync(["range", "get-values", "--session", sessionId,
                "--sheet-name", "Sheet1", "--range-address", "A1"]);
            Assert.Equal(JsonValueKind.Null, values.GetProperty("values")[0][0].ValueKind);
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            try
            {
                if (sessionId is not null)
                    await RunAsync(["session", "close", "--session", sessionId, "--save", "false"]);
                var discovery = await RunAsync(["session", "discover"]);
                foreach (var book in discovery.GetProperty("workbooks").EnumerateArray()
                    .Where(book => book.GetProperty("location").GetString() == path))
                {
                    var attached = await RunAsync(["session", "attach", "--workbook-id", book.GetProperty("workbookId").GetString()!]);
                    await RunAsync(["session", "close", "--session", attached.GetProperty("sessionId").GetString()!, "--save", "false"]);
                }
                await RunAsync(["service", "stop"]);
                await CloseCleanupSessionAsync();
                cleanupService.Dispose();
                if (ownedProcess is not null)
                    Assert.True(ownedProcess.WaitForExit(10000),
                        $"The test-owned Excel process is still running: {ownedIdentity}.");
                File.Delete(path);
            }
            catch (Exception cleanup)
            {
                failure = failure is null ? cleanup : new AggregateException(failure, cleanup);
            }
            finally
            {
                ownedProcess?.Dispose();
            }
        }
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();

        async Task CloseCleanupSessionAsync()
        {
            if (cleanupSessionId is null)
                return;
            var closed = await cleanupService.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = cleanupSessionId,
                Args = """{"save":false}"""
            });
            Assert.True(closed.Success, closed.ErrorMessage);
            cleanupSessionId = null;
        }

        async Task<JsonElement> RunAsync(string[] arguments)
        {
            output.WriteLine($"Starting: {string.Join(' ', arguments)}");
            var result = await CliProcessHelper.RunAsync(arguments, timeoutMs: 120000, environmentVariables: environment);
            output.WriteLine($"Finished: {result.Stdout}");
            Assert.True(result.ExitCode == 0, $"{result.Stdout}\n{result.Stderr}");
            using var document = JsonDocument.Parse(result.Stdout);
            Assert.True(document.RootElement.GetProperty("success").GetBoolean(), result.Stdout);
            return document.RootElement.TryGetProperty("data", out var data)
                ? data.Clone()
                : document.RootElement.Clone();
        }
    }

    [DllImport("ntdll.dll")]
    private static extern int NtSuspendProcess(IntPtr processHandle);

    [DllImport("ntdll.dll")]
    private static extern int NtResumeProcess(IntPtr processHandle);
}
