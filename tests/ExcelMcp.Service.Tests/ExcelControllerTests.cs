using System.Text.Json;
using Sbroenne.ExcelMcp.Service.Rpc;
using Sbroenne.ExcelMcp.Service.Workers;
using StreamJsonRpc;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("Sequential")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("Speed", "Medium")]
[Trait("RequiresExcel", "false")]
public sealed class ExcelControllerTests
{
    private static ProductProcessLauncher Launcher => new(Path.GetFullPath(Path.Combine(
        AppContext.BaseDirectory, @"..\..\..\..\..\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe")), []);

    [Theory(Timeout = 30000)]
    [InlineData("session.save", null, "InvalidInput")]
    [InlineData("session.save", "missing", "SessionNotFound")]
    [InlineData("service.restart", null, "InvalidInput")]
    [InlineData("unknown.action", null, "InvalidInput")]
    public async Task Controller_PreservesExistingValidation(string command, string? session, string expectedCategory)
    {
        var request = new ServiceRequest { Command = command, SessionId = session };
        using var expectedService = new ExcelMcpService();
        var expected = await expectedService.ProcessAsync(request);
        using var controller = new ExcelController(launcher: Launcher);
        var actual = await controller.ProcessAsync(request);
        Assert.False(actual.Success);
        Assert.Equal(expectedCategory, actual.ErrorCategory);
        Assert.Equal(expected.ErrorCategory, actual.ErrorCategory);
        Assert.Equal(expected.ErrorMessage, actual.ErrorMessage);
        Assert.Equal(expected.Command, actual.Command);
        Assert.Equal(expected.SessionId, actual.SessionId);
        Assert.Empty(controller.ExcelProcesses);
    }

    [Fact(Timeout = 30000)]
    public async Task Controller_InvalidCloseDoesNotBreakLaterSessionlessCommands()
    {
        using var controller = new ExcelController(launcher: Launcher);
        var closed = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = "missing"
        });
        Assert.False(closed.Success);
        Assert.Equal("SessionNotFound", closed.ErrorCategory);
        var echoed = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "diag.echo",
            Args = """{"message":"after-invalid-close"}"""
        });
        Assert.True(echoed.Success, echoed.ErrorMessage);
        using var result = JsonDocument.Parse(echoed.Result!);
        Assert.Equal("after-invalid-close", result.RootElement.GetProperty("message").GetString());
    }

    [Fact(Timeout = 30000)]
    public async Task Controller_CancelledStartupReleasesItsEmptyWorkerAndAllowsRetry()
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelmcp-missing-{Guid.NewGuid():N}.xlsx");
        var registry = new WorkerRegistry();
        var existing = registry.Read().Select(record => record.PipeName).ToHashSet();
        using var controller = new ExcelController(launcher: Launcher);
        using var cancellation = new CancellationTokenSource();
        var opening = controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = path }),
            CallerCancellation = cancellation.Token
        });
        Assert.False(opening.IsCompleted);
        cancellation.Cancel();
        var cancelled = await opening;
        Assert.False(cancelled.Success);
        Assert.Equal("Cancelled", cancelled.ErrorCategory);
        Assert.Empty(controller.Sessions);
        Assert.False(File.Exists(path));
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(8));
        while (registry.Read().Any(record => record.Controller == WorkerProcess.Current
            && !existing.Contains(record.PipeName) && record.Worker.IsAlive()) && !deadline.IsCancellationRequested)
            await Task.Delay(50);
        Assert.DoesNotContain(registry.Read(), record =>
            record.Controller == WorkerProcess.Current && !existing.Contains(record.PipeName) && record.Worker.IsAlive());
        var retried = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = path })
        });
        Assert.False(retried.Success);
        Assert.Equal("FileNotFoundException", retried.ExceptionType);
        Assert.Empty(controller.Sessions);
        Assert.False(File.Exists(path));
    }

    [Fact(Timeout = 30000)]
    public async Task ControlEndpoint_AnswersWhenAllWorkConnectionsAreOccupied()
    {
        var pipeName = $"excelmcp-controller-test-{Guid.NewGuid():N}";
        using var controller = new ExcelController(launcher: Launcher);
        var run = controller.RunAsync(pipeName);
        var connections = new List<IDisposable>();
        try
        {
            for (var index = 0; index < DaemonHost.MaximumConnections; index++)
            {
                var pipe = ServiceSecurity.CreateClient(pipeName);
                connections.Add(pipe);
                await pipe.ConnectAsync(2000);
                var proxy = JsonRpc.Attach<IExcelDaemonRpc>(pipe);
                connections.Add((IDisposable)proxy);
                Assert.True((await proxy.ProcessCommandAsync(
                    new ServiceRequest { Command = "service.ping" }).WaitAsync(TimeSpan.FromSeconds(2))).Success);
            }
            using var control = new ServiceClient(pipeName, useControlEndpoint: true);
            foreach (var command in new[] { "service.ping", "service.status", "session.list" })
            {
                var response = await control.SendAsync(new ServiceRequest { Command = command }, TimeSpan.FromSeconds(2));
                Assert.True(response.Success, response.ErrorMessage);
                Assert.Null(response.Command);
                if (command == "service.status")
                {
                    using var status = JsonDocument.Parse(response.Result!);
                    Assert.True(status.RootElement.GetProperty("running").GetBoolean());
                    Assert.Equal(Environment.ProcessId, status.RootElement.GetProperty("processId").GetInt32());
                    Assert.Equal(0, status.RootElement.GetProperty("sessionCount").GetInt32());
                }
                else if (command == "session.list")
                {
                    using var list = JsonDocument.Parse(response.Result!);
                    Assert.Equal(0, list.RootElement.GetProperty("count").GetInt32());
                    Assert.Empty(list.RootElement.GetProperty("sessions").EnumerateArray());
                }
            }
            var stopped = await control.SendAsync(new ServiceRequest { Command = "service.shutdown" }, TimeSpan.FromSeconds(2));
            Assert.True(stopped.Success, stopped.ErrorMessage);
            await run.WaitAsync(TimeSpan.FromSeconds(4));
        }
        finally
        {
            foreach (var connection in connections)
                connection.Dispose();
            await controller.RequestShutdownAsync();
            await run.WaitAsync(TimeSpan.FromSeconds(4));
        }
    }
}

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresExcel", "true")]
public sealed class ExcelControllerWorkbookTests
{
    [Fact(Timeout = 90000)]
    public async Task Controller_WorkerPreservesValuesAndReleasesClosedSessions()
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelmcp-controller-{Guid.NewGuid():N}.xlsx");
        var savedAsPath = Path.ChangeExtension(path, ".saved-as.xlsx");
        var launcher = new ProductProcessLauncher(Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, @"..\..\..\..\..\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe")), []);
        using var controller = new ExcelController(launcher: launcher);
        string? sessionId = null;
        try
        {
            var created = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.create",
                Args = JsonSerializer.Serialize(new { filePath = path })
            });
            Assert.True(created.Success, created.ErrorMessage);
            sessionId = Assert.IsType<string>(created.SessionId);
            Assert.Single(controller.Sessions);
            Assert.Single(controller.ExcelProcesses);
            var sheet = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "sheet.create",
                SessionId = sessionId,
                Args = """{"sheetName":"WorkerMarker"}"""
            });
            Assert.True(sheet.Success, sheet.ErrorMessage);
            var written = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "range.set-values",
                SessionId = sessionId,
                Args = """{"sheetName":"WorkerMarker","rangeAddress":"A1","values":[["worker-owned-marker"]]}"""
            });
            Assert.True(written.Success, written.ErrorMessage);
            var read = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = sessionId,
                Args = """{"sheetName":"WorkerMarker","rangeAddress":"A1"}"""
            });
            Assert.True(read.Success, read.ErrorMessage);
            using var result = JsonDocument.Parse(read.Result!);
            Assert.Equal("worker-owned-marker",
                result.RootElement.GetProperty("values")[0][0].GetString());
            var savedAs = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "workbook.save-as",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new { targetPath = savedAsPath })
            });
            Assert.True(savedAs.Success, savedAs.ErrorMessage);
            Assert.Equal(savedAsPath, Assert.Single(controller.Sessions).FilePath);
            var conflicting = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = savedAsPath })
            });
            Assert.False(conflicting.Success);
            Assert.Equal("Conflict", conflicting.ErrorCategory);
            var original = await controller.ProcessAsync(new ServiceRequest
            {
                Command = File.Exists(path) ? "session.open" : "session.create",
                Args = JsonSerializer.Serialize(new { filePath = path })
            });
            Assert.True(original.Success, original.ErrorMessage);
            var originalClosed = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = original.SessionId
            });
            Assert.True(originalClosed.Success, originalClosed.ErrorMessage);
            var afterSaveAs = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = sessionId,
                Args = """{"sheetName":"WorkerMarker","rangeAddress":"A1"}"""
            });
            Assert.True(afterSaveAs.Success, afterSaveAs.ErrorMessage);
            using var afterSaveAsValues = JsonDocument.Parse(afterSaveAs.Result!);
            Assert.Equal("worker-owned-marker", afterSaveAsValues.RootElement.GetProperty("values")[0][0].GetString());
            var traces = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "powerquerytrace.list",
                Args = JsonSerializer.Serialize(new { workbookSessionId = sessionId })
            });
            Assert.True(traces.Success, traces.ErrorMessage);
            using var traceResult = JsonDocument.Parse(traces.Result!);
            Assert.Equal(Assert.Single(controller.ExcelProcesses).ProcessId,
                traceResult.RootElement.GetProperty("processId").GetInt32());
            var closed = await controller.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = sessionId });
            Assert.True(closed.Success, closed.ErrorMessage);
            Assert.Empty(controller.Sessions);
            var again = await controller.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = sessionId });
            Assert.True(again.Success, again.ErrorMessage);
            Assert.Contains("already closed", again.Result);
            sessionId = null;
            Assert.Null(await controller.RequestShutdownAsync());
        }
        finally
        {
            if (sessionId is not null)
            {
                var closed = await controller.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = sessionId });
                Assert.True(closed.Success, closed.ErrorMessage);
            }
            File.Delete(path);
            File.Delete(savedAsPath);
        }
    }
}
