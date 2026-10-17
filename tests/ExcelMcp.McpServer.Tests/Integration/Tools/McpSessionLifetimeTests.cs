using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "SessionLifetime")]
[Trait("RequiresExcel", "true")]
public sealed class McpSessionLifetimeTests(ITestOutputHelper output)
{
    [Fact]
    public async Task CancelledStartup_DisposalRetainsOriginalExecutionAndUnsavedChanges()
    {
        var path = Path.Join(Path.GetTempPath(), $"mcp-pending-startup-{Guid.NewGuid():N}.xlsx");
        using var cancellation = new CancellationTokenSource();
        using var release = new ManualResetEventSlim();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var backend = new ObservedBackend { ObserveCreation = true };
        var bridge = new ServiceBridge.ServiceBridge(() => backend);
        ExcelBatch.AfterWorkbookOpenHookForTests = (_, workbook) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                sheets = ((Excel.Workbook)workbook).Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                cell = sheet.Range["A1"];
                cell.Value2 = "Retained during cancelled startup";
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
            entered.SetResult();
            Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
        };
        ExcelProcessIdentity? identity = null;
        var failure = await Record.ExceptionAsync(async () =>
        {
            var creation = bridge.SendAsync("session.create", args: new { filePath = path },
                cancellationToken: cancellation.Token);
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(20));
            var pending = Assert.Single(backend.SessionManager.GetActiveSessions());
            var batch = Assert.IsType<ExcelBatch>(backend.SessionManager.GetSession(pending.SessionId));
            identity = Assert.Single(SessionManager.GetTrackedExcelProcesses(),
                candidate => candidate.ProcessId == batch.ExcelProcessId);
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => creation);
            var refusal = Assert.Throws<AggregateException>(bridge.Dispose);
            Assert.Contains("pending", refusal.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Same(batch, backend.SessionManager.GetSession(pending.SessionId));
            Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity.Value));
            Assert.Equal(path, Assert.Single(backend.SessionManager.GetActiveSessions()).FilePath);
            ExcelBatch.AfterWorkbookOpenHookForTests = null;
            release.Set();
            var completed = await backend.Finished.Task.WaitAsync(TimeSpan.FromSeconds(20));
            Assert.True(completed.Success, completed.ErrorMessage);
            Assert.Equal(pending.SessionId, await backend.Revealed.Task.WaitAsync(TimeSpan.FromSeconds(10)));
            Assert.True(SpinWait.SpinUntil(
                () => batch.GetRefreshState() == WorkbookRefreshState.Ready,
                TimeSpan.FromSeconds(10)));
            bridge.Dispose();
            Assert.Equal(0, backend.SessionManager.ActiveSessionCount);
            Assert.True(OwnedProcessGuard.TryConfirmExited(identity.Value));

            using var reader = new ExcelMcpService();
            var opened = await reader.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path }, ServiceProtocol.JsonOptions)
            });
            Assert.True(opened.Success, opened.ErrorMessage);
            using var session = JsonDocument.Parse(opened.Result!);
            var values = await reader.ProcessAsync(new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = session.RootElement.GetProperty("sessionId").GetString(),
                Args = """{"sheetName":"Sheet1","rangeAddress":"A1"}"""
            });
            Assert.True(values.Success, values.ErrorMessage);
            using var contents = JsonDocument.Parse(values.Result!);
            Assert.Equal("Retained during cancelled startup",
                contents.RootElement.GetProperty("values")[0][0].GetString());
        });
        release.Set();
        ExcelBatch.AfterWorkbookOpenHookForTests = null;
        var cleanup = await Record.ExceptionAsync(async () =>
        {
            await backend.Finished.Task.WaitAsync(TimeSpan.FromSeconds(30));
            bridge.Dispose();
            if (identity is { } owned)
                Assert.True(OwnedProcessGuard.TryConfirmExited(owned));
            File.Delete(path);
        });
        var failures = new[] { failure, cleanup }.OfType<Exception>().ToArray();
        if (failures.Length > 0)
            throw new AggregateException("Cancelled MCP startup or retained teardown failed.", failures);
    }

    [Fact]
    public async Task CancelledCreation_RetainsLateWorkbookWithoutClosingOtherSession()
    {
        var firstPath = Path.Combine(Path.GetTempPath(), $"mcp-survivor-{Guid.NewGuid():N}.xlsx");
        var secondPath = Path.Combine(Path.GetTempPath(), $"mcp-cancelled-{Guid.NewGuid():N}.xlsx");
        using var shutdown = new CancellationTokenSource();
        using var cancellation = new CancellationTokenSource();
        using var release = new ManualResetEventSlim();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var input = new Pipe();
        var outputPipe = new Pipe();
        var backend = new ObservedBackend();
        var host = await ProgramTransportTestHost.StartAsync(
            input, outputPipe, shutdown.Token, "CreationCancellationClient", () => backend);
        try
        {
            var first = await CallAsync(host.Client, "file", new() { ["action"] = "create", ["file_path"] = firstPath });
            var firstId = first.GetProperty("workbook_session_id").GetString()!;
            await CallAsync(host.Client, "range", new()
            {
                ["action"] = "set-values",
                ["workbook_session_id"] = firstId,
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1",
                ["values"] = new object[][] { ["keep"] }
            });
            backend.ObserveCreation = true;
            ExcelBatch.AfterWorkbookOpenHookForTests = (_, _) =>
            {
                entered.SetResult();
                Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
            };
            var request = new JsonRpcRequest
            {
                Id = new RequestId("cancelled-creation"),
                Method = RequestMethods.ToolsCall,
                Params = JsonSerializer.SerializeToNode(new
                {
                    name = "file",
                    arguments = new { action = "create", file_path = secondPath }
                })
            };
            var creation = host.Client.SendRequestAsync(request, cancellation.Token);
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(30));
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => creation);
            // Cancelling a client's local wait does not prove it sent the protocol notification.
            await host.Client.SendMessageAsync(new JsonRpcNotification
            {
                Method = NotificationMethods.CancelledNotification,
                Params = JsonSerializer.SerializeToNode(new { requestId = "cancelled-creation" })
            });
            await host.Client.ListToolsAsync();
            release.Set();
            var completed = await backend.Finished.Task.WaitAsync(TimeSpan.FromSeconds(150));
            Assert.True(completed.Success, completed.ErrorMessage);
            var retainedId = await backend.Revealed.Task.WaitAsync(TimeSpan.FromSeconds(30));
            Assert.NotEqual(firstId, retainedId);

            var sessions = await CallAsync(host.Client, "file_read", new() { ["action"] = "list" });
            var entries = sessions.GetProperty("sessions").EnumerateArray().ToArray();
            Assert.Equal(2, entries.Length);
            Assert.Contains(entries, session => session.GetProperty("workbook_session_id").GetString() == firstId);
            Assert.True(Assert.Single(entries, session =>
                session.GetProperty("workbook_session_id").GetString() == retainedId)
                .GetProperty("isExcelVisible").GetBoolean());
            var values = await CallAsync(host.Client, "range_read", new()
            {
                ["action"] = "get-values",
                ["workbook_session_id"] = firstId,
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1"
            });
            Assert.Equal("keep", values.GetProperty("values")[0][0].GetString());
        }
        finally
        {
            release.Set();
            ExcelBatch.AfterWorkbookOpenHookForTests = null;
            await ProgramTransportTestHost.StopAsync(host.Client, input, outputPipe, host.ServerTask, output, shutdown);
            File.Delete(firstPath);
            File.Delete(secondPath);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Shutdown_AutoSavesUnlessExplicitCloseDiscardsChanges(bool discard)
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcp-persistence-{Guid.NewGuid():N}.xlsx");
        try
        {
            await CheckPersistenceAsync(path, discard);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private async Task CheckPersistenceAsync(string path, bool discard)
    {
        using var shutdown = new CancellationTokenSource();
        var input = new Pipe();
        var outputPipe = new Pipe();
        var host = await ProgramTransportTestHost.StartAsync(input, outputPipe, shutdown.Token, "PersistenceWriter");
        try
        {
            var created = await CallAsync(host.Client, "file", new() { ["action"] = "create", ["file_path"] = path });
            var id = created.GetProperty("workbook_session_id").GetString()!;
            await CallAsync(host.Client, "range", new()
            {
                ["action"] = "set-values",
                ["workbook_session_id"] = id,
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1",
                ["values"] = new object[][] { ["saved-on-shutdown"] }
            });
            if (discard)
                await CallAsync(host.Client, "file", new() { ["action"] = "close", ["workbook_session_id"] = id, ["save"] = false });
        }
        finally
        {
            await ProgramTransportTestHost.StopAsync(host.Client, input, outputPipe, host.ServerTask, output, shutdown);
        }

        using var readerShutdown = new CancellationTokenSource();
        var readerInput = new Pipe();
        var readerOutput = new Pipe();
        var reader = await ProgramTransportTestHost.StartAsync(readerInput, readerOutput, readerShutdown.Token, "PersistenceReader");
        try
        {
            var opened = await CallAsync(reader.Client, "file", new() { ["action"] = "open", ["file_path"] = path });
            var values = await CallAsync(reader.Client, "range_read", new()
            {
                ["action"] = "get-values",
                ["workbook_session_id"] = opened.GetProperty("workbook_session_id").GetString(),
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1"
            });
            Assert.Equal(discard ? null : "saved-on-shutdown", values.GetProperty("values")[0][0].GetString());
        }
        finally
        {
            await ProgramTransportTestHost.StopAsync(reader.Client, readerInput, readerOutput, reader.ServerTask, output, readerShutdown);
        }
    }

    private static async Task<JsonElement> CallAsync(McpClient client, string tool, Dictionary<string, object?> args)
    {
        var result = await client.CallToolAsync(tool, args);
        Assert.NotEqual(true, result.IsError);
        Assert.NotNull(result.StructuredContent);
        Assert.True(result.StructuredContent.Value.GetProperty("success").GetBoolean());
        return result.StructuredContent.Value;
    }

    private sealed class ObservedBackend : IServiceBridgeBackend
    {
        private readonly ExcelMcpService _service = new();
        internal SessionManager SessionManager => _service.SessionManager;
        internal bool ObserveCreation { get; set; }
        internal TaskCompletionSource<ServiceResponse> Finished { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource<string> Revealed { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public async Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            var observe = ObserveCreation && request.Command == "session.create";
            var response = await _service.ProcessAsync(request);
            if (observe)
                Finished.TrySetResult(response);
            return response;
        }

        public void RevealSession(string sessionId)
        {
            _service.SessionManager.RevealForRecovery(sessionId, waitExpired: true);
            Revealed.TrySetResult(sessionId);
        }

        public void Dispose() => _service.Dispose();
    }
}
