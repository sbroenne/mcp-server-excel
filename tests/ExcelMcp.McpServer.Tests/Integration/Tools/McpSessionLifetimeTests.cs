using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Xunit.Abstractions;

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
    public async Task CancelledCreation_ReclaimsLateWorkbookWithoutClosingOtherSession()
    {
        var firstPath = Path.Combine(Path.GetTempPath(), $"mcp-survivor-{Guid.NewGuid():N}.xlsx");
        var secondPath = Path.Combine(Path.GetTempPath(), $"mcp-cancelled-{Guid.NewGuid():N}.xlsx");
        using var shutdown = new CancellationTokenSource();
        using var cancellation = new CancellationTokenSource();
        var input = new Pipe();
        var outputPipe = new Pipe();
        var backend = new ObservedBackend();
        var host = await ProgramTransportTestHost.StartAsync(
            input, outputPipe, shutdown.Token, "CreationCancellationClient", () => backend);
        try
        {
            var first = await CallAsync(host.Client, "file", new() { ["action"] = "create", ["path"] = firstPath });
            var firstId = first.GetProperty("session_id").GetString()!;
            await CallAsync(host.Client, "range", new()
            {
                ["action"] = "set-values",
                ["session_id"] = firstId,
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1",
                ["values"] = new object[][] { ["keep"] }
            });
            backend.ObserveCreation = true;
            var request = new JsonRpcRequest
            {
                Id = new RequestId("cancelled-creation"),
                Method = RequestMethods.ToolsCall,
                Params = JsonSerializer.SerializeToNode(new
                {
                    name = "file",
                    arguments = new { action = "create", path = secondPath }
                })
            };
            var creation = host.Client.SendRequestAsync(request, cancellation.Token);
            await backend.Started.Task.WaitAsync(TimeSpan.FromSeconds(30));
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => creation);
            // Cancelling a client's local wait does not prove it sent the protocol notification.
            await host.Client.SendMessageAsync(new JsonRpcNotification
            {
                Method = NotificationMethods.CancelledNotification,
                Params = JsonSerializer.SerializeToNode(new { requestId = "cancelled-creation" })
            });
            await host.Client.ListToolsAsync();
            backend.ContinueCreation.TrySetResult();
            var completed = await backend.Finished.Task.WaitAsync(TimeSpan.FromSeconds(150));
            Assert.True(completed.Success, completed.ErrorMessage);
            var closedId = await backend.Closed.Task.WaitAsync(TimeSpan.FromSeconds(30));
            Assert.NotEqual(firstId, closedId);

            var sessions = await CallAsync(host.Client, "file", new() { ["action"] = "list" });
            Assert.Equal(firstId, Assert.Single(sessions.GetProperty("sessions").EnumerateArray())
                .GetProperty("session_id").GetString());
            var values = await CallAsync(host.Client, "range", new()
            {
                ["action"] = "get-values",
                ["session_id"] = firstId,
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1"
            });
            Assert.Equal("keep", values.GetProperty("values")[0][0].GetString());
        }
        finally
        {
            backend.ContinueCreation.TrySetResult();
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
            var created = await CallAsync(host.Client, "file", new() { ["action"] = "create", ["path"] = path });
            var id = created.GetProperty("session_id").GetString()!;
            await CallAsync(host.Client, "range", new()
            {
                ["action"] = "set-values",
                ["session_id"] = id,
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1",
                ["values"] = new object[][] { ["saved-on-shutdown"] }
            });
            if (discard)
                await CallAsync(host.Client, "file", new() { ["action"] = "close", ["session_id"] = id, ["save"] = false });
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
            var opened = await CallAsync(reader.Client, "file", new() { ["action"] = "open", ["path"] = path });
            var values = await CallAsync(reader.Client, "range", new()
            {
                ["action"] = "get-values",
                ["session_id"] = opened.GetProperty("session_id").GetString(),
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
        internal bool ObserveCreation { get; set; }
        internal TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource ContinueCreation { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource<ServiceResponse> Finished { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource<string> Closed { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public async Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            var observe = ObserveCreation && request.Command == "session.create";
            if (observe)
            {
                Started.TrySetResult();
                await ContinueCreation.Task.WaitAsync(TimeSpan.FromSeconds(30));
            }
            var response = await _service.ProcessAsync(request);
            if (observe)
                Finished.TrySetResult(response);
            return response;
        }

        public bool ForceCloseSession(string sessionId)
        {
            var closed = _service.SessionManager.CloseSession(sessionId, save: false, force: true);
            if (closed)
                Closed.TrySetResult(sessionId);
            return closed;
        }

        public void Dispose() => _service.Dispose();
    }
}
