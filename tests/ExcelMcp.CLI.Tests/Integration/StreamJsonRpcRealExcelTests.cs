using System.Text.Json;
using Nerdbank.Streams;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Rpc;
using StreamJsonRpc;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "StreamJsonRpc")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class StreamJsonRpcRealExcelTests : IClassFixture<TempDirectoryFixture>
{
    private readonly string _tempRoot;

    public StreamJsonRpcRealExcelTests(TempDirectoryFixture fixture)
    {
        _tempRoot = Path.Combine(fixture.TempDir, $"R-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempRoot);
    }

    [Fact]
    public async Task DroppedClientConnection_DoesNotDestroyHealthyReusableExcelSession()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Rpc",
            "reconnect",
            _tempRoot,
            ".xlsx");
        using var service = new ExcelMcpService(Path.Combine(_tempRoot, "safety-state"));

        var create = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = workbookPath, show = false }, ServiceProtocol.JsonOptions)
        });
        Assert.True(create.Success, create.ErrorMessage);
        using var createDocument = JsonDocument.Parse(create.Result!);
        var sessionId = createDocument.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));

        var (firstServerStream, firstClientStream) = FullDuplexStream.CreatePair();
        using var firstServerRpc = JsonRpc.Attach(firstServerStream, new DaemonRpcTarget(service));
        var firstClient = JsonRpc.Attach<IExcelDaemonRpc>(firstClientStream);
        var beforeDrop = await firstClient.ProcessCommandAsync(new ServiceRequest { Command = "session.list" });
        Assert.True(beforeDrop.Success, beforeDrop.ErrorMessage);
        Assert.Contains(sessionId!, beforeDrop.Result ?? string.Empty, StringComparison.Ordinal);

        ((IDisposable)firstClient).Dispose();
        var completion = await Task.WhenAny(firstServerRpc.Completion, Task.Delay(TimeSpan.FromSeconds(5)));
        Assert.Same(firstServerRpc.Completion, completion);
        Assert.Equal(1, service.SessionCount);

        var (secondServerStream, secondClientStream) = FullDuplexStream.CreatePair();
        using var secondServerRpc = JsonRpc.Attach(secondServerStream, new DaemonRpcTarget(service));
        var secondClient = JsonRpc.Attach<IExcelDaemonRpc>(secondClientStream);
        try
        {
            var afterDrop = await secondClient.ProcessCommandAsync(new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = sessionId,
                Args = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"A1\"}"
            });
            Assert.True(afterDrop.Success, afterDrop.ErrorMessage);

            var close = await secondClient.ProcessCommandAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = sessionId,
                Args = "{\"save\":false}"
            });
            Assert.True(close.Success, close.ErrorMessage);
        }
        finally
        {
            ((IDisposable)secondClient).Dispose();
        }
    }

}
