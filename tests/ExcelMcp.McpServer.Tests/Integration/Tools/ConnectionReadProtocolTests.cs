using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Connection")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ConnectionReadProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task Test_MapsConnectionInspectionThroughReadEndpoint()
    {
        var call = await fixture.CallToolAsync("connection_read", new()
        {
            ["action"] = "test",
            ["workbook_session_id"] = "session-1",
            ["connection_name"] = "Sales"
        }, RecordingToolTest.Success("""{"success":true}"""), "connection.test",
            JsonSerializer.Serialize(new { connectionName = "Sales" }, ServiceProtocol.JsonOptions));

        Assert.False(call.Result.IsError);
        using var document = JsonDocument.Parse(call.JsonResult);
        Assert.True(document.RootElement.GetProperty("success").GetBoolean());
    }
}
