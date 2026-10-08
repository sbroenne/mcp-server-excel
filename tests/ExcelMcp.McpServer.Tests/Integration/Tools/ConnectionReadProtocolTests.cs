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
    [Theory]
    [InlineData("connection_read", "get-account-settings", """{"success":true,"accountHintPresent":true,"passwordPresent":false,"impersonationPresent":false,"identityMode":"Connection"}""")]
    [InlineData("connection", "clear-account-hint", """{"success":true,"changed":true,"accountHintPresent":false}""")]
    public async Task AccountSettings_MapExactConnectionAndPreserveResults(string tool, string action, string result)
    {
        var call = await fixture.CallToolAsync(tool, new()
        {
            ["action"] = action,
            ["workbook_session_id"] = "session-1",
            ["connection_name"] = "Sales"
        }, RecordingToolTest.Success(result), "connection." + action,
            JsonSerializer.Serialize(new { connectionName = "Sales" }, ServiceProtocol.JsonOptions));

        Assert.False(call.Result.IsError);
        using var expected = JsonDocument.Parse(result);
        using var actual = JsonDocument.Parse(call.JsonResult);
        foreach (var property in expected.RootElement.EnumerateObject())
            Assert.Equal(property.Value.GetRawText(), actual.RootElement.GetProperty(property.Name).GetRawText());
    }

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
