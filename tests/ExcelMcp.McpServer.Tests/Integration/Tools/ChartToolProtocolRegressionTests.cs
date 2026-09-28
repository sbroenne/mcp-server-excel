using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Charts")]
[Trait("RequiresExcel", "false")]
public sealed class ChartToolProtocolRegressionTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task ChartList_EmptyWorkbook_ReturnsStructuredEmptyList_AndSessionRemainsUsable()
    {
        const string sessionId = "recording-session";
        var listCall = await _fixture.CallToolAsync(
            "chart",
            new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["session_id"] = sessionId
            },
            RecordingToolTest.Success("""{"success":true,"charts":[]}"""),
            "chart.list",
            null);

        Assert.Equal("chart.list", listCall.Request.Command);
        Assert.Equal(sessionId, listCall.Request.SessionId);
        using (var result = JsonDocument.Parse(listCall.JsonResult))
        {
            Assert.Empty(result.RootElement.GetProperty("charts").EnumerateArray());
        }

        var worksheetCall = await _fixture.CallToolAsync(
            "worksheet",
            new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["session_id"] = sessionId
            },
            RecordingToolTest.Success(
                """{"success":true,"worksheets":[{"name":"Sheet1"}]}"""),
            "sheet.list",
            "{}");

        Assert.Equal("sheet.list", worksheetCall.Request.Command);
        Assert.Equal(sessionId, worksheetCall.Request.SessionId);
        using var worksheetResult = JsonDocument.Parse(worksheetCall.JsonResult);
        Assert.Equal(
            "Sheet1",
            worksheetResult.RootElement.GetProperty("worksheets")[0]
                .GetProperty("name").GetString());
    }
}
