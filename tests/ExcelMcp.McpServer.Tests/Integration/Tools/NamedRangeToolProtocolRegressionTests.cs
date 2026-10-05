using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "false")]
public sealed class NamedRangeToolProtocolRegressionTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task NamedRangeList_ReturnsStructuredResult_AndSessionRemainsUsable()
    {
        const string sessionId = "recording-session";
        var listCall = await _fixture.CallToolAsync(
            "namedrange_read",
            new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["session_id"] = sessionId
            },
            RecordingToolTest.Success(
                """{"success":true,"namedRanges":[{"name":"CsvFolder","refersTo":"=Sheet1!$B$4","value":"C:\\Data"}]}"""),
            "namedrange.list",
            null);

        Assert.Equal("namedrange.list", listCall.Request.Command);
        Assert.Equal(sessionId, listCall.Request.SessionId);
        using (var result = JsonDocument.Parse(listCall.JsonResult))
        {
            var range = Assert.Single(
                result.RootElement.GetProperty("namedRanges").EnumerateArray());
            Assert.Equal("CsvFolder", range.GetProperty("name").GetString());
            Assert.Contains(
                "$B$4",
                range.GetProperty("refersTo").GetString(),
                StringComparison.OrdinalIgnoreCase);
            Assert.Equal("C:\\Data", range.GetProperty("value").GetString());
        }

        var worksheetCall = await _fixture.CallToolAsync(
            "worksheet_read",
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
        using var worksheetResult = JsonDocument.Parse(worksheetCall.JsonResult);
        Assert.Equal(
            "Sheet1",
            worksheetResult.RootElement.GetProperty("worksheets")[0]
                .GetProperty("name").GetString());
    }
}
