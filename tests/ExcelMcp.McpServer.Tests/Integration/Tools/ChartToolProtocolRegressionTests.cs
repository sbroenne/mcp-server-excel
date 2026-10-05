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
    public async Task ChartRead_PreservesPivotLinkAndActualPlottedArrays()
    {
        const string response = """
            {"success":true,"name":"RevenueChart","isPivotChart":true,"linkedPivotTable":"RevenuePivot","series":[
              {"name":"Alpha","valuesRange":"","categoryRange":null,"values":[10,40],"categories":["Q1","Q2"]},
              {"name":"Beta","valuesRange":"","categoryRange":null,"values":[20,50],"categories":["Q1","Q2"]},
              {"name":"Gamma","valuesRange":"","categoryRange":null,"values":[30,60],"categories":["Q1","Q2"]}
            ]}
            """;
        var call = await _fixture.CallToolAsync(
            "chart_read",
            new Dictionary<string, object?>
            {
                ["action"] = "read",
                ["workbook_session_id"] = "recording-session",
                ["chart_name"] = "RevenueChart"
            },
            RecordingToolTest.Success(response),
            "chart.read",
            """{"chartName":"RevenueChart"}""");

        using var result = JsonDocument.Parse(call.JsonResult);
        var root = result.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.True(root.GetProperty("isPivotChart").GetBoolean());
        Assert.Equal("RevenuePivot", root.GetProperty("linkedPivotTable").GetString());
        var series = root.GetProperty("series");
        Assert.Equal(3, series.GetArrayLength());
        Assert.Equal(60, series[2].GetProperty("values")[1].GetInt32());
        Assert.Equal("Q2", series[2].GetProperty("categories")[1].GetString());
    }

    [Fact]
    public async Task ChartList_EmptyWorkbook_ReturnsStructuredEmptyList_AndSessionRemainsUsable()
    {
        const string sessionId = "recording-session";
        var listCall = await _fixture.CallToolAsync(
            "chart_read",
            new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["workbook_session_id"] = sessionId
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
            "worksheet_read",
            new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["workbook_session_id"] = sessionId
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
