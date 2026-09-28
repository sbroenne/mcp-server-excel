using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class RangeFormatFormatRangesTests(
    RecordingProgramTransportFixture fixture)
{
    private static readonly string[] TargetRanges = ["A1:A2", "C1:C2"];
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task FormatRanges_NumberFormat_RoundTripsViaMcpProtocol()
    {
        const string sessionId = "recording-session";
        var call = await _fixture.CallToolAsync(
            "range_format",
            new Dictionary<string, object?>
            {
                ["action"] = "format-ranges",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Formatting",
                ["range_addresses"] = TargetRanges,
                ["number_format"] = "0.00%"
            },
            RecordingToolTest.Success(
                """{"success":true,"rangeCount":2,"numberFormat":"0.00%"}"""),
            "rangeformat.format-ranges",
            """{"sheetName":"Formatting","rangeAddresses":["A1:A2","C1:C2"],"numberFormat":"0.00%"}""");

        using var args = RecordingToolTest.ParseArgs(
            call.Request,
            "rangeformat.format-ranges",
            sessionId);
        var root = args.RootElement;
        Assert.Equal("Formatting", root.GetProperty("sheetName").GetString());
        var ranges = root.GetProperty("rangeAddresses");
        Assert.Equal(2, ranges.GetArrayLength());
        Assert.Equal("A1:A2", ranges[0].GetString());
        Assert.Equal("C1:C2", ranges[1].GetString());
        Assert.Equal("0.00%", root.GetProperty("numberFormat").GetString());

        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.Equal(2, result.RootElement.GetProperty("rangeCount").GetInt32());
        Assert.Equal(
            "0.00%",
            result.RootElement.GetProperty("numberFormat").GetString());
    }
}
