using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeSpillProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task GetSpillInfo_MapsScopeAndPreservesCompleteRelationships()
    {
        var response = JsonSerializer.Serialize(new
        {
            success = true,
            capability = "supported",
            cellCount = 64,
            cells = Enumerable.Range(1, 64).Select(row => new
            {
                address = $"$A${row}",
                state = row == 1 ? "source" : "result",
                sourceAddress = "$A$1",
                spillAddress = "$A$1:$A$64"
            }).ToArray()
        });
        var expectedArgs = JsonSerializer.Serialize(new
        {
            sheetName = "Sheet1",
            rangeAddress = "A1:A64"
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("range_read", new()
        {
            ["action"] = "get-spill-info",
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1:A64"
        }, RecordingToolTest.Success(response), "range.get-spill-info", expectedArgs);
        Assert.False(call.Result.IsError);
        using var output = JsonDocument.Parse(call.JsonResult);
        Assert.Equal(64, output.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(64, output.RootElement.GetProperty("cells").GetArrayLength());
        Assert.Equal("$A$1", output.RootElement.GetProperty("cells")[63].GetProperty("sourceAddress").GetString());
    }

    [Fact]
    public async Task Discovery_ExplainsNativeRelationshipsAndUnsupportedSessions()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "range_read");
        Assert.Contains("get-spill-info", tool.Description, StringComparison.Ordinal);
        Assert.Contains("Unsupported Excel sessions fail explicitly", tool.Description, StringComparison.Ordinal);
    }
}
