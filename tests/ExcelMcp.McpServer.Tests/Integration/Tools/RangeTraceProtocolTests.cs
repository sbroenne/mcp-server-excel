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
public sealed class RangeTraceProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData("trace-precedents", true)]
    [InlineData("trace-precedents", false)]
    [InlineData("trace-dependents", true)]
    [InlineData("trace-dependents", false)]
    public async Task Trace_MapsLimitsAndPreservesPartialResults(string action, bool useDefaults)
    {
        Dictionary<string, object?> expectedArguments = new()
        {
            ["sheetName"] = "Sheet1",
            ["rangeAddress"] = "A1,C3"
        };
        const string response =
            """{"success":true,"coverage":{"scope":"same-worksheet-only","workbookComplete":false,"nativeTraversalComplete":false},"limits":{"maxDepth":1,"maxNodes":2,"maxPreviewCharacters":8},"truncation":{"any":true,"depthLimitReached":true,"nodeLimitReached":false,"startingCellsOmitted":false,"previewTextTruncated":true},"nodes":[{"address":"$A$1","depth":0,"formula":"=SUM(B1:","formulaTruncated":true}],"edges":[],"cycles":[],"unresolved":[],"boundaries":[{"address":"$A$1","reason":"max-depth"}]}""";
        Dictionary<string, object?> arguments = new()
        {
            ["action"] = action,
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1,C3"
        };
        if (!useDefaults)
        {
            arguments["max_depth"] = expectedArguments["maxDepth"] = 1;
            arguments["max_nodes"] = expectedArguments["maxNodes"] = 2;
            arguments["max_preview_characters"] = expectedArguments["maxPreviewCharacters"] = 8;
        }
        var expected = JsonSerializer.Serialize(expectedArguments, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("range_read", arguments,
            RecordingToolTest.Success(response), $"range.{action}", expected);
        Assert.False(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        using var expectedResult = JsonDocument.Parse(response);
        Assert.True(JsonElement.DeepEquals(expectedResult.RootElement, result.RootElement));
    }

    [Fact]
    public async Task Discovery_ExplainsUnresolvedAndWorkbookCoverageBoundaries()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "range_read");
        Assert.Contains("coverage.workbookComplete is always false", tool.Description, StringComparison.Ordinal);
        Assert.Contains("unresolved, not fabricated empty coverage", tool.Description, StringComparison.Ordinal);
        Assert.Contains("explicit truncation", tool.Description, StringComparison.Ordinal);
        Assert.Contains("truncation.any", tool.Description, StringComparison.Ordinal);
        Assert.Contains("trace-precedents", tool.Description, StringComparison.Ordinal);
        Assert.Contains("trace-dependents", tool.Description, StringComparison.Ordinal);
        var properties = tool.JsonSchema.GetProperty("properties");
        Assert.Contains("0-20", properties.GetProperty("max_depth").GetProperty("description").GetString(), StringComparison.Ordinal);
        Assert.Contains("default 3", properties.GetProperty("max_depth").GetProperty("description").GetString(), StringComparison.Ordinal);
        Assert.Contains("default 50", properties.GetProperty("max_nodes").GetProperty("description").GetString(), StringComparison.Ordinal);
        Assert.Contains("default 256", properties.GetProperty("max_preview_characters").GetProperty("description").GetString(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("trace-precedents")]
    [InlineData("trace-dependents")]
    [InlineData("get-values")]
    public async Task Trace_InvalidTypeOrUnrelatedActionDoesNotDispatch(string action)
    {
        var result = await fixture.CallResultWithoutDispatchAsync("range_read", new()
        {
            ["action"] = action,
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1",
            ["max_depth"] = action == "get-values" ? 1 : "not-a-number"
        });
        Assert.True(result.IsError);
    }
}
