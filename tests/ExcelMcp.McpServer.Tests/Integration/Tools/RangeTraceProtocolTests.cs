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
    [InlineData("trace-precedents")]
    [InlineData("trace-dependents")]
    public async Task Trace_MapsExactScopeWithoutHiddenLimits(string action)
    {
        var expected = JsonSerializer.Serialize(new
        {
            sheetName = "Sheet1",
            rangeAddress = "A1,C3"
        }, ServiceProtocol.JsonOptions);
        const string response =
            """{"success":true,"coverage":{"scope":"same-worksheet-only","workbookComplete":false},"nodes":[],"edges":[],"unresolved":[]}""";
        var call = await fixture.CallToolAsync("range_read", new()
        {
            ["action"] = action,
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1,C3"
        }, RecordingToolTest.Success(response), $"range.{action}", expected);
        Assert.False(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.False(result.RootElement.GetProperty("coverage").GetProperty("workbookComplete").GetBoolean());
    }

    [Fact]
    public async Task Discovery_ExplainsUnresolvedAndWorkbookCoverageBoundaries()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "range_read");
        Assert.Contains("coverage.workbookComplete is always false", tool.Description, StringComparison.Ordinal);
        Assert.Contains("unresolved, not fabricated empty coverage", tool.Description, StringComparison.Ordinal);
        Assert.Contains("without a depth/output cap", tool.Description, StringComparison.Ordinal);
        Assert.Contains("trace-precedents", tool.Description, StringComparison.Ordinal);
        Assert.Contains("trace-dependents", tool.Description, StringComparison.Ordinal);
    }
}
