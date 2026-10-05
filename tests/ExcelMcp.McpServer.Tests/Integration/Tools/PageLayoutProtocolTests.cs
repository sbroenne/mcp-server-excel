using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "PageLayout")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PageLayoutProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task Setup_MapsPointMarginsAndExplicitClearingWithoutOrientation()
    {
        var options = new PageSetupOptions { PrintArea = "", LeftMargin = 36, ZoomPercent = 90 };
        var call = await fixture.CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "set-page-setup",
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Report",
            ["page_setup_options"] = new { printArea = "", leftMargin = 36, zoomPercent = 90 }
        }, RecordingToolTest.Success("""{"success":true}"""), "sheet.set-page-setup",
            JsonSerializer.Serialize(new { sheetName = "Report", pageSetupOptions = options }, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Breaks_MapExplicitEmptyListWithoutProtectionOptions()
    {
        var options = new PageBreakOptions { Rows = [10], Columns = [] };
        var call = await fixture.CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "set-page-breaks",
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Report",
            ["page_break_options"] = new { rows = new List<int> { 10 }, columns = new List<int>() }
        }, RecordingToolTest.Success("""{"success":true}"""), "sheet.set-page-breaks",
            JsonSerializer.Serialize(new { sheetName = "Report", pageBreakOptions = options }, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_SeparatesPrintAndProtectionPayloads()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "worksheet_style");
        var properties = tool.JsonSchema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("page_setup_options", out _));
        Assert.True(properties.TryGetProperty("page_break_options", out _));
        Assert.True(properties.TryGetProperty("options", out _));
    }
}
