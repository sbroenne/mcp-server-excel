using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "ChartDepth")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ChartTextProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task SetTitle_ForwardsPositionAndAlignment()
    {
        var call = await fixture.CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-title",
            ["workbook_session_id"] = "session-1",
            ["chart_name"] = "Sales",
            ["title"] = "Revenue",
            ["left"] = 16,
            ["top"] = 12,
            ["alignment"] = "Left"
        }, RecordingToolTest.Success("""{"success":true}"""), "chartconfig.set-title",
            """{"chartName":"Sales","title":"Revenue","left":16,"top":12,"alignment":"Left"}""");
        Assert.False(call.Result.IsError);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LineFormatting_ForwardsVisibilityAndSecondaryAxis(bool area)
    {
        var action = area ? "set-area-format" : "set-axis-format";
        var call = await fixture.CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = action,
            ["workbook_session_id"] = "session-1",
            ["chart_name"] = "Sales",
            [area ? "area" : "axis"] = area ? "Plot" : "ValueSecondary",
            ["line_color"] = "#D0D0D0",
            ["line_weight"] = 1.5,
            ["line_visible"] = false
        }, RecordingToolTest.Success("""{"success":true}"""), $"chartconfig.{action}",
            area
                ? """{"chartName":"Sales","area":"Plot","lineColor":"#D0D0D0","lineWeight":1.5,"lineVisible":false}"""
                : """{"chartName":"Sales","axis":"ValueSecondary","lineColor":"#D0D0D0","lineWeight":1.5,"lineVisible":false}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task TitleAndLineFormatting_OmittedSettingsStayOmitted()
    {
        var title = await fixture.CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-title",
            ["workbook_session_id"] = "session-1",
            ["chart_name"] = "Sales",
            ["title"] = "Revenue"
        }, RecordingToolTest.Success("""{"success":true}"""), "chartconfig.set-title",
            """{"chartName":"Sales","title":"Revenue"}""");
        Assert.False(title.Result.IsError);
        var axis = await fixture.CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-axis-format",
            ["workbook_session_id"] = "session-1",
            ["chart_name"] = "Sales",
            ["axis"] = "Value",
            ["line_color"] = "#D0D0D0"
        }, RecordingToolTest.Success("""{"success":true}"""), "chartconfig.set-axis-format",
            """{"chartName":"Sales","axis":"Value","lineColor":"#D0D0D0"}""");
        Assert.False(axis.Result.IsError);
    }

    [Fact]
    public async Task SetTextFormat_PreservesExplicitFalseAndFontSettings()
    {
        var call = await fixture.CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-text-format",
            ["workbook_session_id"] = "session-1",
            ["chart_name"] = "Sales",
            ["element"] = "Title",
            ["font_name"] = "Arial",
            ["font_size"] = 14.5,
            ["bold"] = false,
            ["italic"] = true,
            ["color"] = "#2468AC"
        }, RecordingToolTest.Success("""{"success":true,"formats":[]}"""), "chartconfig.set-text-format",
            """{"chartName":"Sales","element":"Title","fontName":"Arial","fontSize":14.5,"bold":false,"italic":true,"color":"#2468AC"}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task ReadTextFormat_UsesReadToolAndPreservesNativeResults()
    {
        const string result = """{"success":true,"chartName":"Sales","element":"DataLabels","formats":[{"seriesIndex":2,"fontName":"Arial","fontSize":14.5,"bold":false,"italic":true,"color":"#2468AC","mixedFields":[]}]}""";
        var call = await fixture.CallToolAsync("chart_config_read", new Dictionary<string, object?>
        {
            ["action"] = "get-text-format",
            ["workbook_session_id"] = "session-1",
            ["chart_name"] = "Sales",
            ["element"] = "DataLabels",
            ["series_index"] = 2
        }, RecordingToolTest.Success(result), "chartconfig.get-text-format",
            """{"chartName":"Sales","element":"DataLabels","seriesIndex":2}""");
        Assert.False(call.Result.IsError);
        using var expected = JsonDocument.Parse(result);
        using var actual = JsonDocument.Parse(call.JsonResult);
        Assert.True(JsonElement.DeepEquals(expected.RootElement, actual.RootElement));
    }

    [Fact]
    public async Task TextFormat_FailureIsProtocolErrorWithOriginalContext()
    {
        var call = await fixture.CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-text-format",
            ["workbook_session_id"] = "session-1",
            ["chart_name"] = "Sales",
            ["element"] = "Title",
            ["font_size"] = 14
        }, new ServiceResponse
        {
            Success = false,
            ErrorCategory = "OperationFailed",
            ExceptionType = "InvalidOperationException",
            ErrorMessage = "The chart has no title.",
            HResult = "0x800A03EC",
            InnerError = "Excel rejected the operation."
        }, "chartconfig.set-text-format", """{"chartName":"Sales","element":"Title","fontSize":14}""");
        Assert.True(call.Result.IsError);
        using var state = JsonDocument.Parse(call.JsonResult);
        Assert.False(state.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("The chart has no title.", state.RootElement.GetProperty("errorMessage").GetString());
        Assert.Equal("0x800A03EC", state.RootElement.GetProperty("hresult").GetString());
        Assert.Equal("Excel rejected the operation.", state.RootElement.GetProperty("innerError").GetString());
    }

    [Fact]
    public async Task Discovery_ExposesFontSettingsAndReadOnlyAction()
    {
        var tools = await fixture.ListToolsAsync();
        var write = Assert.Single(tools, tool => tool.Name == "chart_config");
        var read = Assert.Single(tools, tool => tool.Name == "chart_config_read");
        foreach (var field in new[] { "element", "font_name", "font_size", "bold", "italic", "color",
            "left", "top", "alignment", "line_visible" })
            Assert.True(write.JsonSchema.GetProperty("properties").TryGetProperty(field, out _), field);
        Assert.Contains("set-axis-format", write.Description, StringComparison.Ordinal);
        Assert.Contains("get-text-format", read.Description, StringComparison.Ordinal);
        Assert.Contains("mixedFields", read.Description, StringComparison.Ordinal);
        Assert.Contains("set-text-format", write.Description, StringComparison.Ordinal);
    }
}
