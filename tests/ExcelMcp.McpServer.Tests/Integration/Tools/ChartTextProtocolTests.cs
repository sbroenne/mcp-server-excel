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
        foreach (var field in new[] { "element", "font_name", "font_size", "bold", "italic", "color" })
            Assert.True(write.JsonSchema.GetProperty("properties").TryGetProperty(field, out _), field);
        Assert.Contains("get-text-format", read.Description, StringComparison.Ordinal);
        Assert.Contains("mixedFields", read.Description, StringComparison.Ordinal);
        Assert.Contains("set-text-format", write.Description, StringComparison.Ordinal);
    }
}
