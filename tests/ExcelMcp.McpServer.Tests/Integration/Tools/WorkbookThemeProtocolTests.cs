using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "WorkbookTheme")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class WorkbookThemeProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task ApplyTheme_ForwardsNativePathAndWorkbookScope()
    {
        var call = await fixture.CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "apply-theme",
            ["session_id"] = "session-1",
            ["theme_path"] = "selected.thmx"
        }, RecordingToolTest.Success("""{"success":true}"""), "workbook.apply-theme",
            """{"themePath":"selected.thmx"}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_ExplainsNativeThemeScopeAndHonestFontReads()
    {
        var tools = await fixture.ListToolsAsync();
        var workbook = Assert.Single(tools, tool => tool.Name == "workbook");
        Assert.Contains("theme-sensitive", workbook.Description, StringComparison.Ordinal);
        var read = Assert.Single(tools, tool => tool.Name == "workbook_read");
        Assert.Contains("get-theme", read.Description, StringComparison.Ordinal);
        Assert.Contains("no fallback font", read.Description, StringComparison.Ordinal);
        Assert.True(workbook.JsonSchema.GetProperty("properties").TryGetProperty("theme_path", out _));
    }
}
