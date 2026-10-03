using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "TableStyles")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class TableStyleProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData("get-table-style")]
    [InlineData("delete-table-style")]
    public async Task Selection_PreservesName(string action)
    {
        var call = await fixture.CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = action,
            ["session_id"] = "session-1",
            ["style_name"] = "Custom"
        }, RecordingToolTest.Success("""{"success":true}"""), $"workbook.{action}", """{"styleName":"Custom"}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Creation_PreservesNativeSourceName()
    {
        var call = await fixture.CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "create-table-style",
            ["session_id"] = "session-1",
            ["style_name"] = "Custom",
            ["source_style_name"] = "TableStyleMedium2"
        }, RecordingToolTest.Success("""{"success":true}"""), "workbook.create-table-style",
            """{"styleName":"Custom","sourceStyleName":"TableStyleMedium2"}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Update_PreservesFalseAndNestedElementNames()
    {
        var options = new TableStyleOptions
        {
            ShowAsAvailableTableStyle = false,
            Elements = [new() { ElementType = "xlHeaderRow", Bold = false }]
        };
        var call = await fixture.CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "update-table-style",
            ["session_id"] = "session-1",
            ["style_name"] = "Custom",
            ["table_style_options"] = new { showAsAvailableTableStyle = false, elements = new[] { new { elementType = "xlHeaderRow", bold = false } } }
        }, RecordingToolTest.Success("""{"success":true}"""), "workbook.update-table-style",
            JsonSerializer.Serialize(new { styleName = "Custom", tableStyleOptions = options }, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_ExplainsNativeElementLimitsAndCrossWorkbookEffects()
    {
        var tool = Assert.Single(await fixture.ListToolsAsync(), item => item.Name == "workbook");
        Assert.Contains("list-table-styles", tool.Description, StringComparison.Ordinal);
        Assert.Contains("existing users", tool.Description, StringComparison.Ordinal);
        Assert.Contains("Font name/size", tool.Description, StringComparison.Ordinal);
        Assert.True(tool.JsonSchema.GetProperty("properties").TryGetProperty("table_style_options", out _));
    }
}
