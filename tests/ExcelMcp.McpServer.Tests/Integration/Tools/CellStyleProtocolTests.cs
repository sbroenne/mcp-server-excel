using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "CellStyles")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class CellStyleProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData("get-cell-style")]
    [InlineData("delete-cell-style")]
    public async Task Selection_PreservesExactStyleName(string action)
    {
        var toolName = action == "get-cell-style" ? "workbook_read" : "workbook";
        var call = await fixture.CallToolAsync(toolName, new Dictionary<string, object?>
        {
            ["action"] = action,
            ["session_id"] = "session-1",
            ["style_name"] = "Custom"
        }, RecordingToolTest.Success("""{"success":true}"""), $"workbook.{action}", """{"styleName":"Custom"}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Creation_PreservesSourceCellWithoutApplyingTheStyle()
    {
        var call = await fixture.CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "create-cell-style",
            ["session_id"] = "session-1",
            ["style_name"] = "Custom",
            ["source_sheet_name"] = "Data",
            ["source_cell_address"] = "A1"
        }, RecordingToolTest.Success("""{"success":true}"""), "workbook.create-cell-style",
            """{"styleName":"Custom","sourceSheetName":"Data","sourceCellAddress":"A1"}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Update_PreservesNestedKeysAndExplicitFalse()
    {
        var options = new CellStyleOptions { IncludeFont = false, FormatOptions = new() { Bold = false, FillThemeColor = 5 } };
        var call = await fixture.CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "update-cell-style",
            ["session_id"] = "session-1",
            ["style_name"] = "Custom",
            ["style_options"] = new { includeFont = false, formatOptions = new { bold = false, fillThemeColor = 5 } }
        }, RecordingToolTest.Success("""{"success":true}"""), "workbook.update-cell-style",
            JsonSerializer.Serialize(new { styleName = "Custom", styleOptions = options }, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_ExplainsEffectsOnExistingUsersAndBuiltInRestrictions()
    {
        var tools = await fixture.ListToolsAsync();
        var writeTool = Assert.Single(tools, item => item.Name == "workbook");
        Assert.Contains("existing users", writeTool.Description, StringComparison.Ordinal);
        Assert.Contains("Built-in styles are read-only", writeTool.Description, StringComparison.Ordinal);
        Assert.True(writeTool.JsonSchema.GetProperty("properties").TryGetProperty("style_options", out _));
        var readTool = Assert.Single(tools, item => item.Name == "workbook_read");
        Assert.Contains("list-cell-styles", readTool.Description, StringComparison.Ordinal);
    }
}
