using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Workbook")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class WorkbookOverviewProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task Inspect_ForwardsSelectedSectionsAndLimits()
    {
        var expectedArgs = JsonSerializer.Serialize(new
        {
            sheetName = "Summary",
            includeSheets = true,
            includeTables = false,
            includeDefinedNames = true,
            includePreview = true,
            rangeAddress = "A1:D20",
            maxItems = 7,
            maxPreviewRows = 3,
            maxPreviewColumns = 4,
            maxCellCharacters = 20,
            maxPreviewCharacters = 123
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("workbook_read", new()
        {
            ["action"] = "inspect",
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Summary",
            ["include_sheets"] = true,
            ["include_tables"] = false,
            ["include_defined_names"] = true,
            ["include_preview"] = true,
            ["range_address"] = "A1:D20",
            ["max_items"] = 7,
            ["max_preview_rows"] = 3,
            ["max_preview_columns"] = 4,
            ["max_cell_characters"] = 20,
            ["max_preview_characters"] = 123
        }, RecordingToolTest.Success("""{"success":true}"""), "workbook.inspect", expectedArgs);

        Assert.False(call.Result.IsError);
        using var output = JsonDocument.Parse(call.JsonResult);
        Assert.True(output.RootElement.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task Discovery_AdvertisesInspectAsReadOnlyWithBounds()
    {
        var tools = await fixture.ListToolsAsync();
        var readTool = Assert.Single(tools, tool => tool.Name == "workbook_read");
        Assert.Contains("inspect", readTool.Description, StringComparison.Ordinal);
        Assert.Contains("bounded", readTool.Description, StringComparison.OrdinalIgnoreCase);
        var properties = readTool.JsonSchema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("max_preview_rows", out _));
        Assert.True(properties.TryGetProperty("max_preview_columns", out _));
        Assert.True(properties.TryGetProperty("max_preview_characters", out _));
    }
}
