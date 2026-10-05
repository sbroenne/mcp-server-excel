using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "NativeDataCleanup")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class NativeCleanupProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task RemoveDuplicates_MapsKeysAndHeadersToSharedService()
    {
        int[] columns = [1, 3];
        var expected = JsonSerializer.Serialize(new
        {
            sheetName = "Data",
            rangeAddress = "A1:C10",
            keyColumns = columns,
            hasHeaders = false
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("range_edit", new()
        {
            ["action"] = "remove-duplicates",
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Data",
            ["range_address"] = "A1:C10",
            ["key_columns"] = columns,
            ["has_headers"] = false
        }, RecordingToolTest.Success("""{"success":true,"removedRows":2,"remainingRows":8}"""),
            "rangeedit.remove-duplicates", expected);
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task TextToColumns_PreservesCamelCaseNestedOptionsAndPermission()
    {
        var options = new
        {
            comma = true,
            trailingMinusNumbers = false,
            fields = new[] { new { position = 1, dataType = "Text" } }
        };
        var expected = JsonSerializer.Serialize(new
        {
            sheetName = "Data",
            sourceRange = "A1:A10",
            destinationCell = "D1",
            options = new TextToColumnsOptions
            {
                Comma = true,
                TrailingMinusNumbers = false,
                Fields = [new TextColumnField { Position = 1, DataType = TextFieldType.Text }]
            },
            overwritePolicy = "allow"
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("range_edit", new()
        {
            ["action"] = "text-to-columns",
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Data",
            ["source_range"] = "A1:A10",
            ["destination_cell"] = "D1",
            ["options"] = options,
            ["overwrite_policy"] = "allow"
        }, RecordingToolTest.Success("""{"success":true,"destinationRange":"$D$1:$F$10","outputColumns":3}"""),
            "rangeedit.text-to-columns", expected);
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_ExposesNativeCleanupActionsAndInputNames()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "range_edit");
        Assert.Contains("remove-duplicates", tool.Description, StringComparison.Ordinal);
        Assert.Contains("text-to-columns", tool.Description, StringComparison.Ordinal);
        var properties = tool.JsonSchema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("key_columns", out _));
        Assert.True(properties.TryGetProperty("has_headers", out _));
        Assert.True(properties.TryGetProperty("destination_cell", out _));
        Assert.True(properties.TryGetProperty("options", out _));
    }
}
