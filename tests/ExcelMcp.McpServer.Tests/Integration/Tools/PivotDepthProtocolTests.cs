using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "PivotDepth")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PivotDepthProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task ValueFilter_PreservesExactCaptionAndTypedOptions()
    {
        var options = new PivotFilterOptions { Type = PivotFilterType.ValueIsGreaterThan, Number1 = 150d, DataFieldName = "Total Sales" };
        var call = await fixture.CallToolAsync("pivottable_field", new Dictionary<string, object?>
        {
            ["action"] = "add-field-filter",
            ["session_id"] = "session-1",
            ["pivot_table_name"] = "Sales",
            ["field_name"] = "Region",
            ["filter_options"] = new { type = "ValueIsGreaterThan", number1 = 150d, dataFieldName = "Total Sales" }
        }, RecordingToolTest.Success("""{"success":true}"""), "pivottablefield.add-field-filter",
            JsonSerializer.Serialize(new { pivotTableName = "Sales", fieldName = "Region", filterOptions = options }, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Layout_PreservesTypedFlagsIncludingFalse()
    {
        var options = new PivotLayoutOptions { RowLayout = 1, RepeatLabels = true, PreserveFormatting = false, StyleName = "PivotStyleMedium9" };
        var call = await fixture.CallToolAsync("pivottable_calc", new Dictionary<string, object?>
        {
            ["action"] = "set-layout-options",
            ["session_id"] = "session-1",
            ["pivot_table_name"] = "Sales",
            ["layout_options"] = new { rowLayout = 1, repeatLabels = true, preserveFormatting = false, styleName = "PivotStyleMedium9" }
        }, RecordingToolTest.Success("""{"success":true}"""), "pivottablecalc.set-layout-options",
            JsonSerializer.Serialize(new { pivotTableName = "Sales", layoutOptions = options }, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task SourceSelection_KeepsNativeTopLevelNames()
    {
        var call = await fixture.CallToolAsync("pivottable", new Dictionary<string, object?>
        {
            ["action"] = "set-source",
            ["session_id"] = "session-1",
            ["pivot_table_name"] = "Sales",
            ["source_sheet_name"] = "Data",
            ["table_name"] = "Source"
        }, RecordingToolTest.Success("""{"success":true}"""), "pivottable.set-source",
            """{"pivotTableName":"Sales","sourceSheetName":"Data","tableName":"Source"}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Expansion_KeepsExplicitFalse()
    {
        var call = await fixture.CallToolAsync("pivottable_field", new Dictionary<string, object?>
        {
            ["action"] = "set-item-expansion",
            ["session_id"] = "session-1",
            ["pivot_table_name"] = "Sales",
            ["field_name"] = "Region",
            ["item_name"] = "North",
            ["expanded"] = false
        }, RecordingToolTest.Success("""{"success":true}"""), "pivottablefield.set-item-expansion",
            """{"pivotTableName":"Sales","fieldName":"Region","itemName":"North","expanded":false}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_AdvertisesNativeStyleAndSourceSafety()
    {
        var tools = await fixture.ListToolsAsync();
        var pivot = Assert.Single(tools, tool => tool.Name == "pivottable");
        Assert.Contains("set-source", pivot.Description, StringComparison.Ordinal);
        Assert.Contains("set-layout-options", pivot.Description, StringComparison.Ordinal);
        Assert.DoesNotContain("styles are not supported", pivot.Description, StringComparison.Ordinal);
        var fields = Assert.Single(tools, tool => tool.Name == "pivottable_field");
        Assert.True(fields.JsonSchema.GetProperty("properties").TryGetProperty("filter_options", out _));
        var layout = Assert.Single(tools, tool => tool.Name == "pivottable_calc");
        Assert.True(layout.JsonSchema.GetProperty("properties").TryGetProperty("layout_options", out _));
    }
}
