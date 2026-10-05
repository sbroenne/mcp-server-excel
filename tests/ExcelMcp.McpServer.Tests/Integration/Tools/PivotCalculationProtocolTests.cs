using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "PivotCalculation")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PivotCalculationProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData("Named", "North")]
    [InlineData("Previous", null)]
    [InlineData("Next", null)]
    public async Task AdditionalCalculation_MapsExplicitBaseSettings(string kind, string? itemName)
    {
        var expected = JsonSerializer.Serialize(new
        {
            pivotTableName = "SalesPivot",
            fieldName = "Total Sales",
            calculation = "DifferenceFrom",
            baseFieldName = "Region",
            baseItemKind = kind,
            baseItemName = itemName
        }, ServiceProtocol.JsonOptions);
        Dictionary<string, object?> arguments = new()
        {
            ["action"] = "set-field-calculation",
            ["workbook_session_id"] = "session-1",
            ["pivot_table_name"] = "SalesPivot",
            ["field_name"] = "Total Sales",
            ["calculation"] = "DifferenceFrom",
            ["base_field_name"] = "Region",
            ["base_item_kind"] = kind
        };
        if (itemName is not null)
            arguments["base_item_name"] = itemName;
        var call = await fixture.CallToolAsync("pivottable_field", arguments,
            RecordingToolTest.Success("""{"success":true}"""),
            "pivottablefield.set-field-calculation", expected);
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task NormalReset_DoesNotInventBaseSettings()
    {
        var expected = JsonSerializer.Serialize(new
        {
            pivotTableName = "SalesPivot",
            fieldName = "Average Sales",
            calculation = "Normal"
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("pivottable_field", new()
        {
            ["action"] = "set-field-calculation",
            ["workbook_session_id"] = "session-1",
            ["pivot_table_name"] = "SalesPivot",
            ["field_name"] = "Average Sales",
            ["calculation"] = "Normal"
        }, RecordingToolTest.Success("""{"success":true}"""),
            "pivottablefield.set-field-calculation", expected);
        Assert.False(call.Result.IsError);
    }

    [Theory]
    [InlineData("Missing")]
    [InlineData(null)]
    public async Task MissingOrUnknownCalculation_DoesNotDispatch(string? calculation)
    {
        Dictionary<string, object?> arguments = new()
        {
            ["action"] = "set-field-calculation",
            ["workbook_session_id"] = "session-1",
            ["pivot_table_name"] = "SalesPivot",
            ["field_name"] = "Total Sales"
        };
        if (calculation is not null)
            arguments["calculation"] = calculation;
        var call = await fixture.CallResultWithoutDispatchAsync("pivottable_field", arguments);
        Assert.True(call.IsError);
    }

    [Fact]
    public async Task Discovery_DescribesDisplayedInstancesAndNativeRestrictions()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "pivottable_field");
        Assert.Contains("SHOW VALUES AS", tool.Description, StringComparison.Ordinal);
        Assert.Contains("BaseField/BaseItem are unavailable", tool.Description, StringComparison.Ordinal);
        var properties = tool.JsonSchema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("calculation", out _));
        Assert.True(properties.TryGetProperty("base_field_name", out _));
        Assert.True(properties.TryGetProperty("base_item_kind", out _));
        Assert.True(properties.TryGetProperty("base_item_name", out _));
    }
}
