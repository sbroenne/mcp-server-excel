using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeFillProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData("down")]
    [InlineData("up")]
    [InlineData("left")]
    [InlineData("right")]
    public async Task Fill_MapsDirectionWithoutInventingDefaults(string direction)
    {
        var expected = JsonSerializer.Serialize(new
        {
            sheetName = "Sheet1",
            rangeAddress = "A1:B3",
            direction
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("range_edit", new()
        {
            ["action"] = "fill",
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1:B3",
            ["direction"] = direction
        }, RecordingToolTest.Success("""{"success":true}"""), "rangeedit.fill", expected);
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task AutoFill_MapsNativePatternOptions()
    {
        var expected = JsonSerializer.Serialize(new
        {
            sheetName = "Sheet1",
            sourceRange = "A1:A2",
            destinationRange = "A1:A10",
            fillType = "series"
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("range_edit", new()
        {
            ["action"] = "auto-fill",
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["source_range"] = "A1:A2",
            ["destination_range"] = "A1:A10",
            ["fill_type"] = "series"
        }, RecordingToolTest.Success("""{"success":true}"""), "rangeedit.auto-fill", expected);
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Series_MapsNumericAndDateOptions()
    {
        var expected = JsonSerializer.Serialize(new
        {
            sheetName = "Sheet1",
            rangeAddress = "A1:A10",
            orientation = "columns",
            seriesType = "date",
            stepValue = 2.5,
            stopValue = 50000d,
            dateUnit = "month",
            trend = false,
            overwritePolicy = "allow"
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("range_edit", new()
        {
            ["action"] = "create-series",
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1:A10",
            ["orientation"] = "columns",
            ["series_type"] = "date",
            ["step_value"] = 2.5,
            ["stop_value"] = 50000d,
            ["date_unit"] = "month",
            ["trend"] = false,
            ["overwrite_policy"] = "allow"
        }, RecordingToolTest.Success("""{"success":true}"""), "rangeedit.create-series", expected);
        Assert.False(call.Result.IsError);
    }

    [Theory]
    [InlineData("get-formulas")]
    [InlineData("set-formulas")]
    public async Task Formulas_MapsNativeR1C1Notation(string action)
    {
        Dictionary<string, object?> arguments = new()
        {
            ["action"] = action,
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "B1",
            ["reference_style"] = "r1c1"
        };
        Dictionary<string, object?> expected = new()
        {
            ["sheetName"] = "Sheet1",
            ["rangeAddress"] = "B1",
            ["referenceStyle"] = "r1c1"
        };
        if (action == "set-formulas")
        {
            List<List<string>> formulas = [["=RC[-1]*2"]];
            arguments["formulas"] = formulas;
            expected["formulas"] = formulas;
        }
        var call = await fixture.CallToolAsync("range", arguments,
            RecordingToolTest.Success("""{"success":true}"""), $"range.{action}",
            JsonSerializer.Serialize(expected, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_AdvertisesSourceAndNotationBoundaries()
    {
        var tools = await fixture.ListToolsAsync();
        var edit = Assert.Single(tools, tool => tool.Name == "range_edit");
        Assert.Contains("exactly one direction", edit.Description, StringComparison.Ordinal);
        Assert.Contains("requires allow", edit.Description, StringComparison.Ordinal);
        var properties = edit.JsonSchema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("direction", out _));
        Assert.True(properties.TryGetProperty("source_range", out _));
        Assert.True(properties.TryGetProperty("stop_value", out _));
        var range = Assert.Single(tools, tool => tool.Name == "range");
        Assert.Contains("range addresses stay A1", range.Description, StringComparison.Ordinal);
        Assert.True(range.JsonSchema.GetProperty("properties").TryGetProperty("reference_style", out _));
    }
}
