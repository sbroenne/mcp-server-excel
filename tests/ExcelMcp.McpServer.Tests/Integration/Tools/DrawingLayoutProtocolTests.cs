using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "DrawingLayout")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class DrawingLayoutProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData("group-objects", "group_name", "Together", "groupName")]
    [InlineData("align-objects", "alignment", "Left", "alignment")]
    [InlineData("distribute-objects", "distribution", "Horizontal", "distribution")]
    [InlineData("duplicate-object", "new_name", "Copy", "newName")]
    [InlineData("set-z-order", "z_order", "BringToFront", "zOrder")]
    [InlineData("ungroup-object", "object_name", "Together", "objectName")]
    public async Task Layout_PreservesNativeArgumentsAndDefaults(string action, string input, string value, string output)
    {
        var args = new Dictionary<string, object?>
        {
            ["action"] = action,
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Dashboard",
            [input] = value
        };
        var expected = new Dictionary<string, object?> { ["sheetName"] = "Dashboard", [output] = value };
        if (action is "group-objects" or "align-objects" or "distribute-objects")
        {
            args["object_names"] = """["First","Second","Third"]""";
            expected["objectNames"] = new[] { "First", "Second", "Third" };
        }
        else if (action != "ungroup-object")
        {
            args["object_name"] = "First";
            expected["objectName"] = "First";
        }
        var call = await fixture.CallToolAsync("drawing", args,
            RecordingToolTest.Success("""{"success":true,"drawingObjects":[{"name":"Actual"}]}"""),
            $"drawing.{action}", JsonSerializer.Serialize(expected, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_AdvertisesLayoutNamesAndNativeEnumChoices()
    {
        var tools = await fixture.ListToolsAsync();
        var drawing = Assert.Single(tools, tool => tool.Name == "drawing");
        var properties = drawing.JsonSchema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("object_names", out _));
        Assert.True(properties.TryGetProperty("group_name", out _));
        Assert.True(properties.TryGetProperty("z_order", out _));
        Assert.Contains("group-objects", properties.GetProperty("action").GetProperty("enum").EnumerateArray().Select(item => item.GetString()));
        Assert.Contains("BringToFront", properties.GetProperty("z_order").ToString(), StringComparison.Ordinal);
        Assert.Contains("group", drawing.Description, StringComparison.OrdinalIgnoreCase);
    }
}
