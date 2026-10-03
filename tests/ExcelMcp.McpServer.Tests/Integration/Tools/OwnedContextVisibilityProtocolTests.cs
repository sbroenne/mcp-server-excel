using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Window")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class OwnedContextVisibilityProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task Context_UsesSessionWithoutAnApplicationSelectionInput()
    {
        var call = await fixture.CallToolAsync("window", new()
        {
            ["action"] = "get-context",
            ["session_id"] = "session-1"
        }, RecordingToolTest.Success("""{"success":true,"availability":"available","windows":[]}"""),
            "window.get-context", null);
        Assert.False(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.Equal("available", result.RootElement.GetProperty("availability").GetString());
    }

    [Theory]
    [InlineData("get-visibility", "rows", null)]
    [InlineData("set-visibility", "columns", true)]
    [InlineData("set-visibility", "rows", false)]
    public async Task Visibility_MapsExactScopeAndExplicitHiddenState(string action, string axis, bool? hidden)
    {
        Dictionary<string, object?> arguments = new()
        {
            ["action"] = action,
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A2,C4",
            ["axis"] = axis
        };
        Dictionary<string, object?> expected = new()
        {
            ["sheetName"] = "Sheet1",
            ["rangeAddress"] = "A2,C4",
            ["axis"] = axis
        };
        if (hidden.HasValue)
        {
            arguments["hidden"] = hidden.Value;
            expected["hidden"] = hidden.Value;
        }
        var call = await fixture.CallToolAsync("range_format", arguments,
            RecordingToolTest.Success("""{"success":true}"""), $"rangeformat.{action}",
            JsonSerializer.Serialize(expected, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Theory]
    [InlineData("get-visibility", "axis")]
    [InlineData("set-visibility", "axis")]
    [InlineData("set-visibility", "hidden")]
    public async Task Visibility_MissingRequiredInputDoesNotDispatch(string action, string missing)
    {
        Dictionary<string, object?> arguments = new()
        {
            ["action"] = action,
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A2",
            ["axis"] = "rows",
            ["hidden"] = false
        };
        arguments.Remove(missing);
        var result = await fixture.CallResultWithoutDispatchAsync("range_format", arguments);
        Assert.True(result.IsError);
    }

    [Fact]
    public async Task Discovery_ExplainsOwnedContextAndHiddenCauseLimits()
    {
        var tools = await fixture.ListToolsAsync();
        var window = Assert.Single(tools, tool => tool.Name == "window");
        Assert.Contains("get-context", window.Description, StringComparison.Ordinal);
        Assert.Contains("without activation or selection", window.Description, StringComparison.Ordinal);
        var format = Assert.Single(tools, tool => tool.Name == "range_format");
        Assert.Contains("hidden cause is undetermined", format.Description, StringComparison.Ordinal);
        Assert.Contains("Disjoint gaps remain unchanged", format.Description, StringComparison.Ordinal);
        Assert.True(format.JsonSchema.GetProperty("properties").TryGetProperty("axis", out _));
        Assert.True(format.JsonSchema.GetProperty("properties").TryGetProperty("hidden", out _));
    }
}
