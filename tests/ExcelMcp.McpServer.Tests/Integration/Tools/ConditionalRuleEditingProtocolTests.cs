using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "ConditionalRuleEditing")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ConditionalRuleEditingProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task Update_MapsSelectionAndNestedTypedOptions()
    {
        var options = new ConditionalRuleUpdateOptions { Formula1 = "15", StopIfTrue = false };
        var expected = JsonSerializer.Serialize(new
        {
            sheetName = "Data",
            rulePriority = 3,
            expectedFingerprint = "rule-fingerprint",
            options
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("conditionalformat", new()
        {
            ["action"] = "update-rule",
            ["session_id"] = "session-1",
            ["sheet_name"] = "Data",
            ["rule_priority"] = 3,
            ["expected_fingerprint"] = "rule-fingerprint",
            ["options"] = new { formula1 = "15", stopIfTrue = false }
        }, RecordingToolTest.Success("""{"success":true,"rules":[]}"""), "conditionalformat.update-rule", expected);
        Assert.False(call.Result.IsError);
    }

    [Theory]
    [InlineData("delete-rule")]
    [InlineData("set-rule-priority")]
    public async Task SelectedMutation_MapsCanonicalParameterNames(string action)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = action,
            ["session_id"] = "session-1",
            ["sheet_name"] = "Data",
            ["rule_priority"] = 3,
            ["expected_fingerprint"] = "rule-fingerprint"
        };
        var expectedArgs = new Dictionary<string, object?>
        {
            ["sheetName"] = "Data",
            ["rulePriority"] = 3,
            ["expectedFingerprint"] = "rule-fingerprint"
        };
        if (action == "set-rule-priority")
        {
            arguments["new_priority"] = 1;
            expectedArgs["newPriority"] = 1;
        }
        var call = await fixture.CallToolAsync("conditionalformat", arguments,
            RecordingToolTest.Success("""{"success":true,"rules":[]}"""),
            $"conditionalformat.{action}", JsonSerializer.Serialize(expectedArgs, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_ExposesStaleGuardAndCreationControls()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "conditionalformat");
        Assert.Contains("fingerprint", tool.Description, StringComparison.Ordinal);
        var properties = tool.JsonSchema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("rule_priority", out _));
        Assert.True(properties.TryGetProperty("expected_fingerprint", out _));
        Assert.True(properties.TryGetProperty("options", out _));
        Assert.True(properties.TryGetProperty("priority", out _));
        Assert.True(properties.TryGetProperty("stop_if_true", out _));
    }
}
