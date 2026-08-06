using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Verifies generated mutating tools publish the shared review/checkpoint contract.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Safety")]
[Trait("RequiresExcel", "false")]
public sealed class SafetySurfaceSchemaTests : McpIntegrationTestBase
{
    public SafetySurfaceSchemaTests(ITestOutputHelper output)
        : base(output, "SafetySurfaceSchemaClient")
    {
    }

    [Fact]
    public async Task ListTools_MutatingToolSchema_ExposesReviewAndCheckpointOptions()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var rangeTool = tools.Single(tool => tool.Name == "range");
        var rangeSchema = rangeTool.JsonSchema;
        var rangeProperties = rangeSchema.GetProperty("properties");

        Output.WriteLine(rangeSchema.GetRawText());

        AssertSafetyOptions(rangeProperties);

        var worksheetTool = tools.Single(tool => tool.Name == "worksheet");
        var worksheetProperties = worksheetTool.JsonSchema.GetProperty("properties");
        AssertSafetyOptions(worksheetProperties);
    }

    private static void AssertSafetyOptions(JsonElement properties)
    {
        AssertSafetyOption(properties, "review_only", JsonValueKind.False);
        AssertSafetyOption(properties, "review_id", JsonValueKind.Null);
        AssertSafetyOption(properties, "checkpoint", JsonValueKind.False);
    }

    [Fact]
    public async Task ListTools_FileSchema_ExposesSafetyAndRecoveryActionsWithClosedOptions()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var fileTool = tools.Single(tool => tool.Name == "file");
        var properties = fileTool.JsonSchema.GetProperty("properties");

        var actions = properties.GetProperty("action").GetProperty("enum")
            .EnumerateArray().Select(value => value.GetString()).ToArray();
        Assert.Contains("configure-safety", actions);
        Assert.Contains("journal", actions);
        Assert.Contains("recoveries", actions);
        Assert.Contains("recover", actions);

        AssertClosedEnum(properties, "review_mode", "off", "optional", "required");
        AssertClosedEnum(properties, "checkpoint_mode", "off", "onRequest", "required");
        AssertClosedEnum(properties, "journal_mode", "off", "on");
        AssertClosedEnum(properties, "verification_mode", "off", "on");
        AssertClosedEnum(properties, "abnormal_shutdown_policy", "legacyAutoSave", "discardWithRecoveryEvidence");
        Assert.True(properties.TryGetProperty("recovery_id", out _));
    }

    private static void AssertSafetyOption(JsonElement properties, string name, JsonValueKind expectedDefaultKind)
    {
        Assert.True(properties.TryGetProperty(name, out var property), $"tool schema is missing {name}");
        Assert.True(property.TryGetProperty("description", out var description), $"{name} is missing a description");
        Assert.False(string.IsNullOrWhiteSpace(description.GetString()), $"{name} description is empty");

        if (property.TryGetProperty("default", out var defaultValue))
        {
            Assert.Equal(expectedDefaultKind, defaultValue.ValueKind);
        }
    }

    private static void AssertClosedEnum(JsonElement properties, string name, params string[] expected)
    {
        Assert.True(properties.TryGetProperty(name, out var property), $"file schema is missing {name}");
        var actual = property.GetProperty("enum").EnumerateArray()
            .Where(value => value.ValueKind == JsonValueKind.String)
            .Select(value => value.GetString()).ToArray();
        Assert.Equal(expected, actual);
    }
}
