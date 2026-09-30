using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
[Trait("RequiresExcel", "false")]
public sealed class StandardToolRegistrationTests(ITestOutputHelper output)
    : McpIntegrationTestBase(output, "StandardRegistrationClient")
{
    [Fact]
    public async Task Discovery_PreservesToolsWithoutGeneratedGuidePrompts()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        Assert.Equal(McpToolSurface.ToolCount, tools.Count);
        Assert.Null(Client.ServerCapabilities.Prompts);
    }

    [Theory]
    [InlineData("range", "values")]
    [InlineData("table", "rows")]
    public async Task NativeCellValues_AreNotAdvertisedAsStringsOnly(string tool, string parameter)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var schema = tools.Single(t => t.Name == tool).JsonSchema;
        var cellSchema = schema.GetProperty("properties").GetProperty(parameter)
            .GetProperty("items").GetProperty("items");
        Assert.False(cellSchema.TryGetProperty("type", out var type)
            && type.ValueKind == JsonValueKind.String && type.GetString() == "string");
    }

    [Fact]
    public async Task InjectedParameters_AreNotAdvertisedAsInputs()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        foreach (var tool in tools)
        {
            var properties = tool.JsonSchema.GetProperty("properties");
            foreach (var name in new[] { "bridge", "cancellationToken", "progress", "requestContext" })
            {
                Assert.False(properties.TryGetProperty(name, out _), $"{tool.Name} advertises {name}.");
            }
        }
    }

    [Fact]
    public async Task EveryTool_AdvertisesStructuredOutputSchema()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);

        foreach (var tool in tools)
        {
            var schema = Assert.IsType<JsonElement>(tool.ReturnJsonSchema);
            Assert.Equal("object", schema.GetProperty("type").GetString());
            var properties = schema.GetProperty("properties");
            AssertSchemaAllowsType(properties.GetProperty("success"), "boolean");
        }
    }

    [Theory]
    [InlineData("range", "values")]
    [InlineData("range", "rowCount")]
    [InlineData("table", "tables")]
    [InlineData("worksheet", "worksheets")]
    [InlineData("file", "session_id")]
    [InlineData("screenshot", "mimeType")]
    public async Task OutputSchemas_DescribeActionSpecificFields(string toolName, string propertyName)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var schema = Assert.IsType<JsonElement>(tools.Single(t => t.Name == toolName).ReturnJsonSchema);

        Assert.True(
            schema.GetProperty("properties").TryGetProperty(propertyName, out _),
            $"{toolName} output schema does not describe '{propertyName}'.");
    }

    private static void AssertSchemaAllowsType(JsonElement schema, string expectedType)
    {
        var type = schema.GetProperty("type");
        var allowsType = type.ValueKind == JsonValueKind.String
            ? type.GetString() == expectedType
            : type.EnumerateArray().Any(item => item.GetString() == expectedType);
        Assert.True(allowsType, $"Schema does not allow '{expectedType}'.");
    }
}
