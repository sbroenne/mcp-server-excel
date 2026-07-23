using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Guards against invalid empty-string enum sentinels in generated MCP tool schemas.
/// Strict clients like Gemini reject schemas when any enum member is "".
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
public sealed class GeneratedToolSchemaEnumRegressionTests : McpIntegrationTestBase
{
    public GeneratedToolSchemaEnumRegressionTests(ITestOutputHelper output)
        : base(output, "GeneratedToolSchemaEnumRegressionClient")
    {
    }

    [Fact]
    public async Task ListTools_AllToolSchemas_ExcludeEmptyEnumValues()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var enumSurfaces = new List<EnumSurface>();

        foreach (var tool in tools)
        {
            CollectEnumSurfaces(tool.Name, tool.JsonSchema, "$", propertyName: null, enumSurfaces);
        }

        Assert.NotEmpty(enumSurfaces);

        var distinctPropertyNames = enumSurfaces
            .Select(surface => surface.PropertyName)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToList();

        Assert.Contains("action", distinctPropertyNames);
        Assert.Contains(distinctPropertyNames, name => !string.Equals(name, "action", StringComparison.Ordinal));

        var failures = enumSurfaces
            .SelectMany(surface => surface.Values.Select((value, index) => new
            {
                surface.ToolName,
                surface.SchemaPath,
                surface.PropertyName,
                ValueIndex = index,
                value.Kind,
                value.StringValue,
                value.RawText
            }))
            .Where(item => item.Kind != JsonValueKind.String || string.IsNullOrWhiteSpace(item.StringValue))
            .Select(item => $"{item.ToolName} {item.SchemaPath}.enum[{item.ValueIndex}] ({item.PropertyName ?? "<unknown>"}) = {item.RawText}")
            .ToList();

        Output.WriteLine($"Scanned {tools.Count} tools, {enumSurfaces.Count} enum surfaces, and {enumSurfaces.Sum(surface => surface.Values.Count)} enum values.");
        Output.WriteLine($"Enum-bearing properties: {string.Join(", ", distinctPropertyNames)}");

        Assert.True(
            failures.Count == 0,
            "MCP schemas must not publish empty-string enum members. " + string.Join(Environment.NewLine, failures));
    }

    [Fact]
    public async Task ListTools_CalculationModeSchema_KeepsOptionalActionSpecificEnumsOptionalWithoutEnumSentinels()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var calculationTool = tools.Single(tool => tool.Name == "calculation_mode");

        var properties = calculationTool.JsonSchema.GetProperty("properties");
        var required = GetRequiredPropertyNames(calculationTool.JsonSchema);

        AssertOptionalStringPropertyWithoutEnumSentinel(properties, required, "mode");
        AssertOptionalStringPropertyWithoutEnumSentinel(properties, required, "scope");
    }

    private static void AssertOptionalStringPropertyWithoutEnumSentinel(
        JsonElement properties,
        HashSet<string> required,
        string propertyName)
    {
        Assert.True(properties.TryGetProperty(propertyName, out var property), $"Schema is missing {propertyName}: {properties.GetRawText()}");
        Assert.False(required.Contains(propertyName), $"{propertyName} should stay optional: {string.Join(", ", required)}");
        var typeProperty = property.GetProperty("type");
        Assert.True(
            TypeIncludesString(typeProperty),
            $"{propertyName} should be string-compatible in schema: {property.GetRawText()}");
        Assert.False(property.TryGetProperty("enum", out _), $"{propertyName} should not publish an enum sentinel surface: {property.GetRawText()}");
    }

    private static bool TypeIncludesString(JsonElement typeProperty)
    {
        return typeProperty.ValueKind switch
        {
            JsonValueKind.String => typeProperty.GetString() == "string",
            JsonValueKind.Array => typeProperty.EnumerateArray().Any(value => value.ValueKind == JsonValueKind.String && value.GetString() == "string"),
            _ => false
        };
    }

    private static HashSet<string> GetRequiredPropertyNames(JsonElement schema)
    {
        if (!schema.TryGetProperty("required", out var required))
        {
            return [];
        }

        return new HashSet<string>(
            required.EnumerateArray()
                .Select(value => value.GetString())
                .Where(static value => !string.IsNullOrWhiteSpace(value))
                .Select(static value => value!),
            StringComparer.Ordinal);
    }

    private static void CollectEnumSurfaces(
        string toolName,
        JsonElement node,
        string jsonPath,
        string? propertyName,
        List<EnumSurface> enumSurfaces)
    {
        switch (node.ValueKind)
        {
            case JsonValueKind.Object:
                if (node.TryGetProperty("enum", out var enumProperty) && enumProperty.ValueKind == JsonValueKind.Array)
                {
                    enumSurfaces.Add(new EnumSurface(
                        toolName,
                        jsonPath,
                        propertyName,
                        [.. enumProperty.EnumerateArray().Select(ToEnumValue)]));
                }

                foreach (var property in node.EnumerateObject())
                {
                    CollectEnumSurfaces(toolName, property.Value, $"{jsonPath}.{property.Name}", property.Name, enumSurfaces);
                }

                break;

            case JsonValueKind.Array:
                var index = 0;
                foreach (var item in node.EnumerateArray())
                {
                    CollectEnumSurfaces(toolName, item, $"{jsonPath}[{index}]", propertyName, enumSurfaces);
                    index++;
                }

                break;
        }
    }

    private static EnumValue ToEnumValue(JsonElement element)
    {
        return element.ValueKind == JsonValueKind.String
            ? new EnumValue(element.ValueKind, element.GetString(), element.GetRawText())
            : new EnumValue(element.ValueKind, null, element.GetRawText());
    }

    private sealed record EnumSurface(string ToolName, string SchemaPath, string? PropertyName, IReadOnlyList<EnumValue> Values);

    private sealed record EnumValue(JsonValueKind Kind, string? StringValue, string RawText);
}
