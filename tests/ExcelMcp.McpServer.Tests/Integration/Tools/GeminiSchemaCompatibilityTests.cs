using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Guards against JSON Schema constructs that Google's Gemini function-calling API rejects.
///
/// Gemini accepts only a subset of the OpenAPI 3.0.3 Schema object
/// (https://ai.google.dev/gemini-api/docs/function-calling — "select subset of the OpenAPI schema").
/// OpenAPI 3.0 requires <c>type</c> to be a SINGLE scalar value and expresses nullability with a
/// separate <c>nullable: true</c> field. It does NOT allow the JSON Schema 2020-12 union form
/// <c>"type": ["array","null"]</c>.
///
/// The MCP SDK's default schema generator emits that union form for nullable .NET parameters
/// (e.g. <c>List&lt;List&lt;object?&gt;&gt;?</c> for range values/formulas/formats and table rows),
/// which Gemini rejects with HTTP 400:
///   parameters.properties[formulas].items: field predicate failed: == Type.ARRAY
///
/// This test walks every generated tool schema and asserts no <c>type</c> keyword is a JSON array,
/// which is the exact constraint Gemini enforces. See issue #672.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
public sealed class GeminiSchemaCompatibilityTests : McpIntegrationTestBase
{
    public GeminiSchemaCompatibilityTests(ITestOutputHelper output)
        : base(output, "GeminiSchemaCompatibilityClient")
    {
    }

    [Fact]
    public async Task ListTools_AllToolSchemas_DoNotUseUnionTypeArrays()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        Assert.NotEmpty(tools);

        var violations = new List<string>();

        foreach (var tool in tools)
        {
            CollectUnionTypeViolations(tool.Name, tool.JsonSchema, "$", violations);
        }

        Output.WriteLine($"Scanned {tools.Count} tools for Gemini-incompatible union 'type' arrays.");

        Assert.True(
            violations.Count == 0,
            "Gemini (OpenAPI 3.0 subset) rejects union 'type' arrays like [\"array\",\"null\"]. " +
            "Nullability must use 'nullable: true' with a single scalar 'type'. Violations:" +
            Environment.NewLine + string.Join(Environment.NewLine, violations));
    }

    private static void CollectUnionTypeViolations(
        string toolName,
        JsonElement node,
        string jsonPath,
        List<string> violations)
    {
        switch (node.ValueKind)
        {
            case JsonValueKind.Object:
                if (node.TryGetProperty("type", out var typeProperty) && typeProperty.ValueKind == JsonValueKind.Array)
                {
                    violations.Add($"{toolName} {jsonPath}.type = {typeProperty.GetRawText()}");
                }

                foreach (var property in node.EnumerateObject())
                {
                    CollectUnionTypeViolations(toolName, property.Value, $"{jsonPath}.{property.Name}", violations);
                }

                break;

            case JsonValueKind.Array:
                var index = 0;
                foreach (var item in node.EnumerateArray())
                {
                    CollectUnionTypeViolations(toolName, item, $"{jsonPath}[{index}]", violations);
                    index++;
                }

                break;
        }
    }
}
