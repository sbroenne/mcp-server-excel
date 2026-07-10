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
    public async Task ListTools_AllToolSchemas_MeetGeminiStrictConstraints()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        Assert.NotEmpty(tools);

        var violations = new List<string>();

        foreach (var tool in tools)
        {
            CollectViolations(tool.Name, tool.JsonSchema, "$", violations);
        }

        Output.WriteLine($"Scanned {tools.Count} tools for Gemini-incompatible schema constructs.");

        Assert.True(
            violations.Count == 0,
            "Gemini (OpenAPI 3.0 subset) enforces strict schema rules. Violations:" +
            Environment.NewLine + string.Join(Environment.NewLine, violations));
    }

    private static void CollectViolations(
        string toolName,
        JsonElement node,
        string jsonPath,
        List<string> violations)
    {
        switch (node.ValueKind)
        {
            case JsonValueKind.Object:
                // Rule 1: No union types. type must be a string.
                if (node.TryGetProperty("type", out var typeProperty))
                {
                    if (typeProperty.ValueKind == JsonValueKind.Array)
                    {
                        violations.Add($"{toolName} {jsonPath}.type = {typeProperty.GetRawText()} (Must be scalar string)");
                    }
                }

                // Rule 2: No `nullable: true` on array nodes. 
                // Client translates this to `anyOf` which breaks Gemini's `items` predicate validator.
                if (typeProperty.ValueKind == JsonValueKind.String && typeProperty.GetString() == "array")
                {
                    if (node.TryGetProperty("nullable", out var nullableProperty) && nullableProperty.GetBoolean())
                    {
                        violations.Add($"{toolName} {jsonPath} has both type:\"array\" and nullable:true (Gemini rejects due to anyOf translation)");
                    }
                }

                // Rule 3: items schema must not be empty. It must have a type.
                if (jsonPath.EndsWith(".items", StringComparison.Ordinal) || jsonPath.EndsWith("]items", StringComparison.Ordinal))
                {
                    if (!node.TryGetProperty("type", out _) && !node.TryGetProperty("anyOf", out _) && !node.TryGetProperty("enum", out _))
                    {
                        violations.Add($"{toolName} {jsonPath} is empty/untyped (Gemini requires explicit type on all items)");
                    }
                }

                foreach (var property in node.EnumerateObject())
                {
                    CollectViolations(toolName, property.Value, $"{jsonPath}.{property.Name}", violations);
                }

                break;

            case JsonValueKind.Array:
                var index = 0;
                foreach (var item in node.EnumerateArray())
                {
                    CollectViolations(toolName, item, $"{jsonPath}[{index}]", violations);
                    index++;
                }

                break;
        }
    }
}
