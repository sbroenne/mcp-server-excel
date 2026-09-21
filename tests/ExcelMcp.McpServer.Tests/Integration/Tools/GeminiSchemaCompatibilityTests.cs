using System.Text.Json;
using ModelContextProtocol;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Protects the schema shape used with Gemini client adapters, not a universal Gemini restriction.
/// In particular, @ai-sdk/google 3.0.73 rewrites type arrays to anyOf with sibling items (#672).
/// Keep scalar type keywords and typed items while preserving nullable annotations and the
/// actual mixed scalar cell contract. These tests do not make live Gemini requests.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
[Trait("RequiresExcel", "false")]
public sealed class GeminiSchemaCompatibilityTests : McpIntegrationTestBase
{
    public GeminiSchemaCompatibilityTests(ITestOutputHelper output)
        : base(output, "GeminiSchemaCompatibilityClient")
    {
    }

    [Theory]
    [InlineData("range", "values")]
    [InlineData("range", "formulas")]
    [InlineData("range", "formats")]
    [InlineData("table", "rows")]
    public async Task ListTools_OptionalNestedArrays_PreserveNullableKeyword(string toolName, string parameterName)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var schema = Assert.Single(tools, tool => tool.Name == toolName).JsonSchema;
        var parameter = schema.GetProperty("properties").GetProperty(parameterName);

        Assert.Equal("array", parameter.GetProperty("type").GetString());
        Assert.True(parameter.TryGetProperty("nullable", out var nullable),
            $"{toolName}.{parameterName} must retain the SDK's nullable annotation.");
        Assert.True(nullable.GetBoolean());
        Assert.Equal("array", parameter.GetProperty("items").GetProperty("type").GetString());
        Assert.DoesNotContain(schema.GetProperty("required").EnumerateArray(),
            required => required.GetString() == parameterName);
        Assert.Contains(schema.GetProperty("required").EnumerateArray(),
            required => required.GetString() == "action");
        Assert.Contains(schema.GetProperty("required").EnumerateArray(),
            required => required.GetString() == "session_id");
        Assert.Equal("string", schema.GetProperty("properties").GetProperty("action").GetProperty("type").GetString());
    }

    [Theory]
    [InlineData("range", "values", true)]
    [InlineData("table", "rows", true)]
    [InlineData("analysis", "values", false)]
    public async Task ListTools_ObjectCollectionItems_DescribeMixedScalarsAndNull(
        string toolName, string parameterName, bool nested)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var parameter = Assert.Single(tools, tool => tool.Name == toolName).JsonSchema
            .GetProperty("properties").GetProperty(parameterName);
        var item = parameter.GetProperty("items");
        if (nested)
        {
            item = item.GetProperty("items");
        }

        Assert.True(item.TryGetProperty("anyOf", out var alternatives),
            $"{toolName}.{parameterName} must advertise scalar alternatives, not string coercion.");
        Assert.Collection(alternatives.EnumerateArray(),
            branch =>
            {
                Assert.Equal("string", branch.GetProperty("type").GetString());
                Assert.True(branch.GetProperty("nullable").GetBoolean());
            },
            branch => Assert.Equal("number", branch.GetProperty("type").GetString()),
            branch => Assert.Equal("boolean", branch.GetProperty("type").GetString()),
            branch => Assert.Equal("null", branch.GetProperty("type").GetString()));
        Assert.False(item.TryGetProperty("type", out _));
        Assert.False(item.TryGetProperty("description", out _));
    }

    [Theory]
    [InlineData("formulas")]
    [InlineData("formats")]
    public async Task ListTools_StringMatrices_KeepStringItems(string parameterName)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var cell = Assert.Single(tools, tool => tool.Name == "range").JsonSchema
            .GetProperty("properties").GetProperty(parameterName).GetProperty("items").GetProperty("items");

        Assert.Equal("string", cell.GetProperty("type").GetString());
        Assert.False(cell.TryGetProperty("anyOf", out _));
    }

    [Fact]
    public void CellValues_MixedScalars_DeserializeWithoutStringCoercion()
    {
        var values = JsonSerializer.Deserialize<List<List<object?>>>(
            """[["text",42,2.5,true,false,null]]""", McpJsonUtilities.DefaultOptions);

        var row = Assert.Single(Assert.IsType<List<List<object?>>>(values));
        Assert.Collection(row,
            value => Assert.Equal("text", RangeHelpers.ConvertToCellValue(value)),
            value => Assert.Equal(42d, Assert.IsType<double>(RangeHelpers.ConvertToCellValue(value))),
            value => Assert.Equal(2.5d, Assert.IsType<double>(RangeHelpers.ConvertToCellValue(value))),
            value => Assert.True(Assert.IsType<bool>(RangeHelpers.ConvertToCellValue(value))),
            value => Assert.False(Assert.IsType<bool>(RangeHelpers.ConvertToCellValue(value))),
            value =>
            {
                Assert.Null(value);
                // Null binds as null, then the existing cell converter represents a blank cell.
                Assert.Equal(string.Empty, RangeHelpers.ConvertToCellValue(value));
            });
    }

    [Theory]
    [InlineData("range", "set-values", "values")]
    [InlineData("table", "append", "rows")]
    public async Task CallTool_MixedScalarCells_BindBeforeSessionLookup(
        string toolName, string action, string parameterName)
    {
        using var payload = JsonDocument.Parse("""[["text",42,2.5,true,false,null]]""");
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = action,
            ["session_id"] = "synthetic-unknown-session",
            [parameterName] = payload.RootElement
        };
        if (toolName == "range")
        {
            arguments["sheet_name"] = "Sheet1";
            arguments["range_address"] = "A1:F1";
        }
        else
        {
            arguments["table_name"] = "Table1";
        }
        var json = await CallToolAsync(toolName, arguments, TimeSpan.FromSeconds(30));

        using var response = JsonDocument.Parse(json);
        Assert.False(response.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("not found", response.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ListTools_AllToolSchemas_AvoidAdapterUnionTypesAndUntypedItems()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        Assert.NotEmpty(tools);

        var violations = new List<string>();

        foreach (var tool in tools)
        {
            CollectViolations(tool.Name, tool.JsonSchema, "$", violations);
        }

        Output.WriteLine($"Scanned {tools.Count} tools for adapter-sensitive schema constructs.");

        Assert.True(
            violations.Count == 0,
            "Schemas must retain scalar type keywords and typed items for Gemini adapters. Violations:" +
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
                // Avoid adapter rewrites of type arrays that separate arrays from their items.
                if (node.TryGetProperty("type", out var typeProperty))
                {
                    if (typeProperty.ValueKind == JsonValueKind.Array)
                    {
                        violations.Add($"{toolName} {jsonPath}.type = {typeProperty.GetRawText()} (Must be scalar string)");
                    }
                }

                if (typeProperty.ValueKind == JsonValueKind.String && typeProperty.GetString() == "array")
                {
                    if (!node.TryGetProperty("items", out _))
                    {
                        violations.Add($"{toolName} {jsonPath} is an array without items");
                    }
                }

                // Keep item types explicit, including alternatives for mixed scalar cells.
                if (jsonPath.EndsWith(".items", StringComparison.Ordinal) || jsonPath.EndsWith("]items", StringComparison.Ordinal))
                {
                    if (!node.TryGetProperty("type", out _) && !node.TryGetProperty("anyOf", out _) && !node.TryGetProperty("enum", out _))
                    {
                        violations.Add($"{toolName} {jsonPath} is empty/untyped");
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
