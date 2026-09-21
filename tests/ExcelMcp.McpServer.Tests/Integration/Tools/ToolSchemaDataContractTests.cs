using System.Text.Json;
using ModelContextProtocol.Protocol;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
[Trait("RequiresExcel", "false")]
public sealed class ToolSchemaDataContractTests(ITestOutputHelper output)
    : McpIntegrationTestBase(output, "ToolSchemaDataContractClient")
{
    [Theory]
    [InlineData("range", "values")]
    [InlineData("table", "rows")]
    public async Task ListTools_MixedCellValues_AreNotRestrictedToStrings(string toolName, string parameterName)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var schema = Assert.Single(tools, tool => tool.Name == toolName).JsonSchema;
        var cellSchema = schema.GetProperty("properties").GetProperty(parameterName)
            .GetProperty("items").GetProperty("items");

        // object? cells accept numbers, booleans, strings and null without coercion.
        Assert.Equal(JsonValueKind.Object, cellSchema.ValueKind);
        Assert.Empty(cellSchema.EnumerateObject());
    }

    [Theory]
    [InlineData("range", "values")]
    [InlineData("range", "formulas")]
    [InlineData("table", "rows")]
    public async Task ListTools_OptionalMatrices_PreserveJsonSchemaNullability(string toolName, string parameterName)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var schema = Assert.Single(tools, tool => tool.Name == toolName).JsonSchema;
        var parameter = schema.GetProperty("properties").GetProperty(parameterName);
        var types = parameter.GetProperty("type");

        Assert.Equal(JsonValueKind.Array, types.ValueKind);
        Assert.Equal(["array", "null"], types.EnumerateArray().Select(type => type.GetString()));
        Assert.False(parameter.TryGetProperty("nullable", out _));
        Assert.DoesNotContain(schema.GetProperty("required").EnumerateArray(),
            required => required.GetString() == parameterName);
        Assert.True(parameter.TryGetProperty("items", out _));
    }

    [Fact]
    public async Task ListTools_DiscoveryMatchesAdvertisedToolsAndOperations()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);

        Assert.Equal(
            McpToolSurface.Tools.Select(tool => tool.Name).Order(StringComparer.Ordinal),
            tools.Select(tool => tool.Name).Order(StringComparer.Ordinal));
        foreach (var expected in McpToolSurface.Tools)
        {
            var tool = Assert.Single(tools, tool => tool.Name == expected.Name);
            var schema = tool.JsonSchema;
            Assert.Equal("object", schema.GetProperty("type").GetString());
            Assert.Contains(schema.GetProperty("required").EnumerateArray(),
                required => required.GetString() == "action");
            Assert.Equal(expected.OperationCount, schema.GetProperty("properties")
                .GetProperty("action").GetProperty("enum").GetArrayLength());
        }
    }

    [Theory]
    [InlineData("range", "set-values", "values", "[[42,3.5,true,false,\"text\",null]]")]
    [InlineData("table", "append", "rows", "[[42,3.5,true,false,\"text\",null]]")]
    [InlineData("range", "set-values", "values", "null")]
    [InlineData("range", "set-formulas", "formulas", "null")]
    [InlineData("table", "append", "rows", "null")]
    public async Task CallTool_MixedValuesAndNullMatrices_BindBeforeSessionLookup(
        string toolName, string action, string parameterName, string json)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = action,
            ["session_id"] = "synthetic-unknown-session",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1:F1",
            ["table_name"] = "Table1",
            [parameterName] = JsonSerializer.Deserialize<JsonElement>(json)
        };

        // A missing session proves binding reached our service without opening Excel.
        var response = await Client!.CallToolAsync(toolName, arguments,
            cancellationToken: TestCancellationToken);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var document = JsonDocument.Parse(text);
        Assert.True(response.IsError);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("not found", document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
    }
}
