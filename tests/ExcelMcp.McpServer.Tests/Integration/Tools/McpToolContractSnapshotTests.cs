using System.Text;
using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "GeneratedContracts")]
[Trait("RequiresExcel", "false")]
public sealed class McpToolContractSnapshotTests(ITestOutputHelper output)
    : McpIntegrationTestBase(output, "ToolContractSnapshotClient")
{
    [Fact]
    public async Task ListTools_PublishedContracts_MatchReviewedSnapshot()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        Assert.NotEmpty(tools);
        Assert.Equal(tools.Count, tools.Select(tool => tool.Name).Distinct(StringComparer.Ordinal).Count());

        var contracts = JsonSerializer.SerializeToElement(
            tools.OrderBy(tool => tool.Name, StringComparer.Ordinal).Select(tool => new
            {
                name = tool.Name,
                description = tool.Description,
                inputSchema = tool.JsonSchema,
                outputSchema = Assert.IsType<JsonElement>(tool.ReturnJsonSchema)
            }));
        var actual = ToolContractSnapshot.Normalize(contracts);
        var snapshotDirectory = Path.Combine(AppContext.BaseDirectory, "TestData");
        var expectedPath = Path.Combine(snapshotDirectory, "mcp-tool-contract.json");
        var actualPath = Path.Combine(snapshotDirectory, "mcp-tool-contract.actual.json");
        Directory.CreateDirectory(snapshotDirectory);
        await File.WriteAllTextAsync(actualPath, actual, TestCancellationToken);
        Output.WriteLine($"Published contracts for {tools.Count} tools: {actualPath}");

        Assert.True(File.Exists(expectedPath),
            $"Reviewed MCP contract snapshot is missing: {expectedPath}. Actual contracts: {actualPath}. " +
            "Review the generated file before explicitly copying it to the source snapshot.");
        using var expected = JsonDocument.Parse(await File.ReadAllTextAsync(expectedPath, TestCancellationToken));
        Assert.True(
            string.Equals(ToolContractSnapshot.Normalize(expected.RootElement), actual, StringComparison.Ordinal),
            $"Published MCP contracts changed. Compare {expectedPath} with {actualPath}. " +
            "Review tool names, descriptions, input/output schemas, defaults, and required fields before updating " +
            "tests\\ExcelMcp.McpServer.Tests\\TestData\\mcp-tool-contract.json.");
    }
}

internal static class ToolContractSnapshot
{
    internal static string Normalize(JsonElement contracts)
    {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true }))
        {
            WriteSorted(writer, contracts);
        }

        return Encoding.UTF8.GetString(stream.ToArray()) + "\n";
    }

    private static void WriteSorted(Utf8JsonWriter writer, JsonElement element)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                writer.WriteStartObject();
                foreach (var property in element.EnumerateObject().OrderBy(property => property.Name, StringComparer.Ordinal))
                {
                    writer.WritePropertyName(property.Name);
                    WriteSorted(writer, property.Value);
                }
                writer.WriteEndObject();
                break;
            case JsonValueKind.Array:
                writer.WriteStartArray();
                foreach (var item in element.EnumerateArray())
                {
                    WriteSorted(writer, item);
                }
                writer.WriteEndArray();
                break;
            default:
                element.WriteTo(writer);
                break;
        }
    }
}
