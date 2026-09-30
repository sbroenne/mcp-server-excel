using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "GeneratedContracts")]
[Trait("RequiresExcel", "false")]
public sealed class ToolContractSnapshotTests
{
    [Fact]
    public void Normalize_IgnoresWhitespaceAndObjectPropertyOrder()
    {
        using var expected = JsonDocument.Parse("""{"name":"range","inputSchema":{"type":"object","properties":{"values":{"type":"array"}}}}""");
        using var actual = JsonDocument.Parse("""
            {
                "inputSchema": { "properties": { "values": { "type": "array" } }, "type": "object" },
                "name": "range"
            }
            """);

        Assert.Equal(ToolContractSnapshot.Normalize(expected.RootElement), ToolContractSnapshot.Normalize(actual.RootElement));
    }

    [Theory]
    [InlineData("""[{"name":"range"}]""", """[]""")]
    [InlineData("""{"name":"range"}""", """{"name":"range_edit"}""")]
    [InlineData("""{"description":"Read values"}""", """{"description":"Write values"}""")]
    [InlineData("""{"inputSchema":{"properties":{"values":{"type":"array"}}}}""", """{"inputSchema":{"properties":{"rows":{"type":"array"}}}}""")]
    [InlineData("""{"inputSchema":{"required":["action"]}}""", """{"inputSchema":{"required":["action","values"]}}""")]
    [InlineData("""{"inputSchema":{"properties":{"save":{"default":false}}}}""", """{"inputSchema":{"properties":{"save":{"default":true}}}}""")]
    [InlineData("""{"inputSchema":{"properties":{"action":{"enum":["read","write"]}}}}""", """{"inputSchema":{"properties":{"action":{"enum":["read"]}}}}""")]
    [InlineData("""{"inputSchema":{"properties":{"values":{"items":{"type":"array","items":{"type":"number"}}}}}}""", """{"inputSchema":{"properties":{"values":{"items":{"type":"array","items":{"type":"string"}}}}}}""")]
    [InlineData("""{"inputSchema":{"properties":{"values":{"type":["array","null"]}}}}""", """{"inputSchema":{"properties":{"values":{"type":"array"}}}}""")]
    [InlineData("""{"outputSchema":{"properties":{"success":{"type":"boolean"}}}}""", """{"outputSchema":{"properties":{"success":{"type":"string"}}}}""")]
    public void Normalize_PreservesPublishedContractChanges(string expectedJson, string actualJson)
    {
        using var expected = JsonDocument.Parse(expectedJson);
        using var actual = JsonDocument.Parse(actualJson);

        Assert.NotEqual(ToolContractSnapshot.Normalize(expected.RootElement), ToolContractSnapshot.Normalize(actual.RootElement));
    }
}
