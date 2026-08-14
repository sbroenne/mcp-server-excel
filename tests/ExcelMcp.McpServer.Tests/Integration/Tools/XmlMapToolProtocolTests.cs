using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// End-to-end XML map tool coverage through the MCP protocol.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "XmlMap")]
[Trait("RequiresExcel", "true")]
public sealed class XmlMapToolProtocolTests : McpIntegrationTestBase
{
    private readonly string _tempDir;

    public XmlMapToolProtocolTests(ITestOutputHelper output)
        : base(output, "XmlMapToolProtocolClient")
    {
        _tempDir = CreateTempDirectory("XmlMapToolProtocolTests");
    }

    [Fact]
    public async Task ListTools_XmlMapSchema_ExposesCompleteActionAndParameterContract()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var tool = Assert.Single(tools, candidate => candidate.Name == "xmlmap");
        var properties = tool.JsonSchema.GetProperty("properties");

        var actionValues = properties.GetProperty("action").GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString() ?? string.Empty)
            .ToArray();

        Assert.Equal(
            ["list", "add", "map-range", "import-xml", "export-xml", "delete"],
            actionValues);
        Assert.True(properties.TryGetProperty("schema", out _));
        Assert.True(properties.TryGetProperty("schema_file", out _));
        Assert.True(properties.TryGetProperty("xml_data", out _));
        Assert.True(properties.TryGetProperty("xml_data_file", out _));
        Assert.True(properties.TryGetProperty("map_name", out _));
        Assert.True(properties.TryGetProperty("sheet_name", out _));
        Assert.True(properties.TryGetProperty("range_address", out _));
        Assert.True(properties.TryGetProperty("xpath", out _));
        Assert.True(properties.TryGetProperty("start_cell", out _));
        Assert.True(properties.TryGetProperty("overwrite", out _));
    }

    [Fact]
    public async Task ImportExportDelete_ThroughMcp_RoundTripsXmlData()
    {
        const string xmlData = """
            <customers>
              <customer><name>Ada</name><score>42</score></customer>
              <customer><name>Grace</name><score>99</score></customer>
            </customers>
            """;
        var workbookPath = Path.Join(_tempDir, $"XmlMap_{Guid.NewGuid():N}.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);

        var importJson = await CallToolAsync("xmlmap", new Dictionary<string, object?>
        {
            ["action"] = "import-xml",
            ["session_id"] = sessionId,
            ["xml_data"] = xmlData,
            ["sheet_name"] = "Sheet1",
            ["start_cell"] = "B2"
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(importJson, "xmlmap import-xml");
        using var importDocument = JsonDocument.Parse(importJson);
        var mapName = importDocument.RootElement.GetProperty("mapName").GetString();
        Assert.False(string.IsNullOrWhiteSpace(mapName));

        var exportJson = await CallToolAsync("xmlmap", new Dictionary<string, object?>
        {
            ["action"] = "export-xml",
            ["session_id"] = sessionId,
            ["map_name"] = mapName
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(exportJson, "xmlmap export-xml");
        using var exportDocument = JsonDocument.Parse(exportJson);
        var exportedXml = exportDocument.RootElement.GetProperty("xmlData").GetString();
        Assert.Contains("Ada", exportedXml, StringComparison.Ordinal);
        Assert.Contains("Grace", exportedXml, StringComparison.Ordinal);

        var deleteJson = await CallToolAsync("xmlmap", new Dictionary<string, object?>
        {
            ["action"] = "delete",
            ["session_id"] = sessionId,
            ["map_name"] = mapName
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(deleteJson, "xmlmap delete");

        var listJson = await CallToolAsync("xmlmap", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(listJson, "xmlmap list");
        using var listDocument = JsonDocument.Parse(listJson);
        Assert.Empty(listDocument.RootElement.GetProperty("maps").EnumerateArray());

        await CloseSessionAsync(sessionId, save: false);
    }
}
