using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "XmlMap")]
[Trait("RequiresExcel", "false")]
public sealed class XmlMapToolProtocolTests(
    RecordingProgramTransportFixture fixture)
{
    private const string SessionId = "recording-session";
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task ImportExportDelete_ThroughMcp_RoundTripsXmlData()
    {
        const string xmlData = """
            <customers>
              <customer><name>Ada</name><score>42</score></customer>
              <customer><name>Grace</name><score>99</score></customer>
            </customers>
            """;
        const string mapName = "CustomersMap";

        var importJson = await CallAsync(
            new()
            {
                ["action"] = "import-xml",
                ["workbook_session_id"] = SessionId,
                ["xml_data"] = xmlData,
                ["sheet_name"] = "Sheet1",
                ["start_cell"] = "B2"
            },
            "xmlmap.import-xml",
            $$"""{"xmlData":{{JsonSerializer.Serialize(xmlData)}},"sheetName":"Sheet1","startCell":"B2"}""",
            $$"""{"success":true,"mapName":"{{mapName}}"}""",
            args =>
            {
                Assert.Equal(xmlData, args.GetProperty("xmlData").GetString());
                Assert.Equal("Sheet1", args.GetProperty("sheetName").GetString());
                Assert.Equal("B2", args.GetProperty("startCell").GetString());
            });
        using (var import = JsonDocument.Parse(importJson))
        {
            Assert.Equal(
                mapName,
                import.RootElement.GetProperty("mapName").GetString());
        }

        var exportJson = await CallAsync(
            new()
            {
                ["action"] = "export-xml",
                ["workbook_session_id"] = SessionId,
                ["map_name"] = mapName
            },
            "xmlmap.export-xml",
            """{"mapName":"CustomersMap"}""",
            $$"""{"success":true,"xmlData":{{JsonSerializer.Serialize(xmlData)}}}""",
            args => Assert.Equal(
                mapName,
                args.GetProperty("mapName").GetString()));
        using (var export = JsonDocument.Parse(exportJson))
        {
            var exportedXml = export.RootElement.GetProperty("xmlData").GetString();
            Assert.Contains("Ada", exportedXml, StringComparison.Ordinal);
            Assert.Contains("Grace", exportedXml, StringComparison.Ordinal);
        }

        await CallAsync(
            new()
            {
                ["action"] = "delete",
                ["workbook_session_id"] = SessionId,
                ["map_name"] = mapName
            },
            "xmlmap.delete",
            """{"mapName":"CustomersMap"}""",
            """{"success":true}""",
            args => Assert.Equal(
                mapName,
                args.GetProperty("mapName").GetString()));

        var listJson = await CallAsync(
            new()
            {
                ["action"] = "list",
                ["workbook_session_id"] = SessionId
            },
            "xmlmap.list",
            null,
            """{"success":true,"maps":[]}""");
        using var list = JsonDocument.Parse(listJson);
        Assert.Empty(list.RootElement.GetProperty("maps").EnumerateArray());
    }

    private async Task<string> CallAsync(
        Dictionary<string, object?> arguments,
        string command,
        string? expectedArgsJson,
        string responseJson,
        Action<JsonElement>? assertArgs = null)
    {
        var call = await _fixture.CallToolAsync(
            arguments["action"] is "list" or "export-xml" ? "xmlmap_read" : "xmlmap",
            arguments,
            RecordingToolTest.Success(responseJson),
            command,
            expectedArgsJson);
        Assert.Equal(command, call.Request.Command);
        Assert.Equal(SessionId, call.Request.SessionId);
        if (assertArgs is not null)
        {
            Assert.NotNull(call.Request.Args);
            using var args = JsonDocument.Parse(call.Request.Args);
            assertArgs(args.RootElement);
        }
        return call.JsonResult;
    }
}
