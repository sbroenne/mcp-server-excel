using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// MCP arguments and output for Power Query privacy settings and load-to.
/// Workbook behavior for these actions is covered by the Service tests.
/// </summary>
[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "false")]
public sealed class PowerQueryPrivacyAndLoadProtocolTests(
    RecordingProgramTransportFixture fixture)
{
    private const string SessionId = "recording-session";

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task PowerQueryPrivacySettings_MapSetAndGet(bool ignorePrivacyLevels)
    {
        var flag = ignorePrivacyLevels ? "true" : "false";
        var resultJson = $$"""{"success":true,"ignorePrivacyLevels":{{flag}},"message":"global privacy option may override"}""";

        var set = await fixture.CallToolAsync(
            "powerquery",
            new Dictionary<string, object?>
            {
                ["action"] = "set-privacy-settings",
                ["workbook_session_id"] = SessionId,
                ["ignore_privacy_levels"] = ignorePrivacyLevels
            },
            RecordingToolTest.Success(resultJson),
            "powerquery.set-privacy-settings",
            $$"""{"ignorePrivacyLevels":{{flag}}}""");
        var get = await fixture.CallToolAsync(
            "powerquery_read",
            new Dictionary<string, object?>
            {
                ["action"] = "get-privacy-settings",
                ["workbook_session_id"] = SessionId
            },
            RecordingToolTest.Success(resultJson),
            "powerquery.get-privacy-settings",
            null);

        foreach (var output in new[] { set.JsonResult, get.JsonResult })
        {
            using var document = JsonDocument.Parse(output);
            Assert.Equal(ignorePrivacyLevels, document.RootElement.GetProperty("ignorePrivacyLevels").GetBoolean());
        }
    }

    [Fact]
    public async Task PowerQueryLoadTo_DataModelDestinationReachesService()
    {
        var call = await fixture.CallToolAsync(
            "powerquery",
            new Dictionary<string, object?>
            {
                ["action"] = "load-to",
                ["workbook_session_id"] = SessionId,
                ["query_name"] = "CsvData",
                ["load_destination"] = "load-to-data-model"
            },
            RecordingToolTest.Success("""{"success":true,"queryName":"CsvData","durationMs":18.5,"rowsLoaded":7,"modelRowsLoaded":7}"""),
            "powerquery.load-to",
            """{"queryName":"CsvData","loadDestination":"load-to-data-model"}""");

        using var document = JsonDocument.Parse(call.JsonResult);
        Assert.Equal(18.5, document.RootElement.GetProperty("durationMs").GetDouble());
        Assert.Equal(7, document.RootElement.GetProperty("rowsLoaded").GetInt64());
        Assert.Equal(7, document.RootElement.GetProperty("modelRowsLoaded").GetInt64());
    }
}
