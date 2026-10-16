using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// MCP arguments and output for actions whose workbook behavior is covered by
/// the Service tests (Power Query privacy and loading, DAX measures, and
/// cross-file worksheet copies).
/// </summary>
[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "false")]
public sealed class ModelAndTransferProtocolTests(
    RecordingProgramTransportFixture fixture)
{
    private const string SessionId = "recording-session";
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task PowerQueryPrivacySettings_MapSetAndGet(bool ignorePrivacyLevels)
    {
        var flag = ignorePrivacyLevels ? "true" : "false";
        var resultJson = $$"""{"success":true,"ignorePrivacyLevels":{{flag}},"message":"global privacy option may override"}""";

        var set = await _fixture.CallToolAsync(
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
        var get = await _fixture.CallToolAsync(
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
        var call = await _fixture.CallToolAsync(
            "powerquery",
            new Dictionary<string, object?>
            {
                ["action"] = "load-to",
                ["workbook_session_id"] = SessionId,
                ["query_name"] = "CsvData",
                ["load_destination"] = "load-to-data-model"
            },
            RecordingToolTest.Success("""{"success":true,"queryName":"CsvData","isLoadedToDataModel":true}"""),
            "powerquery.load-to",
            """{"queryName":"CsvData","loadDestination":"load-to-data-model"}""");

        using var document = JsonDocument.Parse(call.JsonResult);
        Assert.True(document.RootElement.GetProperty("isLoadedToDataModel").GetBoolean());
    }

    [Fact]
    public async Task DaxMeasureWrites_SendFormulaUnchanged()
    {
        const string formula = "DIVIDE(SUM(SalesTable[Amount]), 1000)";
        await _fixture.CallToolAsync(
            "datamodel",
            new Dictionary<string, object?>
            {
                ["action"] = "create-measure",
                ["workbook_session_id"] = SessionId,
                ["table_name"] = "SalesTable",
                ["measure_name"] = "Total",
                ["dax_formula"] = formula,
                ["format_type"] = "Decimal"
            },
            RecordingToolTest.Success("""{"success":true}"""),
            "datamodel.create-measure",
            $$"""{"tableName":"SalesTable","measureName":"Total","daxFormula":{{JsonSerializer.Serialize(formula)}},"formatType":"Decimal"}""");
        await _fixture.CallToolAsync(
            "datamodel",
            new Dictionary<string, object?>
            {
                ["action"] = "update-measure",
                ["workbook_session_id"] = SessionId,
                ["measure_name"] = "Total",
                ["dax_formula"] = formula,
                ["format_type"] = "Decimal"
            },
            RecordingToolTest.Success("""{"success":true}"""),
            "datamodel.update-measure",
            $$"""{"measureName":"Total","daxFormula":{{JsonSerializer.Serialize(formula)}},"formatType":"Decimal"}""");
    }

    [Fact]
    public async Task WorksheetCopyToFile_WithoutSessionId_DispatchesWithoutSession()
    {
        var call = await _fixture.CallToolAsync(
            "worksheet",
            new Dictionary<string, object?>
            {
                ["action"] = "copy-to-file",
                ["source_file"] = "source.xlsx",
                ["source_sheet"] = "Sheet1",
                ["target_file"] = "target.xlsx",
                ["target_sheet_name"] = "CopiedSheet"
            },
            RecordingToolTest.Success("""{"success":true}"""),
            "sheet.copy-to-file",
            string.Empty,
            """{"sourceFile":"source.xlsx","sourceSheet":"Sheet1","targetFile":"target.xlsx","targetSheetName":"CopiedSheet"}""");

        Assert.False(call.Result.IsError);
    }
}
