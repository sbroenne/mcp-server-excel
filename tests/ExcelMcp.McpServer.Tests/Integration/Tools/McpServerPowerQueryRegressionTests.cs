using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Real MCP transport smoke for the synchronous Power Query load path.
/// Workbook behavior is covered by the corresponding Service tests.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed class McpServerPowerQueryRegressionTests(
    McpProgramTransportFixture fixture) :
    IClassFixture<McpProgramTransportFixture>
{
    private static readonly TimeSpan ToolTimeout = TimeSpan.FromSeconds(90);
    private readonly McpProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task PowerQuery_LoadToDataModel_CompletesViaMcpProtocol()
    {
        var workbookPath = _fixture.CreateTempPath(
            "LoadToDataModel",
            ".xlsx");
        var sessionId = await _fixture.CreateWorkbookSessionAsync(workbookPath);
        var csvPath = _fixture.CreateTempPath("loadtodm", ".csv");
        await File.WriteAllTextAsync(
            csvPath,
            "Product,Quantity\nWidget,10\nGadget,20");

        try
        {
            var createResult = await _fixture.CallToolAsync(
                "powerquery",
                new Dictionary<string, object?>
                {
                    ["action"] = "create",
                    ["session_id"] = sessionId,
                    ["query_name"] = "CsvData",
                    ["m_code"] = BuildCsvMCode(csvPath),
                    ["load_destination"] = "connection-only"
                },
                ToolTimeout);
            AssertSuccess(createResult, "powerquery.create");

            var loadResult = await _fixture.CallToolAsync(
                "powerquery",
                new Dictionary<string, object?>
                {
                    ["action"] = "load-to",
                    ["session_id"] = sessionId,
                    ["query_name"] = "CsvData",
                    ["load_destination"] = "load-to-data-model"
                },
                ToolTimeout);
            AssertSuccess(loadResult, "powerquery.load-to data-model");

            var listTablesResult = await _fixture.CallToolAsync(
                "datamodel",
                new Dictionary<string, object?>
                {
                    ["action"] = "list-tables",
                    ["session_id"] = sessionId
                },
                ToolTimeout);
            AssertSuccess(
                listTablesResult,
                "datamodel.list-tables after powerquery.load-to");
            Assert.Contains("CsvData", listTablesResult);

            var listSessionsResult = await _fixture.CallToolAsync(
                "file",
                new Dictionary<string, object?> { ["action"] = "list" },
                ToolTimeout);
            AssertSuccess(
                listSessionsResult,
                "file.list after powerquery.load-to");
        }
        finally
        {
            await _fixture.CloseSessionAsync(sessionId);
        }
    }

    private static string BuildCsvMCode(string csvPath) =>
        $"""
        let
            Source = Csv.Document(File.Contents("{csvPath.Replace("\\", "\\\\")}"), [Delimiter = ",", Columns = 2, Encoding = 1252, QuoteStyle = QuoteStyle.None]),
            PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
        in
            PromotedHeaders
        """;

    private static void AssertSuccess(
        string jsonResult,
        string operationName)
    {
        using var json = JsonDocument.Parse(jsonResult);
        Assert.True(
            json.RootElement.GetProperty("success").GetBoolean(),
            $"{operationName} failed: {jsonResult}");
    }
}
