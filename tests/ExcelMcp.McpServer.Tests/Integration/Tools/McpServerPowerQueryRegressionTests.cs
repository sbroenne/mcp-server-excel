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
    private static readonly string[] ExpectedColumns = ["CsvData[Product]", "CsvData[Quantity]"];
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
                    ["workbook_session_id"] = sessionId,
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
                    ["workbook_session_id"] = sessionId,
                    ["query_name"] = "CsvData",
                    ["load_destination"] = "load-to-data-model"
                },
                ToolTimeout);
            AssertSuccess(loadResult, "powerquery.load-to data-model");

            var listTablesResult = await _fixture.CallToolAsync(
                "datamodel_read",
                new Dictionary<string, object?>
                {
                    ["action"] = "list-tables",
                    ["workbook_session_id"] = sessionId
                },
                ToolTimeout);
            AssertSuccess(
                listTablesResult,
                "datamodel_read.list-tables after powerquery.load-to");
            using (var tables = JsonDocument.Parse(listTablesResult))
            {
                var table = Assert.Single(tables.RootElement.GetProperty("tables").EnumerateArray(),
                    item => item.GetProperty("name").GetString() == "CsvData");
                Assert.Equal(2, table.GetProperty("recordCount").GetInt32());
            }

            var evaluated = await _fixture.CallToolAsync("datamodel_read", new Dictionary<string, object?>
            {
                ["action"] = "evaluate",
                ["workbook_session_id"] = sessionId,
                ["dax_query"] = "EVALUATE CsvData ORDER BY CsvData[Product]"
            }, ToolTimeout);
            AssertSuccess(evaluated, "datamodel_read.evaluate loaded CSV");
            using (var data = JsonDocument.Parse(evaluated))
            {
                Assert.Equal(2, data.RootElement.GetProperty("rowCount").GetInt32());
                Assert.Equal(2, data.RootElement.GetProperty("columnCount").GetInt32());
                Assert.Equal(ExpectedColumns,
                    data.RootElement.GetProperty("columns").EnumerateArray().Select(column => column.GetString()));
                var rows = data.RootElement.GetProperty("rows");
                Assert.Equal(2, rows.GetArrayLength());
                Assert.Equal("Gadget", rows[0][0].GetString());
                Assert.Equal(20, rows[0][1].GetInt32());
                Assert.Equal("Widget", rows[1][0].GetString());
                Assert.Equal(10, rows[1][1].GetInt32());
            }

            var listSessionsResult = await _fixture.CallToolAsync(
                "file_read",
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
        $$$"""
        let
            Source = Csv.Document(File.Contents("{{{csvPath.Replace("\"", "\"\"")}}}"), [Delimiter = ",", Columns = 2, Encoding = 1252, QuoteStyle = QuoteStyle.None]),
            PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
            TypedColumns = Table.TransformColumnTypes(PromotedHeaders, {{"Product", type text}, {"Quantity", Int64.Type}})
        in
            TypedColumns
        """;

    private static void AssertSuccess(
        string jsonResult,
        string operationName)
    {
        McpResponseAssertions.AssertSuccess(jsonResult, operationName);
    }
}
