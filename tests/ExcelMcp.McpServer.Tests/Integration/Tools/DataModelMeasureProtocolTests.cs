using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// MCP arguments for DAX measure writes.
/// Workbook behavior for these actions is covered by the Service tests.
/// </summary>
[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "DataModel")]
[Trait("RequiresExcel", "false")]
public sealed class DataModelMeasureProtocolTests(
    RecordingProgramTransportFixture fixture)
{
    private const string SessionId = "recording-session";

    [Fact]
    public async Task DaxMeasureWrites_SendFormulaUnchanged()
    {
        const string formula = "DIVIDE(SUM(SalesTable[Amount]), 1000)";
        await fixture.CallToolAsync(
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
        await fixture.CallToolAsync(
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
}
