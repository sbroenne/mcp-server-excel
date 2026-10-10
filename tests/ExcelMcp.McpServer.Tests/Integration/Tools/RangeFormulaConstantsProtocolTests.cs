using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeFormulaConstantsProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData("range", "set-formulas")]
    [InlineData("range_read", "validate-formulas")]
    public async Task MixedFormulaCells_AreAcceptedAndForwarded(string tool, string action)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = action,
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Checks",
            ["range_address"] = "A1:E1",
            ["formulas"] = new List<List<object?>> { new() { "Label", 5.86, true, null, "=1+1" } }
        };

        var call = await fixture.CallToolAsync(tool, arguments,
            RecordingToolTest.Success("""{"success":true}"""),
            $"range.{action}",
            """{"sheetName":"Checks","rangeAddress":"A1:E1","formulas":[["Label",5.86,true,null,"=1+1"]]}""");

        Assert.False(call.Result.IsError, call.JsonResult);
        using var json = JsonDocument.Parse(call.JsonResult);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task Discovery_FormulaCellsAcceptSameKindsAsValueCells()
    {
        var tools = await fixture.ListToolsAsync();
        var range = Assert.Single(tools, tool => tool.Name == "range");
        var valueCell = range.JsonSchema.GetProperty("properties")
            .GetProperty("values").GetProperty("items").GetProperty("items");

        foreach (var toolName in new[] { "range", "range_read" })
        {
            var tool = Assert.Single(tools, candidate => candidate.Name == toolName);
            var formulaCell = tool.JsonSchema.GetProperty("properties")
                .GetProperty("formulas").GetProperty("items").GetProperty("items");

            Assert.True(JsonElement.DeepEquals(valueCell, formulaCell),
                $"Expected {toolName} formula cell schema {valueCell.GetRawText()}, but found {formulaCell.GetRawText()}.");
        }
    }
}
