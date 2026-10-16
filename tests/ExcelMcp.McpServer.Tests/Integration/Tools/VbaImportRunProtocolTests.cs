using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// MCP arguments for VBA import and run. Macro execution and its saved workbook
/// effect are covered by the Service VBA tests.
/// </summary>
[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "VBA")]
[Trait("RequiresExcel", "false")]
public sealed class VbaImportRunProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task ImportAndRun_MapModuleCodeAndProcedure()
    {
        const string code = "Sub WriteProof()\r\n    ThisWorkbook.Sheets(1).Range(\"A1\").Value = \"ok\"\r\nEnd Sub";
        var import = await fixture.CallToolAsync("vba", new()
        {
            ["action"] = "import",
            ["workbook_session_id"] = "session-1",
            ["module_name"] = "Proof",
            ["vba_code"] = code
        }, RecordingToolTest.Success("""{"success":true}"""), "vba.import",
            $$"""{"moduleName":"Proof","vbaCode":{{JsonSerializer.Serialize(code)}}}""");
        var run = await fixture.CallToolAsync("vba", new()
        {
            ["action"] = "run",
            ["workbook_session_id"] = "session-1",
            ["procedure_name"] = "Proof.WriteProof"
        }, RecordingToolTest.Success("""{"success":true}"""), "vba.run",
            """{"procedureName":"Proof.WriteProof"}""");

        Assert.False(import.Result.IsError);
        Assert.False(run.Result.IsError);
    }
}
