using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// MCP arguments for cross-file worksheet copies without a session.
/// Workbook behavior for these actions is covered by the Service tests.
/// </summary>
[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "false")]
public sealed class WorksheetCopyToFileProtocolTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

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
