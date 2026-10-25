using System.Text.Json;
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
    [Theory]
    [InlineData("worksheet", "create", """{"sheet_name":"New","file_path":"book.xlsx"}""", """{"sheetName":"New","filePath":"book.xlsx"}""")]
    [InlineData("worksheet", "rename", """{"old_name":"Old","new_name":"New"}""", """{"oldName":"Old","newName":"New"}""")]
    [InlineData("worksheet", "copy", """{"source_name":"Old","target_name":"New"}""", """{"sourceName":"Old","targetName":"New"}""")]
    [InlineData("worksheet", "delete", """{"sheet_name":"Old"}""", """{"sheetName":"Old"}""")]
    [InlineData("worksheet", "move", """{"sheet_name":"Old","after_sheet":"Summary"}""", """{"sheetName":"Old","afterSheet":"Summary"}""")]
    [InlineData("worksheet_read", "list", """{"file_path":"book.xlsx"}""", """{"filePath":"book.xlsx"}""")]
    public async Task SessionActions_PreserveExistingCommandAndArguments(
        string tool, string action, string inputs, string expected)
    {
        var arguments = JsonSerializer.Deserialize<Dictionary<string, object?>>(inputs)!;
        arguments["action"] = action;
        arguments["workbook_session_id"] = "session-1";
        var call = await fixture.CallToolAsync(tool, arguments,
            RecordingToolTest.Success("""{"success":true}"""), $"sheet.{action}", expected);
        Assert.False(call.Result.IsError);
    }

    [Theory]
    [InlineData("copy-to-file")]
    [InlineData("move-to-file")]
    public async Task CrossFileActions_PreservePlacementWithoutSession(string action)
    {
        var call = await fixture.CallToolAsync("worksheet", new()
        {
            ["action"] = action,
            ["source_file"] = "source.xlsx",
            ["source_sheet"] = "Data",
            ["target_file"] = "target.xlsx",
            ["before_sheet"] = "Summary"
        }, RecordingToolTest.Success("""{"success":true}"""), $"sheet.{action}", string.Empty,
            """{"sourceFile":"source.xlsx","sourceSheet":"Data","targetFile":"target.xlsx","beforeSheet":"Summary"}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task WorksheetCopyToFile_WithoutSessionId_DispatchesWithoutSession()
    {
        var call = await fixture.CallToolAsync(
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
