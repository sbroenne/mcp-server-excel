using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "false")]
public sealed class WorksheetRenameParameterTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task Rename_WithOldNameAndNewName_Succeeds()
    {
        const string sessionId = "recording-session";
        var call = await _fixture.CallToolAsync(
            "worksheet",
            new Dictionary<string, object?>
            {
                ["action"] = "rename",
                ["workbook_session_id"] = sessionId,
                ["old_name"] = "OriginalSheet",
                ["new_name"] = "RenamedSheet"
            },
            RecordingToolTest.Success(
                """{"success":true,"oldName":"OriginalSheet","newName":"RenamedSheet"}"""),
            "sheet.rename",
            """{"oldName":"OriginalSheet","newName":"RenamedSheet"}""");

        using (var args = RecordingToolTest.ParseArgs(
            call.Request,
            "sheet.rename",
            sessionId))
        {
            Assert.Equal(
                "OriginalSheet",
                args.RootElement.GetProperty("oldName").GetString());
            Assert.Equal(
                "RenamedSheet",
                args.RootElement.GetProperty("newName").GetString());
        }

        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(
            "RenamedSheet",
            result.RootElement.GetProperty("newName").GetString());
    }

    [Theory]
    [InlineData("LegacySheet", "RenamedLegacy", "sheet_name", "sheet_name", "target_name")]
    [InlineData("TargetSheetAlias", "RenamedViaTargetSheet", "sheet_name", "sheet_name", "target_sheet_name")]
    [InlineData("SourceNameAlias", "RenamedViaSourceName", "source_name", "source_name", "target_name")]
    [InlineData("SourceSheetAlias", "RenamedViaSourceSheet", "source_sheet", "source_sheet", "target_sheet_name")]
    [InlineData("MixedCanonicalAlias", "MixedAliasTarget", "target_name", "old_name", "target_name")]
    [InlineData("MixedAliasCanonical", "MixedAliasCanonicalRenamed", "sheet_name", "sheet_name", "new_name")]
    public async Task Rename_WithAliasOnlyPayload_RejectsInapplicableParameter(
        string originalSheetName,
        string renamedSheetName,
        string expectedErrorFragment,
        string oldNameParameter,
        string newNameParameter)
    {
        var resultText = await _fixture.CallToolWithoutDispatchAsync(
            "worksheet",
            new Dictionary<string, object?>
            {
                [oldNameParameter] = originalSheetName,
                [newNameParameter] = renamedSheetName,
                ["action"] = "rename",
                ["workbook_session_id"] = "recording-session"
            });

        using var result = JsonDocument.Parse(resultText);
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains(
            expectedErrorFragment,
            result.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task Rename_WithNoNameParameters_FailsWithOldNameRequired()
    {
        var resultText = await _fixture.CallToolWithoutDispatchAsync(
            "worksheet",
            new Dictionary<string, object?>
            {
                ["action"] = "rename",
                ["workbook_session_id"] = "recording-session"
            });

        using var result = JsonDocument.Parse(resultText);
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains(
            "Parameter 'old_name' is required for worksheet.rename",
            result.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.Ordinal);
    }
}
