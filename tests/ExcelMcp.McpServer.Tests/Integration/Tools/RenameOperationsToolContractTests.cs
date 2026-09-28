using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "RenameContract")]
[Trait("RequiresExcel", "false")]
public sealed class RenameOperationsToolContractTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task PowerQueryRename_MissingQuery_ReturnsJsonBusinessError()
    {
        const string sessionId = "recording-session";
        var call = await _fixture.CallToolAsync(
            "powerquery",
            new Dictionary<string, object?>
            {
                ["action"] = "rename",
                ["session_id"] = sessionId,
                ["old_name"] = "NonExistentQuery",
                ["new_name"] = "NewName"
            },
            RecordingToolTest.Success(
                """{"success":false,"errorMessage":"Query 'NonExistentQuery' not found"}"""),
            "powerquery.rename",
            """{"oldName":"NonExistentQuery","newName":"NewName"}""");

        using (var args = RecordingToolTest.ParseArgs(
            call.Request,
            "powerquery.rename",
            sessionId))
        {
            Assert.Equal(
                "NonExistentQuery",
                args.RootElement.GetProperty("oldName").GetString());
            Assert.Equal("NewName", args.RootElement.GetProperty("newName").GetString());
        }

        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains(
            "not found",
            result.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
        Assert.False(result.RootElement.TryGetProperty("exceptionType", out _));
    }

    [Fact]
    public async Task PowerQueryRename_Success_ReturnsCompleteRenameResult()
    {
        const string sessionId = "recording-session";
        var call = await _fixture.CallToolAsync(
            "powerquery",
            new Dictionary<string, object?>
            {
                ["action"] = "rename",
                ["session_id"] = sessionId,
                ["old_name"] = "OriginalQuery",
                ["new_name"] = "RenamedQuery"
            },
            RecordingToolTest.Success(
                """{"success":true,"objectType":"power-query","oldName":"OriginalQuery","newName":"RenamedQuery"}"""),
            "powerquery.rename",
            """{"oldName":"OriginalQuery","newName":"RenamedQuery"}""");

        using (var args = RecordingToolTest.ParseArgs(
            call.Request,
            "powerquery.rename",
            sessionId))
        {
            Assert.Equal(
                "OriginalQuery",
                args.RootElement.GetProperty("oldName").GetString());
            Assert.Equal(
                "RenamedQuery",
                args.RootElement.GetProperty("newName").GetString());
        }

        using var result = JsonDocument.Parse(call.JsonResult);
        var root = result.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.Equal("power-query", root.GetProperty("objectType").GetString());
        Assert.Equal("OriginalQuery", root.GetProperty("oldName").GetString());
        Assert.Equal("RenamedQuery", root.GetProperty("newName").GetString());
    }
}
