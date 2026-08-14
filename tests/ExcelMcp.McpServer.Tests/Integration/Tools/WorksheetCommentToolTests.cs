using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "true")]
public sealed class WorksheetCommentToolTests : McpIntegrationTestBase
{
    public WorksheetCommentToolTests(ITestOutputHelper output)
        : base(output, "WorksheetCommentClient")
    {
    }

    [Fact]
    public async Task WorksheetStyle_SetAndClearComment_RoundsTripThroughMcp()
    {
        var tempDir = CreateTempDirectory("WorksheetComments");
        var workbookPath = Path.Combine(tempDir, "comments.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);
        await CreateWorksheetAsync(sessionId, "CommentSheet");

        var setCommentJson = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "set-comment",
            ["session_id"] = sessionId,
            ["sheet_name"] = "CommentSheet",
            ["cell_address"] = "A1",
            ["text"] = "Quarterly update"
        });
        AssertSuccess(setCommentJson, "worksheet_style.set-comment");

        using var setCommentDoc = JsonDocument.Parse(setCommentJson);
        Assert.True(setCommentDoc.RootElement.GetProperty("success").GetBoolean());

        var getCommentJson = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "get-comment",
            ["session_id"] = sessionId,
            ["sheet_name"] = "CommentSheet",
            ["cell_address"] = "A1"
        });
        AssertSuccess(getCommentJson, "worksheet_style.get-comment");

        using var getCommentDoc = JsonDocument.Parse(getCommentJson);
        Assert.True(getCommentDoc.RootElement.GetProperty("hasComment").GetBoolean());
        Assert.Equal("Quarterly update", getCommentDoc.RootElement.GetProperty("text").GetString());

        var clearCommentJson = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "clear-comment",
            ["session_id"] = sessionId,
            ["sheet_name"] = "CommentSheet",
            ["cell_address"] = "A1"
        });
        AssertSuccess(clearCommentJson, "worksheet_style.clear-comment");

        using var clearCommentDoc = JsonDocument.Parse(clearCommentJson);
        Assert.True(clearCommentDoc.RootElement.GetProperty("success").GetBoolean());

        await TryCloseSessionAsync(sessionId, save: true);
    }
}
