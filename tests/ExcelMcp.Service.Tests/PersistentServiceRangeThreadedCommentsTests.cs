using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceRangeThreadedCommentsTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void ThreadedComments_AddReplyListDelete_RoundTrips()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var addResult = _commands.AddThreadedComment(
            batch, sheetName, "B2", "Review this value");
        Assert.True(addResult.Success);

        var replyResult = _commands.AddThreadedCommentReply(
            batch, sheetName, "B2", "Reviewed");
        Assert.True(replyResult.Success);

        var listResult = _commands.ListThreadedComments(batch, sheetName, "B2");
        Assert.True(listResult.Success);
        var comment = Assert.Single(listResult.Comments);
        Assert.Equal("B2", comment.CellAddress);
        Assert.Equal("Review this value", comment.Text);
        Assert.Equal(["Reviewed"], comment.Replies.Select(reply => reply.Text));

        var deleteResult = _commands.DeleteThreadedComment(batch, sheetName, "B2");
        Assert.True(deleteResult.Success);

        var finalList = _commands.ListThreadedComments(batch, sheetName, "B2");
        Assert.True(finalList.Success);
        Assert.Empty(finalList.Comments);
    }
}
