using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

public partial class RangeCommandsTests
{
    [Fact]
    public void ThreadedComments_AddReplyListDelete_RoundTrips()
    {
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);
        const string sheetName = "ThreadedComments";
        batch.Execute((ctx, ct) =>
        {
            var sheet = ctx.Book.Worksheets.Add();
            sheet.Name = sheetName;
            return 0;
        });

        var addResult = _commands.AddThreadedComment(batch, sheetName, "B2", "Review this value");
        Assert.True(addResult.Success);

        var replyResult = _commands.AddThreadedCommentReply(batch, sheetName, "B2", "Reviewed");
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
