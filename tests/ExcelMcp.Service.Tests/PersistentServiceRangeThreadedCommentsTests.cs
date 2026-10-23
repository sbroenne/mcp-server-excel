using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using RangeCommands = Sbroenne.ExcelMcp.Core.Commands.Range.RangeCommands;

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
    [Theory]
    [InlineData("add")]
    [InlineData("list")]
    [InlineData("reply")]
    [InlineData("delete")]
    [InlineData("multiple-cells")]
    public void ThreadedComments_InspectionFailure_ReleasesCellAndPreservesComments(string operation)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.AddThreadedComment(batch, sheetName, "B2", "Original comment").Success);
        object? acquiredCell = null;
        var inspectionFailure = Assert.IsType<COMException>(
            Marshal.GetExceptionForHR(unchecked((int)0x80004005), new IntPtr(-1)));
        RangeCommands.CommentCellAcquiredForTests = cell =>
        {
            acquiredCell = cell;
            if (operation != "multiple-cells")
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(inspectionFailure).Throw();
        };

        PersistentServiceCleanupFailures.Run(() =>
        {
            var failure = Assert.Throws(
                operation == "multiple-cells" ? typeof(ArgumentException) : typeof(InvalidOperationException), () =>
            {
                switch (operation)
                {
                    case "add":
                        _commands.AddThreadedComment(batch, sheetName, "B3", "New comment");
                        break;
                    case "multiple-cells":
                        _commands.AddThreadedComment(batch, sheetName, "B2:B3", "New comment");
                        break;
                    case "list":
                        _commands.ListThreadedComments(batch, sheetName, "B2");
                        break;
                    case "reply":
                        _commands.AddThreadedCommentReply(batch, sheetName, "B2", "New reply");
                        break;
                    case "delete":
                        _commands.DeleteThreadedComment(batch, sheetName, "B2");
                        break;
                }
            });
            Assert.Contains(operation == "multiple-cells" ? "single-cell address" : inspectionFailure.Message,
                failure.Message, StringComparison.Ordinal);
            Assert.NotNull(acquiredCell);
            RangeCommands.CommentCellAcquiredForTests = null;

            var unchanged = _commands.ListThreadedComments(batch, sheetName, "B2");
            Assert.True(unchanged.Success);
            var original = Assert.Single(unchanged.Comments);
            Assert.Equal("B2", original.CellAddress);
            Assert.Equal("Original comment", original.Text);
            Assert.Empty(original.Replies);
            var empty = _commands.ListThreadedComments(batch, sheetName, "B3");
            Assert.True(empty.Success);
            Assert.Empty(empty.Comments);
            _fixture.ExecuteRawVerification((_, _) =>
                Assert.Throws<InvalidComObjectException>(() => _ = ((Excel.Range)acquiredCell).CountLarge));
        },
        () =>
        {
            RangeCommands.CommentCellAcquiredForTests = null;
            _fixture.ExecuteRawVerification((_, _) => ComUtilities.Release(ref acquiredCell));
        });
    }

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
