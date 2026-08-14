using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    /// <inheritdoc />
    public OperationResult AddThreadedComment(IExcelBatch batch, string sheetName, string cellAddress, string text)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? cell = ResolveCommentCell(ctx.Book, sheetName, cellAddress);
            Excel.CommentThreaded? comment = null;
            try
            {
                comment = cell.CommentThreaded;
                if (comment != null)
                {
                    throw new InvalidOperationException($"Cell '{sheetName}!{cellAddress}' already has a threaded comment.");
                }

                comment = cell.AddCommentThreaded(text);
                return new OperationResult
                {
                    Success = true,
                    Action = "add-threaded-comment",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref comment);
                ComUtilities.Release(ref cell);
            }
        });
    }

    /// <inheritdoc />
    public ThreadedCommentsResult ListThreadedComments(IExcelBatch batch, string sheetName, string cellAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            var result = new ThreadedCommentsResult
            {
                FilePath = batch.WorkbookPath,
                SheetName = sheetName,
                CellAddress = cellAddress
            };

            Excel.Range? cell = ResolveCommentCell(ctx.Book, sheetName, cellAddress);
            Excel.CommentThreaded? comment = null;
            try
            {
                comment = cell.CommentThreaded;
                if (comment != null)
                {
                    result.Comments.Add(ReadThreadedComment(comment, cellAddress));
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref comment);
                ComUtilities.Release(ref cell);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult AddThreadedCommentReply(IExcelBatch batch, string sheetName, string cellAddress, string text)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? cell = ResolveCommentCell(ctx.Book, sheetName, cellAddress);
            Excel.CommentThreaded? comment = null;
            Excel.CommentThreaded? reply = null;
            try
            {
                comment = cell.CommentThreaded
                    ?? throw new InvalidOperationException($"Cell '{sheetName}!{cellAddress}' has no threaded comment.");
                reply = comment.AddReply(text);
                return new OperationResult
                {
                    Success = true,
                    Action = "add-threaded-comment-reply",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref reply);
                ComUtilities.Release(ref comment);
                ComUtilities.Release(ref cell);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult DeleteThreadedComment(IExcelBatch batch, string sheetName, string cellAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? cell = ResolveCommentCell(ctx.Book, sheetName, cellAddress);
            Excel.CommentThreaded? comment = null;
            try
            {
                comment = cell.CommentThreaded
                    ?? throw new InvalidOperationException($"Cell '{sheetName}!{cellAddress}' has no threaded comment.");
                comment.Delete();
                return new OperationResult
                {
                    Success = true,
                    Action = "delete-threaded-comment",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref comment);
                ComUtilities.Release(ref cell);
            }
        });
    }

    private static Excel.Range ResolveCommentCell(Excel.Workbook workbook, string sheetName, string cellAddress)
    {
        var cell = RangeHelpers.ResolveRange(workbook, sheetName, cellAddress, out string? specificError) as Excel.Range;
        if (cell == null)
        {
            throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, cellAddress));
        }

        if (cell.CountLarge != 1)
        {
            ComUtilities.Release(ref cell);
            throw new ArgumentException("Threaded comment operations require a single-cell address.", nameof(cellAddress));
        }

        return cell;
    }

    private static ThreadedCommentInfo ReadThreadedComment(Excel.CommentThreaded comment, string cellAddress)
    {
        Excel.Author? author = null;
        Excel.CommentsThreaded? replies = null;
        try
        {
            author = comment.Author;
            var result = new ThreadedCommentInfo
            {
                CellAddress = cellAddress,
                Text = comment.Text(),
                AuthorName = author?.Name ?? string.Empty,
                Date = ConvertCommentDate(comment.Date)
            };

            replies = comment.Replies;
            for (int index = 1; index <= replies.Count; index++)
            {
                Excel.CommentThreaded? reply = null;
                Excel.Author? replyAuthor = null;
                try
                {
                    reply = replies.Item(index);
                    replyAuthor = reply.Author;
                    result.Replies.Add(new ThreadedCommentReplyInfo
                    {
                        Text = reply.Text(),
                        AuthorName = replyAuthor?.Name ?? string.Empty,
                        Date = ConvertCommentDate(reply.Date)
                    });
                }
                finally
                {
                    ComUtilities.Release(ref replyAuthor);
                    ComUtilities.Release(ref reply);
                }
            }

            return result;
        }
        finally
        {
            ComUtilities.Release(ref replies);
            ComUtilities.Release(ref author);
        }
    }

    private static DateTime? ConvertCommentDate(object? value)
    {
        return value switch
        {
            DateTime date => date,
            double serial => DateTime.FromOADate(serial),
            _ => null
        };
    }
}
