namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// A threaded comment attached to an Excel cell.
/// </summary>
public sealed class ThreadedCommentInfo
{
    /// <summary>Cell address containing the comment.</summary>
    public string CellAddress { get; init; } = string.Empty;

    /// <summary>Top-level comment text.</summary>
    public string Text { get; init; } = string.Empty;

    /// <summary>Display name of the comment author.</summary>
    public string AuthorName { get; init; } = string.Empty;

    /// <summary>Date reported by Excel for the comment.</summary>
    public DateTime? Date { get; init; }

    /// <summary>Replies in thread order.</summary>
    public List<ThreadedCommentReplyInfo> Replies { get; init; } = [];
}
