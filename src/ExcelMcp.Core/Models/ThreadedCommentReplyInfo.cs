namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// A reply within an Excel threaded comment.
/// </summary>
public sealed class ThreadedCommentReplyInfo
{
    /// <summary>Reply text.</summary>
    public string Text { get; init; } = string.Empty;

    /// <summary>Display name of the reply author.</summary>
    public string AuthorName { get; init; } = string.Empty;

    /// <summary>Date reported by Excel for the reply.</summary>
    public DateTime? Date { get; init; }
}
