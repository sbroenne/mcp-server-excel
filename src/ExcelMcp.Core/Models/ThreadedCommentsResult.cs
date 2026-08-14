namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result containing the threaded comment for a cell, when present.
/// </summary>
public sealed class ThreadedCommentsResult : ResultBase
{
    /// <summary>Worksheet name.</summary>
    public string SheetName { get; init; } = string.Empty;

    /// <summary>Cell address inspected.</summary>
    public string CellAddress { get; init; } = string.Empty;

    /// <summary>Top-level comments found at the cell.</summary>
    public List<ThreadedCommentInfo> Comments { get; init; } = [];
}
