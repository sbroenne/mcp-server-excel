namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result returned after creating a workbook savepoint.
/// </summary>
public sealed class WorkbookSavepointCreateResult : ResultBase
{
    /// <summary>Session that owns the savepoint.</summary>
    public string SessionId { get; set; } = string.Empty;

    /// <summary>Created savepoint metadata.</summary>
    public WorkbookSavepointInfo Savepoint { get; set; } = new();
}
