namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result returned after releasing a workbook savepoint.
/// </summary>
public sealed class WorkbookSavepointReleaseResult : ResultBase
{
    /// <summary>Session that owned the savepoint.</summary>
    public string SessionId { get; set; } = string.Empty;

    /// <summary>Requested savepoint name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Whether a savepoint was found and released.</summary>
    public bool Released { get; set; }
}
