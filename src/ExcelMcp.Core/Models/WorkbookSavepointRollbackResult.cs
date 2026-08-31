namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result returned after restoring a workbook savepoint.
/// </summary>
public sealed class WorkbookSavepointRollbackResult : ResultBase
{
    /// <summary>Session retained after rollback.</summary>
    public string SessionId { get; set; } = string.Empty;

    /// <summary>Rolled-back savepoint name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>UTC rollback completion time.</summary>
    public DateTime RestoredAtUtc { get; set; }

    /// <summary>Whether the savepoint remains available for another rollback.</summary>
    public bool SavepointRetained { get; set; }

    /// <summary>Whether the workbook was reopened under the same session ID.</summary>
    public bool SessionReopened { get; set; }
}
