namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// One actionable result from a table-creation preflight.
/// </summary>
public sealed class TablePreflightFinding
{
    /// <summary>
    /// Finding type.
    /// </summary>
    public TablePreflightFindingKind Kind { get; set; }

    /// <summary>
    /// Whether the finding blocks creation or is advisory.
    /// </summary>
    public TablePreflightSeverity Severity { get; set; }

    /// <summary>
    /// True when the finding is based on a heuristic rather than a deterministic Excel constraint.
    /// </summary>
    public bool IsHeuristic { get; set; }

    /// <summary>
    /// Cells or ranges related to the finding.
    /// </summary>
    public List<string> Addresses { get; set; } = [];

    /// <summary>
    /// Plain-English explanation of the problem.
    /// </summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>
    /// Plain-English action that resolves or reviews the finding.
    /// </summary>
    public string Remediation { get; set; } = string.Empty;
}
