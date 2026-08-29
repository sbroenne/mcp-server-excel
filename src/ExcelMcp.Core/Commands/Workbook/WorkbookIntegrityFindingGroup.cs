namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Integrity findings sharing one severity and category.</summary>
public sealed class WorkbookIntegrityFindingGroup
{
    /// <summary>Severity shared by the group.</summary>
    public WorkbookIntegritySeverity Severity { get; set; }

    /// <summary>Category shared by the group.</summary>
    public WorkbookIntegrityCategory Category { get; set; }

    /// <summary>Total matching findings, including details omitted by the result limit.</summary>
    public int Count { get; set; }

    /// <summary>Retained finding details for this group.</summary>
    public List<WorkbookIntegrityFinding> Findings { get; set; } = [];
}
