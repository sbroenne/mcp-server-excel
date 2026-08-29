using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Result of read-only workbook integrity validation.</summary>
public sealed class WorkbookIntegrityResult : ResultBase
{
    /// <summary>Overall integrity outcome based on the highest finding severity.</summary>
    public WorkbookIntegrityStatus OverallStatus { get; set; }

    /// <summary>Calculation mode observed without changing it.</summary>
    public string CalculationMode { get; set; } = string.Empty;

    /// <summary>Calculation state observed without waiting or recalculating.</summary>
    public string CalculationState { get; set; } = string.Empty;

    /// <summary>Checks that were run.</summary>
    public List<WorkbookIntegrityCheck> CheckedChecks { get; set; } = [];

    /// <summary>Worksheets inspected by scoped checks or control totals.</summary>
    public List<string> CheckedWorksheets { get; set; } = [];

    /// <summary>Total number of findings, including omitted details.</summary>
    public int FindingCount { get; set; }

    /// <summary>Number of error findings.</summary>
    public int ErrorCount { get; set; }

    /// <summary>Number of warning findings.</summary>
    public int WarningCount { get; set; }

    /// <summary>Number of informational findings.</summary>
    public int InformationCount { get; set; }

    /// <summary>Whether finding details were omitted because the result limit was reached.</summary>
    public bool FindingsTruncated { get; set; }

    /// <summary>Findings grouped by severity and category.</summary>
    public List<WorkbookIntegrityFindingGroup> Groups { get; set; } = [];
}
