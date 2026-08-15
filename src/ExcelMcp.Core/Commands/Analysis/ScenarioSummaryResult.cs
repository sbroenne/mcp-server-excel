using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Analysis;

/// <summary>
/// Result of creating an Excel scenario summary report.
/// </summary>
public sealed class ScenarioSummaryResult : OperationResult
{
    /// <summary>Name of the worksheet Excel created for the report.</summary>
    public string ReportSheetName { get; set; } = string.Empty;

    /// <summary>Created report type.</summary>
    public string ReportType { get; set; } = string.Empty;
}
