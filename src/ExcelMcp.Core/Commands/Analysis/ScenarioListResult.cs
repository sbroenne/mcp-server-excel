using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Analysis;

/// <summary>
/// Result containing all scenarios on a worksheet.
/// </summary>
public sealed class ScenarioListResult : OperationResult
{
    /// <summary>Worksheet containing the scenarios.</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Scenarios defined on the worksheet.</summary>
    public List<ScenarioInfo> Scenarios { get; set; } = [];
}
