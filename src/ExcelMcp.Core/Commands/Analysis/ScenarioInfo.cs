namespace Sbroenne.ExcelMcp.Core.Commands.Analysis;

/// <summary>
/// Metadata and stored inputs for an Excel worksheet scenario.
/// </summary>
public sealed class ScenarioInfo
{
    /// <summary>Scenario name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Absolute A1 address of the changing cells.</summary>
    public string ChangingCells { get; set; } = string.Empty;

    /// <summary>Stored values in changing-cell order.</summary>
    public List<object?> Values { get; set; } = [];

    /// <summary>Scenario comment, including Excel's author/date prefix when present.</summary>
    public string Comment { get; set; } = string.Empty;

    /// <summary>Whether the scenario is locked.</summary>
    public bool Locked { get; set; }

    /// <summary>Whether the scenario is hidden.</summary>
    public bool Hidden { get; set; }
}
