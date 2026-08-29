namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Post-creation validation details for a converted table.
/// </summary>
public sealed class TableConversionValidationResult
{
    /// <summary>
    /// Whether every deterministic validation passed.
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// Number of formula cells inspected.
    /// </summary>
    public int FormulaCellsChecked { get; set; }

    /// <summary>
    /// Names of consistent calculated columns.
    /// </summary>
    public List<string> CalculatedColumns { get; set; } = [];

    /// <summary>
    /// Whether the converted table shows an Excel totals row.
    /// </summary>
    public bool ShowTotals { get; set; }

    /// <summary>
    /// Deterministic validation failures.
    /// </summary>
    public List<TableConversionValidationFinding> Findings { get; set; } = [];
}
