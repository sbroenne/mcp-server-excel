namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Sparse diagnostics for formula cells whose calculated values are Excel errors.
/// </summary>
public sealed class RangeFormulaErrorResult : ResultBase
{
    /// <summary>Worksheet containing the source range.</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Resolved source range address.</summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>Total formula error cells found in the source range.</summary>
    public long TotalErrorCount { get; set; }

    /// <summary>Number of diagnostics returned.</summary>
    public int ReturnedErrorCount { get; set; }

    /// <summary>Maximum diagnostics requested by the caller.</summary>
    public int MaxErrors { get; set; }

    /// <summary>Whether additional formula errors were omitted.</summary>
    public bool IsTruncated { get; set; }

    /// <summary>Sparse formula error diagnostics in deterministic worksheet order.</summary>
    public List<RangeCellError> Errors { get; set; } = [];
}
