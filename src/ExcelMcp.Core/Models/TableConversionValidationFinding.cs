namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// One deterministic post-creation table validation failure.
/// </summary>
public sealed class TableConversionValidationFinding
{
    /// <summary>
    /// Finding type.
    /// </summary>
    public TableConversionValidationFindingKind Kind { get; set; }

    /// <summary>
    /// Cells or ranges related to the finding.
    /// </summary>
    public List<string> Addresses { get; set; } = [];

    /// <summary>
    /// Plain-English explanation.
    /// </summary>
    public string Message { get; set; } = string.Empty;
}
