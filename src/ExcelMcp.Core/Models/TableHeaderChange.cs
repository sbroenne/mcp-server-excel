namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// One header normalization performed during table conversion.
/// </summary>
public sealed class TableHeaderChange
{
    /// <summary>
    /// Absolute address of the changed header cell.
    /// </summary>
    public string Address { get; set; } = string.Empty;

    /// <summary>
    /// Header value before normalization.
    /// </summary>
    public string? OriginalValue { get; set; }

    /// <summary>
    /// Header value after normalization.
    /// </summary>
    public string NewValue { get; set; } = string.Empty;

    /// <summary>
    /// Why the header was changed.
    /// </summary>
    public TableHeaderChangeReason Reason { get; set; }
}
