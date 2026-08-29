namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Composite row-key preservation check.
/// </summary>
public sealed class TableRowKeyCheckResult
{
    /// <summary>
    /// Columns used to form the composite key.
    /// </summary>
    public List<string> KeyColumns { get; set; } = [];

    /// <summary>
    /// Number of unique keys before sorting.
    /// </summary>
    public int BeforeCount { get; set; }

    /// <summary>
    /// Number of unique keys after sorting.
    /// </summary>
    public int AfterCount { get; set; }

    /// <summary>
    /// Keys whose row content was missing or changed.
    /// </summary>
    public List<string> MismatchedKeys { get; set; } = [];

    /// <summary>
    /// Whether all keyed rows were preserved.
    /// </summary>
    public bool Passed { get; set; }
}
