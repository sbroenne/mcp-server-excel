namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Explicit date value accepted as one cell inside a range set-values matrix.
/// </summary>
public sealed class TypedCellValue
{
    /// <summary>The date value type.</summary>
    public TypedCellValueType? Type { get; init; }

    /// <summary>The ISO 8601 value.</summary>
    public string? Value { get; init; }

    /// <summary>Optional Excel number format. A type-specific ISO format is used when omitted.</summary>
    public string? NumberFormat { get; init; }
}
