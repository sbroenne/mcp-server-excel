namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>
/// Workbook document property value.
/// </summary>
public sealed class DocumentPropertyInfo
{
    /// <summary>Property name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Invariant string representation of the property value.</summary>
    public string? Value { get; set; }

    /// <summary>Office document property value type.</summary>
    public string ValueType { get; set; } = string.Empty;

    /// <summary>Property collection: built-in or custom.</summary>
    public string Scope { get; set; } = string.Empty;
}
