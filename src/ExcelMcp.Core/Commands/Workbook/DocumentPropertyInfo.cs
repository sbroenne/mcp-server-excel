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

    /// <summary>Whether Excel successfully returned the property value. False does not establish that the value is undefined.</summary>
    public bool ValueAvailable { get; set; } = true;

    /// <summary>Native value-read error and uncertainty when ValueAvailable is false; otherwise null.</summary>
    public string? ValueError { get; set; }

    /// <summary>Office document property value type.</summary>
    public string ValueType { get; set; } = string.Empty;

    /// <summary>Property collection: built-in or custom.</summary>
    public string Scope { get; set; } = string.Empty;
}
