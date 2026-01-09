using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Information about a Data Model column
/// </summary>
public class DataModelColumnInfo
{
    /// <summary>
    /// Column name
    /// </summary>
    [JsonPropertyName("n")]
    public string Name { get; init; } = "";

    /// <summary>
    /// Column data type
    /// </summary>
    [JsonPropertyName("dt")]
    public string DataType { get; init; } = "";

    /// <summary>
    /// Whether this is a calculated column (has DAX formula)
    /// </summary>
    [JsonPropertyName("ic")]
    public bool IsCalculated { get; init; }
}
