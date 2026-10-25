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
    public string Name { get; init; } = "";

    /// <summary>
    /// Column data type
    /// </summary>
    public string DataType { get; init; } = "";

    /// <summary>
    /// Raw OLE DB type code returned by Excel's ModelTableColumn.DataType.
    /// </summary>
    public int DataTypeValue { get; init; }

    /// <summary>
    /// Human-readable name for an observed model column OLE DB type, or Unknown with its raw code.
    /// </summary>
    public string DataTypeName { get; init; } = "";

    /// <summary>
    /// Whether this is a calculated column, or null when Excel cannot expose that information.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public bool? IsCalculated { get; init; }
}
