
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
    /// Excel XlParameterDataType numeric value.
    /// </summary>
    public int DataTypeValue { get; init; }

    /// <summary>
    /// Human-readable XlParameterDataType name.
    /// </summary>
    public string DataTypeName { get; init; } = "";

    /// <summary>
    /// Whether this is a calculated column (has DAX formula)
    /// </summary>
    public bool IsCalculated { get; init; }
}

