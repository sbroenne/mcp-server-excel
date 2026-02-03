
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
    /// Whether this is a calculated column (has DAX formula)
    /// </summary>
    public bool IsCalculated { get; init; }
}
