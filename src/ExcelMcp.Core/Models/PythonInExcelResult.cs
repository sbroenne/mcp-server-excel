using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result of reading a Python in Excel (=PY()) cell's computed value.
/// </summary>
public class PythonInExcelResult : ResultBase
{
    /// <summary>
    /// Worksheet name containing the cell
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Resolved cell/range address
    /// </summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>
    /// The full =PY(...) formula text found in the cell
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Formula { get; set; }

    /// <summary>
    /// The computed value, when the formula's return type is "Excel Value" (returnType=0).
    /// Null when the cell holds a "Python Object" (see <see cref="IsPythonObject"/>) or an error.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Value { get; set; }

    /// <summary>
    /// True if the cell holds a "Python Object" (returnType=1) rich data type (e.g. a DataFrame).
    /// Such cells cannot expose their underlying data via COM Value2 - only a type-name label is
    /// available via <see cref="TypeName"/>. Switch the formula to returnType=0 to read actual data.
    /// </summary>
    public bool IsPythonObject { get; set; }

    /// <summary>
    /// The Python object's type name (e.g. "DataFrame"), only populated when <see cref="IsPythonObject"/> is true.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? TypeName { get; set; }

    /// <summary>
    /// True if the Python code itself raised an error (surfaces in Excel as #PYTHON!).
    /// </summary>
    public bool IsPythonError { get; set; }

    /// <summary>
    /// Informational message (e.g. explaining Python Object limitations or polling timeout).
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }
}
