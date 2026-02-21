using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result of adding an Excel Table to the Data Model.
/// Extends OperationResult with information about bracket-escaped column names.
/// </summary>
public class AddToDataModelResult : OperationResult
{
    /// <summary>
    /// Column names that contain literal bracket characters and cannot be referenced in DAX without escaping.
    /// Populated when stripBracketColumnNames is false and bracket column names are found.
    /// Empty when no bracket column names are present or when stripBracketColumnNames is true.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string[]? BracketColumnsFound { get; set; }

    /// <summary>
    /// Column names that were renamed (bracket characters removed) before adding to the Data Model.
    /// Populated only when stripBracketColumnNames is true and bracket column names were found.
    /// Each entry is the original column name (with brackets). The renamed version removes the brackets.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string[]? BracketColumnsRenamed { get; set; }
}
