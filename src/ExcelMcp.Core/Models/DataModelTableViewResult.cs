using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result for viewing table details
/// </summary>
public class DataModelTableViewResult : ResultBase
{
    /// <summary>
    /// Table name
    /// </summary>
    [JsonPropertyName("tn")]
    public string TableName { get; set; } = "";

    /// <summary>
    /// Source query or connection name
    /// </summary>
    [JsonPropertyName("src")]
    public string SourceName { get; set; } = "";

    /// <summary>
    /// Number of rows in the table
    /// </summary>
    [JsonPropertyName("rc")]
    public int RecordCount { get; set; }

    /// <summary>
    /// List of columns in the table
    /// </summary>
    [JsonPropertyName("col")]
    public List<DataModelColumnInfo> Columns { get; set; } = [];

    /// <summary>
    /// Number of measures defined in this table
    /// </summary>
    [JsonPropertyName("mc")]
    public int MeasureCount { get; set; }
}
