using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result for listing table columns
/// </summary>
public class DataModelTableColumnsResult : ResultBase
{
    /// <summary>
    /// Table name
    /// </summary>
    [JsonPropertyName("tn")]
    public string TableName { get; set; } = "";

    /// <summary>
    /// List of columns in the table
    /// </summary>
    [JsonPropertyName("col")]
    public List<DataModelColumnInfo> Columns { get; set; } = [];
}
