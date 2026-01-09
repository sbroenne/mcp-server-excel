using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result for getting Data Model summary information
/// </summary>
public class DataModelInfoResult : ResultBase
{
    /// <summary>
    /// Number of tables in the model
    /// </summary>
    [JsonPropertyName("tc")]
    public int TableCount { get; set; }

    /// <summary>
    /// Number of DAX measures in the model
    /// </summary>
    [JsonPropertyName("mc")]
    public int MeasureCount { get; set; }

    /// <summary>
    /// Number of relationships in the model
    /// </summary>
    [JsonPropertyName("rlc")]
    public int RelationshipCount { get; set; }

    /// <summary>
    /// Total number of rows across all tables
    /// </summary>
    [JsonPropertyName("tr")]
    public int TotalRows { get; set; }

    /// <summary>
    /// List of table names
    /// </summary>
    [JsonPropertyName("tn")]
    public List<string> TableNames { get; set; } = [];
}
