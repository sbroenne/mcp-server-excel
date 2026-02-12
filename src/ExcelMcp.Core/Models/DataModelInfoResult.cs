
namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result for getting Data Model summary information
/// </summary>
public class DataModelInfoResult : ResultBase
{
    /// <summary>
    /// Number of tables in the model
    /// </summary>
    public int TableCount { get; set; }

    /// <summary>
    /// Number of DAX measures in the model
    /// </summary>
    public int MeasureCount { get; set; }

    /// <summary>
    /// Number of relationships in the model
    /// </summary>
    public int RelationshipCount { get; set; }

    /// <summary>
    /// Total number of rows across all tables
    /// </summary>
    public int TotalRows { get; set; }

    /// <summary>
    /// List of table names
    /// </summary>
    public List<string> TableNames { get; set; } = [];
}


