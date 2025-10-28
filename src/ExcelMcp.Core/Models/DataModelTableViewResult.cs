namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result for viewing table details
/// </summary>
public class DataModelTableViewResult : ResultBase
{
    /// <summary>
    /// Table name
    /// </summary>
    public string TableName { get; set; } = "";

    /// <summary>
    /// Source query or connection name
    /// </summary>
    public string SourceName { get; set; } = "";

    /// <summary>
    /// Number of rows in the table
    /// </summary>
    public int RecordCount { get; set; }

    /// <summary>
    /// Last refresh date/time (if available)
    /// </summary>
    public DateTime? RefreshDate { get; set; }

    /// <summary>
    /// List of columns in the table
    /// </summary>
    public List<DataModelColumnInfo> Columns { get; set; } = new();

    /// <summary>
    /// Number of measures defined in this table
    /// </summary>
    public int MeasureCount { get; set; }
}
