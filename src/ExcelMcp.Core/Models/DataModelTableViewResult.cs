
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
    /// Name of the workbook connection that supplies this table.
    /// </summary>
    public string SourceConnectionName { get; set; } = "";

    /// <summary>
    /// Description of the source workbook connection.
    /// </summary>
    public string SourceConnectionDescription { get; set; } = "";

    /// <summary>
    /// Human-readable source connection type.
    /// </summary>
    public string SourceConnectionType { get; set; } = "";

    /// <summary>
    /// Excel XlConnectionType numeric value for the source connection.
    /// </summary>
    public int SourceConnectionTypeValue { get; set; }

    /// <summary>
    /// Whether Excel reports the source connection as participating in the model.
    /// </summary>
    public bool SourceConnectionInModel { get; set; }

    /// <summary>
    /// Number of rows in the table
    /// </summary>
    public int RecordCount { get; set; }

    /// <summary>
    /// List of columns in the table
    /// </summary>
    public List<DataModelColumnInfo> Columns { get; set; } = [];

    /// <summary>
    /// Number of measures defined in this table
    /// </summary>
    public int MeasureCount { get; set; }
}
