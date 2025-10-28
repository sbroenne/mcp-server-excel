namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result for listing table columns
/// </summary>
public class DataModelTableColumnsResult : ResultBase
{
    /// <summary>
    /// Table name
    /// </summary>
    public string TableName { get; set; } = "";

    /// <summary>
    /// List of columns in the table
    /// </summary>
    public List<DataModelColumnInfo> Columns { get; set; } = new();
}
