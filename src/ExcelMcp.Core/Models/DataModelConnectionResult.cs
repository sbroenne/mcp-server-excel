namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Metadata for the workbook's embedded Data Model connection.
/// </summary>
public class DataModelConnectionResult : ResultBase
{
    /// <summary>
    /// Name of the embedded model.
    /// </summary>
    public string ModelName { get; set; } = string.Empty;

    /// <summary>
    /// Name of the workbook connection that represents the embedded model.
    /// </summary>
    public string ConnectionName { get; set; } = string.Empty;

    /// <summary>
    /// Connection description.
    /// </summary>
    public string Description { get; set; } = string.Empty;

    /// <summary>
    /// Human-readable connection type.
    /// </summary>
    public string ConnectionType { get; set; } = string.Empty;

    /// <summary>
    /// Excel XlConnectionType numeric value.
    /// </summary>
    public int ConnectionTypeValue { get; set; }

    /// <summary>
    /// Whether Excel reports the connection as participating in the model.
    /// </summary>
    public bool InModel { get; set; }

    /// <summary>
    /// Human-readable ModelConnection command type.
    /// </summary>
    public string CommandType { get; set; } = string.Empty;

    /// <summary>
    /// Excel XlCmdType numeric value.
    /// </summary>
    public int CommandTypeValue { get; set; }

    /// <summary>
    /// ModelConnection command text, when Excel exposes one.
    /// </summary>
    public string? CommandText { get; set; }

    /// <summary>
    /// Tables exposed through the model workbook connection.
    /// </summary>
    public List<string> TableNames { get; set; } = [];
}
