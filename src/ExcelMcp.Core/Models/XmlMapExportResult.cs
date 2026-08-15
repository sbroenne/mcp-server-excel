namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result from exporting mapped cells to in-memory XML.
/// </summary>
public sealed class XmlMapExportResult : ResultBase
{
    /// <summary>
    /// XML map that was exported.
    /// </summary>
    public string MapName { get; set; } = string.Empty;

    /// <summary>
    /// Exported XML data.
    /// </summary>
    public string XmlData { get; set; } = string.Empty;

    /// <summary>
    /// Excel XML export result.
    /// </summary>
    public string ExportStatus { get; set; } = string.Empty;
}
