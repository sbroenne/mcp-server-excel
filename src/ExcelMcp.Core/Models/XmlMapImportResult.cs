namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result from importing in-memory XML data.
/// </summary>
public sealed class XmlMapImportResult : ResultBase
{
    /// <summary>
    /// XML map used or created by Excel.
    /// </summary>
    public string MapName { get; set; } = string.Empty;

    /// <summary>
    /// Excel XML import result.
    /// </summary>
    public string ImportStatus { get; set; } = string.Empty;

    /// <summary>
    /// Destination worksheet for an automatically mapped import.
    /// </summary>
    public string? SheetName { get; set; }

    /// <summary>
    /// Top-left destination cell for an automatically mapped import.
    /// </summary>
    public string? StartCell { get; set; }
}
