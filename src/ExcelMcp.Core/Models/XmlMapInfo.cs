namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Describes an XML map in an Excel workbook.
/// </summary>
public sealed class XmlMapInfo
{
    /// <summary>
    /// XML map name.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Root element defined by the map schema.
    /// </summary>
    public string RootElementName { get; set; } = string.Empty;

    /// <summary>
    /// Whether Excel can export the current map.
    /// </summary>
    public bool IsExportable { get; set; }
}
