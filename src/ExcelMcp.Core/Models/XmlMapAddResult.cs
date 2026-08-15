namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result from adding an XML map.
/// </summary>
public sealed class XmlMapAddResult : ResultBase
{
    /// <summary>
    /// Name assigned to the XML map.
    /// </summary>
    public string MapName { get; set; } = string.Empty;

    /// <summary>
    /// Root element defined by the map schema.
    /// </summary>
    public string RootElementName { get; set; } = string.Empty;
}
