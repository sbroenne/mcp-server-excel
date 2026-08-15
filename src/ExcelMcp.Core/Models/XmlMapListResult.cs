namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result containing workbook XML maps.
/// </summary>
public sealed class XmlMapListResult : ResultBase
{
    /// <summary>
    /// XML maps in workbook order.
    /// </summary>
    public List<XmlMapInfo> Maps { get; set; } = [];
}
