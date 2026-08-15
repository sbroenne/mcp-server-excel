using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>
/// Result containing workbook document properties.
/// </summary>
public sealed class DocumentPropertyListResult : ResultBase
{
    /// <summary>Workbook document properties.</summary>
    public List<DocumentPropertyInfo> Properties { get; set; } = [];
}
