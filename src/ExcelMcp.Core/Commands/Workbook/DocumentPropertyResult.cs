using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>
/// Result containing one workbook document property.
/// </summary>
public sealed class DocumentPropertyResult : ResultBase
{
    /// <summary>Requested document property.</summary>
    public DocumentPropertyInfo Property { get; set; } = new();
}
