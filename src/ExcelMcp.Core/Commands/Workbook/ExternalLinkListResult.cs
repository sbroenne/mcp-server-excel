using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>
/// Result containing external workbook links.
/// </summary>
public sealed class ExternalLinkListResult : ResultBase
{
    /// <summary>External Excel workbook links.</summary>
    public List<ExternalLinkInfo> Links { get; set; } = [];
}
