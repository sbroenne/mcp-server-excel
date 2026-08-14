namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>
/// External workbook link.
/// </summary>
public sealed class ExternalLinkInfo
{
    /// <summary>External workbook source path as reported by Excel.</summary>
    public string Source { get; set; } = string.Empty;

    /// <summary>External link type.</summary>
    public string LinkType { get; set; } = "excel";
}
