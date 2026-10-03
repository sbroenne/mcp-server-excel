using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>Additional native page settings. Null leaves a setting unchanged; empty text clears it. Margins use points.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class PageSetupOptions
{
    /// <summary>Print range, including disjoint areas; empty clears the explicit print area.</summary>
    public string? PrintArea { get; set; }
    /// <summary>Repeated complete rows such as 1:2; empty clears them.</summary>
    public string? PrintTitleRows { get; set; }
    /// <summary>Repeated complete columns such as A:B; empty clears them.</summary>
    public string? PrintTitleColumns { get; set; }
    /// <summary>Left margin in points.</summary>
    public double? LeftMargin { get; set; }
    /// <summary>Right margin in points.</summary>
    public double? RightMargin { get; set; }
    /// <summary>Top margin in points.</summary>
    public double? TopMargin { get; set; }
    /// <summary>Bottom margin in points.</summary>
    public double? BottomMargin { get; set; }
    /// <summary>Header margin in points.</summary>
    public double? HeaderMargin { get; set; }
    /// <summary>Footer margin in points.</summary>
    public double? FooterMargin { get; set; }
    /// <summary>Native left header text with Excel header codes; empty clears.</summary>
    public string? LeftHeader { get; set; }
    /// <summary>Native center header text with Excel header codes; empty clears.</summary>
    public string? CenterHeader { get; set; }
    /// <summary>Native right header text with Excel header codes; empty clears.</summary>
    public string? RightHeader { get; set; }
    /// <summary>Native left footer text with Excel footer codes; empty clears.</summary>
    public string? LeftFooter { get; set; }
    /// <summary>Native center footer text with Excel footer codes; empty clears.</summary>
    public string? CenterFooter { get; set; }
    /// <summary>Native right footer text with Excel footer codes; empty clears.</summary>
    public string? RightFooter { get; set; }
    /// <summary>Native XlPaperSize name, for example xlPaperA4. Driver support is required.</summary>
    public string? PaperSize { get; set; }
    /// <summary>Native XlOrder name: xlDownThenOver or xlOverThenDown.</summary>
    public string? PageOrder { get; set; }
    /// <summary>Print gridlines independently of on-screen visibility.</summary>
    public bool? PrintGridlines { get; set; }
    /// <summary>Print row/column headings independently of on-screen visibility.</summary>
    public bool? PrintHeadings { get; set; }
    /// <summary>Fixed scale from 10 through 400; cannot accompany fit-to-page parameters.</summary>
    public int? ZoomPercent { get; set; }
    /// <summary>Print in black and white.</summary>
    public bool? BlackAndWhite { get; set; }
    /// <summary>Enable draft printing.</summary>
    public bool? Draft { get; set; }
    /// <summary>First printed page number, or zero for Excel automatic numbering.</summary>
    public int? FirstPageNumber { get; set; }
    /// <summary>Native XlPrintLocation name controlling legacy notes.</summary>
    public string? PrintComments { get; set; }
    /// <summary>Native XlPrintErrors name controlling error display.</summary>
    public string? PrintErrors { get; set; }
    /// <summary>Scale headers and footers with the document.</summary>
    public bool? ScaleWithDocHeaderFooter { get; set; }
    /// <summary>Align headers and footers with page margins.</summary>
    public bool? AlignMarginsHeaderFooter { get; set; }
}

/// <summary>Replace all worksheet manual page breaks. Both lists are required; empty lists clear manual breaks.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class PageBreakOptions
{
    /// <summary>One-based rows before which to break, from 2 through 1048576.</summary>
    public required List<int> Rows { get; set; }
    /// <summary>One-based columns before which to break, from 2 through 16384.</summary>
    public required List<int> Columns { get; set; }
}

/// <summary>One native page break.</summary>
public sealed class PageBreakInfo
{
    /// <summary>One-based row or column before which the break occurs.</summary>
    public int Position { get; set; }
    /// <summary>Native location address.</summary>
    public string Address { get; set; } = string.Empty;
    /// <summary>Whether the break is manual rather than automatically calculated.</summary>
    public bool IsManual { get; set; }
    /// <summary>Native full or partial break extent.</summary>
    public string Extent { get; set; } = string.Empty;
}

/// <summary>All page breaks exposed by Excel for the worksheet's current print scope.</summary>
public sealed class SheetPageBreaksResult : ResultBase
{
    /// <summary>Worksheet name.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Current explicit print area; empty means native automatic scope.</summary>
    public string PrintArea { get; set; } = string.Empty;
    /// <summary>Native read coverage, not a claim to discover breaks outside the print scope.</summary>
    public string Coverage { get; } = "Excel exposes page breaks within the current print scope; automatic breaks depend on printer and scaling.";
    /// <summary>All exposed horizontal page breaks.</summary>
    public List<PageBreakInfo> Horizontal { get; set; } = [];
    /// <summary>All exposed vertical page breaks.</summary>
    public List<PageBreakInfo> Vertical { get; set; } = [];
}
