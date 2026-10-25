using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Bounded traversal of relationships returned by native worksheet getters.</summary>
public sealed class RangeFormulaTraceResult : OperationResult
{
    /// <summary>Resolved worksheet.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Exact starting scope, including disjoint areas.</summary>
    public string RangeAddress { get; set; } = string.Empty;
    /// <summary>precedents or dependents.</summary>
    public string Direction { get; set; } = string.Empty;
    /// <summary>Included cells and bounded formula/value previews, including retained starting cells.</summary>
    public List<FormulaTraceNode> Nodes { get; set; } = [];
    /// <summary>Distinct native directed relationships.</summary>
    public List<FormulaTraceEdge> Edges { get; set; } = [];
    /// <summary>Cycles in returned edges, including self-references; omitted edges can hide other cycles.</summary>
    public List<List<string>> Cycles { get; set; } = [];
    /// <summary>Native lookups that could not establish whether relationships exist.</summary>
    public List<FormulaTraceUnresolved> Unresolved { get; set; } = [];
    /// <summary>Effective caller-selected limits.</summary>
    public FormulaTraceLimits Limits { get; set; } = new(3, 50, 256);
    /// <summary>Separate graph and preview omissions; reaching a bound alone is not an omission.</summary>
    public FormulaTraceTruncation Truncation { get; set; } = new();
    /// <summary>Where enumeration stopped, not an exact count of excluded cells or relationships.</summary>
    public List<FormulaTraceBoundary> Boundaries { get; set; } = [];
    /// <summary>Native scope and semantic limitations; not a complete workbook dependency graph.</summary>
    public FormulaTraceCoverage Coverage { get; set; } = new();
}

/// <summary>An included cell with minimum depth from retained roots and flagged formula/value previews.</summary>
public sealed record FormulaTraceNode(
    string Address, int Row, int Column, string? Formula, object? Value, bool IsRoot,
    int Depth, bool FormulaTruncated, bool ValueTruncated);

/// <summary>Limits shared by both entry points; there is no unlimited setting.</summary>
public sealed record FormulaTraceLimits(int MaxDepth, int MaxNodes, int MaxPreviewCharacters);

/// <summary>Observed omissions, separate from Excel's inherent native coverage limitations.</summary>
public sealed class FormulaTraceTruncation
{
    /// <summary>At least one new related cell was excluded by depth.</summary>
    public bool DepthLimitReached { get; set; }
    /// <summary>At least one new cell was excluded by the returned-cell budget.</summary>
    public bool NodeLimitReached { get; set; }
    /// <summary>Enumeration of starting cells stopped at an excluded distinct cell.</summary>
    public bool StartingCellsOmitted { get; set; }
    /// <summary>At least one returned formula or text value was shortened.</summary>
    public bool PreviewTextTruncated { get; set; }
    /// <summary>True if graph or preview output was shortened.</summary>
    public bool Any => DepthLimitReached || NodeLimitReached || PreviewTextTruncated;
}

/// <summary>A stopped starting scope or source lookup, with max-depth or max-nodes as its reason.</summary>
public sealed record FormulaTraceBoundary(string Address, string Reason);

/// <summary>A relationship in the requested traversal direction.</summary>
public sealed record FormulaTraceEdge(string FromAddress, string ToAddress);

/// <summary>An ambiguous native absence or unavailable lookup, not a fabricated empty relation.</summary>
public sealed record FormulaTraceUnresolved(string Address, string Reason, string NativeErrorCode);

/// <summary>Explicit limitations of Excel's native direct relationship getters.</summary>
public sealed class FormulaTraceCoverage
{
    /// <summary>Only references on the resolved worksheet can be returned.</summary>
    public string Scope { get; } = "same-worksheet-only";
    /// <summary>Always false: native getters cannot establish workbook-wide coverage.</summary>
    public bool WorkbookComplete { get; }
    /// <summary>False when graph traversal was shortened or a native lookup remained unresolved; not affected by preview clipping.</summary>
    public bool NativeTraversalComplete { get; set; }
    /// <summary>Missing coverage inherent in Excel's getters, not output truncation.</summary>
    public IReadOnlyList<string> Limitations { get; } =
    [
        "Cross-worksheet and external-workbook references are not returned.",
        "Dynamic references are not fully resolved: INDIRECT can be omitted and OFFSET can return only its anchor.",
        "A native no-range error cannot distinguish no relationships from unavailable relationships.",
        "Values and relationships reflect current calculated state; no recalculation, refresh, or external workbook opening is performed."
    ];
}
