using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Complete traversal of relationships returned by native worksheet getters.</summary>
public sealed class RangeFormulaTraceResult : OperationResult
{
    /// <summary>Resolved worksheet.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Exact starting scope, including disjoint areas.</summary>
    public string RangeAddress { get; set; } = string.Empty;
    /// <summary>precedents or dependents.</summary>
    public string Direction { get; set; } = string.Empty;
    /// <summary>Every reached cell, including all starting cells.</summary>
    public List<FormulaTraceNode> Nodes { get; set; } = [];
    /// <summary>Distinct native directed relationships.</summary>
    public List<FormulaTraceEdge> Edges { get; set; } = [];
    /// <summary>Strongly connected components with cycles, including self-references.</summary>
    public List<List<string>> Cycles { get; set; } = [];
    /// <summary>Native lookups that could not establish whether relationships exist.</summary>
    public List<FormulaTraceUnresolved> Unresolved { get; set; } = [];
    /// <summary>Native scope and semantic limitations; not a complete workbook dependency graph.</summary>
    public FormulaTraceCoverage Coverage { get; set; } = new();
}

/// <summary>A reached cell and its current native formula/value.</summary>
public sealed record FormulaTraceNode(
    string Address, int Row, int Column, string? Formula, object? Value, bool IsRoot);

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
    /// <summary>False when an ambiguous native lookup remained unresolved.</summary>
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
