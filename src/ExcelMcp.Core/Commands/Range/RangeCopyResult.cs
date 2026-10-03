using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>The resolved native paste bounds and selected behavior.</summary>
public sealed class RangeCopyResult : OperationResult
{
    /// <summary>The source worksheet.</summary>
    public string SourceSheet { get; set; } = string.Empty;
    /// <summary>The resolved absolute source.</summary>
    public string SourceAddress { get; set; } = string.Empty;
    /// <summary>The destination worksheet.</summary>
    public string TargetSheet { get; set; } = string.Empty;
    /// <summary>The complete expanded/repeated destination bounds.</summary>
    public string DestinationAddress { get; set; } = string.Empty;
    /// <summary>The explicit native paste kind.</summary>
    public PasteKind PasteKind { get; set; }
    /// <summary>Whether source rows and columns were exchanged.</summary>
    public bool Transpose { get; set; }
    /// <summary>Whether native blank source cells were skipped.</summary>
    public bool SkipBlanks { get; set; }
}
