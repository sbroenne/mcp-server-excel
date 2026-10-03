using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native duplicate removal counts, excluding the optional header row.</summary>
public sealed class RemoveDuplicatesResult : OperationResult
{
    /// <summary>Original selected rectangle.</summary>
    public string SourceRange { get; set; } = string.Empty;
    /// <summary>Removed data rows, including duplicate blank records.</summary>
    public int RemovedRows { get; set; }
    /// <summary>Retained data rows, including retained blank records.</summary>
    public int RemainingRows { get; set; }
    /// <summary>Retained rectangle, including the optional header.</summary>
    public string RemainingRange { get; set; } = string.Empty;
}

/// <summary>Complete output geometry established by native parsing before the customer write.</summary>
public sealed class TextToColumnsResult : OperationResult
{
    /// <summary>Original single-column input rectangle.</summary>
    public string SourceRange { get; set; } = string.Empty;
    /// <summary>Complete output rectangle, including empty fields.</summary>
    public string DestinationRange { get; set; } = string.Empty;
    /// <summary>Output column count.</summary>
    public int OutputColumns { get; set; }
}
