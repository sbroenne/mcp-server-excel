using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Results from Excel's available Document Inspector modules, not a full privacy audit.</summary>
public sealed class DocumentInspectionResult : ResultBase
{
    /// <summary>Coverage boundaries of Excel's callable inspector collection.</summary>
    public string ReadLimitations { get; set; } =
        "Only modules exposed by Excel's DocumentInspectors collection are checked. " +
        "Comments/revisions and document properties are not included. " +
        "Modules depend on the installed Excel version and custom inspectors; this is not a full privacy audit.";

    /// <summary>Every available inspector's native identity and inspection outcome.</summary>
    public List<DocumentInspectorFinding> Inspectors { get; set; } = [];
    /// <summary>Excel's saved flag before running the native inspector modules.</summary>
    public bool SavedBefore { get; set; }
    /// <summary>Excel's saved flag after inspection; native modules can mark the workbook unsaved without requesting a fix.</summary>
    public bool SavedAfter { get; set; }
}

/// <summary>A native inspector outcome, with the original status code and localized result text.</summary>
/// <param name="StatusCode">Native status: 0 means ok, 1 means issue found, 2 means error; unknown codes are retained.</param>
/// <param name="Results">Excel's native localized results, not a parsed list of removable items.</param>
public sealed record DocumentInspectorOutcome(int StatusCode, string Results)
{
    /// <summary>Stable status label, independent of Excel's display language.</summary>
    public string Status => StatusCode switch
    {
        0 => "ok",
        1 => "issue-found",
        2 => "error",
        _ => "unknown"
    };
}

/// <summary>One callable Document Inspector module and its native inspection outcome.</summary>
/// <param name="Index">One-based position in the live collection; not a stable category identifier.</param>
/// <param name="Name">Native localized module name, including any accelerator ampersands.</param>
/// <param name="StatusCode">Native inspection status.</param>
/// <param name="Results">Native localized inspection results.</param>
public sealed record DocumentInspectorFinding(int Index, string Name,
    int StatusCode, string Results)
{
    /// <summary>Stable status label, independent of Excel's display language.</summary>
    public string Status => new DocumentInspectorOutcome(StatusCode, Results).Status;
}

/// <summary>Native category-wide fix outcome and the selected module's subsequent inspection.</summary>
public sealed class DocumentInspectorFixResult : OperationResult
{
    /// <summary>Selected inspector's one-based position in the live collection.</summary>
    public int InspectorIndex { get; set; }
    /// <summary>Exact native name verified before attempting the fix.</summary>
    public string InspectorName { get; set; } = string.Empty;
    /// <summary>Native fix status and result text.</summary>
    public DocumentInspectorOutcome Fix { get; set; } = new(-1, string.Empty);
    /// <summary>Native inspection after the fix, including any remaining findings.</summary>
    public DocumentInspectorOutcome Inspection { get; set; } = new(-1, string.Empty);
}
