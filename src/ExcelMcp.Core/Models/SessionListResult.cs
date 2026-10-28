using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>Managed workbook sessions and their current operation state.</summary>
public sealed class SessionListResult : ResultBase
{
    /// <summary>CLI daemon state when inspected by the command-line entry point.</summary>
    [ResultContext]
    public string? DaemonState { get; set; }
    /// <summary>Sessions known to the current host.</summary>
    public List<SessionListEntry> Sessions { get; set; } = [];

    /// <summary>Number of sessions in the complete listing.</summary>
    [ResultContext]
    public int Count => Sessions.Count;
}

/// <summary>Read-only workbook restrictions and explicit inspection gaps.</summary>
public sealed class SessionNoticesResult : ResultBase
{
    /// <summary>Restrictions, unknown states, and guidance that must remain visible.</summary>
    [ResultContext]
    public WorkbookNoticeStatus? Notices { get; set; }
}

/// <summary>Identity and safety context for one managed workbook.</summary>
public sealed class SessionListEntry
{
    /// <summary>Session identifier for subsequent operations.</summary>
    [ResultContext]
    public string SessionId { get; init; } = "";
    /// <summary>Current workbook location.</summary>
    public string? FilePath { get; init; }
    /// <summary>Locations retained by the session owner.</summary>
    public IReadOnlyList<string> FilePaths { get; init; } = [];
    /// <summary>Workbook name reported by Excel.</summary>
    public string? WorkbookName { get; init; }
    /// <summary>Whether the session attaches to an existing workbook.</summary>
    public bool IsAttached { get; init; }
    /// <summary>Whether Excel is visible.</summary>
    public bool IsExcelVisible { get; init; }
    /// <summary>Number of active operations.</summary>
    [ResultContext]
    public int ActiveOperations { get; init; }
    /// <summary>Whether Excel execution remains pending.</summary>
    [ResultContext]
    public bool OperationPending { get; init; }
    /// <summary>Whether the owner can currently close the workbook.</summary>
    [ResultContext]
    public bool CanClose { get; init; }
    /// <summary>Current Excel readiness state.</summary>
    [ResultContext]
    public string ExcelState { get; init; } = "";
    /// <summary>Reason the workbook cannot currently be used or closed.</summary>
    [ResultContext]
    public string? BlockingReason { get; init; }
    /// <summary>Visible blocking context, when readable.</summary>
    [ResultContext]
    public string? BlockingMessage { get; init; }
    /// <summary>Pending or retained operation state.</summary>
    [ResultContext]
    public ExcelOperationStatus? Operation { get; init; }
}
