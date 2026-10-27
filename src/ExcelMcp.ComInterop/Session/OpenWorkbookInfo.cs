namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>An already-open workbook selection valid within the discovering host.</summary>
/// <param name="SessionId">Existing session identifier when the workbook is already connected to this host.</param>
/// <param name="ProtectedView">Whether Excel is displaying the workbook in Protected View.</param>
public sealed record OpenWorkbookInfo(
    string WorkbookId,
    string Name,
    string? Location,
    int ProcessId,
    bool Saved,
    bool? AutoSaveOn,
    string? SessionId = null,
    bool ProtectedView = false);

/// <summary>Details for a workbook attached to an active session.</summary>
/// <param name="SessionId">The session identifier for subsequent workbook operations.</param>
/// <param name="Name">The workbook's current Excel name.</param>
/// <param name="FilePath">The workbook location, or null if it has never been saved.</param>
/// <param name="AutoSaveOn">The current Excel AutoSave setting, or null if unavailable.</param>
public sealed record AttachedWorkbookSession(
    string SessionId,
    string Name,
    string? FilePath,
    bool? AutoSaveOn);

/// <summary>An Excel application or workbook that could not be inspected.</summary>
public sealed record WorkbookDiscoveryError(int? ProcessId, string Message);

/// <summary>Open workbooks and explicit discovery gaps in the current user's session.</summary>
public sealed record WorkbookDiscoveryResult(
    IReadOnlyList<OpenWorkbookInfo> Workbooks,
    IReadOnlyList<WorkbookDiscoveryError> Errors)
{
    /// <summary>Whether all located Excel applications were inspected successfully.</summary>
    public bool Complete => Errors.Count == 0;
}

internal sealed record AttachedWorkbookDiscovery(
    IReadOnlyList<OpenWorkbookInfo> Workbooks,
    IReadOnlyList<WorkbookDiscoveryError> Errors);
