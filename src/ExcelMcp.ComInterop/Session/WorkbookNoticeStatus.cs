using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>Verified workbook restrictions and explicit limits of notice inspection.</summary>
/// <param name="ProtectedView">Whether Excel is displaying this workbook in Protected View; null when unknown.</param>
/// <param name="ReadOnly">The workbook's current read-only setting; null when unknown.</param>
/// <param name="HasVbaProject">Whether a VBA project exists, not whether it can run.</param>
/// <param name="VbaState">Known opening state, or unknown for workbooks opened outside this session.</param>
/// <param name="ConnectionCount">Number of workbook connections, not evidence that they are enabled.</param>
/// <param name="ConnectionsState">Whether enabling was explicitly requested in this session, otherwise unknown.</param>
/// <param name="Complete">Whether every notice could be inspected.</param>
/// <param name="Messages">Findings, inspection limits, and manual steps. Not telemetry.</param>
public sealed record WorkbookNoticeStatus(
    [property: JsonIgnore(Condition = JsonIgnoreCondition.Never)] bool? ProtectedView,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.Never)] bool? ReadOnly,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.Never)] bool? HasVbaProject,
    string VbaState,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.Never)] int? ConnectionCount,
    string ConnectionsState,
    bool Complete,
    IReadOnlyList<string> Messages);

/// <summary>The observed result of an explicitly requested workbook security action.</summary>
/// <param name="Action">The action that completed.</param>
/// <param name="Changed">Whether the opening state changed; null when Excel does not expose a before/after comparison.</param>
/// <param name="TrustMayBeRemembered">Whether Excel can remember trust for this workbook.</param>
/// <param name="Reopened">Whether the saved workbook was closed and reopened.</param>
/// <param name="Message">Consequences of the completed action.</param>
public sealed record WorkbookSecurityResult(
    string Action,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.Never)] bool? Changed,
    bool TrustMayBeRemembered, bool Reopened, string Message);
