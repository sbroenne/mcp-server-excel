using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Settings;

/// <summary>
/// Per-user Excel settings that apply to every workbook. No open workbook or session is needed.
/// Changes apply to Excel processes started afterwards; running Excel processes keep their current behavior.
/// </summary>
[ServiceCategory("Settings")]
[NoSession]
[McpTool("settings", Title = "Excel Settings", Destructive = true, Category = "settings",
    Description = "Read and change per-user Excel settings that apply to every workbook, including Power Query tracing and the global requirement for approval before new native database queries run. Changing native-query approval requires explicit confirmation and Excel to be closed; turning approval off reduces security. No workbook session is needed. Restart Excel for tracing changes to take effect. Any Excel process started while tracing is on may write trace files. Use powerquery_trace_read to list or summarize trace files, which can contain data-source names and native query text.")]
[McpReadOnlyActions("get-pq-tracing", "get-native-query-approval")]
public interface ISettingsCommands
{
    /// <summary>
    /// Reads whether Power Query tracing is switched on, its level, the trace folder, and how many trace files it holds.
    /// </summary>
    [ServiceAction("get-pq-tracing")]
    PowerQueryTracingStatusResult GetPqTracing();

    /// <summary>
    /// Switches Power Query tracing on or off for Excel processes started after Excel is restarted.
    /// Returns the previous values and reminds the caller to restart Excel.
    /// </summary>
    /// <param name="enabled">True to switch tracing on, false to switch it off</param>
    /// <param name="level">Optional tracing level to store (zero or greater). Omit to keep the current level; Excel's default is 15.</param>
    [ServiceAction("set-pq-tracing")]
    PowerQueryTracingStatusResult SetPqTracing([RequiredParameter] bool enabled, int? level = null);

    /// <summary>
    /// Reads whether new native database queries require approval in Power Query.
    /// This applies to all workbooks for the current Windows user.
    /// </summary>
    [ServiceAction("get-native-query-approval")]
    NativeQueryApprovalResult GetNativeQueryApproval();

    /// <summary>
    /// Changes the current user's global native database query approval setting.
    /// Excel must be closed, and the caller must explicitly confirm the security impact.
    /// </summary>
    /// <param name="required">True to require approval, false to allow native database queries without approval.</param>
    /// <param name="confirm">Must be true to authorize changing this global security setting.</param>
    [ServiceAction("set-native-query-approval")]
    NativeQueryApprovalResult SetNativeQueryApproval(
        [RequiredParameter] bool? required,
        [RequiredParameter] bool confirm);
}
