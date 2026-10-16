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
    Description = "Read and change per-user Excel settings that apply to every workbook, including Power Query tracing (the same setting as Query Options, Diagnostics, Enable tracing); no workbook session is needed. Changing tracing affects only Excel processes started afterwards, so open a new session after enabling it. Open/create with pq_tracing=true (CLI: --pq-tracing) to temporarily enable tracing and restore the prior setting when the last opted-in session closes. This is a per-user switch, so any Excel process started while it is on may also write trace files. Use powerquery_trace_read to list or summarize trace files. Trace files can contain data-source names and native query text.")]
[McpReadOnlyActions("get-pq-tracing")]
public interface ISettingsCommands
{
    /// <summary>
    /// Reads whether Power Query tracing is switched on, its level, the trace folder, and how many trace files it holds.
    /// </summary>
    [ServiceAction("get-pq-tracing")]
    PowerQueryTracingStatusResult GetPqTracing();

    /// <summary>
    /// Switches Power Query tracing on or off for Excel processes started afterwards.
    /// Returns the previous values.
    /// </summary>
    /// <param name="enabled">True to switch tracing on, false to switch it off</param>
    /// <param name="level">Optional tracing level to store (zero or greater). Omit to keep the current level; Excel's default is 15.</param>
    [ServiceAction("set-pq-tracing")]
    PowerQueryTracingStatusResult SetPqTracing([RequiredParameter] bool enabled, int? level = null);
}
