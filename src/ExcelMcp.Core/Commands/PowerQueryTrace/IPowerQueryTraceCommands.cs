using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;

/// <summary>
/// Lists and summarizes Power Query trace files without requiring a workbook.
/// </summary>
[ServiceCategory("PowerQueryTrace")]
[NoSession]
[McpTool("powerquery_trace", Title = "Power Query Trace Files", Destructive = false, Category = "query",
    Description = "List and summarize local Power Query trace files. Tracing is switched on and off with Excel Settings. Trace files may contain data-source names and native query text; summaries hide query text and exception details unless include_query_text is true.")]
[McpReadOnlyActions("list", "summarize")]
public interface IPowerQueryTraceCommands
{
    /// <summary>Lists trace files, newest first. A session filter includes its Excel and container trace files.</summary>
    /// <param name="workbookSessionId">Excel workbook session to list traces for</param>
    /// <param name="processId">Excel process ID to list traces for</param>
    /// <param name="since">Include files written, started, or still active since this ISO time or age (for example 15m)</param>
    [ServiceAction("list")]
    PowerQueryTraceListResult List(string? workbookSessionId = null, int? processId = null, string? since = null);

    /// <summary>Summarizes Power Query evaluations, data-source work, warnings, errors, and quiet periods.</summary>
    /// <param name="workbookSessionId">Excel workbook session to summarize traces for</param>
    /// <param name="processId">Excel process ID to summarize traces for</param>
    /// <param name="activityId">Limit the summary to one Power Query activity</param>
    /// <param name="since">Include entries ending after this ISO time, age, or traceCursor</param>
    /// <param name="until">Include entries starting before this ISO time or age</param>
    /// <param name="top">Maximum detailed items to return (1 through 50; default 5)</param>
    /// <param name="includeQueryText">Include native query text and exception messages; these may contain sensitive information</param>
    [ServiceAction("summarize")]
    PowerQueryTraceSummaryResult Summarize(
        string? workbookSessionId = null,
        int? processId = null,
        string? activityId = null,
        string? since = null,
        string? until = null,
        int top = PowerQueryTraceAnalyzer.DefaultTop,
        bool includeQueryText = false);
}
