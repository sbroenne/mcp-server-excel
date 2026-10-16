using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;

/// <summary>
/// Commands for reading Power Query trace files.
/// </summary>
public sealed class PowerQueryTraceCommands : IPowerQueryTraceCommands
{
    private readonly PowerQueryTraceReader _reader;
    private readonly Func<string, (int ProcessId, DateTime StartedUtc)?> _resolveSessionProcess;
    private readonly Func<DateTime> _utcNow;

    /// <summary>Creates commands with explicit local dependencies.</summary>
    public PowerQueryTraceCommands(
        string traceFolder,
        Func<string, (int ProcessId, DateTime StartedUtc)?> resolveSessionProcess,
        Func<DateTime>? utcNow = null)
    {
        _reader = new PowerQueryTraceReader(traceFolder);
        _resolveSessionProcess = resolveSessionProcess ?? throw new ArgumentNullException(nameof(resolveSessionProcess));
        _utcNow = utcNow ?? (() => DateTime.UtcNow);
    }

    /// <inheritdoc />
    public PowerQueryTraceListResult List(string? workbookSessionId = null, int? processId = null, string? since = null)
    {
        var targetProcess = ResolveProcess(workbookSessionId, processId);
        DateTime? sinceUtc = PowerQueryTraceTime.Parse(since, _utcNow(), nameof(since));
        return _reader.List(targetProcess.ProcessId, sinceUtc, targetProcess.StartedUtc);
    }

    /// <inheritdoc />
    public PowerQueryTraceSummaryResult Summarize(
        string? workbookSessionId = null,
        int? processId = null,
        string? activityId = null,
        string? since = null,
        string? until = null,
        int top = PowerQueryTraceAnalyzer.DefaultTop,
        bool includeQueryText = false)
    {
        var targetProcess = ResolveProcess(workbookSessionId, processId);
        DateTime now = _utcNow();
        DateTime? sinceUtc = PowerQueryTraceTime.Parse(since, now, nameof(since));
        DateTime? untilUtc = PowerQueryTraceTime.Parse(until, now, nameof(until));
        if (sinceUtc is DateTime start && untilUtc is DateTime end && end < start)
        {
            throw new ArgumentException("until must not be earlier than since.", nameof(until));
        }

        return _reader.Summarize(new PowerQueryTraceFilter(
            ProcessId: targetProcess.ProcessId,
            ActivityId: activityId,
            SinceUtc: sinceUtc,
            UntilUtc: untilUtc,
            Top: top,
            IncludeQueryText: includeQueryText,
            ProcessStartedUtc: targetProcess.StartedUtc));
    }

    private (int? ProcessId, DateTime? StartedUtc) ResolveProcess(string? workbookSessionId, int? processId)
    {
        if (!string.IsNullOrWhiteSpace(workbookSessionId) && processId.HasValue)
        {
            throw new ArgumentException("Specify either workbookSessionId or processId, not both.");
        }

        if (string.IsNullOrWhiteSpace(workbookSessionId))
        {
            return (processId, null);
        }

        var resolved = _resolveSessionProcess(workbookSessionId);
        return resolved is { } sessionProcess
            ? (sessionProcess.ProcessId, sessionProcess.StartedUtc)
            : throw new ArgumentException(
            $"Workbook session '{workbookSessionId}' was not found or its Excel process identity is unavailable.",
            nameof(workbookSessionId));
    }
}
