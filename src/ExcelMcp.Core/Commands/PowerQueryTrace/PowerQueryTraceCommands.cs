using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;

/// <summary>
/// Commands for reading Power Query trace files.
/// </summary>
public sealed class PowerQueryTraceCommands : IPowerQueryTraceCommands
{
    private readonly PowerQueryTraceReader _reader;
    private readonly Func<string, int?> _resolveSessionProcessId;
    private readonly Func<DateTime> _utcNow;

    /// <summary>Creates commands using the default trace folder.</summary>
    public PowerQueryTraceCommands()
        : this(new RegistryPowerQueryTraceSettings().TraceFolder, _ => null, () => DateTime.UtcNow)
    {
    }

    /// <summary>Creates commands with explicit local dependencies.</summary>
    public PowerQueryTraceCommands(
        string traceFolder,
        Func<string, int?> resolveSessionProcessId,
        Func<DateTime>? utcNow = null)
    {
        _reader = new PowerQueryTraceReader(traceFolder);
        _resolveSessionProcessId = resolveSessionProcessId ?? throw new ArgumentNullException(nameof(resolveSessionProcessId));
        _utcNow = utcNow ?? (() => DateTime.UtcNow);
    }

    /// <inheritdoc />
    public PowerQueryTraceListResult List(string? workbookSessionId = null, int? processId = null, string? since = null)
    {
        int? targetProcessId = ResolveProcessId(workbookSessionId, processId);
        DateTime? sinceUtc = PowerQueryTraceTime.Parse(since, _utcNow(), nameof(since));
        return _reader.List(targetProcessId, sinceUtc);
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
        int? targetProcessId = ResolveProcessId(workbookSessionId, processId);
        DateTime now = _utcNow();
        DateTime? sinceUtc = PowerQueryTraceTime.Parse(since, now, nameof(since));
        DateTime? untilUtc = PowerQueryTraceTime.Parse(until, now, nameof(until));
        if (sinceUtc is DateTime start && untilUtc is DateTime end && end < start)
        {
            throw new ArgumentException("until must not be earlier than since.", nameof(until));
        }

        return _reader.Summarize(new PowerQueryTraceFilter(
            targetProcessId,
            activityId,
            sinceUtc,
            untilUtc,
            top,
            includeQueryText));
    }

    private int? ResolveProcessId(string? workbookSessionId, int? processId)
    {
        if (!string.IsNullOrWhiteSpace(workbookSessionId) && processId.HasValue)
        {
            throw new ArgumentException("Specify either workbookSessionId or processId, not both.");
        }

        if (string.IsNullOrWhiteSpace(workbookSessionId))
        {
            return processId;
        }

        int? resolved = _resolveSessionProcessId(workbookSessionId);
        return resolved ?? throw new ArgumentException(
            $"Workbook session '{workbookSessionId}' was not found or has no Excel process.",
            nameof(workbookSessionId));
    }
}
