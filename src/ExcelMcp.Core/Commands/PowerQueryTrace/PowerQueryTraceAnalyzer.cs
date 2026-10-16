using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;

/// <summary>
/// Filters applied while summarizing trace entries.
/// </summary>
internal sealed record PowerQueryTraceFilter(
    int? ProcessId = null,
    string? ActivityId = null,
    DateTime? SinceUtc = null,
    DateTime? UntilUtc = null,
    int Top = PowerQueryTraceAnalyzer.DefaultTop,
    bool IncludeQueryText = false);

/// <summary>
/// Builds a compact summary from Power Query trace entries.
/// </summary>
internal sealed partial class PowerQueryTraceAnalyzer
{
    public const int DefaultTop = 5;
    public const int MaxTop = 50;
    internal const double QuietGapSeconds = 10;
    private const int MaxCommandTextLength = 4000;
    private const int MaxSampleMessageLength = 300;
    private const string EvaluateAction = "SimpleDocumentEvaluator/GetResult/Evaluate";

    private readonly PowerQueryTraceFilter _filter;
    private readonly int _top;
    private readonly List<(DateTime Start, DateTime End)> _intervals = [];
    private readonly Dictionary<string, ActivityTotals> _activities = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<PowerQueryTraceEvaluation> _evaluations = [];
    private readonly Dictionary<(string Kind, string Path), SourceTotals> _sources = [];
    private readonly List<CommandRecord> _commands = [];
    private readonly Dictionary<(string Kind, string Action, string? ExceptionType), PowerQueryTraceIssue> _issues = [];
    private DateTime? _firstStart;
    private DateTime? _lastEnd;
    private int _entries;
    private int _files;
    private int _unreadable;

    public PowerQueryTraceAnalyzer(PowerQueryTraceFilter filter)
    {
        _filter = filter;
        _top = Math.Clamp(filter.Top <= 0 ? DefaultTop : filter.Top, 1, MaxTop);
    }

    public void AddFileRead() => _files++;

    public void AddUnreadableLine() => _unreadable++;

    public bool Matches(PowerQueryTraceEntry entry)
    {
        if (_filter.ProcessId is int pid && entry.ProcessId != pid && entry.HostProcessId != pid)
        {
            return false;
        }

        if (_filter.ActivityId is { Length: > 0 } activityId &&
            !string.Equals(entry.ActivityId, activityId, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (_filter.SinceUtc is DateTime since && entry.EndUtc <= since)
        {
            return false;
        }

        if (_filter.UntilUtc is DateTime until && entry.StartUtc > until)
        {
            return false;
        }

        return true;
    }

    public void Add(PowerQueryTraceEntry entry)
    {
        if (!Matches(entry))
        {
            return;
        }

        _entries++;
        var start = entry.StartUtc;
        var end = entry.EndUtc;
        if (_firstStart is null || start < _firstStart)
        {
            _firstStart = start;
        }

        if (_lastEnd is null || end > _lastEnd)
        {
            _lastEnd = end;
        }

        _intervals.Add((start, end));

        ActivityTotals? activity = null;
        if (!string.IsNullOrEmpty(entry.ActivityId) && entry.ActivityId != Guid.Empty.ToString())
        {
            if (!_activities.TryGetValue(entry.ActivityId, out activity))
            {
                activity = new ActivityTotals(entry.ActivityId, start, end);
                _activities.Add(entry.ActivityId, activity);
            }

            activity.Include(start, end);
            if (entry.HostProcessId is int host && host != entry.ProcessId)
            {
                activity.Containers.Add(entry.ProcessId);
            }
        }

        if (entry.Action.EndsWith(EvaluateAction, StringComparison.OrdinalIgnoreCase))
        {
            _evaluations.Add(new PowerQueryTraceEvaluation
            {
                ActivityId = entry.ActivityId,
                StartUtc = start,
                Seconds = Seconds(entry.Duration),
                ProcessId = entry.ProcessId
            });
            if (activity != null)
            {
                activity.Evaluations++;
            }
        }

        if (!string.IsNullOrEmpty(entry.ResourceKind))
        {
            AddDataSourceEntry(entry, activity);
        }

        if (!string.IsNullOrEmpty(entry.Exception) ||
            entry.Cancelled ||
            entry.Level.Equals("Error", StringComparison.OrdinalIgnoreCase) ||
            entry.Level.Equals("Warning", StringComparison.OrdinalIgnoreCase))
        {
            AddIssue(entry);
        }
    }

    private void AddDataSourceEntry(PowerQueryTraceEntry entry, ActivityTotals? activity)
    {
        var key = (entry.ResourceKind!, entry.ResourcePath ?? string.Empty);
        if (!_sources.TryGetValue(key, out var source))
        {
            source = new SourceTotals();
            _sources.Add(key, source);
        }

        if (entry.Action.EndsWith("/Connection/Open", StringComparison.OrdinalIgnoreCase))
        {
            source.ConnectionOpens++;
            return;
        }

        if (!IsDataSourceCall(entry))
        {
            return;
        }

        double seconds = Seconds(entry.Duration);
        source.Calls++;
        source.TotalSeconds += seconds;
        source.MaxSeconds = Math.Max(source.MaxSeconds, seconds);
        if (activity != null)
        {
            activity.DataSourceCalls++;
        }

        string hash = string.Empty;
        int length = 0;
        if (entry.CommandText != null)
        {
            hash = HashCommand(entry.CommandText);
            length = entry.CommandText.Length;
            source.CommandCounts[hash] = source.CommandCounts.GetValueOrDefault(hash) + 1;
        }

        _commands.Add(new CommandRecord(
            key.Item1,
            key.Item2,
            entry.ActivityId,
            entry.StartUtc,
            seconds,
            hash,
            length,
            _filter.IncludeQueryText ? Shorten(entry.CommandText, MaxCommandTextLength) : null));
    }

    private static bool IsDataSourceCall(PowerQueryTraceEntry entry)
    {
        if (entry.CommandText != null)
        {
            return true;
        }

        int lastSlash = entry.Action.LastIndexOf('/');
        string lastSegment = lastSlash >= 0 ? entry.Action[(lastSlash + 1)..] : entry.Action;
        return lastSegment.StartsWith("Execute", StringComparison.OrdinalIgnoreCase) ||
            entry.Action.Contains("/Request/", StringComparison.OrdinalIgnoreCase);
    }

    private void AddIssue(PowerQueryTraceEntry entry)
    {
        string? exceptionType = null;
        if (entry.Exception != null)
        {
            var typeMatch = ExceptionTypePattern().Match(entry.Exception);
            if (typeMatch.Success)
            {
                exceptionType = typeMatch.Groups["type"].Value.Trim();
            }
        }

        string kind = ClassifyIssue(entry, exceptionType);
        var key = (kind, entry.Action, exceptionType);
        if (!_issues.TryGetValue(key, out var issue))
        {
            issue = new PowerQueryTraceIssue
            {
                Kind = kind,
                Action = entry.Action,
                ExceptionType = exceptionType,
                FirstUtc = entry.StartUtc,
                LastUtc = entry.StartUtc,
                SampleMessage = _filter.IncludeQueryText ? SampleMessage(entry.Exception) : null
            };
            _issues.Add(key, issue);
        }

        issue.Count++;
        if (entry.StartUtc < issue.FirstUtc)
        {
            issue.FirstUtc = entry.StartUtc;
        }

        if (entry.StartUtc > issue.LastUtc)
        {
            issue.LastUtc = entry.StartUtc;
        }
    }

    internal static string ClassifyIssue(PowerQueryTraceEntry entry, string? exceptionType)
    {
        string text = entry.Exception ?? string.Empty;
        if (text.Contains("Firewall", StringComparison.OrdinalIgnoreCase))
        {
            return "firewall";
        }

        if (text.Contains("credential", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("authentication", StringComparison.OrdinalIgnoreCase))
        {
            return "credentials";
        }

        if (text.Contains("permission", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("access denied", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("UnauthorizedAccessException", StringComparison.Ordinal))
        {
            return "permission";
        }

        if (entry.Cancelled ||
            (exceptionType != null &&
             (exceptionType.EndsWith("OperationCanceledException", StringComparison.Ordinal) ||
              exceptionType.EndsWith("TaskCanceledException", StringComparison.Ordinal))))
        {
            return "cancelled";
        }

        return "error";
    }

    public PowerQueryTraceSummaryResult Build()
    {
        var result = new PowerQueryTraceSummaryResult
        {
            Success = true,
            ProcessId = _filter.ProcessId,
            ActivityId = string.IsNullOrEmpty(_filter.ActivityId) ? null : _filter.ActivityId,
            FilesRead = _files,
            EntriesMatched = _entries,
            UnreadableLines = _unreadable,
            Evaluations = _evaluations.Count,
            DataSourceCalls = _sources.Values.Sum(s => s.Calls)
        };

        DateTime? windowStart = _filter.SinceUtc ?? _firstStart;
        DateTime? windowEnd = _filter.UntilUtc ?? _lastEnd;
        if (windowStart is DateTime ws && windowEnd is DateTime we && we > ws)
        {
            result.WindowStartUtc = ws;
            result.WindowEndUtc = we;
            result.WindowSeconds = Round((we - ws).TotalSeconds);
        }

        if (_entries == 0)
        {
            result.Findings.Add(
                "No Power Query trace entries matched. Tracing may be switched off, the Excel process may have started before tracing was switched on, or nothing ran in Power Query during this time.");
            AddUnreadableFinding(result);
            return result;
        }

        result.TraceCursor = _lastEnd!.Value.ToString("o", CultureInfo.InvariantCulture);

        var merged = MergeIntervals(result.WindowStartUtc, result.WindowEndUtc);
        double busy = merged.Sum(i => (i.End - i.Start).TotalSeconds);
        result.PowerQueryBusySeconds = Round(busy);
        result.OutsidePowerQuerySeconds = Round(Math.Max(0, result.WindowSeconds - busy));

        var gaps = FindGaps(merged, result.WindowStartUtc, result.WindowEndUtc);
        result.QuietGaps = NullIfEmpty(gaps
            .OrderByDescending(g => g.Seconds)
            .Take(_top)
            .OrderBy(g => g.StartUtc)
            .ToList());

        result.Activities = NullIfEmpty(_activities.Values
            .OrderByDescending(a => a.End - a.Start)
            .Take(_top)
            .Select(a => new PowerQueryTraceActivity
            {
                ActivityId = a.Id,
                StartUtc = a.Start,
                EndUtc = a.End,
                Seconds = Round((a.End - a.Start).TotalSeconds),
                Evaluations = a.Evaluations,
                DataSourceCalls = a.DataSourceCalls,
                Containers = a.Containers.Count
            })
            .ToList());

        result.SlowestEvaluations = NullIfEmpty(_evaluations
            .OrderByDescending(e => e.Seconds)
            .Take(_top)
            .ToList());

        result.DataSources = NullIfEmpty(_sources
            .Where(s => s.Value.Calls > 0 || s.Value.ConnectionOpens > 0)
            .OrderByDescending(s => s.Value.TotalSeconds)
            .Take(_top)
            .Select(s => new PowerQueryTraceDataSource
            {
                Kind = s.Key.Kind,
                Path = s.Key.Path,
                Calls = s.Value.Calls,
                TotalSeconds = Round(s.Value.TotalSeconds),
                AverageSeconds = s.Value.Calls == 0 ? 0 : Round(s.Value.TotalSeconds / s.Value.Calls),
                MaxSeconds = Round(s.Value.MaxSeconds),
                ConnectionOpens = s.Value.ConnectionOpens,
                DistinctCommands = s.Value.CommandCounts.Count,
                RepeatedExecutions = s.Value.CommandCounts.Values.Sum(c => c - 1)
            })
            .ToList());

        result.SlowestCommands = NullIfEmpty(_commands
            .OrderByDescending(c => c.Seconds)
            .Take(_top)
            .Select(c => new PowerQueryTraceCommand
            {
                Kind = c.Kind,
                Path = c.Path,
                ActivityId = c.ActivityId,
                StartUtc = c.Start,
                Seconds = Round(c.Seconds),
                CommandHash = c.Hash,
                CommandLength = c.Length,
                TimesExecuted = c.Hash.Length == 0
                    ? 1
                    : _sources[(c.Kind, c.Path)].CommandCounts.GetValueOrDefault(c.Hash, 1),
                CommandText = c.Text
            })
            .ToList());

        result.Issues = NullIfEmpty(_issues.Values
            .OrderByDescending(i => i.Count)
            .ThenBy(i => i.FirstUtc)
            .Take(_top)
            .ToList());

        AddFindings(result, gaps);
        return result;
    }

    private void AddFindings(PowerQueryTraceSummaryResult result, List<PowerQueryTraceGap> gaps)
    {
        var findings = result.Findings;
        if (result.WindowSeconds > 0)
        {
            double share = result.PowerQueryBusySeconds / result.WindowSeconds * 100;
            findings.Add(
                $"Power Query was active for {Format(result.PowerQueryBusySeconds)} s of the {Format(result.WindowSeconds)} s window ({share.ToString("0", CultureInfo.InvariantCulture)}%). " +
                $"During the other {Format(result.OutsidePowerQuerySeconds)} s Power Query recorded nothing, so that time was spent elsewhere (for example Excel loading data, the Data Model, PivotTables, or a dialog waiting for an answer).");
        }

        var sources = _sources.Where(s => s.Value.Calls > 0).ToList();
        if (sources.Count == 0)
        {
            findings.Add("No data-source commands were recorded.");
        }
        else
        {
            double total = sources.Sum(s => s.Value.TotalSeconds);
            int calls = sources.Sum(s => s.Value.Calls);
            var slowest = sources.OrderByDescending(s => s.Value.TotalSeconds).First();
            findings.Add(
                $"{calls} data-source commands took {Format(Round(total))} s in total. " +
                $"Slowest source: {slowest.Key.Kind} with {slowest.Value.Calls} commands, {Format(Round(slowest.Value.TotalSeconds))} s in total, longest {Format(Round(slowest.Value.MaxSeconds))} s.");

            foreach (var source in sources.Where(s => s.Value.CommandCounts.Values.Any(c => c > 1)))
            {
                int repeated = source.Value.CommandCounts.Values.Sum(c => c - 1);
                findings.Add(
                    $"{repeated} of {source.Value.Calls} {source.Key.Kind} commands repeated a command that had already run (same text, see commandHash). " +
                    "Power Query often runs a native query twice: once to read the column layout and once for the data.");
            }

            foreach (var source in sources.Where(s => s.Value.Calls >= 5 && s.Value.ConnectionOpens >= s.Value.Calls))
            {
                findings.Add(
                    $"{source.Value.ConnectionOpens} {source.Key.Kind} connections were opened for {source.Value.Calls} commands, so every command opened its own connection.");
            }
        }

        var longestGap = gaps.OrderByDescending(g => g.Seconds).FirstOrDefault();
        if (longestGap != null && longestGap.Seconds >= 30)
        {
            findings.Add(
                $"Longest time without Power Query activity: {Format(longestGap.Seconds)} s from {longestGap.StartUtc.ToString("HH:mm:ss", CultureInfo.InvariantCulture)} UTC.");
        }

        var problems = _issues.Values.Where(i => i.Kind != "cancelled").ToList();
        if (problems.Count > 0)
        {
            string kinds = string.Join(", ", problems.Select(i => i.Kind).Distinct(StringComparer.Ordinal));
            findings.Add($"{problems.Sum(i => i.Count)} warnings or errors were recorded ({kinds}); see issues.");
        }

        int cancelled = _issues.Values.Where(i => i.Kind == "cancelled").Sum(i => i.Count);
        if (cancelled > 0)
        {
            findings.Add(
                $"{cancelled} operations were cancelled. Excel often cancels previews and background work, so cancellations alone do not mean the refresh failed.");
        }

        AddUnreadableFinding(result);
    }

    private void AddUnreadableFinding(PowerQueryTraceSummaryResult result)
    {
        if (_unreadable > 0)
        {
            result.Findings.Add($"{_unreadable} lines could not be read as trace entries and were skipped.");
        }
    }

    private List<(DateTime Start, DateTime End)> MergeIntervals(DateTime? windowStart, DateTime? windowEnd)
    {
        var merged = new List<(DateTime Start, DateTime End)>();
        foreach (var (rawStart, rawEnd) in _intervals.OrderBy(i => i.Start))
        {
            var start = windowStart is DateTime ws && rawStart < ws ? ws : rawStart;
            var end = windowEnd is DateTime we && rawEnd > we ? we : rawEnd;
            if (end <= start)
            {
                continue;
            }

            if (merged.Count > 0 && start <= merged[^1].End)
            {
                if (end > merged[^1].End)
                {
                    merged[^1] = (merged[^1].Start, end);
                }
            }
            else
            {
                merged.Add((start, end));
            }
        }

        return merged;
    }

    private static List<PowerQueryTraceGap> FindGaps(
        List<(DateTime Start, DateTime End)> merged,
        DateTime? windowStart,
        DateTime? windowEnd)
    {
        var gaps = new List<PowerQueryTraceGap>();
        if (windowStart is not DateTime cursor || windowEnd is not DateTime end)
        {
            return gaps;
        }

        foreach (var interval in merged)
        {
            AddGap(gaps, cursor, interval.Start);
            if (interval.End > cursor)
            {
                cursor = interval.End;
            }
        }

        AddGap(gaps, cursor, end);
        return gaps;
    }

    private static void AddGap(List<PowerQueryTraceGap> gaps, DateTime start, DateTime end)
    {
        double seconds = (end - start).TotalSeconds;
        if (seconds >= QuietGapSeconds)
        {
            gaps.Add(new PowerQueryTraceGap { StartUtc = start, EndUtc = end, Seconds = Round(seconds) });
        }
    }

    internal static string HashCommand(string commandText)
    {
        byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(commandText));
        return Convert.ToHexString(hash, 0, 6).ToLowerInvariant();
    }

    private static string? SampleMessage(string? exception)
    {
        if (string.IsNullOrWhiteSpace(exception))
        {
            return null;
        }

        var message = MessagePattern().Match(exception);
        string line = message.Success
            ? message.Groups["message"].Value
            : exception.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).FirstOrDefault() ?? string.Empty;
        return Shorten(line.Trim(), MaxSampleMessageLength);
    }

    private static string? Shorten(string? text, int maxLength) =>
        text is null || text.Length <= maxLength ? text : text[..maxLength] + "…";

    private static double Seconds(TimeSpan duration) => duration.TotalSeconds;

    private static double Round(double seconds) => Math.Round(seconds, 3, MidpointRounding.AwayFromZero);

    private static string Format(double seconds) => seconds.ToString("0.0", CultureInfo.InvariantCulture);

    private static List<T>? NullIfEmpty<T>(List<T> items) => items.Count == 0 ? null : items;

    [GeneratedRegex(@"ExceptionType:\s*(?<type>[^,\r\n]+)", RegexOptions.CultureInvariant)]
    private static partial Regex ExceptionTypePattern();

    [GeneratedRegex(@"^\s*Message:\s*(?<message>.+)$", RegexOptions.CultureInvariant | RegexOptions.Multiline)]
    private static partial Regex MessagePattern();

    private sealed class ActivityTotals(string id, DateTime start, DateTime end)
    {
        public string Id { get; } = id;
        public DateTime Start { get; private set; } = start;
        public DateTime End { get; private set; } = end;
        public int Evaluations { get; set; }
        public int DataSourceCalls { get; set; }
        public HashSet<int> Containers { get; } = [];

        public void Include(DateTime start, DateTime end)
        {
            if (start < Start)
            {
                Start = start;
            }

            if (end > End)
            {
                End = end;
            }
        }
    }

    private sealed class SourceTotals
    {
        public int Calls { get; set; }
        public double TotalSeconds { get; set; }
        public double MaxSeconds { get; set; }
        public int ConnectionOpens { get; set; }
        public Dictionary<string, int> CommandCounts { get; } = new(StringComparer.Ordinal);
    }

    private sealed record CommandRecord(
        string Kind,
        string Path,
        string ActivityId,
        DateTime Start,
        double Seconds,
        string Hash,
        int Length,
        string? Text);
}
