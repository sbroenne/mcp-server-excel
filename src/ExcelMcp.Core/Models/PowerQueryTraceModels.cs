using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Current or changed Power Query tracing setting.
/// </summary>
public sealed class PowerQueryTracingStatusResult : ResultBase
{
    /// <summary>Whether Power Query tracing is switched on for Excel processes started from now on.</summary>
    public bool Enabled { get; set; }

    /// <summary>Effective tracing level, using Excel's default of 15 when no value is stored.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Level { get; set; }

    /// <summary>Tracing state before set-tracing changed it.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? PreviousEnabled { get; set; }

    /// <summary>Tracing level before set-tracing changed it.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? PreviousLevel { get; set; }

    /// <summary>Folder where Power Query writes trace files.</summary>
    public string TraceFolder { get; set; } = string.Empty;

    /// <summary>Number of trace files currently in the trace folder.</summary>
    public int TraceFileCount { get; set; }

    /// <summary>Explains when the setting takes effect.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; set; }
}

/// <summary>
/// One Power Query trace file.
/// </summary>
public sealed class PowerQueryTraceFileInfo
{
    /// <summary>File name inside the trace folder.</summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>"excel" for an Excel process log or "container" for a Power Query evaluation container log.</summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>Process ID that wrote the file.</summary>
    public int ProcessId { get; set; }

    /// <summary>Excel process that owns a container file, when it could be read from the file.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? HostProcessId { get; set; }

    /// <summary>File start time in UTC, taken from the file name.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTime? StartedUtc { get; set; }

    /// <summary>Last time the file was written, in UTC.</summary>
    public DateTime LastWriteUtc { get; set; }

    /// <summary>File size in bytes.</summary>
    public long SizeBytes { get; set; }
}

/// <summary>
/// Trace files found in the Power Query trace folder.
/// </summary>
public sealed class PowerQueryTraceListResult : ResultBase
{
    /// <summary>Folder where Power Query writes trace files.</summary>
    public string TraceFolder { get; set; } = string.Empty;

    /// <summary>Excel process ID used to filter files, when a filter was given.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ProcessId { get; set; }

    /// <summary>Matching trace files, newest first.</summary>
    public List<PowerQueryTraceFileInfo> Files { get; set; } = [];

    /// <summary>Total size of the matching files in bytes.</summary>
    public long TotalSizeBytes { get; set; }
}

/// <summary>
/// A time span where no Power Query trace activity was recorded.
/// </summary>
public sealed class PowerQueryTraceGap
{
    /// <summary>Start of the quiet period in UTC.</summary>
    public DateTime StartUtc { get; set; }

    /// <summary>End of the quiet period in UTC.</summary>
    public DateTime EndUtc { get; set; }

    /// <summary>Length of the quiet period in seconds.</summary>
    public double Seconds { get; set; }
}

/// <summary>
/// One Power Query activity (a refresh or evaluation request) identified by ActivityId.
/// </summary>
public sealed class PowerQueryTraceActivity
{
    /// <summary>Activity identifier from the trace; pass it to summarize to focus on this activity.</summary>
    public string ActivityId { get; set; } = string.Empty;

    /// <summary>First recorded start in UTC.</summary>
    public DateTime StartUtc { get; set; }

    /// <summary>Last recorded end in UTC.</summary>
    public DateTime EndUtc { get; set; }

    /// <summary>Elapsed seconds from first start to last end.</summary>
    public double Seconds { get; set; }

    /// <summary>Number of query evaluations in this activity.</summary>
    public int Evaluations { get; set; }

    /// <summary>Number of native data-source commands in this activity.</summary>
    public int DataSourceCalls { get; set; }

    /// <summary>Number of distinct evaluation container processes used.</summary>
    public int Containers { get; set; }
}

/// <summary>
/// One query evaluation.
/// </summary>
public sealed class PowerQueryTraceEvaluation
{
    /// <summary>Activity identifier the evaluation belongs to.</summary>
    public string ActivityId { get; set; } = string.Empty;

    /// <summary>Start time in UTC.</summary>
    public DateTime StartUtc { get; set; }

    /// <summary>Duration in seconds.</summary>
    public double Seconds { get; set; }

    /// <summary>Process ID that ran the evaluation.</summary>
    public int ProcessId { get; set; }
}

/// <summary>
/// Calls made to one data source.
/// </summary>
public sealed class PowerQueryTraceDataSource
{
    /// <summary>Data source kind, for example AnalysisServices or SQL.</summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>Data source path as reported by Power Query (for example server and database).</summary>
    public string Path { get; set; } = string.Empty;

    /// <summary>Number of native commands executed.</summary>
    public int Calls { get; set; }

    /// <summary>Total command time in seconds.</summary>
    public double TotalSeconds { get; set; }

    /// <summary>Average command time in seconds.</summary>
    public double AverageSeconds { get; set; }

    /// <summary>Longest command time in seconds.</summary>
    public double MaxSeconds { get; set; }

    /// <summary>Number of connections opened.</summary>
    public int ConnectionOpens { get; set; }

    /// <summary>Number of distinct command texts.</summary>
    public int DistinctCommands { get; set; }

    /// <summary>Number of executions that repeated a command text already executed in the same window.</summary>
    public int RepeatedExecutions { get; set; }
}

/// <summary>
/// One native data-source command execution. Command text is included only on request.
/// </summary>
public sealed class PowerQueryTraceCommand
{
    /// <summary>Data source kind.</summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>Data source path.</summary>
    public string Path { get; set; } = string.Empty;

    /// <summary>Activity identifier the command belongs to.</summary>
    public string ActivityId { get; set; } = string.Empty;

    /// <summary>Start time in UTC.</summary>
    public DateTime StartUtc { get; set; }

    /// <summary>Duration in seconds.</summary>
    public double Seconds { get; set; }

    /// <summary>Short SHA-256 fingerprint of the command text; equal values mean the same command.</summary>
    public string CommandHash { get; set; } = string.Empty;

    /// <summary>Length of the command text in characters.</summary>
    public int CommandLength { get; set; }

    /// <summary>How many times this exact command ran in the window.</summary>
    public int TimesExecuted { get; set; }

    /// <summary>Command text, only when includeQueryText was requested. Long text is shortened.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CommandText { get; set; }
}

/// <summary>
/// A group of warnings or errors recorded in the trace.
/// </summary>
public sealed class PowerQueryTraceIssue
{
    /// <summary>Issue kind: cancelled, firewall, credentials, permission, or error.</summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>Trace action that reported the issue.</summary>
    public string Action { get; set; } = string.Empty;

    /// <summary>Exception type name, when the trace records one.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ExceptionType { get; set; }

    /// <summary>Number of occurrences.</summary>
    public int Count { get; set; }

    /// <summary>First occurrence in UTC.</summary>
    public DateTime FirstUtc { get; set; }

    /// <summary>Last occurrence in UTC.</summary>
    public DateTime LastUtc { get; set; }

    /// <summary>First line of one exception message, only when includeQueryText was requested.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SampleMessage { get; set; }
}

/// <summary>
/// Summary of Power Query trace activity.
/// </summary>
public sealed class PowerQueryTraceSummaryResult : ResultBase
{
    /// <summary>Excel process ID used to filter entries, when a filter was given.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ProcessId { get; set; }

    /// <summary>Activity identifier used to filter entries, when a filter was given.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ActivityId { get; set; }

    /// <summary>Start of the summarized time window in UTC.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTime? WindowStartUtc { get; set; }

    /// <summary>End of the summarized time window in UTC.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTime? WindowEndUtc { get; set; }

    /// <summary>Number of trace files read.</summary>
    public int FilesRead { get; set; }

    /// <summary>Number of trace entries that matched the filters.</summary>
    public int EntriesMatched { get; set; }

    /// <summary>Number of lines that could not be read as trace entries.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int UnreadableLines { get; set; }

    /// <summary>Length of the window in seconds.</summary>
    public double WindowSeconds { get; set; }

    /// <summary>Seconds in the window where Power Query recorded activity.</summary>
    public double PowerQueryBusySeconds { get; set; }

    /// <summary>Seconds in the window without Power Query activity (time spent elsewhere, for example in Excel or the Data Model).</summary>
    public double OutsidePowerQuerySeconds { get; set; }

    /// <summary>Number of query evaluations.</summary>
    public int Evaluations { get; set; }

    /// <summary>Number of native data-source commands.</summary>
    public int DataSourceCalls { get; set; }

    /// <summary>Plain-language findings, most important first.</summary>
    public List<string> Findings { get; set; } = [];

    /// <summary>Longest quiet periods of at least 10 seconds.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<PowerQueryTraceGap>? QuietGaps { get; set; }

    /// <summary>Longest activities.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<PowerQueryTraceActivity>? Activities { get; set; }

    /// <summary>Slowest evaluations.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<PowerQueryTraceEvaluation>? SlowestEvaluations { get; set; }

    /// <summary>Calls per data source, slowest total first.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<PowerQueryTraceDataSource>? DataSources { get; set; }

    /// <summary>Slowest native data-source commands.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<PowerQueryTraceCommand>? SlowestCommands { get; set; }

    /// <summary>Warnings and errors, most frequent first.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<PowerQueryTraceIssue>? Issues { get; set; }

    /// <summary>
    /// End time of the newest matched entry in UTC (ISO 8601). Pass it as since on the next
    /// summarize call to read only newer activity.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? TraceCursor { get; set; }
}
