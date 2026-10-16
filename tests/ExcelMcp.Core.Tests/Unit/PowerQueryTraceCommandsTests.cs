using System.Globalization;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;
using Sbroenne.ExcelMcp.Core.Commands.Settings;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "false")]
public sealed class PowerQueryTraceCommandsTests
{
    [Fact]
    public void SettingsCommands_GetAndSet_ReportsCurrentAndPreviousValues()
    {
        using var folder = new TemporaryFolder();
        File.WriteAllText(Path.Combine(folder.Path, "EXCEL.123.2026-10-15T07-27-19-661979.log"), "trace");
        File.WriteAllText(Path.Combine(folder.Path, "notes.log"), "ignore");
        var settings = new FakeTraceSettings(folder.Path, new PowerQueryTracingSnapshot(0, 15));
        var commands = new SettingsCommands(settings);

        var before = commands.GetPqTracing();
        var after = commands.SetPqTracing(enabled: true);

        Assert.True(before.Success);
        Assert.False(before.Enabled);
        Assert.Equal(15, before.Level);
        Assert.Equal(1, before.TraceFileCount);
        Assert.False(after.PreviousEnabled);
        Assert.Equal(15, after.PreviousLevel);
        Assert.True(after.Enabled);
        Assert.Equal(15, after.Level);
        Assert.Equal("1", settings.Read().EnabledValue?.ToString(CultureInfo.InvariantCulture));
        Assert.Contains("Restart Excel", after.Note, StringComparison.Ordinal);
    }

    [Fact]
    public void SettingsCommands_WhenLevelIsNotSet_ReportsExcelDefault()
    {
        using var folder = new TemporaryFolder();
        var settings = new FakeTraceSettings(folder.Path, new PowerQueryTracingSnapshot(null, null));
        var commands = new SettingsCommands(settings);

        var result = commands.SetPqTracing(enabled: true);

        Assert.Equal(15, result.PreviousLevel);
        Assert.Equal(15, result.Level);
    }

    [Fact]
    public void TraceCommands_ListAndSummarize_FilterByExcelProcessAndContainerHost()
    {
        using var folder = new TemporaryFolder();
        string timestamp = "2026-10-15T07-27-19-661979";
        File.WriteAllText(
            Path.Combine(folder.Path, $"EXCEL.42.{timestamp}.log"),
            TraceLine("run/excel", 42, null, "00:00:02", null));
        File.WriteAllText(
            Path.Combine(folder.Path, $"Microsoft.Mashup.Container.Loader.100.{timestamp}.log"),
            TraceLine("Engine/IO/AnalysisServices/Command/ExecuteReader", 100, "42", "00:00:03", "SELECT 1"));
        File.WriteAllText(
            Path.Combine(folder.Path, $"Microsoft.Mashup.Container.Loader.101.{timestamp}.log"),
            TraceLine("Engine/IO/AnalysisServices/Command/ExecuteReader", 101, "99", "00:00:03", "SELECT 2"));
        var commands = new PowerQueryTraceCommands(folder.Path, _ => null, () => new DateTime(2026, 10, 15, 8, 0, 0, DateTimeKind.Utc));

        var listed = commands.List(processId: 42);
        var summary = commands.Summarize(processId: 42, includeQueryText: true);

        Assert.Equal(2, listed.Files.Count);
        Assert.Contains(listed.Files, file => file.Kind == "excel" && file.ProcessId == 42);
        Assert.Contains(listed.Files, file => file.Kind == "container" && file.ProcessId == 100 && file.HostProcessId == 42);
        Assert.Equal(2, summary.FilesRead);
        Assert.Equal(2, summary.EntriesMatched);
        Assert.Equal(1, summary.DataSourceCalls);
        Assert.Equal("SELECT 1", Assert.Single(summary.SlowestCommands!).CommandText);
    }

    [Fact]
    public void TraceCommands_SessionFilterExcludesFilesFromPreviousProcessWithReusedPid()
    {
        using var folder = new TemporaryFolder();
        var processStarted = new DateTime(2026, 10, 15, 7, 27, 20, DateTimeKind.Utc);
        File.WriteAllText(
            Path.Combine(folder.Path, "EXCEL.42.2026-10-15T07-27-19-661979.log"),
            TraceLine("run/old-excel", 42, null, "00:00:02", null));
        File.WriteAllText(
            Path.Combine(folder.Path, "Microsoft.Mashup.Container.Loader.100.2026-10-15T07-27-19-661979.log"),
            TraceLine("Engine/IO/AnalysisServices/Command/ExecuteReader", 100, "42", "00:00:03", "SELECT previous workbook"));
        File.WriteAllText(
            Path.Combine(folder.Path, "EXCEL.42.2026-10-15T07-27-21-000000.log"),
            TraceLine("run/current-excel", 42, null, "00:00:02", null));
        File.WriteAllText(
            Path.Combine(folder.Path, "Microsoft.Mashup.Container.Loader.101.2026-10-15T07-27-21-000000.log"),
            TraceLine("Engine/IO/AnalysisServices/Command/ExecuteReader", 101, "42", "00:00:03", "SELECT current workbook"));
        var commands = new PowerQueryTraceCommands(
            folder.Path,
            _ => (42, processStarted),
            () => new DateTime(2026, 10, 15, 8, 0, 0, DateTimeKind.Utc));

        var listed = commands.List(workbookSessionId: "current-session");
        var summary = commands.Summarize(workbookSessionId: "current-session", includeQueryText: true);

        Assert.Equal(2, listed.Files.Count);
        Assert.All(listed.Files, file => Assert.True(file.StartedUtc >= processStarted));
        Assert.Equal(2, summary.FilesRead);
        Assert.Equal(1, summary.DataSourceCalls);
        Assert.Equal("SELECT current workbook", Assert.Single(summary.SlowestCommands!).CommandText);
        Assert.DoesNotContain("previous workbook",
            summary.SlowestCommands!.Single().CommandText,
            StringComparison.Ordinal);
    }

    [Fact]
    public void TraceReader_ListSince_IncludesActiveContainerWhenHostExcelIsRunning()
    {
        using var folder = new TemporaryFolder();
        const string timestamp = "2026-10-15T07-27-19-661979";
        string tracePath = Path.Combine(folder.Path, $"Microsoft.Mashup.Container.Loader.100.{timestamp}.log");
        File.WriteAllText(
            tracePath,
            "DataMashup.Trace Information: 24579 : {\"Action\":\"run/container\",\"Start\":\"2026-10-15T07:27:19.661979Z\",\"Duration\":\"00:00:02\",\"Pid\":100,\"HostProcessId\":\"42\"}");
        var now = new DateTime(2026, 10, 15, 8, 0, 0, DateTimeKind.Utc);
        File.SetLastWriteTimeUtc(tracePath, now.AddHours(-1));
        var reader = new PowerQueryTraceReader(folder.Path, processId => processId == 42);

        var result = reader.List(processId: null, sinceUtc: now.AddMinutes(-15));

        var file = Assert.Single(result.Files);
        Assert.Equal(100, file.ProcessId);
        Assert.Equal(42, file.HostProcessId);
    }

    [Fact]
    public void TraceReader_SummarizeSince_SkipsOldInactiveFilesButKeepsActiveFiles()
    {
        using var folder = new TemporaryFolder();
        const string timestamp = "2026-10-15T07-27-19-661979";
        foreach (var (name, line) in new[]
        {
            ($"EXCEL.41.{timestamp}.log", TraceLine("run/excel", 41, null, "00:00:02", null)),
            ($"EXCEL.42.{timestamp}.log", TraceLine("run/excel", 42, null, "00:00:02", null)),
            ($"Microsoft.Mashup.Container.Loader.100.{timestamp}.log", TraceLine("run/container", 100, "43", "00:00:02", null))
        })
        {
            string path = Path.Combine(folder.Path, name);
            File.WriteAllText(path, line);
            File.SetLastWriteTimeUtc(path, new DateTime(2026, 10, 15, 7, 30, 0, DateTimeKind.Utc));
        }

        var now = new DateTime(2026, 10, 15, 8, 0, 0, DateTimeKind.Utc);
        var reader = new PowerQueryTraceReader(folder.Path, processId => processId is 42 or 43);

        var result = reader.Summarize(new PowerQueryTraceFilter(SinceUtc: now.AddMinutes(-15)));

        Assert.Equal(2, result.FilesRead);
        Assert.Equal(0, result.EntriesMatched);
    }

    [Fact]
    public void RefreshTrace_ReadableTraceFile_AttachesMatchingSummaryToBothRefreshResults()
    {
        using var folder = new TemporaryFolder();
        File.WriteAllText(
            Path.Combine(folder.Path, "EXCEL.42.2026-10-15T07-27-19-661979.log"),
            TraceLine("run/excel", 42, null, "00:00:02", null));
        var settings = new FakeTraceSettings(folder.Path, new PowerQueryTracingSnapshot(1, 15));
        var commands = new PowerQueryCommands(new DataModelCommands(), settings);
        var single = new PowerQueryRefreshResult { Success = true };
        var all = new PowerQueryRefreshAllResult { Success = true };
        var start = new DateTime(2026, 10, 15, 7, 27, 19, DateTimeKind.Utc);

        commands.AddRefreshTrace(single, 42, start, start.AddSeconds(5), failed: false);
        commands.AddRefreshTrace(all, 42, start, start.AddSeconds(5), failed: false);

        foreach (var summary in new[] { single.TraceSummary, all.TraceSummary })
        {
            Assert.NotNull(summary);
            Assert.True(summary.Success);
            Assert.Null(summary.ErrorMessage);
            Assert.Equal(42, summary.ProcessId);
            Assert.Equal(1, summary.FilesRead);
            Assert.Equal(1, summary.EntriesMatched);
            Assert.Equal(2, summary.PowerQueryBusySeconds);
        }

        Assert.Null(single.TraceNote);
        Assert.Null(all.TraceNote);
    }

    [Fact]
    public void RefreshTrace_LockedTraceFile_PreservesSingleQueryRefreshOutcome()
    {
        using var folder = new TemporaryFolder();
        var path = Path.Combine(folder.Path, "EXCEL.42.2026-10-15T07-27-19-661979.log");
        using var lockedFile = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        var settings = new FakeTraceSettings(folder.Path, new PowerQueryTracingSnapshot(1, 15));
        var commands = new PowerQueryCommands(new DataModelCommands(), settings);
        var result = new PowerQueryRefreshResult { Success = true, QueryName = "Example" };
        var start = new DateTime(2026, 10, 15, 7, 0, 0, DateTimeKind.Utc);

        commands.AddRefreshTrace(result, 42, start, start.AddHours(1), failed: false);

        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.Equal("Example", result.QueryName);
        Assert.NotNull(result.TraceSummary);
        Assert.False(result.TraceSummary.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.TraceSummary.ErrorMessage));
        Assert.Contains("could not be read", result.TraceNote, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(true, null)]
    [InlineData(false, "Query refresh failed")]
    public void RefreshTrace_LockedTraceFile_PreservesRefreshAllOutcome(bool success, string? error)
    {
        using var folder = new TemporaryFolder();
        var path = Path.Combine(folder.Path, "EXCEL.42.2026-10-15T07-27-19-661979.log");
        using var lockedFile = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        var settings = new FakeTraceSettings(folder.Path, new PowerQueryTracingSnapshot(1, 15));
        var commands = new PowerQueryCommands(new DataModelCommands(), settings);
        var result = new PowerQueryRefreshAllResult { Success = success, ErrorMessage = error };
        var start = new DateTime(2026, 10, 15, 7, 0, 0, DateTimeKind.Utc);

        commands.AddRefreshTrace(result, 42, start, start.AddHours(1), failed: !success);

        Assert.Equal(success, result.Success);
        Assert.Equal(error, result.ErrorMessage);
        Assert.NotNull(result.TraceSummary);
        Assert.False(result.TraceSummary.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.TraceSummary.ErrorMessage));
        Assert.Contains("could not be read", result.TraceNote, StringComparison.Ordinal);
    }

    private static string TraceLine(string action, int processId, string? hostProcessId, string duration, string? commandText)
    {
        string host = hostProcessId == null ? string.Empty : $",\"HostProcessId\":\"{hostProcessId}\"";
        string command = commandText == null ? string.Empty : $",\"ResourceKind\":\"AnalysisServices\",\"ResourcePath\":\"server/model\",\"CommandText\":\"{commandText}\"";
        return $"DataMashup.Trace Information: 24579 : {{\"Action\":\"{action}\",\"ActivityId\":\"activity-1\",\"Start\":\"2026-10-15T07:27:19.661979Z\",\"Duration\":\"{duration}\",\"Pid\":{processId}{host}{command}}}";
    }

    private sealed class FakeTraceSettings(string traceFolder, PowerQueryTracingSnapshot initial) : IPowerQueryTraceSettings
    {
        private int? _enabled = initial.EnabledValue;
        private int? _level = initial.LevelValue;

        public string TraceFolder { get; } = traceFolder;

        public PowerQueryTracingSnapshot Read() => new(_enabled, _level);

        public void Write(bool enabled, int? level)
        {
            _enabled = enabled ? 1 : 0;
            _level = level ?? _level;
        }

    }

    private sealed class TemporaryFolder : IDisposable
    {
        public TemporaryFolder()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"pqtrace-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            Directory.Delete(Path, recursive: true);
        }
    }
}
