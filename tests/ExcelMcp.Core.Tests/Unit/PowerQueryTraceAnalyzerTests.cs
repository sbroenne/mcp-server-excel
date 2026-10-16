using Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "false")]
public sealed class PowerQueryTraceAnalyzerTests
{
    private static readonly DateTime Start = new(2026, 10, 15, 7, 0, 0, DateTimeKind.Utc);

    [Fact]
    public void Build_CombinesIntervalsAndRepeatedCommands_HidesSensitiveTextByDefault()
    {
        var analyzer = new PowerQueryTraceAnalyzer(new PowerQueryTraceFilter(
            SinceUtc: Start,
            UntilUtc: Start.AddSeconds(60)));
        analyzer.AddFileRead();
        analyzer.Add(Entry(
            "SimpleDocumentEvaluator/GetResult/Evaluate",
            Start,
            TimeSpan.FromSeconds(10),
            activityId: "activity-1"));
        analyzer.Add(Entry(
            "Engine/IO/AnalysisServices/Command/ExecuteReader",
            Start.AddSeconds(5),
            TimeSpan.FromSeconds(20),
            activityId: "activity-1",
            hostProcessId: 42,
            resourceKind: "AnalysisServices",
            resourcePath: "server/model",
            commandText: "SELECT secret"));
        analyzer.Add(Entry(
            "Engine/IO/AnalysisServices/Command/ExecuteReader",
            Start.AddSeconds(40),
            TimeSpan.FromSeconds(5),
            activityId: "activity-1",
            hostProcessId: 42,
            resourceKind: "AnalysisServices",
            resourcePath: "server/model",
            commandText: "SELECT secret"));

        var result = analyzer.Build();

        Assert.True(result.Success);
        Assert.Equal(3, result.EntriesMatched);
        Assert.Equal(1, result.FilesRead);
        Assert.Equal(60, result.WindowSeconds);
        Assert.Equal(30, result.PowerQueryBusySeconds);
        Assert.Equal(30, result.OutsidePowerQuerySeconds);
        Assert.Equal(1, result.Evaluations);
        Assert.Equal(2, result.DataSourceCalls);
        Assert.Equal(Start.AddSeconds(45).ToString("o"), result.TraceCursor);
        var source = Assert.Single(result.DataSources!);
        Assert.Equal(1, source.DistinctCommands);
        Assert.Equal(1, source.RepeatedExecutions);
        Assert.Equal(2, result.SlowestCommands!.Count);
        Assert.Equal(2, result.SlowestCommands[0].TimesExecuted);
        Assert.All(result.SlowestCommands, command => Assert.Null(command.CommandText));
        Assert.Contains(result.QuietGaps!, gap => gap.Seconds == 15);
        Assert.Contains(result.Findings, finding => finding.Contains("active for 30.0 s", StringComparison.Ordinal));
    }

    [Fact]
    public void Build_IncludeQueryText_ReturnsCommandAndIssueSample()
    {
        var analyzer = new PowerQueryTraceAnalyzer(new PowerQueryTraceFilter(IncludeQueryText: true));
        analyzer.Add(Entry(
            "Engine/IO/AnalysisServices/Command/ExecuteReader",
            Start,
            TimeSpan.FromSeconds(2),
            resourceKind: "AnalysisServices",
            resourcePath: "server/model",
            commandText: "SELECT secret"));
        analyzer.Add(Entry(
            "Engine/IO/AnalysisServices/Connection/Open",
            Start.AddSeconds(3),
            TimeSpan.Zero,
            resourceKind: "AnalysisServices",
            exception: "Exception:\nExceptionType: System.UnauthorizedAccessException\nMessage: access denied"));

        var result = analyzer.Build();

        Assert.Equal("SELECT secret", Assert.Single(result.SlowestCommands!).CommandText);
        var issue = Assert.Single(result.Issues!);
        Assert.Equal("permission", issue.Kind);
        Assert.Equal(typeof(UnauthorizedAccessException).FullName, issue.ExceptionType);
        Assert.Contains("access denied", issue.SampleMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_IncludesWarningWithoutExceptionText()
    {
        var analyzer = new PowerQueryTraceAnalyzer(new PowerQueryTraceFilter());
        analyzer.Add(Entry("Engine/Warning", Start, TimeSpan.Zero) with { Level = "Warning" });

        var result = analyzer.Build();

        var issue = Assert.Single(result.Issues!);
        Assert.Equal("Engine/Warning", issue.Action);
        Assert.Equal(1, issue.Count);
        Assert.Null(issue.ExceptionType);
        Assert.Contains(result.Findings, finding => finding.Contains("warnings or errors", StringComparison.Ordinal));
    }

    [Fact]
    public void Build_ProcessFilterMatchesContainerHostAndAppliesTimeBounds()
    {
        var analyzer = new PowerQueryTraceAnalyzer(new PowerQueryTraceFilter(
            ProcessId: 42,
            SinceUtc: Start.AddSeconds(5),
            UntilUtc: Start.AddSeconds(15)));
        analyzer.Add(Entry(
            "Engine/IO/AnalysisServices/Command/ExecuteReader",
            Start,
            TimeSpan.FromSeconds(10),
            processId: 101,
            hostProcessId: 42,
            resourceKind: "AnalysisServices",
            commandText: "SELECT 1"));
        analyzer.Add(Entry(
            "Engine/IO/AnalysisServices/Command/ExecuteReader",
            Start.AddSeconds(20),
            TimeSpan.FromSeconds(10),
            processId: 101,
            hostProcessId: 42,
            resourceKind: "AnalysisServices",
            commandText: "SELECT 2"));
        analyzer.Add(Entry(
            "Engine/IO/AnalysisServices/Command/ExecuteReader",
            Start,
            TimeSpan.FromSeconds(10),
            processId: 101,
            hostProcessId: 99,
            resourceKind: "AnalysisServices",
            commandText: "SELECT 3"));

        var result = analyzer.Build();

        Assert.Equal(1, result.EntriesMatched);
        Assert.Equal(1, result.DataSourceCalls);
        Assert.Equal("SELECT 1".Length, Assert.Single(result.SlowestCommands!).CommandLength);
    }

    [Fact]
    public void Build_CursorDoesNotSkipCommandsAfterRequestedUpperBound()
    {
        var entries = new[]
        {
            Entry("SimpleDocumentEvaluator/GetResult/Evaluate", Start, TimeSpan.FromSeconds(100)),
            Entry("Engine/IO/AnalysisServices/Command/ExecuteReader",
                Start.AddSeconds(20), TimeSpan.FromSeconds(2),
                resourceKind: "AnalysisServices", commandText: "SELECT 1")
        };
        var first = new PowerQueryTraceAnalyzer(new PowerQueryTraceFilter(
            SinceUtc: Start, UntilUtc: Start.AddSeconds(10)));
        foreach (var entry in entries)
        {
            first.Add(entry);
        }

        var result = first.Build();

        Assert.Equal(1, result.EntriesMatched);
        Assert.Equal(Start.AddSeconds(10).ToString("o"), result.TraceCursor);
        var next = new PowerQueryTraceAnalyzer(new PowerQueryTraceFilter(
            SinceUtc: PowerQueryTraceTime.Parse(result.TraceCursor, Start, "since"),
            UntilUtc: Start.AddSeconds(30)));
        foreach (var entry in entries)
        {
            next.Add(entry);
        }

        Assert.Equal(1, next.Build().DataSourceCalls);
    }

    private static PowerQueryTraceEntry Entry(
        string action,
        DateTime start,
        TimeSpan duration,
        string activityId = "activity-1",
        int processId = 101,
        int? hostProcessId = null,
        string? resourceKind = null,
        string? resourcePath = null,
        string? commandText = null,
        string? exception = null)
    {
        return new PowerQueryTraceEntry(
            "Information",
            action,
            activityId,
            start,
            duration,
            processId,
            hostProcessId,
            "Mashup",
            resourceKind,
            resourcePath,
            commandText,
            exception,
            Cancelled: false);
    }
}
