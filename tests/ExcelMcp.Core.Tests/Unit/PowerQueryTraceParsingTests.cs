using Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "false")]
public sealed class PowerQueryTraceParsingTests
{
    [Fact]
    public void TryParse_TraceLineWithStringHostProcessId_ReadsEntry()
    {
        const string line =
            "DataMashup.Trace Information: 24579 : {\"Action\":\"Engine/IO/AnalysisServices/Command/ExecuteReader\",\"ActivityId\":\"activity-1\",\"Start\":\"2026-10-15T07:27:19.661979Z\",\"Duration\":\"00:00:02.5000000\",\"Pid\":16772,\"HostProcessId\":\"1234\",\"Process\":\"Mashup\",\"ResourceKind\":\"AnalysisServices\",\"ResourcePath\":\"server/model\",\"CommandText\":\"SELECT 1\"}";

        Assert.True(PowerQueryTraceLineParser.TryParse(line, out var entry));
        Assert.NotNull(entry);
        Assert.Equal("Information", entry.Level);
        Assert.Equal(16772, entry.ProcessId);
        Assert.Equal(1234, entry.HostProcessId);
        Assert.Equal("SELECT 1", entry.CommandText);
        Assert.Equal(TimeSpan.FromSeconds(2.5), entry.Duration);
        Assert.Equal(DateTimeKind.Utc, entry.StartUtc.Kind);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("not a trace")]
    [InlineData("DataMashup.Trace Information: 24579 : {bad json}")]
    [InlineData("DataMashup.Trace Information: 24579 : {\"Start\":\"2026-10-15T07:27:19Z\",\"Pid\":1}")]
    [InlineData("DataMashup.Trace Information: 24579 : {\"Action\":\"run\",\"Start\":\"bad\",\"Pid\":1}")]
    [InlineData("DataMashup.Trace Information: 24579 : {\"Action\":\"run\",\"Start\":\"2026-10-15T07:27:19Z\",\"Pid\":1,\"Duration\":\"bad\"}")]
    public void TryParse_InvalidLine_ReturnsFalse(string? line)
    {
        Assert.False(PowerQueryTraceLineParser.TryParse(line, out var entry));
        Assert.Null(entry);
    }

    [Theory]
    [InlineData("EXCEL.16772.2026-10-15T07-27-19-661979.log", "excel", 16772)]
    [InlineData("Microsoft.Mashup.Container.Loader.24680.2026-10-15T07-27-19-661979.log", "container", 24680)]
    public void FileName_TwoTraceFileKinds_AreRecognized(string fileName, string expectedKind, int expectedProcessId)
    {
        Assert.True(PowerQueryTraceFileName.TryParse(fileName, out var kind, out var processId, out var startedUtc));
        Assert.Equal(expectedKind, kind);
        Assert.Equal(expectedProcessId, processId);
        Assert.Equal(new DateTime(2026, 10, 15, 7, 27, 19, DateTimeKind.Utc).AddTicks(6_619_790), startedUtc);
    }

    [Fact]
    public void FileName_UnrelatedFile_IsRejected()
    {
        Assert.False(PowerQueryTraceFileName.TryParse("notes.log", out _, out _, out _));
    }

    [Fact]
    public void Parse_AgeAndIsoDateTime_ReturnUtcTimes()
    {
        var now = new DateTime(2026, 10, 15, 8, 0, 0, DateTimeKind.Utc);

        Assert.Equal(now.AddMinutes(-15), PowerQueryTraceTime.Parse("15m", now, "since"));
        Assert.Equal(
            new DateTime(2026, 10, 15, 6, 30, 0, DateTimeKind.Utc),
            PowerQueryTraceTime.Parse("2026-10-15T08:30:00+02:00", now, "since"));
    }

    [Theory]
    [InlineData("15 months")]
    [InlineData("-2h")]
    [InlineData("999999999999999999999999d")]
    public void Parse_InvalidOrTooLargeTime_ThrowsArgumentException(string value)
    {
        var now = new DateTime(2026, 10, 15, 8, 0, 0, DateTimeKind.Utc);

        Assert.Throws<ArgumentException>(() => PowerQueryTraceTime.Parse(value, now, "since"));
    }
}
