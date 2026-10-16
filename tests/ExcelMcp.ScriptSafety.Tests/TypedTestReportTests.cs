using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TypedTestReportTests
{
    [Theory]
    [InlineData(1, 1, 1, "Completed", "Passed", true)]
    [InlineData(0, 0, 0, "Completed", "Passed", false)]
    [InlineData(1, 0, 1, "Completed", "Failed", false)]
    [InlineData(1, 0, 0, "Completed", "NotExecuted", false)]
    [InlineData(1, 1, 1, "Failed", "Passed", false)]
    [InlineData(2, 2, 2, "Completed", "Passed", false)]
    public async Task Reports_RequireEverySelectedCaseToPass(int total, int passed, int executed, string summary, string outcome, bool succeeds)
    {
        var run = await CheckAsync($$"""
            <TestRun><Results><UnitTestResult testName="case" outcome="{{outcome}}"/></Results>
            <ResultSummary outcome="{{summary}}"><Counters total="{{total}}" passed="{{passed}}" executed="{{executed}}" failed="0" notExecuted="0"/></ResultSummary></TestRun>
            """);
        Assert.Equal(succeeds, run.ExitCode == 0);
    }

    [Theory]
    [InlineData("failed", "1")]
    [InlineData("notExecuted", "1")]
    [InlineData("error", "1")]
    [InlineData("timeout", "1")]
    [InlineData("aborted", "1")]
    [InlineData("inconclusive", "1")]
    [InlineData("passedButRunAborted", "1")]
    [InlineData("notRunnable", "1")]
    [InlineData("disconnected", "1")]
    [InlineData("warning", "1")]
    [InlineData("inProgress", "1")]
    [InlineData("pending", "1")]
    [InlineData("failed", "invalid")]
    public async Task Reports_RejectNonPassingCounters(string name, string value)
    {
        var counters = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["total"] = "1",
            ["passed"] = "1",
            ["executed"] = "1",
            ["failed"] = "0",
            ["notExecuted"] = "0"
        };
        counters[name] = value;
        var attributes = string.Join(' ', counters.Select(item => $"{item.Key}=\"{item.Value}\""));
        var run = await CheckAsync($"""
            <TestRun><Results><UnitTestResult testName="case" outcome="Passed"/></Results><ResultSummary outcome="Completed"><Counters {attributes}/></ResultSummary></TestRun>
            """);
        Assert.NotEqual(0, run.ExitCode);
    }

    [Fact]
    public async Task Reports_RequireFailureCountersAndRejectCleanupErrors()
    {
        foreach (var counters in new[] { "total=\"1\" passed=\"1\" executed=\"1\"", "total=\"1\" passed=\"1\" executed=\"1\" failed=\"0\" notExecuted=\"0\"" })
        {
            var run = await CheckAsync($"""
                <TestRun><Results><UnitTestResult testName="case" outcome="Passed"/></Results><ResultSummary outcome="Completed"><Counters {counters}/><RunInfos><RunInfo outcome="Error"/></RunInfos></ResultSummary></TestRun>
                """);
            Assert.NotEqual(0, run.ExitCode);
        }
    }

    [Theory]
    [InlineData("'case-one','case-one'", true)]
    [InlineData("'case-one','case-two'", false)]
    [InlineData("'case-one'", false)]
    [InlineData("'case-one','case-one','omitted-case'", false)]
    public async Task Reports_ReconcileCaseNamesAndMultiplicities(string expected, bool succeeds)
    {
        var run = await CheckAsync("""
            <TestRun><Results><UnitTestResult testName="case-one" outcome="Passed"/><UnitTestResult testName="case-one" outcome="Passed"/></Results><ResultSummary outcome="Completed"><Counters total="2" passed="2" executed="2" failed="0" notExecuted="0"/></ResultSummary></TestRun>
            """, $"-ExpectedCases @({expected})");
        Assert.Equal(succeeds, run.ExitCode == 0);
        if (!succeeds) { Assert.Contains("Test case reconciliation failed", run.Output, StringComparison.Ordinal); }
    }

    [Fact]
    public async Task Discovery_RejectsMissingOrEmptyListings()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            . .\scripts\Invoke-TestStage.ps1
            $cases = @(Get-DiscoveredTestCases "Build output`nThe following Tests are available:`n    Suite.Case(1)`n    Suite.Case(2)`n")
            if (($cases -join ',') -cne 'Suite.Case(1),Suite.Case(2)') { throw 'Test names changed.' }
            foreach ($output in @('No tests found.',"The following Tests are available:`n")) {
                try { Get-DiscoveredTestCases $output; throw 'Invalid listing accepted.' }
                catch { if ($_ -notmatch 'no test-list header|no cases') { throw } }
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    private static Task<(int ExitCode, string Output)> CheckAsync(string report, string arguments = "") =>
        ValidationSelectionTests.RunAsync($$"""
            . .\scripts\Invoke-TestStage.ps1
            $file = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.Report.$([Guid]::NewGuid().ToString('N')).trx"
            try {
                '{{report}}' | Set-Content -LiteralPath $file
                Assert-TestReport -Path $file {{arguments}}
            } finally { Remove-Item -LiteralPath $file }
            """);
}
