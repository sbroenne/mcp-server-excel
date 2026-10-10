using Sbroenne.ExcelMcp.Build;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

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
    public void Reports_RequireEverySelectedCaseToPass(int total, int passed, int executed, string summary, string outcome, bool succeeds)
    {
        var file = Path.Combine(Path.GetTempPath(), $"ExcelMcp.Report.{Guid.NewGuid():N}.trx");
        try
        {
            File.WriteAllText(file, $$"""
                <TestRun xmlns="http://microsoft.com/schemas/VisualStudio/TeamTest/2010">
                  <Results><UnitTestResult testName="case" outcome="{{outcome}}"/></Results>
                  <ResultSummary outcome="{{summary}}">
                    <Counters total="{{total}}" passed="{{passed}}" executed="{{executed}}"
                      failed="0" notExecuted="0" error="0" timeout="0" aborted="0"
                      inconclusive="0" passedButRunAborted="0" notRunnable="0"
                      disconnected="0" warning="0" inProgress="0" pending="0"/>
                  </ResultSummary>
                </TestRun>
                """);
            if (succeeds) { Assert.Equal(total, TestReport.Verify(file)); }
            else { Assert.Throws<InvalidOperationException>(() => TestReport.Verify(file)); }
        }
        finally { File.Delete(file); }
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
    public void Reports_RejectNonPassingCounters(string name, string value)
    {
        var counters = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["total"] = "1",
            ["passed"] = "1",
            ["executed"] = "1",
            ["failed"] = "0",
            ["notExecuted"] = "0",
            ["error"] = "0",
            ["timeout"] = "0",
            ["aborted"] = "0",
            ["inconclusive"] = "0",
            ["passedButRunAborted"] = "0",
            ["notRunnable"] = "0",
            ["disconnected"] = "0",
            ["warning"] = "0",
            ["inProgress"] = "0",
            ["pending"] = "0"
        };
        counters[name] = value;
        var file = WriteReport(counters, ["case"]);
        try { Assert.Throws<InvalidOperationException>(() => TestReport.Verify(file)); }
        finally { File.Delete(file); }
    }

    [Fact]
    public void Reports_RequireFailureCounters()
    {
        var file = WriteReport(new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["total"] = "1",
            ["passed"] = "1",
            ["executed"] = "1"
        }, ["case"]);
        try { Assert.Throws<InvalidOperationException>(() => TestReport.Verify(file)); }
        finally { File.Delete(file); }
    }

    [Fact]
    public void Reports_ReconcileCaseNamesAndMultiplicities()
    {
        var file = WriteReport(new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["total"] = "2",
            ["passed"] = "2",
            ["executed"] = "2",
            ["failed"] = "0",
            ["notExecuted"] = "0"
        }, ["case-one", "case-one"]);
        try
        {
            Assert.Throws<InvalidOperationException>(() =>
                TestReport.Verify(file, ["case-one", "case-two"]));
            Assert.Equal(2, TestReport.Verify(file, ["case-one", "case-one"]));
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void Discovery_ParsesTestNamesAndRejectsEmptyListing()
    {
        const string output = """
            Build output
            The following Tests are available:
                Suite.Case(1)
                Suite.Case(2)

            """;
        Assert.Equal(["Suite.Case(1)", "Suite.Case(2)"], TestReport.ParseDiscoveredCases(output));
        Assert.Throws<InvalidOperationException>(() => TestReport.ParseDiscoveredCases("No tests found."));
        Assert.Throws<InvalidOperationException>(() =>
            TestReport.ParseDiscoveredCases("The following Tests are available:\n"));
    }

    [Fact]
    public async Task ReconcileCases_DiscoversThenChecksExecutedTestIdentities()
    {
        var root = Path.Combine(Path.GetTempPath(), $"ExcelMcp.Reconcile.{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var project = Path.Combine(root, "fixture.proj");
        var results = Path.Combine(root, "results");
        await File.WriteAllTextAsync(project, "<Project/>");
        var runner = new ReconcileRunner();
        try
        {
            await new TestExecution(root, runner).RunStageAsync(new TestStageOptions
            {
                Project = project,
                Filter = "fixture",
                ResultsDirectory = results,
                Name = "acceptance",
                ReconcileCases = true
            });
            Assert.Equal(2, runner.Commands.Count);
            Assert.Contains("--list-tests", runner.Commands[0]);
            Assert.DoesNotContain("--list-tests", runner.Commands[1]);
        }
        finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task RunAsync_ReconciliationRejectsAnOmittedDiscoveredCase()
    {
        var root = Path.Combine(Path.GetTempPath(), $"ExcelMcp.Reconcile.{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var project = Path.Combine(root, "tests", "ExcelMcp.CLI.Tests", "ExcelMcp.CLI.Tests.csproj");
        var results = Path.Combine(root, "results");
        Directory.CreateDirectory(Path.GetDirectoryName(project)!);
        await File.WriteAllTextAsync(project, "<Project/>");
        var runner = new ReconcileRunner(includeOmittedCase: true);
        try
        {
            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                new TestExecution(root, runner).RunAsync(
                    "CLI", "fixture", results, excel: true, reconcileCases: true));
            Assert.Equal(2, runner.Commands.Count);
            Assert.Contains("--list-tests", runner.Commands[0]);
            Assert.DoesNotContain("--list-tests", runner.Commands[1]);
        }
        finally { Directory.Delete(root, recursive: true); }
    }

    private static string WriteReport(IReadOnlyDictionary<string, string> counters, IReadOnlyList<string> cases)
    {
        var file = Path.Combine(Path.GetTempPath(), $"ExcelMcp.Report.{Guid.NewGuid():N}.trx");
        var attributes = string.Join(' ', counters.Select(item => $"{item.Key}=\"{item.Value}\""));
        var results = string.Join("", cases.Select(name => $"<UnitTestResult testName=\"{name}\" outcome=\"Passed\"/>"));
        File.WriteAllText(file, $$"""
            <TestRun xmlns="http://microsoft.com/schemas/VisualStudio/TeamTest/2010">
              <Results>{{results}}</Results>
              <ResultSummary outcome="Completed"><Counters {{attributes}}/></ResultSummary>
            </TestRun>
            """);
        return file;
    }

    private sealed class ReconcileRunner(bool includeOmittedCase = false) : IProcessRunner
    {
        public List<string[]> Commands { get; } = [];

        public async Task<ProcessResult> RunAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
            IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false)
        {
            Assert.Equal("dotnet", executable);
            Assert.NotNull(environment);
            Assert.Equal("en", environment["DOTNET_CLI_UI_LANGUAGE"]);
            var command = arguments.ToArray();
            Commands.Add(command);
            if (command.Contains("--list-tests", StringComparer.Ordinal))
            {
                var secondCase = includeOmittedCase ? "\n    Suite.OmittedCase" : "";
                return new ProcessResult(0, $"The following Tests are available:\n    Suite.Case{secondCase}", "");
            }
            var resultsDirectory = command[Array.IndexOf(command, "--results-directory") + 1];
            var logger = command.Single(argument => argument.StartsWith("trx;LogFileName=", StringComparison.Ordinal));
            var report = logger["trx;LogFileName=".Length..];
            await File.WriteAllTextAsync(Path.Combine(resultsDirectory, report), """
                <TestRun xmlns="http://microsoft.com/schemas/VisualStudio/TeamTest/2010">
                  <Results><UnitTestResult testName="Suite.Case" outcome="Passed"/></Results>
                  <ResultSummary outcome="Completed">
                    <Counters total="1" passed="1" executed="1" failed="0" notExecuted="0"/>
                  </ResultSummary>
                </TestRun>
                """);
            return new ProcessResult(0, "passed", "");
        }

        public Task<ProcessResult> CheckedAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
            IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false) =>
            throw new InvalidOperationException("The test stage must use the bounded process runner.");
    }
}
