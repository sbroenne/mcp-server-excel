using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

public sealed partial class AutomationSafetyTests
{
    private static string LoadBehaviorRunnerFunctions => $$"""
        $ast = [Management.Automation.Language.Parser]::ParseFile('{{Quote(Path.Combine(RepoRoot, "scripts", "Test-ExcelBehavior.ps1"))}}', [ref]$null, [ref]$null)
        foreach ($function in $ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
            . ([scriptblock]::Create($function.Extent.Text))
        }
        """;

    [Theory]
    [InlineData("passed", true)]
    [InlineData("failed", false)]
    [InlineData("not-run", false)]
    [InlineData("running", false)]
    public async Task ExcelBehavior_FinalVerdictRejectsAnyFailedOrUnfinishedStage(string status, bool valid)
    {
        var root = NewSandbox();
        try
        {
            var result = await RunAsync(root, LoadBehaviorRunnerFunctions + $$"""

                Assert-ExcelBehaviorStageResults -Stages @(
                    @{ name = 'first'; status = '{{status}}' },
                    @{ name = 'second'; status = 'passed' },
                    @{ name = 'optional'; status = 'not-applicable-no-discovered-cases' })
                """);
            if (valid) { Assert.True(result.ExitCode == 0, result.Output); }
            else { Assert.NotEqual(0, result.ExitCode); }
        }
        finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("@()", "@()")]
    [InlineData("@('a', 'b')", "@('a')")]
    [InlineData("@('a', 'b')", "@('a', 'a')")]
    [InlineData("@('a')", "@('a', 'extra')")]
    [InlineData("@('A')", "@('a')")]
    public async Task ExcelBehavior_ReconciliationRejectsEmptyOmittedDuplicateOrWrongCases(string expected, string actual)
    {
        var root = NewSandbox();
        try
        {
            var result = await RunAsync(root, LoadBehaviorRunnerFunctions +
                $"\nAssert-ExcelBehaviorCases -Expected {expected} -Actual {actual}");
            Assert.NotEqual(0, result.ExitCode);
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task ExcelBehavior_ReconciliationPreservesRepeatedTheoryNames()
    {
        var root = NewSandbox();
        try
        {
            var result = await RunAsync(root, LoadBehaviorRunnerFunctions + """

                Assert-ExcelBehaviorCases -Expected @('theory(row: 1)', 'theory(row: 1)', 'other') -Actual @('other', 'theory(row: 1)', 'theory(row: 1)')
                """);
            Assert.True(result.ExitCode == 0, result.Output);
        }
        finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("Passed", 1, 1, "Completed", true)]
    [InlineData("Failed", 1, 0, "Failed", false)]
    [InlineData("NotExecuted", 1, 0, "Completed", false)]
    [InlineData("Failed", 1, 1, "Completed", false)]
    [InlineData("Passed", 2, 2, "Completed", false)]
    [InlineData("Passed", 1, 1, "Failed", false)]
    public async Task ExcelBehavior_TrxChecksActualResultsAndCounters(
        string outcome, int total, int passed, string summaryOutcome, bool valid)
    {
        var root = NewSandbox();
        try
        {
            File.WriteAllText(Path.Combine(root, "result.trx"), $$"""
                <TestRun xmlns="http://microsoft.com/schemas/VisualStudio/TeamTest/2010">
                  <Results><UnitTestResult testName="case" outcome="{{outcome}}" /></Results>
                  <ResultSummary outcome="{{summaryOutcome}}">
                    <Counters total="{{total}}" executed="{{total}}" passed="{{passed}}" failed="0" notExecuted="0" />
                  </ResultSummary>
                </TestRun>
                """);
            var result = await RunAsync(root, LoadBehaviorRunnerFunctions + """

                Assert-ExcelBehaviorReport -Path result.trx -Discovered @('case')
                """);
            if (valid) { Assert.True(result.ExitCode == 0, result.Output); }
            else { Assert.NotEqual(0, result.ExitCode); }
        }
        finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("failed", 1)]
    [InlineData("notExecuted", 1)]
    [InlineData("executed", 0)]
    [InlineData("error", 1)]
    [InlineData("timeout", 1)]
    [InlineData("aborted", 1)]
    public async Task ExcelBehavior_RejectsContradictoryCountersEvenWhenEveryRowPassed(
        string counter, int value)
    {
        var root = NewSandbox();
        try
        {
            var report = System.Xml.Linq.XDocument.Parse("""
                <TestRun>
                  <Results><UnitTestResult testName="case" outcome="Passed" /></Results>
                  <ResultSummary outcome="Completed">
                    <Counters total="1" executed="1" passed="1" failed="0" notExecuted="0" />
                  </ResultSummary>
                </TestRun>
                """);
            report.Descendants("Counters").Single().SetAttributeValue(counter, value);
            report.Save(Path.Combine(root, "result.trx"));
            var result = await RunAsync(root, LoadBehaviorRunnerFunctions + """

                Assert-ExcelBehaviorReport -Path result.trx -Discovered @('case')
                """);
            Assert.NotEqual(0, result.ExitCode);
        }
        finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("<broken")]
    [InlineData("<TestRun/>")]
    public async Task ExcelBehavior_RejectsMissingOrMalformedReports(string? report)
    {
        var root = NewSandbox();
        try
        {
            if (report != null) { File.WriteAllText(Path.Combine(root, "result.trx"), report); }
            var result = await RunAsync(root, LoadBehaviorRunnerFunctions + """

                Assert-ExcelBehaviorReport -Path result.trx -Discovered @('case')
                """);
            Assert.NotEqual(0, result.ExitCode);
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task ExcelBehavior_DiscoveryDoesNotCountWarningsAsCases()
    {
        var root = NewSandbox();
        try
        {
            var result = await RunAsync(root, LoadBehaviorRunnerFunctions + """

                $cases = @(Get-ExcelBehaviorDiscoveredCases -Output "warning`nThe following Tests are available:`n    case(row: 1)`n    case(row: 1)`n")
                Assert-ExcelBehaviorCases -Expected @('case(row: 1)', 'case(row: 1)') -Actual $cases
                """);
            Assert.True(result.ExitCode == 0, result.Output);
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task ExcelBehavior_HardDeadlineTerminatesTheOwnedChildAndRetainsEvidence()
    {
        var root = NewSandbox();
        try
        {
            var result = await RunAsync(root, LoadBehaviorRunnerFunctions + """

                Invoke-ExcelBehaviorProcess -Executable pwsh -WorkingDirectory (Get-Location).Path -LogBase run -DeadlineSeconds 1 -Arguments @('-NoProfile', '-Command', 'Write-Output ready; Start-Sleep -Seconds 20')
                """);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("Hard deadline exceeded", result.Output, StringComparison.Ordinal);
            Assert.True(File.Exists(Path.Combine(root, "run.stdout.txt")));
            Assert.True(File.Exists(Path.Combine(root, "run.command.json")));
            Assert.True(File.Exists(Path.Combine(root, "run.stderr.txt")));
            using var identity = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "run.process.json")));
            var processId = identity.RootElement.GetProperty("processId").GetInt32();
            var startedAt = identity.RootElement.GetProperty("startedAtUtcFileTime").GetInt64();
            try
            {
                using var process = Process.GetProcessById(processId);
                Assert.True(process.HasExited ||
                    process.StartTime.ToUniversalTime().ToFileTimeUtc() != startedAt,
                    "The exact child process survived its hard deadline.");
            }
            catch (ArgumentException)
            {
                // The recorded PID no longer exists.
            }
        }
        finally { Directory.Delete(root, true); }
    }
}
