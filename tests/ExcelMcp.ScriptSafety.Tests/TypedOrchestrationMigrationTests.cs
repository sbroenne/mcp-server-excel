using Sbroenne.ExcelMcp.Build;
using System.ComponentModel;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TypedOrchestrationMigrationTests
{
    [Fact]
    public async Task CompleteCiBuild_RestoresAndBuildsSolutionInOnlyOneSelectedGroup()
    {
        var plan = new ValidationPolicy(TypedValidationPolicyTests.Root).Select([], full: true);
        var runner = new BuildRunner();
        var execution = new ValidationExecution(TypedValidationPolicyTests.Root, runner);
        foreach (var group in plan.CiTestGroups)
        {
            var projects = execution.BuildProjects(plan, group);
            if (group == plan.SourceChecksGroup) { Assert.Equal(["Sbroenne.ExcelMcp.sln"], projects); }
            else { Assert.All(projects, project => Assert.EndsWith(".Tests.csproj", project, StringComparison.Ordinal)); }
            await execution.BuildAsync(plan, group);
        }
        foreach (var verb in new[] { "restore", "build" })
        {
            var commands = runner.Commands.Where(command => command[0] == verb).ToArray();
            Assert.Single(commands, command => Path.GetFileName(command[1]) == "Sbroenne.ExcelMcp.sln");
            Assert.Contains(commands, command => command[1].EndsWith("ExcelMcp.CLI.Tests.csproj", StringComparison.Ordinal));
            Assert.Contains(commands, command => command[1].EndsWith("ExcelMcp.ScriptSafety.Tests.csproj", StringComparison.Ordinal));
        }
    }

    [Theory]
    [InlineData(true, false, true, false, false, "Fast")]
    [InlineData(false, true, true, false, false, "Process")]
    [InlineData(false, false, true, false, false, "Tooling")]
    [InlineData(true, false, true, false, true, "Tooling")]
    [InlineData(true, false, true, true, false, "Fast")]
    [InlineData(false, true, false, false, true, "Process")]
    public void CompleteGroupedBuild_HasOneDeterministicOwner(
        bool fast, bool process, bool tooling, bool sourceChecks, bool documentationCounts, string expectedOwner)
    {
        var plan = new ValidationPlan
        {
            FullSolutionBuild = true,
            SourceChecks = sourceChecks,
            DocumentationCounts = documentationCounts
        };
        if (fast) { plan.FastFilters["Core"] = "All"; }
        if (process) { plan.ProcessFilters["CLI"] = "All"; }
        if (tooling) { plan.ToolingFilters["ScriptSafety"] = "All"; }
        var execution = new ValidationExecution(TypedValidationPolicyTests.Root, new BuildRunner());
        foreach (var group in plan.CiTestGroups)
        {
            var projects = execution.BuildProjects(plan, group);
            if (group == expectedOwner) { Assert.Equal(["Sbroenne.ExcelMcp.sln"], projects); }
            else { Assert.All(projects, project => Assert.EndsWith(".Tests.csproj", project, StringComparison.Ordinal)); }
        }
        Assert.Equal(["Sbroenne.ExcelMcp.sln"], execution.BuildProjects(plan));
        Assert.Throws<InvalidOperationException>(() => execution.BuildProjects(plan, "Unknown"));
    }

    [Fact]
    public async Task ExcelValidation_CleansUpAndPreservesUnexpectedPrimaryAndCleanupFailures()
    {
        if (!OperatingSystem.IsWindows()) { return; }
        var plan = new ValidationPlan { Excel = true };
        plan.ExcelSelections.Add(new ExcelSelection("CLI", "FullyQualifiedName~Example", "Acceptance"));
        var runner = new CleanupFailureRunner();
        var results = Path.Combine(Path.GetTempPath(), $"ExcelMcp.Cleanup.{Guid.NewGuid():N}");
        try
        {
            var exception = await Assert.ThrowsAsync<AggregateException>(() =>
                new ValidationExecution(TypedValidationPolicyTests.Root, runner).TestAsync(plan, null, results));
            Assert.IsType<UnauthorizedAccessException>(exception.InnerExceptions[0]);
            Assert.IsType<Win32Exception>(exception.InnerExceptions[1]);
            Assert.Equal("pwsh", runner.CleanupExecutable);
            Assert.Contains("Stop-ExcelMcpProcesses.ps1", string.Join(' ', runner.CleanupArguments), StringComparison.Ordinal);
        }
        finally { if (Directory.Exists(results)) { Directory.Delete(results, recursive: true); } }
    }

    [Fact]
    public void FocusedGroupedBuild_KeepsOnlyItsOwningProject()
    {
        var plan = new ValidationPolicy(TypedValidationPolicyTests.Root).Select(
            ["tests/ExcelMcp.Core.Tests/Unit/GeneratedActionContractTests.cs"]);
        var projects = new ValidationExecution(TypedValidationPolicyTests.Root, new BuildRunner()).BuildProjects(plan, "Fast");
        Assert.Equal(Path.Combine("tests", "ExcelMcp.Core.Tests", "ExcelMcp.Core.Tests.csproj"), Assert.Single(projects));
    }

    [Theory]
    [InlineData("Invoke-TestStage.ps1")]
    [InlineData("Invoke-ExcelFreeTests.ps1")]
    [InlineData("Invoke-ExcelTests.ps1")]
    [InlineData("Get-ExcelTestGroups.ps1")]
    [InlineData("Build-CiInputs.ps1")]
    [InlineData("Test-CiCompletion.ps1")]
    public void CompatibilityCommands_UseTheSharedTypedExecution(string name)
    {
        var source = File.ReadAllText(Path.Combine(TypedValidationPolicyTests.Root, "scripts", name));
        Assert.Contains("Invoke-ExcelMcpBuild", source, StringComparison.Ordinal);
        Assert.DoesNotContain("ProcessStartInfo", source, StringComparison.Ordinal);
        Assert.DoesNotContain("dotnet build", source, StringComparison.Ordinal);
        Assert.DoesNotContain("Assert-TestReport -Path $report", source, StringComparison.Ordinal);
    }

    [Fact]
    public void Completion_CannotPassWithASelectedMissingResult()
    {
        Assert.Throws<InvalidOperationException>(() => CiCompletion.Verify(new CiCompletionOptions
        {
            Detection = "success",
            Checks = [
                new("tests", "true", "skipped"), new("packages", "false", "skipped"),
                new("npm", "false", "skipped"), new("lockfiles", "false", "skipped")
            ]
        }));
    }

    [Fact]
    public async Task BuiltInventory_PreservesClassFixtureBoundariesAndMixedMethodGroups()
    {
        var results = Path.Combine(Path.GetTempPath(), $"ExcelMcp.Inventory.{Guid.NewGuid():N}");
        try
        {
            var runner = new InventoryRunner();
            var groups = await new ExcelGroupExecution(TypedValidationPolicyTests.Root, runner).InventoryAsync(results);
            Assert.Equal(["Core", "Service", "CLI", "McpServer", "ComInterop"], runner.Owners);
            Assert.Equal(15, groups.Length);
            foreach (var owner in runner.Owners)
            {
                var selected = groups.Where(item => item.ProjectName == owner).ToArray();
                Assert.Contains(selected, item => item.Filter == $"(FullyQualifiedName~Fixture.{owner}.)" && item.Cases.Length == 2);
                Assert.Contains(selected, item => item.Filter == "(FullyQualifiedName=Mixed.Normal)" && item.Group == "Lifecycle");
                Assert.Contains(selected, item => item.Filter == "(FullyQualifiedName=Mixed.Macro)" && item.Group == "VBA" && item.Cases[0].Prerequisite);
            }
            Assert.Contains("new inventory path", (await Assert.ThrowsAsync<InvalidOperationException>(
                () => new ExcelGroupExecution(TypedValidationPolicyTests.Root, runner).InventoryAsync(results))).Message, StringComparison.Ordinal);
        }
        finally { if (Directory.Exists(results)) { Directory.Delete(results, true); } }
    }

    private sealed class BuildRunner : IProcessRunner
    {
        public List<string[]> Commands { get; } = [];
        public Task<ProcessResult> CheckedAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
            IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false)
        {
            Assert.Equal("dotnet", executable);
            Commands.Add(arguments.ToArray());
            return Task.FromResult(new ProcessResult(0, "", ""));
        }
        public Task<ProcessResult> RunAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
            IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false) =>
            throw new InvalidOperationException("Builds must check command results.");
    }

    private sealed class CleanupFailureRunner : IProcessRunner
    {
        public string? CleanupExecutable { get; private set; }
        public string[] CleanupArguments { get; private set; } = [];

        public Task<ProcessResult> RunAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
            IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false) =>
            throw new UnauthorizedAccessException("test host launch failure");

        public Task<ProcessResult> CheckedAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
            IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false)
        {
            CleanupExecutable = executable;
            CleanupArguments = arguments.ToArray();
            throw new Win32Exception("owned cleanup launch failure");
        }
    }

    private sealed class InventoryRunner : IProcessRunner
    {
        public List<string> Owners { get; } = [];
        public async Task<ProcessResult> RunAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
            IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false)
        {
            Assert.Equal("dotnet", executable);
            Assert.Equal(TimeSpan.FromSeconds(120), deadline);
            Assert.False(preserveGitContext);
            var command = arguments.ToArray();
            Assert.Contains("FullyQualifiedName~ExcelValidationGroups_PreserveClassFixturesAndExportBuiltInventory", command);
            Assert.NotNull(environment);
            var path = environment["EXCELMCP_TEST_SELECTION_OUTPUT"];
            var owner = Path.GetFileName(path).Replace("-inventory.json", "", StringComparison.Ordinal);
            Owners.Add(owner);
            await File.WriteAllTextAsync(path, JsonSerializer.Serialize(new BuiltTestCase[] {
                new($"Fixture.{owner}", $"Fixture.{owner}.One", "Editing", false, false, false),
                new($"Fixture.{owner}", $"Fixture.{owner}.Two", "Editing", false, false, false),
                new("Mixed", "Mixed.Normal", "Lifecycle", false, false, false),
                new("Mixed", "Mixed.Macro", "VBA", true, true, false)
            }));
            await File.WriteAllTextAsync(Path.Combine(Path.GetDirectoryName(path)!, $"{owner}-inventory.trx"),
                """<TestRun><Results><UnitTestResult outcome="Passed"/></Results><ResultSummary outcome="Completed"><Counters total="1" passed="1" executed="1" failed="0" notExecuted="0"/></ResultSummary></TestRun>""");
            return new ProcessResult(0, "fixture", "");
        }
        public Task<ProcessResult> CheckedAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
            IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false) =>
            throw new InvalidOperationException("Inventory must not execute workbook tests or process cleanup.");
    }
}
