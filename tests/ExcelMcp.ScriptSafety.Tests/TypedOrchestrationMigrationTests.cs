using Sbroenne.ExcelMcp.Build;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TypedOrchestrationMigrationTests
{
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
                """<TestRun><Results><UnitTestResult outcome="Passed"/></Results><ResultSummary outcome="Completed"><Counters total="1" passed="1" executed="1"/></ResultSummary></TestRun>""");
            return new ProcessResult(0, "fixture", "");
        }
        public Task<ProcessResult> CheckedAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
            IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false) =>
            throw new InvalidOperationException("Inventory must not execute workbook tests or process cleanup.");
    }
}
