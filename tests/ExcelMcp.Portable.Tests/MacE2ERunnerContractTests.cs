using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Trait("RequiresExcel", "false")]
public sealed class MacE2ERunnerContractTests
{
    [Fact]
    public void RunnerRetainsExistingWorkflowsAndIncludesNativeAcceptanceWithoutCopiedCounts()
    {
        var script = File.ReadAllText(Path.Combine(FindRepository(), "scripts", "Test-MacE2E.ps1"));
        Assert.Contains("FullyQualifiedName~MacExcelE2ETests", script, StringComparison.Ordinal);
        Assert.Contains("FullyQualifiedName~MacRangeEditE2ETests", script, StringComparison.Ordinal);
        Assert.Contains("FullyQualifiedName~MacNativeSessionE2ETests", script, StringComparison.Ordinal);
        Assert.Contains("FullyQualifiedName~MacNativeWorksheetE2ETests", script, StringComparison.Ordinal);
        Assert.Contains("FullyQualifiedName~MacAppleEventDesktopTests", script, StringComparison.Ordinal);
        Assert.Contains("FullyQualifiedName~MacNativeFormulaApiTests", script, StringComparison.Ordinal);
        Assert.Contains("FullyQualifiedName~MacNativeFormulaE2ETests", script, StringComparison.Ordinal);
        Assert.Contains("--list-tests", script, StringComparison.Ordinal);
        Assert.Contains("Invoke-TestStage", script, StringComparison.Ordinal);
        Assert.Contains("Compare-Object", script, StringComparison.Ordinal);
        Assert.DoesNotContain("$expectedPassed = 22", script, StringComparison.Ordinal);
        Assert.DoesNotContain("$summaryPattern", script, StringComparison.Ordinal);
    }

    [Fact]
    public void Runner_LaunchesExcelNormallyBeforeNonPromptingAutomationCheck()
    {
        var script = File.ReadAllText(Path.Combine(
            FindRepository(),
            "scripts",
            "Test-MacE2E.ps1"));

        var launch = script.IndexOf(
            "Invoke-MacTestCommand /usr/bin/open @('-a', 'Microsoft Excel')",
            StringComparison.Ordinal);
        var permissionCheck = script.IndexOf(
            "Assert-MacAutomationAllowed",
            StringComparison.Ordinal);

        Assert.True(launch >= 0, "The runner must launch Excel through LaunchServices.");
        Assert.True(
            permissionCheck > launch,
            "The non-prompting Automation check must run after Excel has launched.");
        Assert.DoesNotContain("System Events", script, StringComparison.Ordinal);
    }

    private static string FindRepository()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null
               && !File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln")))
        {
            directory = directory.Parent;
        }

        return directory?.FullName
            ?? throw new InvalidOperationException("Repository root not found.");
    }
}
