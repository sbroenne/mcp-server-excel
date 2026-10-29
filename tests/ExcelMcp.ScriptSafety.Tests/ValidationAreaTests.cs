using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ValidationAreaTests
{
    [Fact]
    public async Task CodeQl_SelectsActualLanguageInputs()
    {
        var run = await ValidationSelectionTests.RunCasesAsync(
        [
            ["tests/AGENTS.md", ""],
            ["vscode-extension/AGENTS.md", ""],
            ["src/ExcelMcp.CLI/README.md", ""],
            ["Directory.Build.targets", "csharp"],
            ["NuGet.Config", "csharp"],
            ["global.json", "csharp"],
            ["scripts/PluginContent.mjs", "javascript-typescript"],
            ["vscode-extension/vitest.config.mts", "javascript-typescript"],
            ["npm-packages/shared/package-lock.json", "javascript-typescript"],
            ["gh-pages/hooks.py", "python"],
            ["llm-tests/requirements.txt", "python"],
            [".github/workflows/link-check.yml", "actions"],
            [".github/actions/example/action.yaml", "actions"],
            [".github/workflows/codeql.yml", "actions,csharp,javascript-typescript,python"],
            [".github/codeql/codeql-config.yml", "actions,csharp,javascript-typescript,python"],
            ["scripts/Get-ValidationPlan.ps1", "actions,csharp,javascript-typescript,python"]
        ], """
            $path, $languages = $case
            if (((Get-ValidationLanguages $path) -join ',') -cne $languages) { throw 'Wrong language inputs.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task CodeQl_TrackedSourceFilesSelectTheirLanguage()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $paths = @(git -c core.quotepath=false ls-files -- '*.cs' '*.csproj' '*.sln' '*.props' '*.targets' '*.js' '*.jsx' '*.mjs' '*.cjs' '*.ts' '*.tsx' '*.mts' '*.cts' '*.py' '.github/workflows/*.yml')
            if ($LASTEXITCODE -ne 0 -or -not $paths.Count) { throw 'Cannot inspect tracked sources.' }
            foreach ($path in $paths) {
                $extension = [IO.Path]::GetExtension($path)
                $language = switch -Regex ($extension) {
                    '^\.(cs|csproj|sln|props|targets)$' { 'csharp' }
                    '^\.(js|jsx|mjs|cjs|ts|tsx|mts|cts)$' { 'javascript-typescript' }
                    '^\.py$' { 'python' }
                    '^\.yml$' { 'actions' }
                }
                if ($language -notin @(Get-ValidationLanguages $path)) { throw "Language lost: $path" }
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task PreparatoryBuilds_SelectOnlyOwningProjects()
    {
        var run = await ValidationSelectionTests.RunCasesAsync(
        [
            ["tests/ExcelMcp.McpServer.Tests/Integration/Tools/CalculationGuidanceContractTests.cs", "Fast", "tests\\ExcelMcp.McpServer.Tests\\ExcelMcp.McpServer.Tests.csproj"],
            ["scripts/Publish-PreparedPlugins.ps1", "Tooling", "tests\\ExcelMcp.Packaging.Tests\\ExcelMcp.Packaging.Tests.csproj"],
            ["doc-counts.json", "Tooling", "Sbroenne.ExcelMcp.sln"]
        ], """
            $path, $group, $projects = $case
            . .\scripts\ValidationExecution.ps1
            $plan = Get-ValidationPlan -Paths $path
            if (((Get-SelectedBuildProjects $plan $group) -join ',') -cne $projects) { throw 'Wrong preparatory projects.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task IssueTriage_AndCopiedTemplatesKeepTheirConsumers()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            foreach ($path in @('scripts/IssueTriage.mjs','.github/workflows/issue-triage.md','.github/workflows/issue-triage.lock.yml')) {
                $plan = Get-ValidationPlan -Paths $path
                if ($plan.Excel -or $plan.FastProjects.Count -or $plan.ProcessProjects.Count -or
                    $plan.ToolingFilters.ScriptSafety -notmatch 'IssueTriageScriptTests') { throw 'Wrong issue automation consumers.' }
            }
            foreach ($entry in @(@('skills/excel-mcp-report-formatting/AGENTS.md','Packaging,SkillGeneration'), @('.github/plugins/excel-cli/AGENTS.md','Packaging'))) {
                $plan = Get-ValidationPlan -Paths $entry[0]
                if (($plan.ToolingProjects -join ',') -cne $entry[1] -or $plan.Excel) { throw 'Copied template was treated as unrelated guidance.' }
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }
}
