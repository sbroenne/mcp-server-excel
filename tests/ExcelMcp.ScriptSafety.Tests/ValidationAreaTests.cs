using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ValidationAreaTests
{
    [Theory]
    [InlineData("tests/AGENTS.md", "")]
    [InlineData("vscode-extension/AGENTS.md", "")]
    [InlineData("src/ExcelMcp.CLI/README.md", "")]
    [InlineData("Directory.Build.targets", "csharp")]
    [InlineData("NuGet.Config", "csharp")]
    [InlineData("global.json", "csharp")]
    [InlineData("scripts/PluginContent.mjs", "javascript-typescript")]
    [InlineData("vscode-extension/vitest.config.mts", "javascript-typescript")]
    [InlineData("npm-packages/shared/package-lock.json", "javascript-typescript")]
    [InlineData("gh-pages/hooks.py", "python")]
    [InlineData("llm-tests/requirements.txt", "python")]
    [InlineData(".github/workflows/link-check.yml", "actions")]
    [InlineData(".github/actions/example/action.yaml", "actions")]
    [InlineData(".github/workflows/codeql.yml", "actions,csharp,javascript-typescript,python")]
    [InlineData(".github/codeql/codeql-config.yml", "actions,csharp,javascript-typescript,python")]
    [InlineData("scripts/Get-ValidationPlan.ps1", "actions,csharp,javascript-typescript,python")]
    public async Task CodeQl_SelectsActualLanguageInputs(string path, string languages)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            if (((Get-ValidationLanguages '{{path}}') -join ',') -cne '{{languages}}') { throw 'Wrong language inputs.' }
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

    [Theory]
    [InlineData("tests/ExcelMcp.McpServer.Tests/Integration/Tools/CalculationGuidanceContractTests.cs", "Fast", "tests\\ExcelMcp.McpServer.Tests\\ExcelMcp.McpServer.Tests.csproj")]
    [InlineData("scripts/Publish-PreparedPlugins.ps1", "Tooling", "tests\\ExcelMcp.Packaging.Tests\\ExcelMcp.Packaging.Tests.csproj")]
    [InlineData("doc-counts.json", "Tooling", "Sbroenne.ExcelMcp.sln")]
    public async Task PreparatoryBuilds_SelectOnlyOwningProjects(string path, string group, string projects)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            . .\scripts\ValidationExecution.ps1
            $plan = Get-ValidationPlan -Paths '{{path}}'
            if (((Get-SelectedBuildProjects $plan '{{group}}') -join ',') -cne '{{projects}}') { throw 'Wrong preparatory projects.' }
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
            foreach ($entry in @(@('skills/excel-mcp-report-formatting/AGENTS.md','SkillGeneration'), @('.github/plugins/excel-cli/AGENTS.md','Packaging'))) {
                $plan = Get-ValidationPlan -Paths $entry[0]
                if (($plan.ToolingProjects -join ',') -cne $entry[1] -or $plan.Excel) { throw 'Copied template was treated as unrelated guidance.' }
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }
}
