using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ValidationAreaTests
{
    [Theory]
    [InlineData("tests/AGENTS.md", "")]
    [InlineData("tests/README.md", "")]
    [InlineData("vscode-extension/AGENTS.md", "")]
    [InlineData("llm-tests/AGENTS.md", "")]
    [InlineData("src/ExcelMcp.CLI/README.md", "")]
    [InlineData(".github/workflows/docs/publish-plugins-setup.md", "")]
    [InlineData("src/ExcelMcp.Core/Commands/DataModel/DataModelCommands.cs", "csharp")]
    [InlineData("tests/ExcelMcp.Core.Tests/Unit/Example.cs", "csharp")]
    [InlineData("Directory.Build.targets", "csharp")]
    [InlineData("NuGet.Config", "csharp")]
    [InlineData("global.json", "csharp")]
    [InlineData("scripts/PluginContent.mjs", "javascript-typescript")]
    [InlineData("scripts/AwesomeCopilotPolicy.mjs", "javascript-typescript")]
    [InlineData("tests/ExcelMcp.Packaging.Tests/PluginPublicationHistory.test.mjs", "javascript-typescript")]
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
        var result = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths '{{path}}'
            if (($plan.CodeQlLanguages -join ',') -cne '{{languages}}') {
                throw "Wrong CodeQL languages: $($plan.CodeQlLanguages)"
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task CodeQl_TrackedSourceFilesSelectTheirLanguage()
    {
        var result = await ValidationSelectionTests.RunAsync("""
            $paths = @(git -c core.quotepath=false ls-files -- '*.cs' '*.csproj' '*.sln' '*.props' '*.targets' `
                '*.js' '*.jsx' '*.mjs' '*.cjs' '*.ts' '*.tsx' '*.mts' '*.cts' '*.py' '.github/workflows/*.yml')
            if ($LASTEXITCODE -ne 0 -or -not $paths.Count) { throw 'Cannot inventory tracked language inputs.' }
            foreach ($path in $paths) {
                $language = switch -Regex ($path) {
                    '\.(cs|csproj|sln|props|targets)$' { 'csharp'; break }
                    '\.(js|jsx|mjs|cjs|ts|tsx|mts|cts)$' { 'javascript-typescript'; break }
                    '\.py$' { 'python'; break }
                    '^\.github/workflows/[^/]+\.yml$' { 'actions'; break }
                }
                if ($language -and $language -notin (Get-ValidationPlan -Paths $path).CodeQlLanguages) {
                    throw "Tracked $language input has no analysis: $path"
                }
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Theory]
    [InlineData("tests/README.md", "")]
    [InlineData("tests/AGENTS.md", "")]
    [InlineData("vscode-extension/AGENTS.md", "")]
    [InlineData("src/ExcelMcp.CLI/README.md", "Cli")]
    [InlineData("src/ExcelMcp.McpServer/README.md", "Mcp")]
    [InlineData("npm-packages/excelcli/README.md", "Cli")]
    [InlineData("npm-packages/excelcli-win32-x64/README.md", "Cli")]
    [InlineData("npm-packages/excelcli-win32-arm64/README.md", "Cli")]
    [InlineData("npm-packages/mcp-server-excel/README.md", "Mcp")]
    [InlineData("npm-packages/mcp-server-excel-win32-x64/README.md", "Mcp")]
    [InlineData("npm-packages/mcp-server-excel-win32-arm64/README.md", "Mcp")]
    [InlineData("README.md", "Cli,Mcp")]
    [InlineData("docs/AGENT-SKILLS.md", "Skills")]
    [InlineData("CHANGELOG.md", "Cli,Mcp,Extension,Mcpb")]
    [InlineData("LICENSE", "Cli,Mcp,Extension,Mcpb")]
    public async Task Documentation_SelectsPackageContentsWithoutRuntimeTests(string path, string components)
    {
        var result = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths '{{path}}'
            if ($plan.Excel -or $plan.FastProjects.Count -or $plan.ProcessProjects.Count) {
                throw 'Documentation selected runtime validation.'
            }
            $selected = @('Cli','Mcp','Extension','Mcpb','Skills','Plugins') | Where-Object { $plan.$_ }
            if (($selected -join ',') -cne '{{components}}') { throw "Wrong packages: $selected" }
            $owner = if (-not '{{components}}') { '' } elseif ('{{components}}' -eq 'Skills') { 'SkillGeneration' } else { 'Packaging' }
            if (($plan.ToolingProjects -join ',') -cne $owner) { throw 'Package documentation checks were not selected.' }
            if ($owner -and -not $plan.ToolingFilters[$owner]) { throw 'Package documentation has no test filter.' }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task RuntimeChanges_KeepContractsAndBinaryPackagesWithoutPublicationTests()
    {
        var result = await ValidationSelectionTests.RunAsync("""
            $plan = Get-ValidationPlan -Paths 'src/ExcelMcp.Core/Commands/DataModel/DataModelCommands.Read.cs'
            if (($plan.FastProjects -join ',') -ne 'CLI,ComInterop,Core,McpServer,Service') {
                throw 'Shared runtime coverage lost.'
            }
            if (($plan.ProcessProjects -join ',') -ne 'CLI' -or -not $plan.Excel) {
                throw 'Runtime acceptance coverage lost.'
            }
            if ($plan.ToolingProjects.Count) { throw 'Unrelated publication tests selected.' }
            $selected = @('Cli','Mcp','Extension','Mcpb','Skills','Plugins') | Where-Object { $plan.$_ }
            if (($selected -join ',') -ne 'Cli,Mcp,Extension') { throw "Wrong binary packages: $selected" }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Theory]
    [InlineData("skills/excel-mcp-report-formatting/AGENTS.md", "SkillGeneration")]
    [InlineData(".github/plugins/excel-cli/AGENTS.md", "Packaging")]
    public async Task CopiedPackageTemplates_RemainInputsUntilDeveloperFilesAreExcluded(string path, string owner)
    {
        var result = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths '{{path}}'
            if (-not $plan.Packages -or ($plan.ToolingProjects -join ',') -cne '{{owner}}') {
                throw 'A copied package template was treated as unshipped documentation.'
            }
            if ($plan.FastProjects.Count -or $plan.ProcessProjects.Count -or $plan.Excel) {
                throw 'Package template selected unrelated runtime validation.'
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task MixedToolingOwners_DoNotBroadenEachOthersFilters()
    {
        var result = await ValidationSelectionTests.RunAsync("""
            $plan = Get-ValidationPlan -Paths @('doc-counts.json','tests/ExcelMcp.ScriptSafety.Tests/TestSelectionTests.cs')
            if ($plan.ToolingFilters.Packaging -cne 'FullyQualifiedName~DocumentationCounts') {
                throw "Packaging filter broadened: $($plan.ToolingFilters.Packaging)"
            }
            if ($plan.ToolingFilters.ScriptSafety -cne 'RequiresExcel=false') {
                throw 'Owning ScriptSafety coverage lost.'
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task StandaloneToolingProjectInputs_DoNotSelectRuntimeProjects()
    {
        var result = await ValidationSelectionTests.RunAsync("""
            $plan = Get-ValidationPlan -Paths 'tests/ExcelMcp.Packaging.Tests/ExcelMcp.Packaging.Tests.csproj'
            if (($plan.ToolingProjects -join ',') -ne 'Packaging') { throw 'Wrong tooling owners.' }
            if ($plan.FastProjects.Count -or $plan.ProcessProjects.Count -or $plan.ExcelGroups.Count) {
                throw 'Standalone tooling project selected runtime tests.'
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task MixedCodeAndDocumentation_PreserveRelevantLanguageAndProject()
    {
        var result = await ValidationSelectionTests.RunAsync("""
            $plan = Get-ValidationPlan -Paths @(
                'src\ExcelMcp.CLI\README.md',
                'src/ExcelMcp.McpServer/README.md',
                'tests/ExcelMcp.McpServer.Tests/Integration/Tools/CalculationGuidanceContractTests.cs',
                'gh-pages/hooks.py')
            if (($plan.CodeQlLanguages -join ',') -ne 'csharp,python') { throw 'Actual code language missed.' }
            if (($plan.FastProjects -join ',') -ne 'McpServer') { throw 'MCP test selection broadened.' }
            if ($plan.ProcessProjects.Count -or $plan.Excel) { throw 'Documentation selected daemon or Excel tests.' }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task FullSelection_PreservesAllLanguagesAndPerProjectFilters()
    {
        var result = await ValidationSelectionTests.RunAsync("""
            $plan = Get-ValidationPlan -Full
            if (($plan.CodeQlLanguages -join ',') -ne 'actions,csharp,javascript-typescript,python') {
                throw 'Full language coverage lost.'
            }
            foreach ($owner in @('Packaging','ScriptSafety','SkillGeneration')) {
                if ($plan.ToolingFilters[$owner] -cne 'RequiresExcel=false') { throw "$owner coverage narrowed." }
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Theory]
    [InlineData("tests/ExcelMcp.McpServer.Tests/Example.cs", "Fast",
        "tests\\ExcelMcp.McpServer.Tests\\ExcelMcp.McpServer.Tests.csproj")]
    [InlineData("scripts/Publish-PreparedPlugins.ps1", "Tooling",
        "tests\\ExcelMcp.Packaging.Tests\\ExcelMcp.Packaging.Tests.csproj")]
    [InlineData("src/ExcelMcp.Core/Commands/DataModel/Example.cs", "Fast", "Sbroenne.ExcelMcp.sln")]
    [InlineData("src/ExcelMcp.CLI/Program.cs", "Process",
        "tests\\ExcelMcp.CLI.Tests\\ExcelMcp.CLI.Tests.csproj")]
    [InlineData("doc-counts.json", "Tooling", "Sbroenne.ExcelMcp.sln")]
    public async Task PreparatoryBuilds_SelectOnlyNeededProjects(string path, string group, string expected)
    {
        var result = await ValidationSelectionTests.RunAsync($$"""
            $file = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.BuildPlan.$([Guid]::NewGuid().ToString('N')).json"
            try {
                Get-ValidationPlan -Paths '{{path}}' | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $file
                $projects = @(& .\scripts\Build-CiInputs.ps1 -PlanFile $file -Group '{{group}}' -ListProjects)
                if (($projects -join ',') -cne '{{expected}}') { throw "Wrong build projects: $projects" }
            } finally { Remove-Item -LiteralPath $file }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Theory]
    [InlineData("restore", false)]
    [InlineData("build", false)]
    [InlineData("", true)]
    public async Task PreparatoryBuilds_ExecuteSelectedProjectAndPropagateFailures(string failure, bool succeeds)
    {
        var result = await ValidationSelectionTests.RunAsync($$"""
            $file = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.BuildExecution.$([Guid]::NewGuid().ToString('N')).json"
            $global:calls = [Collections.Generic.List[string]]::new()
            function global:dotnet {
                $global:calls.Add($args -join ' ')
                $global:LASTEXITCODE = if ($args[0] -eq '{{failure}}') { 17 } else { 0 }
            }
            try {
                Get-ValidationPlan -Paths 'tests/ExcelMcp.McpServer.Tests/Example.cs' |
                    ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $file
                $caught = $null
                try { & .\scripts\Build-CiInputs.ps1 -PlanFile $file -Group Fast } catch { $caught = $_ }
                if ('{{failure}}') {
                    if (-not $caught -or $caught.Exception.Message -notmatch '{{failure}} failed') {
                        throw 'Native failure was not propagated.'
                    }
                } elseif ($caught -or $LASTEXITCODE -ne 0) { throw 'Successful preparation did not complete.' }
                $expectedCalls = if ('{{failure}}' -eq 'restore') { 1 } else { 2 }
                if ($global:calls.Count -ne $expectedCalls) { throw 'Unexpected native command count.' }
                foreach ($call in $global:calls) {
                    if ($call -notmatch 'tests.ExcelMcp\.McpServer\.Tests.ExcelMcp\.McpServer\.Tests\.csproj' -or
                        $call -match '\.sln') { throw "Unexpected project build: $call" }
                }
                if ($expectedCalls -eq 2 -and $global:calls[1] -notmatch '--disable-build-servers') {
                    throw 'Scoped build retained build servers.'
                }
                if ($caught) { throw $caught }
            } finally {
                Remove-Item -LiteralPath $file
                Remove-Item Function:\dotnet
            }
            """);
        Assert.Equal(succeeds, result.ExitCode == 0);
    }

    [Theory]
    [InlineData("{\"CiTestGroups\":[]}", "was not selected")]
    [InlineData("{\"CiTestGroups\":[\"Fast\"],\"FastProjects\":[]}", "has no build projects")]
    [InlineData("{\"CiTestGroups\":[\"Fast\"],\"FastProjects\":[\"Packaging\"]}", "Unexpected Fast project")]
    public async Task PreparatoryBuilds_RejectInvalidSelections(string plan, string error)
    {
        var result = await ValidationSelectionTests.RunAsync($$"""
            $file = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.InvalidBuild.$([Guid]::NewGuid().ToString('N')).json"
            try {
                '{{plan}}' | Set-Content -LiteralPath $file
                & .\scripts\Build-CiInputs.ps1 -PlanFile $file -Group Fast -ListProjects
            } finally { Remove-Item -LiteralPath $file }
            """);
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains(error, result.Output, StringComparison.Ordinal);
    }
}
