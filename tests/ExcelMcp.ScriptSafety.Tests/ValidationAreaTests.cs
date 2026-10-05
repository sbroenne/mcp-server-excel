using Sbroenne.ExcelMcp.Build;
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
    public void CodeQl_SelectsActualLanguageInputs(string path, string languages)
    {
        Assert.Equal(languages, string.Join(',', ValidationPolicy.LanguagesForPath(path)));
    }

    [Fact]
    public async Task CodeQl_TrackedSourceFilesSelectTheirLanguage()
    {
        var result = await new ProcessRunner(TypedValidationPolicyTests.Root).CheckedAsync("git",
            ["-c", "core.quotepath=false", "ls-files", "--", "*.cs", "*.csproj", "*.sln", "*.props", "*.targets",
             "*.js", "*.jsx", "*.mjs", "*.cjs", "*.ts", "*.tsx", "*.mts", "*.cts", "*.py", ".github/workflows/*.yml"],
            TimeSpan.FromSeconds(30));
        var paths = result.Output.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        Assert.NotEmpty(paths);
        foreach (var path in paths)
        {
            var language = Path.GetExtension(path) switch
            {
                ".cs" or ".csproj" or ".sln" or ".props" or ".targets" => "csharp",
                ".js" or ".jsx" or ".mjs" or ".cjs" or ".ts" or ".tsx" or ".mts" or ".cts" => "javascript-typescript",
                ".py" => "python",
                ".yml" => "actions",
                _ => throw new InvalidOperationException($"Unexpected tracked input: {path}.")
            };
            Assert.Contains(language, ValidationPolicy.LanguagesForPath(path));
        }
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
            if (($plan.FastProjects -join ',') -ne 'Core') {
                throw 'Unchanged adapters selected.'
            }
            if ($plan.ProcessProjects.Count -or -not $plan.Excel) {
                throw 'Owning workbook coverage lost or unchanged process tests selected.'
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
            if ($plan.ToolingFilters.ScriptSafety -notmatch 'TestSelectionTests' -or
                $plan.ToolingFilters.ScriptSafety -match 'PreCommitScriptTests') {
                throw 'Test-only selection was lost or broadened.'
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
                if (-not $plan.ToolingFilters[$owner]) { throw "$owner has no selected cases." }
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Theory]
    [InlineData("tests/ExcelMcp.McpServer.Tests/Integration/Tools/CalculationGuidanceContractTests.cs", "Fast",
        "tests\\ExcelMcp.McpServer.Tests\\ExcelMcp.McpServer.Tests.csproj")]
    [InlineData("scripts/Publish-PreparedPlugins.ps1", "Tooling",
        "tests\\ExcelMcp.Packaging.Tests\\ExcelMcp.Packaging.Tests.csproj")]
    [InlineData("src/ExcelMcp.Core/Commands/DataModel/Example.cs", "Fast",
        "tests\\ExcelMcp.Core.Tests\\ExcelMcp.Core.Tests.csproj")]
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
        var root = TypedValidationPolicyTests.Root;
        var plan = new ValidationPolicy(root).Select(["tests/ExcelMcp.McpServer.Tests/Integration/Tools/CalculationGuidanceContractTests.cs"]);
        var runner = new BuildRunner(failure);
        var error = await Record.ExceptionAsync(() => new ValidationExecution(root, runner).BuildAsync(plan, "Fast"));
        Assert.Equal(succeeds, error is null);
        if (!succeeds) { Assert.Contains(failure + " failed with exit code 17", error!.Message, StringComparison.Ordinal); }
        Assert.Equal(failure == "restore" ? 1 : 2, runner.Commands.Count);
        foreach (var command in runner.Commands)
        {
            Assert.Equal(Path.Combine(root, "tests", "ExcelMcp.McpServer.Tests", "ExcelMcp.McpServer.Tests.csproj"), command[1]);
        }
        Assert.Equal("restore", runner.Commands[0][0]);
        if (runner.Commands.Count == 2)
        {
            Assert.Equal("build", runner.Commands[1][0]);
            Assert.Contains("--disable-build-servers", runner.Commands[1]);
            Assert.Contains("--no-restore", runner.Commands[1]);
        }
    }

    private sealed class BuildRunner(string failure) : IProcessRunner
    {
        public List<string[]> Commands { get; } = [];
        public Task<ProcessResult> CheckedAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
            IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false)
        {
            Assert.Equal("dotnet", executable);
            Assert.Equal(TimeSpan.FromMinutes(20), deadline);
            Assert.Null(environment);
            Assert.False(preserveGitContext);
            var command = arguments.ToArray();
            Commands.Add(command);
            if (command[0] == failure) { throw new InvalidOperationException(command[0] + " failed with exit code 17: native-root-cause"); }
            return Task.FromResult(new ProcessResult(0, "fixture", ""));
        }
        public Task<ProcessResult> RunAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
            IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false) =>
            throw new InvalidOperationException("Unexpected unchecked build command.");
    }

    [Theory]
    [InlineData("{\"SchemaVersion\":1,\"CiTestGroups\":[]}", "was not selected")]
    [InlineData("{\"SchemaVersion\":1,\"FastFilters\":{}}", "was not selected")]
    [InlineData("{\"SchemaVersion\":1,\"FastFilters\":{\"Packaging\":\"RequiresExcel=false\"}}", "Unexpected Fast project")]
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
