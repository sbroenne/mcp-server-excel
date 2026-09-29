using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "AutomationSafety")]
public sealed class AutomationSafetyTests
{
    private static readonly string RepoRoot = FindRepoRoot();

    [Theory]
    [InlineData("check-com-leaks.ps1", "dynamic item = source.Item;")]
    [InlineData("check-success-flag.ps1", "result.Success = true;\nresult.ErrorMessage = \"failure\";")]
    [InlineData("check-dynamic-casts.ps1", "var item = ((dynamic)source).Item;")]
    public async Task SourceGuards_RejectEmptyDiscoveryAndIgnoreGeneratedFiles(string script, string suspicious)
    {
        var root = NewSandbox();
        try
        {
            var scripts = Directory.CreateDirectory(Path.Combine(root, "scripts")).FullName;
            File.Copy(Path.Combine(RepoRoot, "scripts", script), Path.Combine(scripts, script));
            Directory.CreateDirectory(Path.Combine(root, "src", "ExcelMcp.Core", "Commands"));
            Directory.CreateDirectory(Path.Combine(root, "src", "ExcelMcp.ComInterop"));
            var command = $"& '{Quote(Path.Combine(scripts, script))}'";
            var empty = await RunAsync(root, command);
            Assert.NotEqual(0, empty.ExitCode);

            var commands = Path.Combine(root, "src", "ExcelMcp.Core", "Commands");
            var source = Path.Combine(commands, "Example.cs");
            File.WriteAllText(source, suspicious);
            File.WriteAllText(Path.Combine(root, "src", "ExcelMcp.ComInterop", "Example.cs"), "class Example {}");
            var invalid = await RunAsync(root, command);
            Assert.NotEqual(0, invalid.ExitCode);

            File.WriteAllText(source, "class Example {}");
            var generated = Directory.CreateDirectory(Path.Combine(commands, "obj")).FullName;
            File.WriteAllText(Path.Combine(generated, "Generated.cs"), suspicious);
            var valid = await RunAsync(root, command);
            Assert.True(valid.ExitCode == 0, valid.Output);
        }
        finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("configure-analytics-oidc.ps1", "-SubscriptionId requested")]
    [InlineData("deploy-appinsights.ps1", "")]
    public async Task AzurePreview_NeverMutatesAndRejectsFailedReads(string script, string arguments)
    {
        var root = NewSandbox();
        try
        {
            var target = Quote(Path.Combine(RepoRoot, "infrastructure", "azure", script));
            var calls = Path.Combine(root, "calls.txt");
            var setup = $$"""
                function az {
                    $call = $args -join ' '
                    Add-Content -LiteralPath '{{Quote(calls)}}' -Value $call -WhatIf:$false
                    $global:LASTEXITCODE = if ($env:MOCK_AZ_FAILURE -eq 'yes') { 23 } else { 0 }
                    switch -Regex ($call) {
                        '^version' { '{"azure-cli":"fixture"}'; break }
                        '^account show' { '{"id":"requested","tenantId":"tenant","name":"fixture"}'; break }
                        '^monitor' { '{"id":"/subscriptions/requested/resourceGroups/fixture/workspaces/test"}'; break }
                        '^ad app list' { '[{"appId":"application"}]'; break }
                        '^ad sp list' { '[{"id":"principal"}]'; break }
                        '^ad app federated-credential list' { '[{"name":"github-main-usage-analytics"}]'; break }
                        '^role assignment list' { '[{"id":"assignment"}]'; break }
                        default { '{}' }
                    }
                }
                function gh { throw 'Preview invoked a GitHub write.' }
                """;
            var preview = await RunAsync(root, setup + $"\n& '{target}' {arguments} -WhatIf");
            Assert.True(preview.ExitCode == 0, preview.Output);
            var recorded = File.ReadAllText(calls);
            Assert.DoesNotContain("account set", recorded, StringComparison.Ordinal);
            Assert.DoesNotContain(" create", recorded, StringComparison.Ordinal);
            Assert.DoesNotContain("login", recorded, StringComparison.Ordinal);
            Assert.DoesNotContain("Configured read-only", preview.Output, StringComparison.Ordinal);
            if (script == "deploy-appinsights.ps1")
            {
                Assert.Contains("deployment sub what-if --subscription requested", recorded, StringComparison.Ordinal);
            }
            File.Delete(calls);
            var failed = await RunAsync(root, setup + $"\n$env:MOCK_AZ_FAILURE='yes'\n& '{target}' {arguments} -WhatIf");
            Assert.NotEqual(0, failed.ExitCode);
            Assert.Single(File.ReadAllLines(calls));
        }
        finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("Cli", "src/ExcelMcp.CLI/Program.cs")]
    [InlineData("Mcp", "src/ExcelMcp.McpServer/Program.cs")]
    [InlineData("Plugins", ".github/workflows/publish-plugins.yml")]
    [InlineData("Extension", "vscode-extension/package.json")]
    [InlineData("Mcpb", "mcpb/manifest.json")]
    public async Task PackageSelection_ChoosesOwningComponent(string component, string path)
    {
        var root = NewSandbox();
        try
        {
            var result = await RunAsync(root, $$"""
                . '{{Quote(Path.Combine(RepoRoot, "scripts", "Get-ValidationPlan.ps1"))}}'
                $plan = Get-ValidationPlan -Paths '{{path}}'
                if (-not $plan.{{component}}) { throw 'Owning package was not selected.' }
                """);
            Assert.True(result.ExitCode == 0, result.Output);
        }
        finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("src/ExcelMcp.CLI/Program.cs", "Cli,Skills,Plugins")]
    [InlineData("src/ExcelMcp.McpServer/Program.cs", "Mcp,Extension,Mcpb,Skills,Plugins")]
    [InlineData("npm-packages/shared/launcher.js", "Cli,Mcp")]
    [InlineData(".npmrc", "Cli,Mcp,Extension,Mcpb,Skills,Plugins")]
    [InlineData("README.md", "")]
    public async Task PackageSelection_ExcludesUnrelatedDistributions(string path, string expected)
    {
        var root = NewSandbox();
        try
        {
            var result = await RunAsync(root, $$"""
                . '{{Quote(Path.Combine(RepoRoot, "scripts", "Get-ValidationPlan.ps1"))}}'
                $plan = Get-ValidationPlan -Paths '{{path}}'
                $selected = @('Cli','Mcp','Extension','Mcpb','Skills','Plugins') | Where-Object { $plan.$_ }
                if (($selected -join ',') -ne '{{expected}}') { throw "Wrong package selection: $selected" }
                """);
            Assert.True(result.ExitCode == 0, result.Output);
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task CaptureEvidence_NeverSavesAndStillQuitsAfterCloseFails()
    {
        var root = NewSandbox();
        try
        {
            var script = Path.Combine(RepoRoot, "videos", "excel-mcp-intro", "Capture-Evidence.ps1");
            var result = await RunAsync(root, $$"""
                $ast = [Management.Automation.Language.Parser]::ParseFile('{{Quote(script)}}', [ref]$null, [ref]$null)
                $outer = $ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.TryStatementAst] } | Select-Object -Last 1
                $book = [pscustomobject]@{}
                $book | Add-Member ScriptMethod Close { param($save) throw 'simulated close failure' }
                $excel = [pscustomobject]@{}
                $excel | Add-Member ScriptMethod Quit { Write-Output 'quit-after-close-failure' }
                $objects = [Collections.Generic.List[object]]::new()
                & ([scriptblock]::Create('try {} finally ' + $outer.Finally.Extent.Text))
                """);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("quit-after-close-failure", result.Output, StringComparison.Ordinal);
            Assert.DoesNotContain("$book.Save()", File.ReadAllText(script), StringComparison.Ordinal);
        }
        finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(23, true)]
    [InlineData(0, false)]
    public async Task CliWorkflow_CannotPassFailedCommands(int exitCode, bool success)
    {
        var root = NewSandbox();
        try
        {
            var script = Path.Combine(RepoRoot, "scripts", "Test-CliWorkflow.ps1");
            var result = await RunAsync(root, $$"""
                $ast = [Management.Automation.Language.Parser]::ParseFile('{{Quote(script)}}', [ref]$null, [ref]$null)
                $function = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Test-Step' }, $true)
                . ([scriptblock]::Create($function.Extent.Text))
                $script:passed=0
                $script:failed=0
                Test-Step 'fake command' { $global:LASTEXITCODE={{exitCode}}; [pscustomobject]@{ success=${{success.ToString().ToLowerInvariant()}} } } -Verify { $true } | Out-Null
                if ($script:failed -ne 1 -or $script:passed -ne 0) { throw 'Failed command was accepted.' }
                $global:LASTEXITCODE=0
                """);
            Assert.True(result.ExitCode == 0, result.Output);
        }
        finally { Directory.Delete(root, true); }
    }

    private static string NewSandbox()
        => Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"ExcelMcp.Automation.{Guid.NewGuid():N}")).FullName;

    private static string Quote(string value) => value.Replace("'", "''", StringComparison.Ordinal);

    private static async Task<(int ExitCode, string Output)> RunAsync(string root, string body)
    {
        var script = Path.Combine(root, $"test-{Guid.NewGuid():N}.ps1");
        File.WriteAllText(script, "$ErrorActionPreference='Stop'\n$global:LASTEXITCODE=0\n" + body + "\nexit $LASTEXITCODE");
        var info = new ProcessStartInfo("pwsh")
        {
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
            WorkingDirectory = root
        };
        foreach (var argument in new[] { "-NoProfile", "-File", script }) { info.ArgumentList.Add(argument); }
        using var process = Process.Start(info)!;
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        try { await process.WaitForExitAsync(timeout.Token); }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
            throw new TimeoutException("Automation script exceeded its 30-second deadline.");
        }
        return (process.ExitCode, await stdout + await stderr);
    }

    private static string FindRepoRoot()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory != null; directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln"))) { return directory.FullName; }
        }
        throw new DirectoryNotFoundException("Repository root was not found.");
    }
}
