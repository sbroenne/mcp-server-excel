using Xunit;
using static Sbroenne.ExcelMcp.Tests.Infrastructure.PackagingScriptTestHelper;

namespace Sbroenne.ExcelMcp.Packaging.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "Packaging")]
public sealed class PackageOrderingTests
{
    [Fact]
    public async Task NpmCommands_UseTheFirstNodeWhenSeveralInstallationsAreAvailable()
    {
        var result = await RunPowerShellCommandAsync($$"""
            . '{{Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1")}}'
            $global:primaryNode = (Get-Command node.exe -CommandType Application | Select-Object -First 1).Source
            $cli = Join-Path (Split-Path $global:primaryNode -Parent) 'node_modules\npm\bin\npm-cli.js'
            $expected = & $global:primaryNode $cli --version
            if ($LASTEXITCODE -ne 0) { throw 'The fixture npm installation failed.' }
            function Get-Command {
                [pscustomobject]@{ Source = $global:primaryNode }
                [pscustomobject]@{ Source = Join-Path $PSScriptRoot 'unused\node.exe' }
            }
            $actual = Invoke-PackageCommand $PSScriptRoot npm @('--version')
            if ($actual.Trim() -cne $expected.Trim()) { throw 'The first Node installation was not used.' }
            """);
        Assert.True(result.ExitCode == 0, result.CombinedOutput);
    }

    [Fact]
    public async Task NativeCommands_PreserveArgumentBoundariesAndExactOutput()
    {
        var result = await RunPowerShellCommandAsync($$"""
            . '{{Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1")}}'
            $arguments = @('has space', 'quote"inside', '', '{"value":"a b"}', 'trailing\', '%PATH%', '& | < >')
            $harness = Join-Path $PSScriptRoot 'arguments.ps1'
            '[Console]::Write((ConvertTo-Json -InputObject $args -Compress))' | Set-Content $harness
            $actual = Invoke-PackageCommand $PSScriptRoot 'pwsh' (@('-NoProfile', '-File', $harness) + $arguments)
            if ($actual -isnot [string]) { throw "Native output includes non-text values: $actual" }
            $expected = ConvertTo-Json -InputObject $arguments -Compress
            if ($actual -cne $expected) { throw "Expected $expected, got $actual" }
            """);
        Assert.True(result.ExitCode == 0, result.CombinedOutput);
    }

    [Fact]
    public async Task NativeCommands_EnforceDeadlineAndStopTheOwnedProcess()
    {
        var result = await RunPowerShellCommandAsync($$"""
            . '{{Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1")}}'
            $pidPath = Join-Path $PSScriptRoot 'child.pid'
            $harness = Join-Path $PSScriptRoot 'timeout.ps1'
            '[IO.File]::WriteAllText($args[0], "$PID"); Start-Sleep -Seconds 60' | Set-Content $harness
            try {
                Invoke-PackageCommand $PSScriptRoot 'pwsh' @('-NoProfile', '-File', $harness, $pidPath) -TimeoutSeconds 3
                throw 'Native deadline was not enforced.'
            } catch {
                if ($_ -notmatch 'exceeded its hard deadline') { throw }
            }
            $childId = [int][IO.File]::ReadAllText($pidPath)
            if (Get-Process -Id $childId -ErrorAction SilentlyContinue) { throw 'Timed out child is still running.' }
            """);
        Assert.True(result.ExitCode == 0, result.CombinedOutput);
    }

    [Theory]
    [InlineData("run compile")]
    [InlineData("run lint")]
    [InlineData("run typecheck:tests")]
    [InlineData("test")]
    public async Task ExtensionSourceFailure_StopsBeforeRuntimePublishing(string failedCommand)
    {
        var sandbox = CreateSandbox("package-order");
        try
        {
            var scripts = Directory.CreateDirectory(Path.Combine(sandbox, "scripts")).FullName;
            foreach (var name in new[] { "Build-ReleasePackages.ps1", "ExtensionPackaging.ps1" })
            {
                File.Copy(Path.Combine(RepoRoot, "scripts", name), Path.Combine(scripts, name));
            }
            var calls = Path.Combine(sandbox, "calls.txt");
            File.WriteAllText(Path.Combine(scripts, "PackageHelpers.ps1"),
                File.ReadAllText(Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1")) + $$"""

                function Invoke-PackageCommand {
                    param($Directory, $Executable, $Arguments)
                    [IO.File]::AppendAllText('{{calls}}', "$Executable $($Arguments -join ' ')`n")
                    if (($Arguments -join ' ') -eq '{{failedCommand}}') { throw 'source-check-failed' }
                    if ($Executable -ne 'npm') { throw 'Expensive runtime work was reached.' }
                }
                """);
            var extension = Directory.CreateDirectory(Path.Combine(sandbox, "vscode-extension")).FullName;
            File.WriteAllText(Path.Combine(extension, "package.json"), """{"version":"1.2.3","scripts":{"vscode:prepublish":"npm run compile"}}""");
            File.WriteAllText(Path.Combine(sandbox, "CHANGELOG.md"), "fixture");
            var skills = Path.Combine(sandbox, "prepared");
            var skill = Directory.CreateDirectory(Path.Combine(skills, "excel-mcp-report-formatting")).FullName;
            File.WriteAllText(Path.Combine(skill, "VERSION"), "1.2.3");
            File.WriteAllText(Path.Combine(skill, "SKILL.md"), "fixture");
            var output = Path.Combine(sandbox, "artifacts", "packages");
            var result = await RunPowerShellFileAsync(Path.Combine(scripts, "Build-ReleasePackages.ps1"),
                ["-Components", "Extension", "-Version", "1.2.3", "-SkillsDirectory", skills, "-OutputDirectory", output]);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("source-check-failed", result.CombinedOutput, StringComparison.Ordinal);
            var commands = File.ReadAllLines(calls);
            Assert.Equal("npm ci --ignore-scripts", commands[0]);
            Assert.Equal($"npm {failedCommand}", commands[^1]);
            Assert.DoesNotContain(commands, command => !command.StartsWith("npm ", StringComparison.Ordinal));
            Assert.False(Directory.Exists(Path.Combine(output, "runtimes")));
            Assert.False(Directory.Exists(Path.Combine(output, "nuget")));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }
}
