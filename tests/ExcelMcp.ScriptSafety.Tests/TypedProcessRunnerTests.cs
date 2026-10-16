using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TypedProcessRunnerTests
{
    [Fact]
    public async Task NativeFailure_PreservesExitCodeAndOriginalDiagnostic()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            . .\scripts\Invoke-NativeCommand.ps1
            Invoke-NativeCommand -Executable pwsh -Arguments @('-NoProfile','-Command',"[Console]::Error.WriteLine('native-root-cause'); exit 23")
            """);
        Assert.NotEqual(0, run.ExitCode);
        Assert.Contains("exit code 23", run.Output, StringComparison.Ordinal);
        Assert.Contains("native-root-cause", run.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task NativeArguments_AreNotCombinedIntoShellCode()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            . .\scripts\Invoke-NativeCommand.ps1
            $value = 'spaces; ''quotes'' & $variables "double quotes" trailing\'
            $file = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.Arguments.$([Guid]::NewGuid().ToString('N')).ps1"
            try {
                'param([string]$Value); [Console]::Write($Value)' | Set-Content -LiteralPath $file
                $result = Invoke-NativeCommand -Executable pwsh -Arguments @('-NoProfile','-File',$file,$value)
                if ($result.Output -cne $value) { throw 'Argument boundaries changed.' }
            } finally { Remove-Item -LiteralPath $file }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task NativeChildren_DoNotInheritGitIndexOrLoseOtherEnvironment()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            . .\scripts\Invoke-NativeCommand.ps1
            $env:GIT_INDEX_FILE = 'not-the-child-index'
            $result = Invoke-NativeCommand -Executable pwsh -Arguments @('-NoProfile','-Command',
                'if ($env:GIT_INDEX_FILE) { throw "Ambient index inherited" }; if ($env:EXCELMCP_FIXTURE_MARKER -cne "retained") { throw "Marker lost" }') `
                -Environment @{ EXCELMCP_FIXTURE_MARKER='retained' }
            if ($result.ExitCode) { throw 'Native child failed.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }
}
