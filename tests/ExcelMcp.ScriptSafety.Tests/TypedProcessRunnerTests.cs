using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TypedProcessRunnerTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(23)]
    public async Task NativeOutput_IsReadableBeforeCommandExitsAndPreservesFinalResult(int exitCode)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"ExcelMcp.LiveOutput.{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var outputPath = Path.Combine(directory, "stdout.log");
        var errorPath = Path.Combine(directory, "stderr.log");
        var releasePath = Path.Combine(directory, "continue");
        var run = ValidationSelectionTests.RunAsync("""
            . .\scripts\Invoke-NativeCommand.ps1
            $child = @'
            [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
            [Console]::Write('before-' + [char]0x03A9)
            [Console]::Out.Flush()
            [Console]::Error.Write('error-before')
            [Console]::Error.Flush()
            while (-not (Test-Path -LiteralPath $env:EXCELMCP_PROGRESS_RELEASE)) { Start-Sleep -Milliseconds 50 }
            [Console]::Write('-after')
            [Console]::Error.Write('-after')
            exit ([int]$env:EXCELMCP_PROGRESS_EXIT)
            '@
            $result = Invoke-NativeCommand -Executable pwsh -Arguments @('-NoProfile','-Command',$child) `
                -OutputPath $env:EXCELMCP_PROGRESS_OUTPUT -ErrorPath $env:EXCELMCP_PROGRESS_ERROR `
                -ProgressName 'live-output-fixture' -AllowFailure
            if ($result.ExitCode -ne [int]$env:EXCELMCP_PROGRESS_EXIT) { throw 'Exit code changed.' }
            if ($result.Output -cne ('before-' + [char]0x03A9 + '-after')) { throw 'Captured output changed.' }
            if ($result.Error -cne 'error-before-after') { throw 'Captured error changed.' }
            """, new Dictionary<string, string>
        {
            ["EXCELMCP_PROGRESS_OUTPUT"] = outputPath,
            ["EXCELMCP_PROGRESS_ERROR"] = errorPath,
            ["EXCELMCP_PROGRESS_RELEASE"] = releasePath,
            ["EXCELMCP_PROGRESS_EXIT"] = exitCode.ToString(System.Globalization.CultureInfo.InvariantCulture)
        });
        try
        {
            var deadline = System.Diagnostics.Stopwatch.StartNew();
            while (!run.IsCompleted && deadline.Elapsed < TimeSpan.FromSeconds(30) &&
                (ReadLiveLog(outputPath) != "before-\u03A9" || ReadLiveLog(errorPath) != "error-before"))
            {
                await Task.Delay(50);
            }
            Assert.False(run.IsCompleted, run.IsCompleted ? (await run).Output : null);
            Assert.Equal("before-\u03A9", ReadLiveLog(outputPath));
            Assert.Equal("error-before", ReadLiveLog(errorPath));
            await File.WriteAllTextAsync(releasePath, "continue");
            var result = await run;
            Assert.True(result.ExitCode == 0, result.Output);
            Assert.Contains("live-output-fixture", result.Output, StringComparison.Ordinal);
            Assert.Equal("before-\u03A9-after", await File.ReadAllTextAsync(outputPath));
            Assert.Equal("error-before-after", await File.ReadAllTextAsync(errorPath));
        }
        finally
        {
            await File.WriteAllTextAsync(releasePath, "continue");
            await run;
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task NativeTimeout_PreservesLiveLogs()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"ExcelMcp.TimeoutOutput.{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var outputPath = Path.Combine(directory, "stdout.log");
        var errorPath = Path.Combine(directory, "stderr.log");
        try
        {
            var run = await ValidationSelectionTests.RunAsync("""
                . .\scripts\Invoke-NativeCommand.ps1
                Invoke-NativeCommand -Executable pwsh -Arguments @('-NoProfile','-Command',
                    "[Console]::Write('output-before-timeout'); [Console]::Out.Flush(); [Console]::Error.Write('error-before-timeout'); [Console]::Error.Flush(); Start-Sleep -Seconds 30") `
                    -DeadlineSeconds 6 -OutputPath $env:EXCELMCP_PROGRESS_OUTPUT -ErrorPath $env:EXCELMCP_PROGRESS_ERROR
                """, new Dictionary<string, string>
            {
                ["EXCELMCP_PROGRESS_OUTPUT"] = outputPath,
                ["EXCELMCP_PROGRESS_ERROR"] = errorPath
            });
            Assert.NotEqual(0, run.ExitCode);
            Assert.Contains("hard deadline", run.Output, StringComparison.Ordinal);
            Assert.Equal("output-before-timeout", await File.ReadAllTextAsync(outputPath));
            Assert.Equal("error-before-timeout", await File.ReadAllTextAsync(errorPath));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static string ReadLiveLog(string path)
    {
        if (!File.Exists(path)) { return string.Empty; }
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

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
