using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ChangedAreaRegressionTests
{
    [Fact]
    public async Task BuildBootstrap_UsesSdkFreshnessAndPrivateExecutionOutputs()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $root = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.Bootstrap.$([guid]::NewGuid().ToString('N'))"
            New-Item -ItemType Directory -Path "$root\tools\ExcelMcp.Build", "$root\temp" -Force | Out-Null
            try {
                Copy-Item .\build.ps1 "$root\build.ps1"
                foreach ($file in @('Directory.Build.props','Directory.Build.targets','Directory.Packages.props','global.json')) {
                    Set-Content "$root\$file" 'fixture'
                }
                Set-Content "$root\tools\ExcelMcp.Build\ExcelMcp.Build.csproj" 'fixture'
                $env:TEMP = $env:TMP = $env:TMPDIR = "$root\temp"
                $env:BOOTSTRAP_CALLS = "$root\calls.jsonl"
                Remove-Item Env:EXCELMCP_BUILD_DLL
                @'
            $ErrorActionPreference = 'Stop'
            ConvertTo-Json -InputObject @($args) -Compress | Add-Content $env:BOOTSTRAP_CALLS
            if ($args[0] -eq 'build') {
                $output = $args[[Array]::IndexOf($args, '-o') + 1]
                New-Item -ItemType Directory -Path $output -Force | Out-Null
                Set-Content (Join-Path $output 'ExcelMcp.Build.dll') 'fixture'
            }
            elseif (-not (Test-Path -LiteralPath $args[0])) { throw 'The tool was not built before execution.' }
            $global:LASTEXITCODE = 0
            '@ | Set-Content "$root\dotnet.ps1"
                $global:bootstrapDotnet = "$root\dotnet.ps1"
                function Get-Command { [pscustomobject]@{ Source = $global:bootstrapDotnet } }
                & "$root\build.ps1" --help
                & "$root\build.ps1" --help
                $calls = @(Get-Content $env:BOOTSTRAP_CALLS | ForEach-Object { ,(ConvertFrom-Json $_) })
                $builds = @($calls | Where-Object { $_[0] -eq 'build' })
                if ($builds.Count -ne 2) { throw 'Every invocation must let MSBuild decide whether compilation is needed.' }
                $runs = @($calls | Where-Object { $_[0] -ne 'build' })
                if ($runs.Count -ne 2 -or $runs[0][0] -eq $runs[1][0]) { throw 'Executing tools must use distinct outputs.' }
                foreach ($call in $builds) {
                    if ($call -contains '--disable-build-servers' -or $call -contains '--no-incremental') {
                        throw 'Native SDK optimizations were disabled.'
                    }
                }
                foreach ($call in $runs) {
                    if (Test-Path -LiteralPath $call[0]) { throw 'Private execution output was not cleaned up.' }
                }
            }
            finally { Remove-Item -LiteralPath $root -Recurse -Force }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task BuildBootstrap_MultipleDotnetApplications_UsesFirstExecutable()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $global:expectedDotnetExecutable = (Microsoft.PowerShell.Core\Get-Command dotnet -CommandType Application |
                Select-Object -First 1).Source
            function Get-Command {
                param([string]$Name, [string]$CommandType)
                if ($Name -ne 'dotnet' -or $CommandType -ne 'Application') { throw 'Unexpected command lookup.' }
                [pscustomobject]@{ Source = $global:expectedDotnetExecutable }
                [pscustomobject]@{ Source = "$global:expectedDotnetExecutable-secondary" }
            }
            $help = & .\build.ps1 --help
            if ($LASTEXITCODE -ne 0 -or ($help -join "`n") -notmatch 'ExcelMcp development tooling') {
                throw 'The first dotnet executable did not run the tool.'
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BuildCommandFailure_PreservesDiagnosticsWhenCleanupAlsoFails(bool cleanupFails)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            $PSNativeCommandUseErrorActionPreference = $true
            if ({{(cleanupFails ? "$true" : "$false")}}) {
                function Remove-Item {
                    param([string]$LiteralPath, [switch]$Force, [switch]$Recurse)
                    Microsoft.PowerShell.Management\Remove-Item -LiteralPath $LiteralPath -Force -Recurse -ErrorAction Stop
                    throw 'cleanup-root-cause'
                }
            }
            try {
                & .\build.ps1 invalid-target
                throw 'The invalid build target unexpectedly succeeded.'
            }
            catch {
                if ($_.Exception.Message -notmatch "Unknown build target 'invalid-target'") {
                    throw "The build diagnostic was lost: $($_.Exception.Message)"
                }
                if ({{(cleanupFails ? "$true" : "$false")}} -and $_.Exception.Message -notmatch 'cleanup-root-cause') {
                    throw 'The cleanup failure was lost.'
                }
                $global:LASTEXITCODE = 0
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Theory]
    [InlineData("Range", "RangeCommands.Values.cs", "Core")]
    [InlineData("PowerQuery", "PowerQueryCommands.Read.cs", "Core,Service")]
    public async Task ImplementationChange_SelectsItsFeatureWithoutUnrelatedAdapters(string area, string file, string expectedProjects)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths 'src/ExcelMcp.Core/Commands/{{area}}/{{file}}'
            if (($plan.FastProjects -join ',') -ne '{{expectedProjects}}') { throw 'Unrelated runtime projects selected.' }
            if ($plan.ProcessProjects.Count) { throw 'Unchanged CLI process behavior selected.' }
            if (-not $plan.FastFilters.Core.Contains('{{area}}')) { throw 'Owning Core feature filter missing.' }
            if (-not @($plan.ExcelSelections | Where-Object { $_.Project -eq 'Service' -and $_.Area -eq '{{area}}' }).Count) {
                throw 'Owning Service workbook cases missing.'
            }
            if ($plan.FullE2E) { throw 'Unchanged full acceptance workflow selected.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task TestOnlyChange_SelectsItsClassRatherThanItsWholeProject()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $plan = Get-ValidationPlan -Paths 'tests/ExcelMcp.Core.Tests/Unit/GeneratedActionContractTests.cs'
            if (($plan.FastProjects -join ',') -ne 'Core') { throw 'Unrelated project selected.' }
            if ($plan.FastFilters.Core -ne 'FullyQualifiedName~Sbroenne.ExcelMcp.Core.Tests.Unit.GeneratedActionContractTests.') {
                throw 'Test-only selection was broadened beyond its class.'
            }
            if ($plan.Excel -or $plan.ProcessProjects.Count) { throw 'Unrelated integration checks selected.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task UnknownInput_FailsInsteadOfSelectingEverything()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            Get-ValidationPlan -Paths 'unknown-build-input.config'
            """);
        Assert.NotEqual(0, run.ExitCode);
        Assert.Contains("No validation mapping", run.Output, StringComparison.Ordinal);
    }
}
