#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Runs independently reported Excel-dependent acceptance stages before merge.

.DESCRIPTION
    Builds the Release solution unless -SkipBuild is supplied, then runs:
    1. Independent CLI workflow scenarios.
    2. Independent MCP workflow scenarios.
    3. External OLAP schema discovery through the Service boundary.

    Defaults to all stages. The OLAP stage uses EXCELMCP_TEST_OLAP_* when set;
    otherwise it starts the synthetic Atoti cube via Start-OlapTestCube.ps1
    (requires Python) and stops it afterwards.
    A focused -Stages run is not complete acceptance. The script fails if any
    gate fails or if a required filter matches no tests.

.EXAMPLE
    & .\scripts\Test-E2E.ps1
#>

[CmdletBinding()]
param(
    [switch]$SkipBuild,
    [string]$PipeName,
    [ValidateNotNullOrEmpty()]
    [ValidateSet('Cli', 'Mcp', 'Olap')][string[]]$Stages = @('Cli', 'Mcp', 'Olap'),
    [string]$ResultsDirectory,
    [switch]$KeepCliFiles
)

$ErrorActionPreference = 'Stop'
$rootDir = Split-Path -Parent $PSScriptRoot
$cliTestProject = Join-Path $rootDir 'tests\ExcelMcp.CLI.Tests\ExcelMcp.CLI.Tests.csproj'
$mcpTestProject = Join-Path $rootDir 'tests\ExcelMcp.McpServer.Tests\ExcelMcp.McpServer.Tests.csproj'
$serviceTestProject = Join-Path $rootDir 'tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj'
. (Join-Path $PSScriptRoot 'Invoke-TestStage.ps1')
if (-not $ResultsDirectory) { $ResultsDirectory = Join-Path $rootDir "TestResults\e2e-$([Guid]::NewGuid().ToString('N'))" }
$previousPipeName = $env:EXCELMCP_CLI_PIPE
$selectedPipeName = if ([string]::IsNullOrWhiteSpace($PipeName)) {
    "excelmcp-e2e-$PID-$([Guid]::NewGuid().ToString('N'))"
}
else {
    $PipeName
}
$env:EXCELMCP_CLI_PIPE = $selectedPipeName
$failures = [Collections.Generic.List[Exception]]::new()
$olapCube = $null

Push-Location $rootDir
try {
    Write-Host "Using private CLI pipe: $selectedPipeName" -ForegroundColor DarkGray

    if (-not $SkipBuild) {
        Write-Host 'Building Release solution...' -ForegroundColor Cyan
        dotnet build Sbroenne.ExcelMcp.sln --configuration Release --disable-build-servers -p:NuGetAudit=false --verbosity minimal
        if ($LASTEXITCODE -ne 0) {
            throw "Release build failed with exit code $LASTEXITCODE."
        }
    }

    foreach ($stage in $Stages | Select-Object -Unique) {
        $parameters = @{
            ResultsDirectory = $ResultsDirectory
            Name = $stage
            Environment = @{
                EXCELMCP_CLI_PIPE = $selectedPipeName
                EXCELMCP_CLI_WORKFLOW_KEEP_FILE = $KeepCliFiles.ToString().ToLowerInvariant()
            }
        }
        switch ($stage) {
            'Cli' {
                $parameters.Project = $cliTestProject
                $parameters.Filter = 'RequiresExcel=true&FullyQualifiedName~CliWorkflowAcceptanceTests'
                $parameters.DeadlineSeconds = 600
                $parameters.HangTimeout = '5m'
            }
            'Mcp' {
                $parameters.Project = $mcpTestProject
                $parameters.Filter = 'RequiresExcel=true&Acceptance=Required&FullyQualifiedName~McpServerSmokeTests'
                $parameters.DeadlineSeconds = 900
                $parameters.HangTimeout = '15m'
            }
            'Olap' {
                if ([string]::IsNullOrWhiteSpace($env:EXCELMCP_TEST_OLAP_CONNECTION_STRING)) {
                    $olapCube = & (Join-Path $PSScriptRoot 'Start-OlapTestCube.ps1')
                    foreach ($key in $olapCube.Settings.Keys) { $parameters.Environment[$key] = $olapCube.Settings[$key] }
                }
                $parameters.Project = $serviceTestProject
                $parameters.Filter = 'RequiresExcel=true&FullyQualifiedName~ExternalOlapSchema_UsesSelectedCubeAndContinuesThroughService'
                $parameters.DeadlineSeconds = 600
                $parameters.HangTimeout = '5m'
            }
        }
        & (Join-Path $PSScriptRoot 'Stop-ExcelCliService.ps1') -PipeName $selectedPipeName
        if ($LASTEXITCODE -ne 0) { throw "Owned CLI cleanup failed before $stage acceptance." }
        Invoke-TestStage @parameters -ReconcileCases
    }
    Write-Host "Selected acceptance stages passed: $($Stages -join ', ')."
}
catch {
    $failures.Add($_.Exception)
}
finally {
    try {
        & (Join-Path $PSScriptRoot 'Stop-ExcelCliService.ps1') -PipeName $selectedPipeName
        if ($LASTEXITCODE -ne 0) { throw 'Owned CLI cleanup failed after E2E validation.' }
    }
    catch { $failures.Add($_.Exception) }
    finally {
        if ($olapCube -and -not $olapCube.Process.HasExited) {
            # Atoti runs Python and Java child processes; stop the whole owned tree by PID.
            taskkill.exe /PID $olapCube.Process.Id /T /F | Out-Null
        }
        if ($null -eq $previousPipeName) {
            Remove-Item Env:EXCELMCP_CLI_PIPE -ErrorAction SilentlyContinue
        }
        else {
            $env:EXCELMCP_CLI_PIPE = $previousPipeName
        }
        Pop-Location
    }

}

if ($failures.Count) { throw [AggregateException]::new('E2E validation failed.', $failures) }

$global:LASTEXITCODE = 0
