#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Runs all Excel-dependent end-to-end gates required before merge.

.DESCRIPTION
    Builds the Release solution unless -SkipBuild is supplied, then runs:
    1. The CLI workflow smoke test.
    2. The stale-build graceful-save acceptance test.
    3. The MCP all-tools end-to-end smoke test.

    The script fails if either gate fails or if the MCP filter matches no tests.

.EXAMPLE
    & .\scripts\Test-E2E.ps1
#>

[CmdletBinding()]
param(
    [switch]$SkipBuild,
    [string]$PipeName
)

$ErrorActionPreference = 'Stop'
$rootDir = Split-Path -Parent $PSScriptRoot
$cliTestProject = Join-Path $rootDir 'tests\ExcelMcp.CLI.Tests\ExcelMcp.CLI.Tests.csproj'
$mcpTestProject = Join-Path $rootDir 'tests\ExcelMcp.McpServer.Tests\ExcelMcp.McpServer.Tests.csproj'
$staleCleanupAcceptanceFilter = 'FullyQualifiedName~PreBuildGracefulSaveAcceptanceTests.StaleLockedBuildCleanup'
$smokeTestFilter = 'FullyQualifiedName~McpServerSmokeTests.SmokeTest_AllTools_E2EWorkflow'
$previousPipeName = $env:EXCELMCP_CLI_PIPE
$selectedPipeName = if ([string]::IsNullOrWhiteSpace($PipeName)) {
    "excelmcp-e2e-$PID-$([Guid]::NewGuid().ToString('N'))"
}
else {
    $PipeName
}
$env:EXCELMCP_CLI_PIPE = $selectedPipeName

Push-Location $rootDir
try {
    Write-Host "Using private CLI pipe: $selectedPipeName" -ForegroundColor DarkGray

    if (-not $SkipBuild) {
        Write-Host 'Building Release solution...' -ForegroundColor Cyan
        dotnet build Sbroenne.ExcelMcp.sln --configuration Release -p:NuGetAudit=false --verbosity minimal
        if ($LASTEXITCODE -ne 0) {
            throw "Release build failed with exit code $LASTEXITCODE."
        }
    }

    Write-Host ''
    Write-Host 'Running CLI workflow E2E test...' -ForegroundColor Cyan
    & (Join-Path $PSScriptRoot 'Test-CliWorkflow.ps1') -PipeName $selectedPipeName
    if ($LASTEXITCODE -ne 0) {
        throw "CLI workflow E2E test failed with exit code $LASTEXITCODE."
    }

    Write-Host ''
    Write-Host 'Running stale-build graceful-save acceptance test...' -ForegroundColor Cyan
    $staleCleanupOutput = dotnet test $cliTestProject `
        --configuration Release `
        --no-build `
        --filter $staleCleanupAcceptanceFilter `
        --verbosity minimal `
        --blame-hang-timeout 6m `
        -- RunConfiguration.MaxCpuCount=1 2>&1 | Out-String
    $staleCleanupExitCode = $LASTEXITCODE

    Write-Host $staleCleanupOutput

    if ($staleCleanupOutput -notmatch 'Passed!.*Passed:\s*[1-9]') {
        throw "No stale-build graceful-save acceptance test passed. Verify the filter still matches $staleCleanupAcceptanceFilter."
    }

    if ($staleCleanupExitCode -ne 0) {
        throw "Stale-build graceful-save acceptance test failed with exit code $staleCleanupExitCode."
    }

    Write-Host ''
    Write-Host 'Running MCP all-tools E2E test...' -ForegroundColor Cyan
    & (Join-Path $PSScriptRoot 'Stop-ExcelMcpProcesses.ps1') -PipeName $selectedPipeName
    if ($LASTEXITCODE -ne 0) {
        throw "Owned CLI cleanup failed with exit code $LASTEXITCODE before the MCP E2E test."
    }

    $testOutput = dotnet test $mcpTestProject `
        --configuration Release `
        --no-build `
        --filter $smokeTestFilter `
        --verbosity minimal `
        --blame-hang-timeout 15m `
        -- RunConfiguration.MaxCpuCount=1 2>&1 | Out-String
    $testExitCode = $LASTEXITCODE

    Write-Host $testOutput

    if ($testOutput -notmatch 'Passed!.*Passed:\s*[1-9]') {
        throw "No MCP E2E tests passed. Verify the filter still matches $smokeTestFilter."
    }

    if ($testExitCode -ne 0) {
        throw "MCP all-tools E2E test failed with exit code $testExitCode."
    }

    Write-Host 'All Excel-dependent E2E tests passed.' -ForegroundColor Green
}
finally {
    $cleanupExitCode = 0
    try {
        & (Join-Path $PSScriptRoot 'Stop-ExcelMcpProcesses.ps1') -PipeName $selectedPipeName
        $cleanupExitCode = $LASTEXITCODE
    }
    finally {
        if ($null -eq $previousPipeName) {
            Remove-Item Env:EXCELMCP_CLI_PIPE -ErrorAction SilentlyContinue
        }
        else {
            $env:EXCELMCP_CLI_PIPE = $previousPipeName
        }
        Pop-Location
    }

}

if ($cleanupExitCode -ne 0) {
    throw "Owned CLI cleanup failed with exit code $cleanupExitCode after E2E validation."
}

$global:LASTEXITCODE = 0
