#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Validates staged source and local Excel behavior. Never prepares release artifacts for publication.
#>
$ErrorActionPreference = 'Stop'
$rootDir = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Get-ValidationPlan.ps1')

function Invoke-Check {
    param([string]$Name, [scriptblock]$Action)
    Write-Host $Name -ForegroundColor Cyan
    $global:LASTEXITCODE = 0
    & $Action
    if ($LASTEXITCODE -ne 0) { throw "$Name failed with exit code $LASTEXITCODE." }
}

function Read-Git {
    param([string[]]$Arguments)
    $output = & git @Arguments
    if ($LASTEXITCODE -ne 0) { throw "Cannot inspect Git state: git $($Arguments -join ' ')." }
    return $output
}

Push-Location $rootDir
try {
    $branch = Read-Git @('branch', '--show-current')
    if ($branch -eq 'main') { throw "Cannot commit directly to main. Use a feature branch and a pull request." }
    $mergeHead = & git rev-parse --verify --quiet MERGE_HEAD 2>$null
    $base = if ($LASTEXITCODE -eq 0 -and $mergeHead) { $mergeHead } else { 'HEAD' }
    $paths = @(Read-Git @('-c', 'core.quotepath=false', 'diff', '--cached', '--name-only', '--no-renames', $base))
    $plan = Get-ValidationPlan -Paths $paths
    $plan.Reasons | ForEach-Object { Write-Host "  $_" }
    $supportsDesktopExcelValidation = [OperatingSystem]::IsWindows() -or (
        [OperatingSystem]::IsMacOS() -and
        [Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture -eq
            [Runtime.InteropServices.Architecture]::Arm64)

    Invoke-Check 'Checking staged npm lockfiles' {
        & (Join-Path $PSScriptRoot 'check-npm-lockfiles.ps1') -Staged
    }
    if ($plan.Build -or $plan.SourceChecks) {
        $workingPaths = @(
            Read-Git @('-c', 'core.quotepath=false', 'diff', '--name-only', '--no-renames')
            Read-Git @('-c', 'core.quotepath=false', 'ls-files', '--others', '--exclude-standard')
        )
        $inputs = @($workingPaths | Where-Object {
            $_ -match '^(src[\\/]|tests[\\/]|skills[\\/]|scripts[\\/]|docs[\\/]reference[\\/]report-formatting\.md$|Directory\.|\.editorconfig$|global\.json$|NuGet\.Config$|Sbroenne\.ExcelMcp\.sln$)'
        })
        if ($inputs.Count -gt 0) {
            throw "Validation inputs differ from the index. Stage or set aside these changes explicitly: $($inputs -join ', '). No files were staged or stashed."
        }
    }
    if ($plan.Excel -and -not $supportsDesktopExcelValidation) {
        throw 'Desktop Excel validation requires Windows or Apple Silicon macOS; required E2E was not run.'
    }
    if ($plan.SourceChecks) {
        foreach ($script in @('check-com-leaks', 'check-success-flag', 'check-dynamic-casts', 'check-workbook-package-access')) {
            Invoke-Check $script { & (Join-Path $PSScriptRoot "$script.ps1") }
        }
    }
    if ($plan.Build) {
        Invoke-Check 'Building Release solution' {
            [string[]]$platformProperties = if ([OperatingSystem]::IsWindows()) { @() } else { @('-p:EnableWindowsTargeting=true') }
            dotnet build Sbroenne.ExcelMcp.sln -c Release -p:NuGetAudit=false @platformProperties --verbosity minimal
        }
        Invoke-Check 'Running focused Excel-free tests' {
            & (Join-Path $PSScriptRoot 'Invoke-ExcelFreeTests.ps1') -Local -HookTests:$plan.HookTests -Contracts:$plan.Excel -SkillTests:$plan.SkillTests -PackagingTests:$plan.PackagingTests -ChangedPaths $paths
        }
    }
    if ($plan.Excel) {
        Write-Host 'Complete local Excel E2E is required once against the final PR source.' -ForegroundColor Yellow
    }
    Write-Host 'All selected pre-commit checks passed. Release artifact validation belongs to PR CI.' -ForegroundColor Green
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
finally { Pop-Location }
exit 0
