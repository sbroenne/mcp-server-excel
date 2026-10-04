[CmdletBinding()]
param(
    [string]$BaseRef,
    [string]$HeadRef = 'HEAD',
    [ValidateSet('MergeBase', 'Direct')][string]$Comparison = 'MergeBase',
    [switch]$Full,
    [Parameter(Mandatory)][string]$OutputPath
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Get-ValidationPlan.ps1')
if ($Full -or -not $BaseRef -or $BaseRef -match '^0+$') {
    $plan = Get-ValidationPlan -Full
} else {
    $range = if ($Comparison -eq 'MergeBase') { "$BaseRef...$HeadRef" } else { "$BaseRef..$HeadRef" }
    $paths = @(git -c core.quotepath=false diff --name-only --no-renames $range)
    if ($LASTEXITCODE -ne 0) { throw 'Cannot determine changed validation inputs.' }
    $plan = Get-ValidationPlan -Paths $paths
}
$plan.Reasons | ForEach-Object { Write-Host $_ }
$directory = Split-Path -Parent $OutputPath
if ($directory) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
$plan | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $OutputPath -Encoding utf8
if ($env:GITHUB_OUTPUT) {
    $include = @($plan.CiTestGroups | ForEach-Object { @{ group = $_ } })
    $matrix = @{ include = $include } | ConvertTo-Json -Compress -Depth 5
    $languages = @($plan.CodeQlLanguages | ForEach-Object {
        @{
            language = $_
            os = if ($_ -eq 'csharp') { 'windows-latest' } else { 'ubuntu-26.04' }
            'build-mode' = if ($_ -eq 'csharp') { 'manual' } else { 'none' }
        }
    })
    $codeQlMatrix = @{ include = $languages } | ConvertTo-Json -Compress -Depth 5
    @(
        "matrix=$matrix"
        "source_checks_group=$($plan.SourceChecksGroup)"
        "tests=$($plan.CiTestGroups.Count -gt 0)".ToLowerInvariant()
        "packages=$($plan.Packages)".ToLowerInvariant()
        "package_build=$($plan.PackageBuild)".ToLowerInvariant()
        "codeql_matrix=$codeQlMatrix"
        "codeql=$($plan.CodeQlLanguages.Count -gt 0)".ToLowerInvariant()
        "npm=$($plan.NpmTests)".ToLowerInvariant()
        "lockfiles=$($plan.LockfileTests)".ToLowerInvariant()
    ) | Add-Content -LiteralPath $env:GITHUB_OUTPUT -Encoding utf8
}
$global:LASTEXITCODE = 0
