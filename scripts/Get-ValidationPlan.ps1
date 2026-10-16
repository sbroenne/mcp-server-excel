#Requires -Version 7.4
function Get-ValidationPlan {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][string[]]$Paths = @(),
        [switch]$Full,
        [switch]$Staged,
        [string]$BaseRef,
        [string]$HeadRef = 'HEAD',
        [ValidateSet('MergeBase', 'Direct')][string]$Comparison = 'MergeBase'
    )

    $root = if ($env:EXCELMCP_BUILD_ROOT) { $env:EXCELMCP_BUILD_ROOT } else { Split-Path -Parent $PSScriptRoot }
    $arguments = @('plan', '--root', $root)
    if ($Full) { $arguments += '--full' }
    else { $arguments += @('--paths-json', (ConvertTo-Json -InputObject @($Paths) -Compress)) }
    if ($Staged) { $arguments += '--staged' }
    if ($BaseRef) { $arguments += @('--base-ref', $BaseRef, '--head-ref', $HeadRef, '--comparison', $Comparison) }
    $json = & (Join-Path $root 'build.ps1') @arguments
    if ($LASTEXITCODE -ne 0) { throw 'Changed-area selection failed.' }
    $json -join [Environment]::NewLine | ConvertFrom-Json -AsHashtable
}
