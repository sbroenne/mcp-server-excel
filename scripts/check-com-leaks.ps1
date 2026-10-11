#!/usr/bin/env pwsh
param([string[]]$InputPath)

$ErrorActionPreference = 'Stop'
$scanRoot = Split-Path -Parent $PSScriptRoot
$root = if ($env:EXCELMCP_BUILD_ROOT) { $env:EXCELMCP_BUILD_ROOT } else { $scanRoot }
$arguments = @('check-source', '--root', $root, '--scan-root', $scanRoot, '--rule', 'com-leaks')
if ($InputPath.Count -gt 0) {
    $arguments += @('--inputs-json', (ConvertTo-Json -InputObject @($InputPath) -Compress))
}
& (Join-Path $root 'build.ps1') @arguments
if ($LASTEXITCODE -ne 0) { throw 'COM source guard failed.' }
