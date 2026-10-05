#!/usr/bin/env pwsh
param([switch]$Verbose)

$ErrorActionPreference = 'Stop'
$scanRoot = Split-Path -Parent $PSScriptRoot
$root = if ($env:EXCELMCP_BUILD_ROOT) { $env:EXCELMCP_BUILD_ROOT } else { $scanRoot }
& (Join-Path $root 'build.ps1') check-source --root $root --scan-root $scanRoot --rule dynamic-casts
if ($LASTEXITCODE -ne 0) { throw 'Dynamic cast source guard failed.' }
