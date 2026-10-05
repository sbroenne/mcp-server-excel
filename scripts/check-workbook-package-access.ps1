#!/usr/bin/env pwsh
param([string]$RootPath)

$ErrorActionPreference = 'Stop'
$root = if ($env:EXCELMCP_BUILD_ROOT) { $env:EXCELMCP_BUILD_ROOT } else { Split-Path -Parent $PSScriptRoot }
if ([string]::IsNullOrWhiteSpace($RootPath)) { $RootPath = Split-Path -Parent $PSScriptRoot }
& (Join-Path $root 'build.ps1') check-source --root $root --scan-root $RootPath --rule workbook-package-access
if ($LASTEXITCODE -ne 0) { throw 'Workbook package source guard failed.' }
