#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Runs the real-executable CLI workflow acceptance scenarios.
.DESCRIPTION
    Delegates to the CLI stage of Test-E2E.ps1 and preserves failure reporting.
#>
[CmdletBinding()]
param(
    [switch]$KeepFile,
    [string]$PipeName,
    [string]$ResultsDirectory
)
$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'Test-E2E.ps1') -SkipBuild -Stages Cli `
    -PipeName $PipeName -ResultsDirectory $ResultsDirectory -KeepCliFiles:$KeepFile
if ($LASTEXITCODE -ne 0) { throw 'CLI workflow acceptance failed.' }
$global:LASTEXITCODE = 0
