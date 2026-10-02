#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Runs the real-executable CLI workflow acceptance scenarios.
.DESCRIPTION
    Lifecycle/persistence, safe editing, typed formatting, and native API coverage
    are independent tests with their own prerequisites and cleanup. Reports are
    retained in TestResults. This is not the complete Test-E2E acceptance gate.
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
