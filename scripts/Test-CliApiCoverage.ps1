#!/usr/bin/env pwsh
[CmdletBinding()]
param([switch]$KeepFile, [string]$PipeName)

$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'Test-E2E.ps1') -SkipBuild -Stages Cli -PipeName $PipeName -KeepCliFiles:$KeepFile
if ($LASTEXITCODE -ne 0) { throw "Native CLI acceptance failed with exit code $LASTEXITCODE." }
