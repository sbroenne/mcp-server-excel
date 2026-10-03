<#
.SYNOPSIS
Runs one complete GitHub job in the limited desktop, only after a current-boot grant.
.DESCRIPTION
The interactive task pattern is adapted from mcp-windows runner setup
(MIT, Copyright (c) 2025 Sbroenne). There is deliberately no logon trigger:
restarts and maintenance must never start a listener before qualification.
#>
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ExcelRunnerPolicy.ps1')
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
try {
    if ($identity.Name -ine "$env:COMPUTERNAME\excelrunner" -or
        ([Security.Principal.WindowsPrincipal]::new($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -or
        -not [Environment]::UserInteractive -or (Get-Process -Id $PID).SessionId -le 0) {
        throw 'The runner requires the dedicated non-admin interactive desktop.'
    }
}
finally { $identity.Dispose() }
$permit = Get-Content (Join-Path $PSScriptRoot 'permit.json') -Raw | ConvertFrom-Json
$now = [DateTime]::UtcNow
$boot = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime()
if ($permit.state -ne 'admitted' -or (ConvertTo-ExcelRunnerUtc $permit.bootTime) -ne $boot -or
    (ConvertTo-ExcelRunnerUtc $permit.expiresAt) -lt $now -or $permit.runId -notmatch '^\d+$') {
    throw 'The listener has no unexpired current-boot hosted control grant.'
}
Assert-ExcelRunnerPatchState (Get-Content 'C:\ProgramData\ExcelMcp\Maintenance\patched.json' -Raw | ConvertFrom-Json)
Assert-ExcelRunnerReadiness (Get-Content (Join-Path $PSScriptRoot 'ready.json') -Raw | ConvertFrom-Json) -BootTime $boot
if (@(Get-Process -Name Runner.Listener, Runner.Worker, EXCEL -ErrorAction SilentlyContinue).Count) {
    throw 'Another listener, job or workbook is already active.'
}
$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
    [Environment]::GetEnvironmentVariable('Path', 'User')
Set-Location 'C:\actions-runner'
& 'C:\actions-runner\run.cmd' --once
if ($LASTEXITCODE -ne 0) { throw "The one-job runner exited with code $LASTEXITCODE." }
