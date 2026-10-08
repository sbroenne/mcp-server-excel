<#
.SYNOPSIS
    Forcibly stops only CLI background services from this worktree.
.DESCRIPTION
    Development builds do not save workbooks or terminate Excel. Unsaved work
    in stopped services may be lost. An optional pipe limits test-run cleanup.
#>
[CmdletBinding()]
param([string]$PipeName)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$cliPaths = @(
    (Join-Path $root 'src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe'),
    (Join-Path $root 'src\ExcelMcp.CLI\bin\Debug\net10.0-windows\excelcli.exe')
)

foreach ($candidate in Get-CimInstance Win32_Process -Filter "Name = 'excelcli.exe'") {
    if ($cliPaths -notcontains $candidate.ExecutablePath -or
        $candidate.CommandLine -notmatch '^(?:"[^"]+"|\S+)\s+service\s+run(?:\s|$)' -or
        $null -eq $candidate.CreationDate) {
        continue
    }
    if (-not [string]::IsNullOrWhiteSpace($PipeName)) {
        $pipe = [regex]::Match($candidate.CommandLine,
            '(?:^|\s)--pipe-name\s+(?:"(?<pipe>[^"]+)"|(?<pipe>\S+))(?=\s|$)')
        if (-not $pipe.Success -or $pipe.Groups['pipe'].Value -ine $PipeName) {
            continue
        }
    }

    $current = Get-CimInstance Win32_Process -Filter "ProcessId = $($candidate.ProcessId)"
    if ($null -eq $current -or $current.CreationDate -ne $candidate.CreationDate -or
        $current.ExecutablePath -ne $candidate.ExecutablePath -or
        $current.CommandLine -ne $candidate.CommandLine) {
        continue
    }
    try {
        Stop-Process -Id $candidate.ProcessId -Force -ErrorAction Stop
    }
    catch {
        # A service can exit between identity validation and termination.
        if (Get-CimInstance Win32_Process -Filter "ProcessId = $($candidate.ProcessId)") {
            throw
        }
    }
    Write-Host "Stopped development CLI service PID $($candidate.ProcessId)."
}
$global:LASTEXITCODE = 0
