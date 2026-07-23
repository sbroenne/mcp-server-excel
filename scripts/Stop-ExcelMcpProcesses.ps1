<#
.SYNOPSIS
    Stops the ExcelMCP Service gracefully and kills Excel processes before build.
.DESCRIPTION
    Pre-build cleanup script that:
    1. Gracefully stops the ExcelMCP Service via named pipe (service.shutdown)
    2. Kills any remaining Excel (EXCEL.EXE) processes

    This prevents file locking issues during build when the service or Excel
    holds handles to assemblies or workbooks.
.NOTES
    Called from Directory.Build.props as a BeforeBuild target.
    Safe to run when no processes are running (silently succeeds).
#>

param(
    [switch]$Verbose
)

$ErrorActionPreference = 'SilentlyContinue'

function Write-Status($message) {
    if ($Verbose) {
        Write-Host "  [pre-build] $message" -ForegroundColor DarkGray
    }
}

# ----------------------------------------------
# 1. Gracefully stop ExcelMCP Service via CLI
# ----------------------------------------------
function Stop-ExcelMcpService {
    # Look for excelcli in build output directories (Debug/Release)
    $scriptDir = Split-Path -Parent $PSScriptRoot  # repo root
    $cliPaths = @(
        "$scriptDir\src\ExcelMcp.CLI\bin\Debug\net10.0-windows\excelcli.exe",
        "$scriptDir\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe"
    )
    $excelcli = $cliPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($excelcli) {
        Write-Status "Using CLI: $excelcli"
        $output = & $excelcli service stop --quiet 2>&1
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq 0) {
            # Parse JSON to check if service was running
            try {
                $result = $output | ConvertFrom-Json
                if ($result.message -eq 'Service is not running.') {
                    Write-Status "ExcelMCP Service was not running"
                } else {
                    Write-Host "  ExcelMCP Service stopped gracefully" -ForegroundColor Green
                }
            } catch {
                Write-Status "Service stop completed (exit code 0)"
            }
        } else {
            Write-Status "CLI service stop returned exit code $exitCode, falling back to process kill"
            Stop-ExcelMcpServiceFallback
        }
    } else {
        Write-Status "excelcli not found (first build?), using fallback"
        Stop-ExcelMcpServiceFallback
    }
}

function Stop-ExcelMcpServiceFallback {
    # Fallback: direct named pipe shutdown (works without CLI binary)
    $sid = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
    $pipeName = "excelmcp-$sid"

    $pipeExists = Test-Path "\\.\pipe\$pipeName"
    if (-not $pipeExists) {
        Write-Status "ExcelMCP Service not running (no pipe found)"
        return
    }

    Write-Status "ExcelMCP Service detected, sending shutdown via pipe..."
    try {
        $pipe = New-Object System.IO.Pipes.NamedPipeClientStream('.', $pipeName, [System.IO.Pipes.PipeDirection]::InOut)
        $pipe.Connect(3000)

        $writer = New-Object System.IO.StreamWriter($pipe, [System.Text.Encoding]::UTF8, 4096)
        $writer.AutoFlush = $true
        $reader = New-Object System.IO.StreamReader($pipe, [System.Text.Encoding]::UTF8)

        $writer.WriteLine('{"Command":"service.shutdown"}')
        $response = $reader.ReadLine()
        Write-Status "Service response: $response"

        $reader.Dispose()
        $writer.Dispose()
        $pipe.Dispose()

        Start-Sleep -Milliseconds 500
        Write-Host "  ExcelMCP Service stopped gracefully" -ForegroundColor Green
    }
    catch {
        Write-Status "Could not connect to pipe: $($_.Exception.Message)"
        $serviceProcs = Get-Process -Name 'Sbroenne.ExcelMcp.McpServer', 'Sbroenne.ExcelMcp.Service' -ErrorAction SilentlyContinue
        if ($serviceProcs) {
            $serviceProcs | Stop-Process -Force -ErrorAction SilentlyContinue
            Write-Host "  ExcelMCP Service processes killed (pipe unavailable)" -ForegroundColor Yellow
        }
    }
}

# ----------------------------------------------
# 2. Kill Excel processes
# ----------------------------------------------
function Stop-ExcelProcesses {
    $excelProcs = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
    if ($excelProcs) {
        $count = $excelProcs.Count
        $excelProcs | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
        Write-Host "  Killed $count Excel process(es)" -ForegroundColor Yellow
    }
    else {
        Write-Status "No Excel processes running"
    }
}

# ----------------------------------------------
# Run cleanup
# ----------------------------------------------
Stop-ExcelMcpService
Stop-ExcelProcesses
