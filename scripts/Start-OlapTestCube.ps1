#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Starts the synthetic Atoti OLAP cube used by the external OLAP integration test.

.DESCRIPTION
    Creates a private Python environment with a pinned Atoti version (first run only),
    starts scripts\olap_test_cube.py on a free localhost port, waits until
    the cube serves XMLA, and returns the process plus the EXCELMCP_TEST_OLAP_* settings.
    The caller owns the returned process tree (Python plus a Java child) and must
    stop it through the retained process object with Kill($true) and WaitForExit().

    Atoti Community Edition is free for non-commercial use, sends usage telemetry, and
    needs internet access to activate its license. See
    https://docs.activeviam.com/engine/python-sdk/latest/eula.

.EXAMPLE
    $cube = & .\scripts\Start-OlapTestCube.ps1
    try { ... } finally { $cube.Process.Kill($true); $cube.Process.WaitForExit(); $cube.Process.Dispose() }
#>

[CmdletBinding()]
param(
    # Short path avoids Windows long-path failures while pip installs Atoti.
    [string]$ToolDirectory = (Join-Path $env:LOCALAPPDATA 'emo'),
    [ValidateRange(1, 1800)][int]$StartupTimeoutSeconds = 300
)

$ErrorActionPreference = 'Stop'
$atotiVersion = '6.2.2'
$cubeScript = Join-Path $PSScriptRoot 'olap_test_cube.py'
$python = Join-Path $ToolDirectory 'Scripts\python.exe'
$marker = Join-Path $ToolDirectory "atoti-$atotiVersion.installed"

if (-not (Test-Path $marker)) {
    Write-Host "Installing Atoti $atotiVersion into $ToolDirectory..." -ForegroundColor Cyan
    python -m venv $ToolDirectory
    if ($LASTEXITCODE -ne 0) { throw "Creating the OLAP test cube Python environment failed." }
    & $python -m pip install --quiet --disable-pip-version-check "atoti==$atotiVersion"
    if ($LASTEXITCODE -ne 0) { throw "Installing Atoti $atotiVersion failed." }
    New-Item -ItemType File -Path $marker -Force | Out-Null
}

$listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
$listener.Start()
$port = ([Net.IPEndPoint]$listener.LocalEndpoint).Port
$listener.Stop()

$logPath = Join-Path ([IO.Path]::GetTempPath()) "excelmcp-olap-cube-$port.log"
$process = Start-Process -FilePath $python -ArgumentList @('-u', "`"$cubeScript`"", $port) `
    -Environment @{ ATOTI_HIDE_EULA_MESSAGE = 'True' } `
    -RedirectStandardOutput $logPath -RedirectStandardError "$logPath.err" -NoNewWindow -PassThru
$null = $process.Handle
$ready = $false
$failures = [Collections.Generic.List[Exception]]::new()
$deadline = [DateTime]::UtcNow.AddSeconds($StartupTimeoutSeconds)
try {
    while (-not $ready) {
        if ($process.HasExited -or [DateTime]::UtcNow -gt $deadline) {
            $errorText = Get-Content -LiteralPath "$logPath.err" -Tail 20 | Out-String
            throw "The OLAP test cube did not start on port $port. Diagnostics: $logPath. $errorText"
        }
        $ready = (Test-Path -LiteralPath $logPath) -and
            (Select-String -LiteralPath $logPath -Pattern '^READY$' -Quiet)
        if (-not $ready) { Start-Sleep -Seconds 2 }
    }
}
catch { $failures.Add($_.Exception) }
finally {
    if (-not $ready) {
        try {
            if (-not $process.HasExited) {
                $process.Kill($true)
                if (-not $process.WaitForExit(10000)) {
                    throw "OLAP test cube cleanup timed out for PID $($process.Id)."
                }
            }
        }
        catch { $failures.Add($_.Exception) }
        finally { $process.Dispose() }
    }
}
if ($failures.Count) { throw [AggregateException]::new('OLAP test cube startup failed.', $failures) }

[pscustomobject]@{
    Process = $process
    Settings = @{
        EXCELMCP_TEST_OLAP_CONNECTION_STRING = "OLEDB;Provider=MSOLAP;Data Source=http://localhost:$port/xmla;Initial Catalog=atoti"
        EXCELMCP_TEST_OLAP_CUBE = 'SalesCube'
        EXCELMCP_TEST_OLAP_HIERARCHY = '[Sales].[Calendar]'
        EXCELMCP_TEST_OLAP_LEVEL = '[Sales].[Calendar].[Month]'
    }
}
