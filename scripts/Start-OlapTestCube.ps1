#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Starts the synthetic Atoti OLAP cube used by the external OLAP integration test.

.DESCRIPTION
    Creates a private Python environment with a pinned Atoti version (first run only),
    starts scripts\olap_test_cube.py on a free localhost port, waits until
    the cube serves XMLA, and returns the process plus the EXCELMCP_TEST_OLAP_* settings.
    The caller owns the returned process tree (Python plus a Java child) and must
    stop it, for example with taskkill /PID <id> /T /F.

    Atoti Community Edition is free for non-commercial use, sends usage telemetry, and
    needs internet access to activate its license. See
    https://docs.activeviam.com/engine/python-sdk/latest/eula.

.EXAMPLE
    $cube = & .\scripts\Start-OlapTestCube.ps1
    try { ... } finally { taskkill.exe /PID $cube.Process.Id /T /F }
#>

[CmdletBinding()]
param(
    # Short path avoids Windows long-path failures while pip installs Atoti.
    [string]$ToolDirectory = (Join-Path $env:LOCALAPPDATA 'emo'),
    [int]$StartupTimeoutSeconds = 180
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
$env:ATOTI_HIDE_EULA_MESSAGE = 'True'
$process = Start-Process -FilePath $python -ArgumentList @("`"$cubeScript`"", $port) `
    -RedirectStandardOutput $logPath -RedirectStandardError "$logPath.err" -NoNewWindow -PassThru

$deadline = [DateTime]::UtcNow.AddSeconds($StartupTimeoutSeconds)
while (-not ((Test-Path $logPath) -and (Select-String -Path $logPath -Pattern '^READY' -Quiet))) {
    if ($process.HasExited -or [DateTime]::UtcNow -gt $deadline) {
        if (-not $process.HasExited) { taskkill.exe /PID $process.Id /T /F | Out-Null }
        $errorText = if (Test-Path "$logPath.err") { Get-Content "$logPath.err" -Tail 20 | Out-String } else { '' }
        throw "The OLAP test cube did not start on port $port. $errorText"
    }

    Start-Sleep -Seconds 2
}

[pscustomobject]@{
    Process = $process
    Settings = @{
        EXCELMCP_TEST_OLAP_CONNECTION_STRING = "OLEDB;Provider=MSOLAP;Data Source=http://localhost:$port/xmla;Initial Catalog=atoti"
        EXCELMCP_TEST_OLAP_CUBE = 'SalesCube'
        EXCELMCP_TEST_OLAP_HIERARCHY = '[Sales].[Calendar]'
        EXCELMCP_TEST_OLAP_LEVEL = '[Sales].[Calendar].[Month]'
    }
}
