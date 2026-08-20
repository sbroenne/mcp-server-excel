param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"

$PluginDir = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$WrapperPath = Join-Path $PluginDir "bin\start-cli.ps1"
$DownloadScriptPath = Join-Path $PluginDir "bin\download.ps1"
$CopilotDir = Join-Path $env:USERPROFILE ".copilot"
$CopilotBinDir = Join-Path $CopilotDir "bin"
$ShimCmdPath = Join-Path $CopilotBinDir "excelcli.cmd"
$ShimPs1Path = Join-Path $CopilotBinDir "excelcli.ps1"

Write-Host "Excel CLI Global Install Helper" -ForegroundColor Cyan
Write-Host "===============================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $WrapperPath)) {
    Write-Error "❌ Plugin wrapper not found at $WrapperPath"
    exit 1
}

if (-not (Test-Path $DownloadScriptPath)) {
    Write-Error "❌ Plugin bootstrap script not found at $DownloadScriptPath"
    exit 1
}

if (-not (Test-Path $CopilotBinDir)) {
    Write-Host "[Install] Creating $CopilotBinDir ..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $CopilotBinDir -Force | Out-Null
}

$escapedDownloadPath = $DownloadScriptPath.Replace('"', '""')
# Resolve the runtime first, then invoke it with cmd's verbatim %*. Routing arguments through
# "powershell -File" would strip embedded double quotes and corrupt JSON arguments such as
# --values '[["Name","Amount"]]', so the executable is called directly instead.
$cmdShim = @"
@echo off
setlocal
set "EXCELCLI_EXE="
for /f "usebackq delims=" %%i in (``powershell -NoProfile -ExecutionPolicy Bypass -File "$escapedDownloadPath" -PassThru -Quiet``) do set "EXCELCLI_EXE=%%i"
if not defined EXCELCLI_EXE (
    echo excel-cli bootstrap did not resolve a usable excelcli.exe runtime. 1>&2
    exit /b 1
)
"%EXCELCLI_EXE%" %*
exit /b %ERRORLEVEL%
"@

$ps1Shim = @"
& '$WrapperPath' @args
exit `$LASTEXITCODE
"@

if (((Test-Path $ShimCmdPath) -or (Test-Path $ShimPs1Path)) -and -not $Force) {
    Write-Host "✅ CLI shims already exist in $CopilotBinDir" -ForegroundColor Green
    Write-Host "Run again with -Force to overwrite them." -ForegroundColor Yellow
} else {
    Write-Host "[Install] Writing CLI shims..." -ForegroundColor Yellow
    Set-Content -Path $ShimCmdPath -Value $cmdShim -Encoding ASCII
    Set-Content -Path $ShimPs1Path -Value $ps1Shim -Encoding UTF8
}

$userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
$pathEntries = @()
if (-not [string]::IsNullOrWhiteSpace($userPath)) {
    $pathEntries = $userPath -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
}

if ($pathEntries -notcontains $CopilotBinDir) {
    Write-Host "[Install] Adding $CopilotBinDir to user PATH..." -ForegroundColor Yellow
    $newUserPath = if ([string]::IsNullOrWhiteSpace($userPath)) {
        $CopilotBinDir
    } else {
        "$userPath;$CopilotBinDir"
    }

    [Environment]::SetEnvironmentVariable("PATH", $newUserPath, "User")
    $env:PATH = "$env:PATH;$CopilotBinDir"
}

Write-Host ""
Write-Host "✅ excelcli shims are installed." -ForegroundColor Green
Write-Host "   Wrapper: $WrapperPath" -ForegroundColor Gray
Write-Host "   Shim dir: $CopilotBinDir" -ForegroundColor Gray
Write-Host ""
Write-Host "The first real 'excelcli' invocation will auto-download the newest Windows runtime." -ForegroundColor Cyan
Write-Host "Verify installation:" -ForegroundColor Cyan
Write-Host "   excelcli --version" -ForegroundColor Gray
Write-Host "   excelcli --help" -ForegroundColor Gray
