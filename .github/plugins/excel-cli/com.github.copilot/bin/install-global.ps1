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

$escapedDownloadPath = $DownloadScriptPath.Replace("'", "''")
$bootstrapCommand = "[Console]::OutputEncoding = [Text.UTF8Encoding]::new(`$false); & '$escapedDownloadPath' -PassThru -Quiet"
$encodedBootstrap = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($bootstrapCommand))
# Resolve the runtime first, then invoke it with cmd's verbatim %*. Routing arguments through
# "powershell -File" would strip embedded double quotes and corrupt JSON arguments such as
# --values '[["Name","Amount"]]', so the executable is called directly instead.
$cmdShim = @"
@echo off
setlocal
where npx.cmd >nul 2>&1
if not errorlevel 1 (
    call npx.cmd -y @sbroenne/excelcli@latest %*
    exit /b %ERRORLEVEL%
)
set "EXCELCLI_EXE="
for /f "tokens=2 delims=:" %%i in ('chcp') do set "EXCELCLI_CODEPAGE=%%i"
chcp 65001 >nul
for /f "usebackq delims=" %%i in (``powershell -NoProfile -OutputFormat Text -ExecutionPolicy Bypass -EncodedCommand "$encodedBootstrap"``) do set "EXCELCLI_EXE=%%i"
chcp %EXCELCLI_CODEPAGE% >nul
if not defined EXCELCLI_EXE (
    echo excel-cli bootstrap did not resolve a usable excelcli.exe runtime. 1>&2
    exit /b 1
)
"%EXCELCLI_EXE%" %*
exit /b %ERRORLEVEL%
"@

$escapedWrapperPath = $WrapperPath.Replace("'", "''")
$ps1Shim = @"
& '$escapedWrapperPath' @args
exit `$LASTEXITCODE
"@

$updates = @(
    @{ Path = $ShimCmdPath; Content = $cmdShim; Encoding = [Text.Encoding]::ASCII },
    @{ Path = $ShimPs1Path; Content = $ps1Shim; Encoding = [Text.UTF8Encoding]::new($true) }
) | Where-Object { $Force -or -not (Test-Path -LiteralPath $_.Path) }
$prepared = [Collections.Generic.List[object]]::new()
$installationSucceeded = $false
try {
    foreach ($update in $updates) {
        $update.Temp = "$($update.Path).$([Guid]::NewGuid().ToString('N')).tmp"
        $update.Backup = "$($update.Temp).bak"
        $update.Installed = $false
        $prepared.Add($update)
        [IO.File]::WriteAllText($update.Temp, $update.Content, $update.Encoding)
    }
    foreach ($update in $prepared) {
        if (Test-Path -LiteralPath $update.Path) {
            [IO.File]::Replace($update.Temp, $update.Path, $update.Backup)
        } else {
            [IO.File]::Move($update.Temp, $update.Path)
        }
        $update.Installed = $true
    }
    $installationSucceeded = $true
} catch {
    $installationError = $_
    $rollbackErrors = [Collections.Generic.List[string]]::new()
    foreach ($update in $prepared) {
        if (-not $update.Installed) { continue }
        try {
            if (Test-Path -LiteralPath $update.Backup) {
                [IO.File]::Replace($update.Backup, $update.Path, [System.Management.Automation.Language.NullString]::Value)
            } else {
                Remove-Item -LiteralPath $update.Path
            }
        } catch {
            $rollbackErrors.Add($_.Exception.Message)
        }
    }
    if ($rollbackErrors.Count) {
        throw "Launcher installation failed: $($installationError.Exception.Message). Rollback also failed: $($rollbackErrors -join '; '). Original launcher backups have been retained."
    }
    throw $installationError
} finally {
    foreach ($update in $prepared) {
        $cleanupPaths = @($update.Temp)
        if ($installationSucceeded) { $cleanupPaths += $update.Backup }
        foreach ($path in $cleanupPaths) {
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path }
        }
    }
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
Write-Host "'excelcli' uses the latest npm package through npx when Node.js is available." -ForegroundColor Cyan
Write-Host "The plugin release downloader remains available as a fallback." -ForegroundColor Cyan
Write-Host "Verify installation:" -ForegroundColor Cyan
Write-Host "   excelcli --version" -ForegroundColor Gray
Write-Host "   excelcli --help" -ForegroundColor Gray
