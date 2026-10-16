#Requires -Version 7.4
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$directory = Join-Path ([IO.Path]::GetTempPath()) "excel-runner-control-$([Guid]::NewGuid().ToString('N'))"
$originalPath = $env:PATH
$originalInvalidResponse = $env:EXCELMCP_TEST_GH_ALWAYS_INVALID
$originalWhitespaceResponse = $env:EXCELMCP_TEST_GH_WHITESPACE
$originalStderrResponse = $env:EXCELMCP_TEST_GH_STDERR
New-Item -ItemType Directory -Path $directory | Out-Null
try {
    . (Join-Path $root 'ExcelRunnerHost.ps1')
    $deadline = [DateTime]::UtcNow.AddMinutes(1)
    $env:PATH = "$directory;$originalPath"
    $gh = Join-Path $directory 'gh.cmd'
    @'
@echo off
echo call>>"%~dp0calls.txt"
if "%EXCELMCP_TEST_GH_ALWAYS_INVALID%"=="true" goto invalid
if exist "%~dp0recovered" goto success
echo recovered>"%~dp0recovered"
if "%EXCELMCP_TEST_GH_STDERR%"=="true" goto stderrFailure
if "%EXCELMCP_TEST_GH_WHITESPACE%"=="true" goto whitespace
echo {"name":"cut
exit /b 0
:invalid
if "%EXCELMCP_TEST_GH_STDERR%"=="true" goto stderrFailure
if "%EXCELMCP_TEST_GH_WHITESPACE%"=="true" goto whitespace
echo {"partial":
exit /b 0
:stderrFailure
echo unexpected end of JSON input 1>&2
exit /b 1
:whitespace
echo.
exit /b 0
:success
echo {"ok":true}
exit /b 0
'@ | Set-Content -LiteralPath $gh -Encoding ascii

    $result = Invoke-ExcelRunnerGithub 'repos/test/repo/actions/runs' -Paginate
    if (-not $result.ok -or @(Get-Content (Join-Path $directory 'calls.txt')).Count -ne 2) {
        throw 'A truncated JSON string response was not retried successfully.'
    }

    Remove-Item -LiteralPath (Join-Path $directory 'calls.txt'), (Join-Path $directory 'recovered')
    $env:EXCELMCP_TEST_GH_WHITESPACE = 'true'
    $result = Invoke-ExcelRunnerGithub 'repos/test/repo/actions/runs' -Paginate
    if (-not $result.ok -or @(Get-Content (Join-Path $directory 'calls.txt')).Count -ne 2) {
        throw 'A transient whitespace-only response was not retried successfully.'
    }

    Remove-Item -LiteralPath (Join-Path $directory 'calls.txt'), (Join-Path $directory 'recovered')
    Remove-Item Env:EXCELMCP_TEST_GH_WHITESPACE -ErrorAction SilentlyContinue
    $env:EXCELMCP_TEST_GH_STDERR = 'true'
    $result = Invoke-ExcelRunnerGithub 'repos/test/repo/actions/runs' -Paginate
    if (-not $result.ok -or @(Get-Content (Join-Path $directory 'calls.txt')).Count -ne 2) {
        throw 'A transient GitHub CLI stderr failure was not retried successfully.'
    }

    Remove-Item -LiteralPath (Join-Path $directory 'calls.txt'), (Join-Path $directory 'recovered')
    Remove-Item Env:EXCELMCP_TEST_GH_WHITESPACE -ErrorAction SilentlyContinue
    $env:EXCELMCP_TEST_GH_ALWAYS_INVALID = 'true'
    Remove-Item Env:EXCELMCP_TEST_GH_STDERR -ErrorAction SilentlyContinue
    $failure = $null
    try { $null = Invoke-ExcelRunnerGithub 'repos/test/repo/actions/runs' -Paginate }
    catch { $failure = $_.Exception.Message }
    if (-not $failure -or $failure -notmatch 'repos/test/repo/actions/runs' -or
        @(Get-Content (Join-Path $directory 'calls.txt')).Count -ne 2) {
        throw "Persistent incomplete response was not handled as expected. Failure='$failure'; calls=$(@(Get-Content (Join-Path $directory 'calls.txt')).Count)."
    }

    Remove-Item -LiteralPath (Join-Path $directory 'calls.txt')
    $env:EXCELMCP_TEST_GH_WHITESPACE = 'true'
    $failure = $null
    try { $null = Invoke-ExcelRunnerGithub 'repos/test/repo/actions/runs' -Paginate }
    catch { $failure = $_.Exception.Message }
    if (-not $failure -or $failure -notmatch 'repos/test/repo/actions/runs' -or
        @(Get-Content (Join-Path $directory 'calls.txt')).Count -ne 2) {
        throw "Persistent whitespace-only response was not handled as expected. Failure='$failure'; calls=$(@(Get-Content (Join-Path $directory 'calls.txt')).Count)."
    }

    Remove-Item -LiteralPath (Join-Path $directory 'calls.txt')
    Remove-Item Env:EXCELMCP_TEST_GH_WHITESPACE -ErrorAction SilentlyContinue
    $env:EXCELMCP_TEST_GH_STDERR = 'true'
    $failure = $null
    try { $null = Invoke-ExcelRunnerGithub 'repos/test/repo/actions/runs' -Paginate }
    catch { $failure = $_.Exception.Message }
    if (-not $failure -or $failure -notmatch 'repos/test/repo/actions/runs' -or
        @(Get-Content (Join-Path $directory 'calls.txt')).Count -ne 2) {
        throw "Persistent GitHub CLI stderr error was not handled as expected. Failure='$failure'; calls=$(@(Get-Content (Join-Path $directory 'calls.txt')).Count)."
    }

    Remove-Item -LiteralPath (Join-Path $directory 'calls.txt')
    $failure = $null
    try { $null = Invoke-ExcelRunnerGithub 'repos/test/repo' -Method POST }
    catch { $failure = $_.Exception.Message }
    if (-not $failure -or $failure -notmatch 'without retry' -or
        @(Get-Content (Join-Path $directory 'calls.txt')).Count -ne 1) {
        throw 'A non-idempotent GitHub request was retried after an incomplete response.'
    }
}
finally {
    $env:PATH = $originalPath
    if ($null -eq $originalInvalidResponse) {
        Remove-Item Env:EXCELMCP_TEST_GH_ALWAYS_INVALID -ErrorAction SilentlyContinue
    }
    else { $env:EXCELMCP_TEST_GH_ALWAYS_INVALID = $originalInvalidResponse }
    if ($null -eq $originalWhitespaceResponse) {
        Remove-Item Env:EXCELMCP_TEST_GH_WHITESPACE -ErrorAction SilentlyContinue
    }
    else { $env:EXCELMCP_TEST_GH_WHITESPACE = $originalWhitespaceResponse }
    if ($null -eq $originalStderrResponse) {
        Remove-Item Env:EXCELMCP_TEST_GH_STDERR -ErrorAction SilentlyContinue
    }
    else { $env:EXCELMCP_TEST_GH_STDERR = $originalStderrResponse }
    Remove-Item -LiteralPath $directory -Recurse -Force
}
Write-Output 'Excel runner control retries one incomplete GitHub response and fails closed with context.'
