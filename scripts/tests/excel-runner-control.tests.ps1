#Requires -Version 7.4
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$directory = Join-Path ([IO.Path]::GetTempPath()) "excel-runner-control-$([Guid]::NewGuid().ToString('N'))"
$originalPath = $env:PATH
$originalInvalidResponse = $env:EXCELMCP_TEST_GH_ALWAYS_INVALID
New-Item -ItemType Directory -Path $directory | Out-Null
try {
    . (Join-Path $root 'ExcelRunnerHost.ps1')
    $deadline = [DateTime]::UtcNow.AddMinutes(1)
    $env:PATH = "$directory;$originalPath"
    $gh = Join-Path $directory 'gh.cmd'
    @'
@echo off
echo call>>"%~dp0calls.txt"
if "%EXCELMCP_TEST_GH_ALWAYS_INVALID%"=="true" (
  echo {"partial":
  exit /b 0
)
if exist "%~dp0recovered" (
  echo {"ok":true}
  exit /b 0
)
echo recovered>"%~dp0recovered"
echo unexpected end of JSON input 1>&2
exit /b 1
'@ | Set-Content -LiteralPath $gh -Encoding ascii

    $result = Invoke-ExcelRunnerGithub 'repos/test/repo/actions/runs' -Paginate
    if (-not $result.ok -or @(Get-Content (Join-Path $directory 'calls.txt')).Count -ne 2) {
        throw 'An incomplete GitHub response was not retried successfully.'
    }

    Remove-Item -LiteralPath (Join-Path $directory 'calls.txt'), (Join-Path $directory 'recovered')
    $env:EXCELMCP_TEST_GH_ALWAYS_INVALID = 'true'
    $failure = $null
    try { $null = Invoke-ExcelRunnerGithub 'repos/test/repo/actions/runs' -Paginate }
    catch { $failure = $_.Exception.Message }
    if (-not $failure -or $failure -notmatch 'repos/test/repo/actions/runs' -or
        @(Get-Content (Join-Path $directory 'calls.txt')).Count -ne 2) {
        throw "Persistent incomplete response was not handled as expected. Failure='$failure'; calls=$(@(Get-Content (Join-Path $directory 'calls.txt')).Count)."
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
    Remove-Item -LiteralPath $directory -Recurse -Force
}
Write-Output 'Excel runner control retries one incomplete GitHub response and fails closed with context.'
