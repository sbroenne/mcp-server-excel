$ErrorActionPreference = 'Stop'
$permit = Get-Content (Join-Path $PSScriptRoot 'permit.json') -Raw | ConvertFrom-Json
. (Join-Path $PSScriptRoot 'ExcelRunnerPolicy.ps1')
$boot = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime()
if ($env:GITHUB_REPOSITORY -ne 'sbroenne/mcp-server-excel' -or
    $permit.runId -ne $env:GITHUB_RUN_ID -or $permit.job -ne $env:GITHUB_JOB -or
    $permit.state -ne 'admitted' -or (ConvertTo-ExcelRunnerUtc $permit.bootTime) -ne $boot -or
    $env:GITHUB_RUN_ID -notmatch '^\d+$' -or $env:GITHUB_RUN_ATTEMPT -notmatch '^\d+$' -or
    (ConvertTo-ExcelRunnerUtc $permit.expiresAt) -lt [DateTime]::UtcNow -or
    ($permit.kind -eq 'cloud' -and ($env:GITHUB_ACTOR -ne 'Copilot' -or $env:GITHUB_EVENT_NAME -ne 'dynamic')) -or
    ($permit.kind -eq 'validation' -and ($env:GITHUB_ACTOR -ne 'sbroenne' -or
        $env:GITHUB_EVENT_NAME -ne 'workflow_dispatch' -or $env:GITHUB_REF -ne "refs/heads/$($permit.defaultBranch)")) -or
    $permit.kind -notin @('cloud', 'validation')) {
    $context = @{ run = $env:GITHUB_RUN_ID; attempt = $env:GITHUB_RUN_ATTEMPT; job = $env:GITHUB_JOB
        actor = $env:GITHUB_ACTOR; event = $env:GITHUB_EVENT_NAME } | ConvertTo-Json -Compress
    throw "The hosted controller did not admit this GitHub job; refuse it before any setup or agent code runs. GitHub context: $context"
}
$directory = Join-Path $env:LOCALAPPDATA 'ExcelMcp\Jobs'
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$recordPath = Join-Path $directory "$env:GITHUB_RUN_ID-$env:GITHUB_RUN_ATTEMPT.json"
if (Test-Path -LiteralPath $recordPath) { throw 'Previous job cleanup has not completed.' }
$session = (Get-Process -Id $PID).SessionId
$processes = @(
    foreach ($process in Get-CimInstance Win32_Process -Filter "SessionId=$session") {
        @{ id = $process.ProcessId; startedAt = $process.CreationDate.ToUniversalTime().ToString('o') }
    }
)
@{ runId = $env:GITHUB_RUN_ID; attempt = $env:GITHUB_RUN_ATTEMPT; session = $session; processes = $processes
    bootTime = $boot.ToString('o') } |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $recordPath -Encoding UTF8
