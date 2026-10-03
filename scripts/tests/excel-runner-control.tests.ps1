$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'ExcelRunnerPolicy.ps1')
$run = @{
    event = 'dynamic'; path = 'dynamic/copilot-swe-agent/copilot'
    actor = @{ id = 198982749; type = 'Bot' }
    head_repository = @{ full_name = 'synthetic/repository' }
}
if (-not (Test-ExcelRunnerCloudRun $run 'synthetic/repository')) { throw 'Verified platform-owned agent run was rejected.' }
foreach ($field in @('event', 'path', 'actor', 'head_repository')) {
    $candidate = $run.Clone()
    $candidate[$field] = 'untrusted'
    if (Test-ExcelRunnerCloudRun $candidate 'synthetic/repository') { throw 'Untrusted or unrelated workflow was admitted.' }
}
$idle = @{ listeners = 0; workers = 0; excel = 0; taskRunning = $false }
Assert-ExcelRunnerIdle -Jobs @() -Guest $idle
foreach ($field in @('listeners', 'workers', 'excel', 'taskRunning')) {
    $candidate = $idle.Clone()
    $candidate[$field] = 1
    $failed = $false
    try { Assert-ExcelRunnerIdle -Jobs @() -Guest $candidate } catch { $failed = $true }
    if (-not $failed) { throw 'An active guest was treated as idle.' }
}
foreach ($status in @('queued', 'in_progress', 'waiting')) {
    $job = @{ runner_name = 'azure-excel-copilot'; status = $status; labels = @() }
    $failed = $false
    try { Assert-ExcelRunnerIdle -Jobs @($job) -Guest $idle } catch { $failed = $true }
    if (-not $failed) { throw 'Whole-job activity must block shutdown and maintenance.' }
    $job.runner_name = ''
    $job.labels = @('self-hosted', 'excel-copilot')
    $failed = $false
    try { Assert-ExcelRunnerIdle -Jobs @($job) -Guest $idle } catch { $failed = $true }
    if (-not $failed) { throw 'Queued label demand must block maintenance.' }
}
Assert-ExcelRunnerIdle -Jobs @(@{ runner_name = 'azure-excel-copilot'; status = 'completed' }) -Guest $idle
Write-Output 'Platform identity, full-job activity, queue and guest-idle guards passed.'
