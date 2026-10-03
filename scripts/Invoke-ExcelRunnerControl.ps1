<#
.SYNOPSIS
Runs a hosted wake/idle-shutdown check covering complete generated cloud-agent jobs.
.DESCRIPTION
Reuses mcp-windows hosted Azure start/deallocate control and the interactive task.
No Linux VM, Azure daemon or permanent administration token is required.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ResourceGroup = 'rg-excel-copilot-runner',
    [string]$VmName = 'vm-excel-copilot-runner',
    [string]$Repository = 'sbroenne/mcp-server-excel',
    [string]$SubscriptionId
)
$ErrorActionPreference = 'Stop'
if (-not $PSCmdlet.ShouldProcess($VmName, 'Check complete GitHub jobs and safely wake or park the dedicated runner')) { return }
$deadline = [DateTime]::UtcNow.AddMinutes(25)
$subscriptionArguments = if ($SubscriptionId) { @('--subscription', $SubscriptionId) } else { @() }
. (Join-Path $PSScriptRoot 'ExcelRunnerHost.ps1')
function Assert-SelectedExcelRunnerJob {
    param($Job, [string]$JobId, [string]$RunId, [string]$JobName)
    if ("$($Job.id)" -ne $JobId -or "$($Job.run_id)" -ne $RunId -or
        $Job.name -ne $JobName -or -not (Test-ExcelRunnerJobTarget $Job)) {
        throw 'The selected job identity changed during control; refuse admission or completion.'
    }
}

$jobs = @(Get-ExcelRunnerActiveJobs $Repository)
$target = @($jobs | Where-Object { Test-ExcelRunnerJobTarget $_ } | Sort-Object id -Unique)
if (@($target | Where-Object { -not $_.trustedCloudRun -and -not $_.trustedValidationRun }).Count) {
    throw 'Unapproved work requests this runner; do not start its listener.'
}
$vm = Assert-ExcelRunnerOwnedVm
$power = @($vm.instanceView.statuses | Where-Object code -Like 'PowerState/*')
if ($power.Count -ne 1) { throw 'VM power state is uncertain.' }
if ($power[0].code -eq 'PowerState/deallocated' -and $target.Count -eq 0) {
    Write-Output 'No approved demand; VM remains parked.'
    return
}
$failures = [Collections.Generic.List[Exception]]::new()
$started = $false
try {
    if ($power[0].code -eq 'PowerState/deallocated') {
        $repositoryInfo = Invoke-ExcelRunnerGithub "repos/$Repository"
        $history = Invoke-ExcelRunnerGithub "repos/$Repository/actions/workflows/excel-runner-maintenance.yml/runs?per_page=100"
        Assert-ExcelRunnerMaintenanceHistory @($history.workflow_runs) $Repository $repositoryInfo.default_branch
        $null = Invoke-RunnerAzure @(
            'vm', 'auto-shutdown', '--resource-group', $ResourceGroup, '--name', $VmName,
            '--time', [DateTime]::UtcNow.AddHours(6).ToString('HHmm')
        )
        $started = $true
        $null = Invoke-RunnerAzure @('vm', 'start', '--resource-group', $ResourceGroup, '--name', $VmName) 900
    }
    elseif ($power[0].code -ne 'PowerState/running') { throw 'VM is changing power state; do not replay lifecycle operations.' }
    Wait-ExcelRunnerGuestAgent
    $activity = Get-ExcelRunnerGuestActivity
    if ($activity.listeners -gt 0 -and $activity.workers -eq 0) {
        Stop-ExcelRunnerExpiredListener $jobs
        Start-Sleep -Seconds 5
        $activity = Get-ExcelRunnerGuestActivity
    }
    if ($activity.workers -gt 0 -or $activity.listeners -gt 0 -or $activity.taskRunning) {
        Write-Output 'The complete job/listener lifecycle remains active; leaving the VM undisturbed.'
        return
    }
    if ($activity.cleanupPending) {
        Assert-ExcelRunnerIdle $jobs @{ listeners = 0; workers = 0; excel = 0; taskRunning = $false } -AllowQueued
        Invoke-ExcelRunnerJobRecovery
        $activity = Get-ExcelRunnerGuestActivity
    }
    if ($activity.excel -gt 0 -or $activity.cleanupPending) {
        throw 'Previous work left an unowned workbook or incomplete cleanup; keep coding admission quarantined.'
    }
    if ($target.Count -gt 0) {
        $next = @($target | Where-Object status -EQ 'queued' | Sort-Object id | Select-Object -First 1)
        if ($next.Count -ne 1) { throw 'GitHub reports active work but no listener; recovery needs inspection.' }
        $null = Invoke-ExcelRunnerGuest @'
. 'C:\ProgramData\ExcelMcp\Desktop\ExcelRunnerPolicy.ps1'
Assert-ExcelRunnerPatchState (Get-Content 'C:\ProgramData\ExcelMcp\Maintenance\patched.json' -Raw | ConvertFrom-Json)
Write-Output 'EXCELMCP_CONTROL={"state":"patched"}'
'@
        $null = Invoke-ExcelRunnerDesktopHealth
        $kind = if ($next[0].trustedCloudRun) { 'cloud' } else { 'validation' }
        $jobName = if ($kind -eq 'cloud') { 'copilot' } else { 'excel' }
        $repo = Invoke-ExcelRunnerGithub "repos/$Repository"
        $runId = "$($next[0].excelRunId)"
        if ($runId -notmatch '^\d+$') { throw 'Invalid admitted workflow run ID.' }
        $jobId = "$($next[0].id)"
        if ($jobId -notmatch '^\d+$') { throw 'Invalid admitted workflow job ID.' }
        $currentJob = Invoke-ExcelRunnerGithub "repos/$Repository/actions/jobs/$jobId"
        Assert-SelectedExcelRunnerJob $currentJob $jobId $runId $jobName
        if ($currentJob.status -eq 'completed') {
            Write-Output 'Selected work completed or was cancelled during preparation; no listener started.'
            if (@(Get-ExcelRunnerActiveJobs $Repository | Where-Object { Test-ExcelRunnerJobTarget $_ }).Count) {
                Write-Output 'Other complete-job demand remains; leaving admission to the next control check.'
                return
            }
        }
        elseif ($currentJob.status -eq 'queued') {
            $null = Invoke-ExcelRunnerGuest @"
`$state = @{ state = 'admitted'; runId = '$runId'; job = '$jobName'; kind = '$kind'; defaultBranch = '$($repo.default_branch)'
    bootTime = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime().ToString('o')
    expiresAt = [DateTime]::UtcNow.AddMinutes(10).ToString('o') }
`$state | ConvertTo-Json -Compress | Set-Content 'C:\ProgramData\ExcelMcp\Desktop\permit.json' -Encoding UTF8
Start-ScheduledTask -TaskName 'ExcelMcp-GitHub-Runner'
Write-Output 'EXCELMCP_CONTROL={"state":"admitted"}'
"@
            $until = [DateTime]::UtcNow.AddMinutes(2)
            $completedBeforeObservation = $false
            do {
                Start-Sleep -Seconds 10
                $activity = Get-ExcelRunnerGuestActivity
                if ($activity.listeners -gt 0 -and $activity.taskRunning) {
                    Write-Output "Interactive one-job listener admitted workflow $runId."
                    return
                }
                $observedJob = Invoke-ExcelRunnerGithub "repos/$Repository/actions/jobs/$jobId"
                Assert-SelectedExcelRunnerJob $observedJob $jobId $runId $jobName
                if ($observedJob.status -eq 'completed') {
                    if ($activity.workers -gt 0 -or $activity.listeners -gt 0 -or $activity.taskRunning) {
                        Write-Output 'Admitted job completed; its guest cleanup is still active. Leaving the VM undisturbed.'
                        return
                    }
                    if ($activity.cleanupPending) { Invoke-ExcelRunnerJobRecovery }
                    if (@(Get-ExcelRunnerActiveJobs $Repository | Where-Object { Test-ExcelRunnerJobTarget $_ }).Count) {
                        Write-Output 'Admitted job completed; other demand remains for the next control check.'
                        return
                    }
                    $completedBeforeObservation = $true
                    break
                }
            } while ([DateTime]::UtcNow -lt $until)
            if (-not $completedBeforeObservation) { throw 'The admitted interactive listener did not start.' }
            Write-Output 'Admitted job completed before listener observation; qualifying idle cleanup before parking.'
        }
        else { throw 'The selected job is no longer queued; its complete-job ownership needs inspection.' }
    }
    Start-Sleep -Seconds 30
    Assert-ExcelRunnerIdle @(Get-ExcelRunnerActiveJobs $Repository) (Get-ExcelRunnerGuestActivity)
    $null = Invoke-ExcelRunnerDesktopHealth
    Assert-ExcelRunnerIdle @(Get-ExcelRunnerActiveJobs $Repository) (Get-ExcelRunnerGuestActivity)
    $null = Invoke-RunnerAzure @('vm', 'deallocate', '--resource-group', $ResourceGroup, '--name', $VmName) 600
    Write-Output 'Whole-job and desktop cleanup checks passed; VM parked.'
}
catch { $failures.Add($_.Exception) }
finally {
    if ($failures.Count -and ($started -or $power[0].code -eq 'PowerState/running')) {
        $deadline = [DateTime]::UtcNow.AddMinutes(15)
        try {
            Assert-ExcelRunnerIdle @(Get-ExcelRunnerActiveJobs $Repository) (Get-ExcelRunnerGuestActivity) -AllowQueued
            $null = Invoke-RunnerAzure @('vm', 'deallocate', '--resource-group', $ResourceGroup, '--name', $VmName) 600
        }
        catch { $failures.Add($_.Exception) }
    }
}
if ($failures.Count) { throw [AggregateException]::new('Runner control failed; admission remains unsafe.', $failures) }
