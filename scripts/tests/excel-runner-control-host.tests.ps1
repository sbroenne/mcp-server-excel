$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'ExcelRunnerPolicy.ps1')
$source = Get-Content -LiteralPath (Join-Path $root 'Invoke-ExcelRunnerControl.ps1') -Raw
$import = ". (Join-Path `$PSScriptRoot 'ExcelRunnerHost.ps1')"
if (-not $source.Contains($import)) { throw 'The actual controller host import was not found.' }
$controller = [scriptblock]::Create($source.Replace($import, ''))
$script:Calls = [Collections.Generic.List[string]]::new()
function Start-Sleep { param($Seconds) }
function Get-ExcelRunnerActiveJobs {
    if ($script:Mode -eq 'parked' -or ($script:Mode -eq 'cancelled' -and $script:JobRefreshed) -or
        ($script:Mode -like 'short-*' -and $script:Admitted)) { return @() }
    return @(@{
        id = 1; excelRunId = 123; labels = @('excel-copilot'); runner_name = ''
        status = 'queued'; trustedCloudRun = $script:Mode -ne 'untrusted'; trustedValidationRun = $false
    })
}
function Assert-ExcelRunnerOwnedVm {
    return @{ instanceView = @{ statuses = @(@{ code = $(if ($script:Mode -eq 'busy') { 'PowerState/running' } else { 'PowerState/deallocated' }) }) } }
}
function Invoke-RunnerAzure {
    param($Arguments, $TimeoutSeconds)
    $script:Calls.Add(($Arguments -join ' '))
}
function Wait-ExcelRunnerGuestAgent { $script:Calls.Add('wait-agent') }
function Get-ExcelRunnerGuestActivity {
    $script:Calls.Add('guest-activity')
    $finished = $script:Mode -like 'short-*' -and $script:Admitted
    return @{
        listeners = [int]($script:Admitted -and -not $finished)
        workers = [int]($script:Mode -eq 'busy' -or ($finished -and $script:Mode -eq 'short-cleaning')); excel = 0
        taskRunning = $script:Admitted -and (-not $finished -or $script:Mode -eq 'short-cleaning')
        cleanupPending = $finished -and $script:Mode -eq 'short-recovery' -and -not $script:Recovered
    }
}
function Invoke-ExcelRunnerGithub {
    param($Endpoint)
    if ($Endpoint -match '/actions/jobs/1$') {
        $script:Calls.Add('refresh-job')
        $script:JobRefreshed = $true
        return @{
            id = 1; run_id = $(if ($script:Mode -eq 'changed-job' -or ($script:Mode -eq 'short-changed-job' -and $script:Admitted)) { 999 } else { 123 })
            name = 'copilot'; labels = @('excel-copilot')
            status = $(if ($script:Mode -eq 'cancelled' -or ($script:Mode -like 'short-*' -and $script:Admitted)) { 'completed' } else { 'queued' })
        }
    }
    if ($Endpoint -match '/runs\?') {
        return @{ workflow_runs = @(@{
            head_branch = 'main'; head_repository = @{ full_name = 'synthetic/repository' }
            path = '.github/workflows/excel-runner-maintenance.yml'; status = 'completed'
            conclusion = $(if ($script:Mode -eq 'stale-history') { 'failure' } else { 'success' })
            created_at = [DateTime]::UtcNow.ToString('o'); updated_at = [DateTime]::UtcNow.ToString('o')
        }) }
    }
    return @{ default_branch = 'main' }
}
function Invoke-ExcelRunnerGuest {
    param($Script)
    if ($Script -match 'Assert-ExcelRunnerPatchState' -and $script:Mode -eq 'stale-patch') { throw 'Synthetic stale protected patch evidence.' }
    if ($Script -match "Start-ScheduledTask -TaskName 'ExcelMcp-GitHub-Runner'") {
        $script:Calls.Add('admit-listener')
        $script:Admitted = $true
    }
    return @{ state = 'patched' }
}
function Invoke-ExcelRunnerDesktopHealth { $script:Calls.Add('desktop-health') }
function Invoke-ExcelRunnerJobRecovery { $script:Calls.Add('recover-jobs'); $script:Recovered = $true }
foreach ($mode in @('parked', 'untrusted', 'busy', 'stale-history', 'stale-patch', 'admit', 'cancelled', 'changed-job',
    'short-completed', 'short-cleaning', 'short-recovery', 'short-changed-job')) {
    $script:Mode = $mode
    $script:Admitted = $false
    $script:JobRefreshed = $false
    $script:Recovered = $false
    $script:Calls.Clear()
    $failure = $null
    try { & $controller -ResourceGroup synthetic -VmName synthetic -Repository synthetic/repository }
    catch { $failure = $_.Exception }
    $starts = @($script:Calls | Where-Object { $_ -like 'vm start *' }).Count
    $parks = @($script:Calls | Where-Object { $_ -like 'vm deallocate *' }).Count
    if ($mode -in @('parked', 'untrusted', 'stale-history')) {
        if ($script:Calls.Count) { throw "$mode must not contact the guest or change VM power." }
    }
    if ($mode -eq 'busy' -and ($starts -or $parks -or $failure)) { throw 'A complete active job must remain undisturbed.' }
    if ($mode -eq 'stale-patch' -and (-not $failure -or $starts -ne 1 -or $parks -ne 1 -or $script:Admitted)) {
        throw 'A failed guest patch gate must park an idle VM without admitting work.'
    }
    if ($mode -eq 'admit' -and ($failure -or $starts -ne 1 -or $parks -or -not $script:Admitted -or
        -not $script:Calls.Contains('desktop-health'))) { throw 'Trusted queued demand must qualify the desktop before one-job admission.' }
    if ($mode -eq 'cancelled' -and ($failure -or $starts -ne 1 -or $parks -ne 1 -or $script:Admitted -or
        -not $script:JobRefreshed)) { throw 'Work cancelled during desktop preparation must not start a listener and must park when idle.' }
    if ($mode -eq 'changed-job' -and (-not $failure -or $starts -ne 1 -or $parks -ne 1 -or $script:Admitted)) {
        throw 'A changed selected job identity must fail without admitting work.'
    }
    if ($mode -in @('short-completed', 'short-recovery') -and ($failure -or $starts -ne 1 -or $parks -ne 1 -or
        -not $script:Admitted -or @($script:Calls | Where-Object { $_ -eq 'desktop-health' }).Count -ne 2)) {
        throw 'An admitted job completed before listener observation must qualify cleanup and park without a false startup failure.'
    }
    if ($mode -eq 'short-recovery' -and -not $script:Calls.Contains('recover-jobs')) { throw 'Completed work still requires its pending owned cleanup.' }
    if ($mode -eq 'short-cleaning' -and ($failure -or $parks -or -not $script:Admitted)) { throw 'A completed GitHub job with guest cleanup still active must remain undisturbed.' }
    if ($mode -eq 'short-changed-job' -and (-not $failure -or $parks -ne 1 -or -not $script:Admitted)) {
        throw 'A changed job identity after admission must not be accepted as the admitted job completing.'
    }
}
Write-Output 'Actual hosted control preserves active jobs, refuses stale repeated wakes and parks failed admission.'
