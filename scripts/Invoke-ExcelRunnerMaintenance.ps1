<#
.SYNOPSIS
Patches the dedicated VM and proves recovery, following mcp-windows maintenance.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ResourceGroup = 'rg-excel-copilot-runner',
    [string]$VmName = 'vm-excel-copilot-runner',
    [string]$Repository = 'sbroenne/mcp-server-excel',
    [string]$SubscriptionId,
    [string]$ReportPath
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ExcelRunnerHost.ps1')

function Invoke-ExcelRunnerUpdatePass {
    param([ValidateSet('Windows', 'Office')][string]$Component)
    $pass = [Guid]::NewGuid().ToString('N')
    $worker = 'C:\ProgramData\ExcelMcp\Maintenance\update-excel-runner.ps1'
    $state = Invoke-ExcelRunnerGuest "& '$worker' -Action Start -Component $Component -PassId $pass" 'EXCELMCP_UPDATE=' $pass
    if ($state.state -ne 'running') { throw 'Update task did not start.' }
    $until = [DateTime]::UtcNow.AddMinutes(65)
    do {
        Start-Sleep -Seconds 30
        $state = Invoke-ExcelRunnerGuest "& '$worker' -Action Status -Component $Component -PassId $pass" 'EXCELMCP_UPDATE=' $pass
        if ($state.state -eq 'complete') { return $state }
        if ($state.state -ne 'running') { throw "The $Component update pass failed: $($state.error)" }
    } while ([DateTime]::UtcNow -lt $until)
    throw "The $Component update pass exceeded its hard deadline."
}

function Invoke-ExcelRunnerUpdateCycle {
    param([hashtable]$Report, [int]$MaximumPasses = 4)
    $restarted = $false
    $clean = $false
    for ($pass = 0; $pass -lt $MaximumPasses; $pass++) {
        $state = Invoke-ExcelRunnerUpdatePass Windows
        $Report.passes.Add($state)
        if ($state.rebootRequired -or -not $restarted) {
            $null = Invoke-RunnerAzure @('vm', 'restart', '--resource-group', $ResourceGroup, '--name', $VmName) 900
            Wait-ExcelRunnerGuestAgent
            $restarted = $true
        }
        elseif ($state.installed -eq 0) { $clean = $true; break }
    }
    if (-not $clean) { throw "Windows did not reach a clean post-restart scan in $MaximumPasses passes." }
    $office = Invoke-ExcelRunnerUpdatePass Office
    $Report.passes.Add($office)
    $null = Invoke-RunnerAzure @('vm', 'restart', '--resource-group', $ResourceGroup, '--name', $VmName) 900
    Wait-ExcelRunnerGuestAgent
    $Report.desktop = Invoke-ExcelRunnerDesktopHealth
    $stamp = [DateTime]::UtcNow.ToString('o')
    $state = Invoke-ExcelRunnerGuest @"
. 'C:\ProgramData\ExcelMcp\Maintenance\update-excel-runner.ps1'
if (Test-ExcelRunnerRestart) { throw 'A restart is still pending after recovery.' }
`$state = @{ state = 'patched'; checkedAt = '$stamp'; officeVersion = '$($office.officeVersion)'; officeTarget = '$($office.officeTarget)'; pendingRestart = `$false }
. 'C:\ProgramData\ExcelMcp\Maintenance\ExcelRunnerPolicy.ps1'
Assert-ExcelRunnerPatchState `$state
`$state | ConvertTo-Json -Compress | Set-Content 'C:\ProgramData\ExcelMcp\Maintenance\patched.json.tmp' -Encoding UTF8
`$target = 'C:\ProgramData\ExcelMcp\Maintenance\patched.json'
if (Test-Path -LiteralPath `$target) { [IO.File]::Replace("`$target.tmp", `$target, [NullString]::Value) }
else { [IO.File]::Move("`$target.tmp", `$target) }
Write-Output ('EXCELMCP_CONTROL=' + (`$state | ConvertTo-Json -Compress))
"@
    if ($state.state -ne 'patched' -or $state.pendingRestart -ne $false) { throw 'Patch qualification was not persisted.' }
    $Report.state = 'patched'
    $Report.checkedAt = $stamp
    $Report.officeVersion = $office.officeVersion
    $Report.officeTarget = $office.officeTarget
}

if ($MyInvocation.InvocationName -eq '.') { return }
if (-not $PSCmdlet.ShouldProcess($VmName, 'Install Windows/Excel updates, restart, verify Excel and deallocate')) { return }
if ($Repository -notmatch '^[a-zA-Z0-9_.-]+/[a-zA-Z0-9_.-]+$') { throw 'Invalid repository.' }
$deadline = [DateTime]::UtcNow.AddMinutes(175)
$subscriptionArguments = if ($SubscriptionId) { @('--subscription', $SubscriptionId) } else { @() }
if (-not $ReportPath) { $ReportPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'TestResults\excel-runner-maintenance\report.json' }
$report = @{ state = 'failed'; passes = [Collections.Generic.List[object]]::new() }
$vm = Assert-ExcelRunnerOwnedVm
$power = @($vm.instanceView.statuses | Where-Object code -Like 'PowerState/*')
if ($power.Count -ne 1 -or $power[0].code -ne 'PowerState/deallocated') {
    throw 'Maintenance starts only from a parked VM; never interrupt a running coding session.'
}
$jobs = @(Get-ExcelRunnerActiveJobs $Repository)
Assert-ExcelRunnerIdle $jobs @{ listeners = 0; workers = 0; excel = 0; taskRunning = $false } -AllowQueued
$failures = [Collections.Generic.List[Exception]]::new()
$started = $false
try {
    $null = Invoke-RunnerAzure @(
        'vm', 'auto-shutdown', '--resource-group', $ResourceGroup, '--name', $VmName,
        '--time', [DateTime]::UtcNow.AddHours(4).ToString('HHmm')
    )
    $started = $true
    $null = Invoke-RunnerAzure @('vm', 'start', '--resource-group', $ResourceGroup, '--name', $VmName) 900
    Wait-ExcelRunnerGuestAgent
    Assert-ExcelRunnerIdle @(Get-ExcelRunnerActiveJobs $Repository) (Get-ExcelRunnerGuestActivity) -AllowQueued
    Send-ExcelRunnerFiles Maintenance @('update-excel-runner.ps1', 'install-excel-office.ps1', 'ExcelRunnerPolicy.ps1')
    Send-ExcelRunnerFiles Desktop @('test-excel-desktop.ps1', 'ExcelRunnerPolicy.ps1', 'global.json')
    $null = Invoke-ExcelRunnerGuest @'
$state = @{ state = 'quarantined'; reason = 'maintenance-in-progress'; checkedAt = [DateTime]::UtcNow.ToString('o') }
$state | ConvertTo-Json -Compress | Set-Content 'C:\ProgramData\ExcelMcp\Maintenance\patched.json' -Encoding UTF8
Write-Output 'EXCELMCP_CONTROL={"state":"quarantined"}'
'@
    Invoke-ExcelRunnerUpdateCycle $report
}
catch { $failures.Add($_.Exception); $report.error = $_.Exception.Message }
finally {
    $deadline = [DateTime]::UtcNow.AddMinutes(15)
    if ($started) {
        try {
            Assert-ExcelRunnerIdle @(Get-ExcelRunnerActiveJobs $Repository) (Get-ExcelRunnerGuestActivity) -AllowQueued
            $null = Invoke-RunnerAzure @('vm', 'deallocate', '--resource-group', $ResourceGroup, '--name', $VmName) 600
            $report.deallocated = $true
        }
        catch { $failures.Add($_.Exception); $report.deallocated = $false; $report.cleanupError = $_.Exception.Message }
    }
    if ($failures.Count) { $report.state = 'failed' }
    $parent = Split-Path -Parent ([IO.Path]::GetFullPath($ReportPath))
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
}
if ($failures.Count) { throw [AggregateException]::new('Excel runner maintenance failed; keep coding admission disabled.', $failures) }
Write-Output 'Windows and Excel update checks and non-admin desktop recovery passed; VM deallocated.'
