$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-ExcelRunnerMaintenance.ps1')
$script:MaintenanceCalls = [Collections.Generic.List[string]]::new()
$script:MaintenancePass = 0
$script:MaintenanceMode = 'ready'
$ResourceGroup = 'synthetic'
$VmName = 'synthetic'
$Repository = 'synthetic/repository'
$deadline = [DateTime]::UtcNow.AddMinutes(10)
function Start-Sleep { param($Seconds) }
function Wait-ExcelRunnerGuestAgent { $script:MaintenanceCalls.Add('wait-agent') }
function Invoke-RunnerAzure {
    param($Arguments, $TimeoutSeconds)
    $script:MaintenanceCalls.Add(($Arguments -join ' '))
}
function Invoke-ExcelRunnerGuest {
    param($Script, $Marker, $OperationId)
    $script:MaintenanceCalls.Add($Script)
    if ($Script -match '-Action Start -Component Windows') {
        $script:MaintenancePass++
        return @{ state = 'running'; passId = $OperationId }
    }
    if ($Script -match '-Action Status -Component Windows') {
        return @{
            state = $(if ($script:MaintenanceMode -eq 'partial') { 'failed' } else { 'complete' })
            passId = $OperationId; installed = $(if ($script:MaintenancePass -eq 1) { 1 } else { 0 })
            rebootRequired = $script:MaintenanceMode -eq 'never-clean'
            error = 'synthetic partial installation'
        }
    }
    if ($Script -match '-Action Start -Component Office') { return @{ state = 'running'; passId = $OperationId } }
    if ($Script -match '-Action Status -Component Office') {
        return @{ state = 'complete'; passId = $OperationId; officeVersion = '16.0.20000.20000'; officeTarget = '16.0.20000.20000'; rebootRequired = $false }
    }
    return @{ state = 'patched'; pendingRestart = $false }
}
function Invoke-ExcelRunnerDesktopHealth {
    $script:MaintenanceCalls.Add('desktop-health')
    if ($script:MaintenanceMode -eq 'health-failure') { throw 'synthetic licensed desktop failure' }
    return @{ state = 'ready' }
}
foreach ($mode in @('partial', 'never-clean', 'health-failure', 'ready')) {
    $script:MaintenanceCalls.Clear()
    $script:MaintenancePass = 0
    $script:MaintenanceMode = $mode
    $report = @{ passes = [Collections.Generic.List[object]]::new() }
    $failure = $null
    try { Invoke-ExcelRunnerUpdateCycle -Report $report } catch { $failure = $_.Exception }
    if ($mode -eq 'ready') {
        if ($failure -or $report.state -ne 'patched' -or $report.passes.Count -ne 3 -or
            @($script:MaintenanceCalls | Where-Object { $_ -like 'vm restart *' }).Count -ne 2 -or
            $script:MaintenanceCalls[-2] -ne 'desktop-health') {
            throw 'Maintenance must clean-scan Windows, separately update Excel, restart and establish real desktop readiness.'
        }
    }
    elseif (-not $failure -or $report.state -eq 'patched') { throw 'Partial updates or failed recovery must quarantine, not report success.' }
}
Write-Output 'Bounded multi-pass Windows/Office updates and failed-recovery propagation passed.'
