$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'ExcelRunnerHost.ps1')
foreach ($name in @('Invoke-ExcelRunnerDesktopHealth', 'Invoke-ExcelRunnerJobRecovery', 'Stop-ExcelRunnerExpiredListener')) {
    if (-not (Get-Command $name -CommandType Function -ErrorAction SilentlyContinue)) {
        throw "Hosted controller helper is not callable: $name"
    }
}
$script:GuestScripts = [Collections.Generic.List[string]]::new()
function Start-Sleep { param($Seconds) }
function Invoke-ExcelRunnerGuest {
    param([string]$Script, $Marker, $OperationId)
    $script:GuestScripts.Add($Script)
    $parseErrors = $null
    $null = [Management.Automation.Language.Parser]::ParseInput($Script, [ref]$null, [ref]$parseErrors)
    if ($parseErrors.Count) { throw "Generated guest script is invalid: $parseErrors" }
    if ($Script -match 'ExcelMcp-Test-Desktop') {
        return @{
            state = 'ready'; operationId = $OperationId; license = 'Licensed'
            formula = 42; persistedFormula = 42
            bootTime = [DateTime]::UtcNow.AddMinutes(-5).ToString('o')
            checkedAt = [DateTime]::UtcNow.ToString('o')
        }
    }
    if ($Script -match 'ExcelMcp-Recover-Jobs') { return @{ state = 'recovered'; operationId = $OperationId } }
    return @{ state = 'drained' }
}
$null = Invoke-ExcelRunnerDesktopHealth
Invoke-ExcelRunnerJobRecovery
Stop-ExcelRunnerExpiredListener @()
if ($script:GuestScripts.Count -ne 5) { throw 'Desktop qualification, recovery and idle drain did not use the expected guest stages.' }
if (@($script:GuestScripts | Where-Object { $_ -match 'function Invoke-ExcelRunnerJobRecovery|function Stop-ExcelRunnerExpiredListener' }).Count) {
    throw 'Host functions must not be embedded in a guest health command.'
}
$script:GuestScripts.Clear()
Stop-ExcelRunnerExpiredListener @(@{ labels = @('excel-copilot'); status = 'in_progress' })
if ($script:GuestScripts.Count) { throw 'An active whole job must prevent listener draining.' }
Write-Output 'Callable host helpers, valid generated guest commands and whole-job drain protection passed.'
