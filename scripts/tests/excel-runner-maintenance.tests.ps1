$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'infrastructure\azure\update-excel-runner.ps1')
function Expect-Failure {
    param([scriptblock]$Operation)
    $failed = $false
    try { & $Operation } catch { $failed = $true }
    if (-not $failed) { throw 'Unsafe maintenance input was accepted.' }
}
$security = '0fa1201d-4330-4fa8-8ae9-b877473b6441'
$driver = '3689bdc8-b205-4af4-8d4a-a63924c5e9d5'
foreach ($case in @(
    @{ Type = 1; BrowseOnly = $false; Categories = @(@{ CategoryID = $security }); expected = $true },
    @{ Type = 2; BrowseOnly = $false; Categories = @(@{ CategoryID = $security }); expected = $false },
    @{ Type = 1; BrowseOnly = $true; Categories = @(@{ CategoryID = $security }); expected = $false },
    @{ Type = 1; BrowseOnly = $false; Categories = @(@{ CategoryID = $driver }); expected = $false },
    @{ Type = 1; BrowseOnly = $false; Categories = @(); expected = $false }
)) {
    if ((Test-ExcelRunnerUpdate $case) -ne $case.expected) { throw 'Wrong Windows update selection.' }
}
Assert-ExcelRunnerUpdateResult @{ ResultCode = 2; HResult = 0 } 'Synthetic search'
foreach ($code in @(0, 1, 3, 4, 5)) {
    Expect-Failure { Assert-ExcelRunnerUpdateResult @{ ResultCode = $code; HResult = 42 } 'Synthetic install' }
}
$metadata = @(
    [pscustomobject]@{ channelId = 'Current'; latestVersion = '16.0.20000.20000' },
    [pscustomobject]@{ channelId = 'BetaChannel'; latestVersion = '16.0.21000.20000' }
)
if ((Get-ExcelRunnerOfficeTarget $metadata) -ne '16.0.20000.20000') { throw 'Office must stay on Current.' }
Expect-Failure { Get-ExcelRunnerOfficeTarget @() }
Expect-Failure { Get-ExcelRunnerOfficeTarget @($metadata[0], $metadata[0]) }
Expect-Failure { Get-ExcelRunnerOfficeTarget @([pscustomobject]@{ channelId = 'Current'; latestVersion = 'unsafe' }) }

. (Join-Path (Split-Path -Parent $PSScriptRoot) 'ExcelRunnerPolicy.ps1')
$now = [DateTime]::UtcNow
$healthy = @{ state = 'ready'; checkedAt = $now.ToString('o'); bootTime = $now.AddMinutes(-2).ToString('o'); license = 'Licensed'; formula = 42; persistedFormula = 42 }
Assert-ExcelRunnerReadiness $healthy -Now $now
foreach ($field in @('state', 'checkedAt', 'bootTime', 'license', 'formula', 'persistedFormula')) {
    $candidate = $healthy.Clone()
    $candidate[$field] = 'invalid'
    Expect-Failure { Assert-ExcelRunnerReadiness $candidate -Now $now }
}
Expect-Failure { Assert-ExcelRunnerReadiness $healthy -Now $now.AddDays(9) }
Expect-Failure { Assert-ExcelRunnerReadiness $healthy -Now $now.AddMinutes(-5) }
$windows = @{ state = 'patched'; checkedAt = $now.ToString('o'); officeVersion = '16.0.20000.20000'; officeTarget = '16.0.20000.20000'; pendingRestart = $false }
Assert-ExcelRunnerPatchState $windows -Now $now
$newerRetail = $windows.Clone()
$newerRetail.officeVersion = '16.0.20000.20040'
Assert-ExcelRunnerPatchState $newerRetail -Now $now
$olderRetail = $windows.Clone()
$olderRetail.officeVersion = '16.0.19999.20000'
Expect-Failure { Assert-ExcelRunnerPatchState $olderRetail -Now $now }
foreach ($field in @('state', 'checkedAt', 'officeVersion', 'officeTarget', 'pendingRestart')) {
    $candidate = $windows.Clone()
    $candidate[$field] = 'invalid'
    Expect-Failure { Assert-ExcelRunnerPatchState $candidate -Now $now }
}
Expect-Failure { Assert-ExcelRunnerPatchState $windows -Now $now.AddDays(9) }
$worker = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'infrastructure\azure\update-excel-runner.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($worker, [ref]$null, [ref]$null)
$switch = $ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.SwitchStatementAst] }
if (@($switch).Count -ne 1) { throw 'Expected one maintenance action dispatcher.' }
function Test-Path { param($LiteralPath) return $false }
function Get-ScheduledTask { param($TaskName) return $null }
function Get-Process { param($Name) return @() }
function New-Item { param($Path, [switch]$Force) }
function New-ItemProperty { param($Path, $Name, $Value, $PropertyType, [switch]$Force) }
function New-ScheduledTaskAction { param($Execute, $Argument) return 'synthetic-task-action' }
function New-ScheduledTaskPrincipal { param($UserId, $LogonType, $RunLevel) return 'synthetic-principal' }
function New-ScheduledTaskSettingsSet { param($ExecutionTimeLimit) return 'synthetic-settings' }
function Register-ScheduledTask {
    param($TaskName, $Action, $Principal, $Settings, [switch]$Force)
    if ($Action -ne 'synthetic-task-action') { throw 'The task must receive the action object, not the dispatcher string.' }
}
function Start-ScheduledTask { param($TaskName) }
function Write-ExcelRunnerUpdateState { param($State) }
$resultPath = 'synthetic-result'
$taskName = 'synthetic-task'
$PassId = 'a' * 32
$Component = 'Windows'
& ([scriptblock]::Create("param([ValidateSet('Start','Status','Worker')][string]`$Action)`n" + $switch.Extent.Text)) -Action Start
Write-Output 'Windows/Office update policy, partial failures and readiness freshness tests passed.'
