$ErrorActionPreference = 'Stop'
$guest = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'infrastructure\azure\setup-excel-desktop.ps1'
. $guest

$global:ExcelDesktopTestEnabled = $true
$global:ExcelDesktopTestAdmin = $false
function Get-LocalUser {
    return @{ Enabled = $global:ExcelDesktopTestEnabled; SID = @{ Value = 'synthetic-runner-sid' } }
}
function Get-LocalGroupMember {
    if ($global:ExcelDesktopTestAdmin) { return @{ SID = @{ Value = 'synthetic-runner-sid' } } }
    return @()
}

$null = Assert-NonAdminDesktopAccount -Name 'synthetic-runner'
foreach ($case in @('admin', 'disabled')) {
    $global:ExcelDesktopTestAdmin = $case -eq 'admin'
    $global:ExcelDesktopTestEnabled = $case -ne 'disabled'
    $failure = $null
    try { $null = Assert-NonAdminDesktopAccount -Name 'synthetic-runner' }
    catch { $failure = $_.Exception.Message }
    if (-not $failure) { throw "$case desktop account must be rejected." }
}

$boot = [DateTime]::UtcNow
Assert-DesktopProfileRun -TaskInfo @{ LastRunTime = $boot.AddSeconds(1); LastTaskResult = 0 } -BootTime $boot
foreach ($info in @(
    @{ LastRunTime = $boot.AddSeconds(-1); LastTaskResult = 0 },
    @{ LastRunTime = $boot.AddSeconds(1); LastTaskResult = 1 }
)) {
    $failure = $null
    try { Assert-DesktopProfileRun -TaskInfo $info -BootTime $boot }
    catch { $failure = $_.Exception.Message }
    if (-not $failure) { throw 'Stale or failed profile initialization must be rejected.' }
}
$global:ExcelDesktopTestEnabled = $true
$global:ExcelDesktopTestAdmin = $false
$global:ExcelDesktopTestActivationCalls = [Collections.Generic.List[string]]::new()
function Get-Process { return @() }
function Add-LocalGroupMember {
    param($SID, $Member)
    $global:ExcelDesktopTestActivationCalls.Add("group:$SID")
}
function Set-ItemProperty {
    param($LiteralPath, $Name, $Value)
    $global:ExcelDesktopTestActivationCalls.Add("registry:${Name}:$Value")
}
function Enable-NetFirewallRule {
    param($Name)
    $global:ExcelDesktopTestActivationCalls.Add("firewall:$Name")
}
Enable-PrivateActivationDesktop -Name 'synthetic-runner'
foreach ($expected in @(
    'group:S-1-5-32-555',
    'registry:fDenyTSConnections:0',
    'firewall:RemoteDesktop-UserMode-In-TCP'
)) {
    if (-not $global:ExcelDesktopTestActivationCalls.Contains($expected)) {
        throw "Activation access did not configure the expected narrow permission: $expected"
    }
}
Write-Output 'Desktop account, current-boot initialization and activation access tests passed.'
