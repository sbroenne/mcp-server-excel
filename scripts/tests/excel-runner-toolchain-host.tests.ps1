$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Install-ExcelAgentToolchain.ps1'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$global:ExcelToolchainHostSdk = (Get-Content (Join-Path $root 'global.json') -Raw | ConvertFrom-Json).sdk.version
$global:ExcelToolchainHostCalls = [Collections.Generic.List[string]]::new()
$global:ExcelToolchainHostMode = 'ready'
$global:ExcelToolchainHostCleanupFails = $false

function Start-Job {
    param($ScriptBlock, $ArgumentList)
    $command = $ArgumentList[0] -join ' '
    $global:ExcelToolchainHostCalls.Add($command)
    [pscustomobject]@{ State = 'Completed'; Command = $command }
}
function Wait-Job { param($Job, $Timeout) $Job }
function Remove-Job { param($Job, [switch]$Force) }
function Start-Sleep { param($Seconds) }
function Receive-Job {
    param($Job)
    switch -Wildcard ($Job.Command) {
        'vm get-instance-view *' {
            $project = if ($global:ExcelToolchainHostMode -eq 'unowned') { 'unrelated' } else { 'mcp-server-excel' }
            $power = if ($global:ExcelToolchainHostMode -eq 'running') { 'running' } else { 'deallocated' }
            return (@{
                tags = @{ project = $project; purpose = 'excel-cloud-agent-devtest' }
                instanceView = @{ statuses = @(@{ code = "PowerState/$power" }) }
            } | ConvertTo-Json -Depth 5 -Compress)
        }
        'keyvault list *' { return '[{"name":"synthetic-vault","tags":{"purpose":"excel-cloud-agent-devtest"}}]' }
        'keyvault show *' { return '{"properties":{"accessPolicies":[]}}' }
        'vm auto-shutdown *' { return '{}' }
        'vm start *' { return '{}' }
        'vm run-command invoke *' {
            $sdk = if ($global:ExcelToolchainHostMode -eq 'sdk-mismatch') { '1.0.0' } else { $global:ExcelToolchainHostSdk }
            $prefix = if ($global:ExcelToolchainHostMode -eq 'spoofed') { 'unexpected=' } else { 'EXCELMCP_TOOLCHAIN=' }
            $tools = @{ sdk = $sdk; requiredSdk = $global:ExcelToolchainHostSdk; rollForward = 'latestFeature'
                git = 'git version 2.56.0.windows.1'; powershell = '7.6.6'; node = 'v22.23.3' }
            if ($global:ExcelToolchainHostMode -eq 'missing-tools') { $tools.Remove('git') }
            if ($global:ExcelToolchainHostMode -eq 'wrong-node') { $tools.node = 'v24.0.0' }
            $message = $prefix + (@{ state = 'installed'; tools = $tools; runnerRegistered = $false } | ConvertTo-Json -Compress)
            if ($global:ExcelToolchainHostMode -eq 'guest-failure') { $message = 'Synthetic guest upload failure.' }
            return (@{ value = @(@{ message = $message }) } | ConvertTo-Json -Depth 5 -Compress)
        }
        'vm deallocate *' {
            if ($global:ExcelToolchainHostCleanupFails) { throw 'Synthetic toolchain deallocation failure.' }
            return '{}'
        }
        default { throw "Unexpected Azure command: $($Job.Command)" }
    }
}

& $scriptPath -WhatIf
if ($global:ExcelToolchainHostCalls.Count -ne 0) { throw 'WhatIf must not contact Azure.' }
foreach ($mode in @('unowned', 'running')) {
    $global:ExcelToolchainHostCalls.Clear()
    $global:ExcelToolchainHostMode = $mode
    $failed = $false
    try { & $scriptPath -Confirm:$false }
    catch { $failed = $true }
    if (-not $failed -or $global:ExcelToolchainHostCalls.Count -ne 1) {
        throw 'A running or unowned VM must be rejected without changing resources.'
    }
}
foreach ($mode in @('sdk-mismatch', 'missing-tools', 'wrong-node', 'spoofed', 'guest-failure')) {
    $global:ExcelToolchainHostCalls.Clear()
    $global:ExcelToolchainHostMode = $mode
    $global:ExcelToolchainHostCleanupFails = $mode -eq 'guest-failure'
    $failure = $null
    try { & $scriptPath -Confirm:$false }
    catch { $failure = $_.Exception }
    if ($failure -isnot [AggregateException] -or
        @($global:ExcelToolchainHostCalls | Where-Object { $_ -like 'vm deallocate *' }).Count -ne 1) {
        throw 'Invalid results must fail and attempt owned deallocation.'
    }
    if ($mode -eq 'guest-failure' -and
        ($failure.ToString() -notmatch 'Synthetic guest upload failure' -or
         $failure.ToString() -notmatch 'Synthetic toolchain deallocation failure')) {
        throw 'The original guest failure and cleanup failure must both be preserved.'
    }
}
$global:ExcelToolchainHostCalls.Clear()
$global:ExcelToolchainHostMode = 'ready'
$global:ExcelToolchainHostCleanupFails = $false
& $scriptPath -Confirm:$false
if (@($global:ExcelToolchainHostCalls | Where-Object { $_ -like 'keyvault secret *' -or $_ -like 'role assignment create *' }).Count) {
    throw 'Toolchain setup must not retrieve passwords or grant Azure access.'
}
Write-Output 'Toolchain host ownership, SDK, result and cleanup tests passed.'
