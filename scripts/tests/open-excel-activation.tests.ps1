$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Open-ExcelAgentActivation.ps1'
$global:ExcelActivationTestCalls = [Collections.Generic.List[string]]::new()
$global:ExcelActivationTestGuestMode = 'failure'
$global:ExcelActivationTestCleanupFails = $true

function Start-Job {
    param($ScriptBlock, $ArgumentList)
    $command = $ArgumentList[0] -join ' '
    $global:ExcelActivationTestCalls.Add($command)
    return [pscustomobject]@{ State = 'Completed'; Command = $command }
}
function Wait-Job { param($Job, $Timeout) return $Job }
function Remove-Job { param($Job, [switch]$Force) }
function Start-Sleep { param($Seconds) }
function Receive-Job {
    param($Job)
    switch -Wildcard ($Job.Command) {
        'vm get-instance-view *' {
            return '{"id":"/synthetic-vm","tags":{"project":"mcp-server-excel","purpose":"excel-cloud-agent-devtest"},"instanceView":{"statuses":[{"code":"PowerState/deallocated"}]}}'
        }
        'network bastion show *' {
            return '{"sku":{"name":"Developer"},"provisioningState":"Succeeded","tags":{"purpose":"excel-cloud-agent-devtest"}}'
        }
        'keyvault list *' { return '[{"name":"synthetic-vault","tags":{"purpose":"excel-cloud-agent-devtest"}}]' }
        'keyvault show *' { return '{"properties":{"accessPolicies":[]}}' }
        'vm auto-shutdown *' { return '{}' }
        'vm start *' { return '{}' }
        'vm run-command invoke *' {
            if ($global:ExcelActivationTestGuestMode -eq 'failure') { throw 'Synthetic activation probe failure.' }
            $messages = @(
                'EXCELMCP_DESKTOP={"state":"ready","runnerRegistered":false}',
                'EXCELMCP_DESKTOP={"state":"activation-access-enabled","runnerRegistered":false}'
            )
            if ($global:ExcelActivationTestGuestMode -eq 'spoofed') {
                $messages = $messages -replace '^EXCELMCP_DESKTOP=', 'unexpected-output='
            }
            if ($global:ExcelActivationTestGuestMode -eq 'registered') {
                $messages = $messages -replace '"runnerRegistered":false', '"runnerRegistered":true'
            }
            return (@{ value = @(@{ message = $messages -join "`n" }) } | ConvertTo-Json -Depth 5 -Compress)
        }
        'vm deallocate *' {
            if ($global:ExcelActivationTestCleanupFails) { throw 'Synthetic deallocation failure.' }
            return '{}'
        }
        default { throw "Unexpected mocked Azure command: $($Job.Command)" }
    }
}

& $scriptPath -WhatIf
if ($global:ExcelActivationTestCalls.Count -ne 0) { throw 'WhatIf must not contact Azure.' }
$failure = $null
try { & $scriptPath -Confirm:$false }
catch { $failure = $_.Exception }
if ($failure -isnot [AggregateException] -or
    $failure.ToString() -notmatch 'activation probe failure' -or
    $failure.ToString() -notmatch 'deallocation failure') {
    throw 'Activation must preserve both primary and cleanup errors.'
}
if (@($global:ExcelActivationTestCalls | Where-Object { $_ -like 'vm deallocate *' }).Count -ne 1) {
    throw 'Failed activation preparation must attempt owned VM deallocation.'
}
$global:ExcelActivationTestCleanupFails = $false
foreach ($mode in @('spoofed', 'registered')) {
    $global:ExcelActivationTestCalls.Clear()
    $global:ExcelActivationTestGuestMode = $mode
    $failure = $null
    try { & $scriptPath -Confirm:$false }
    catch { $failure = $_.Exception }
    if (-not $failure -or
        @($global:ExcelActivationTestCalls | Where-Object { $_ -like 'vm deallocate *' }).Count -ne 1) {
        throw "Invalid activation results must fail and deallocate: $mode"
    }
}
$global:ExcelActivationTestCalls.Clear()
$global:ExcelActivationTestGuestMode = 'ready'
& $scriptPath -Confirm:$false
if (@($global:ExcelActivationTestCalls | Where-Object { $_ -like 'vm deallocate *' -or $_ -like 'keyvault secret show *' }).Count -ne 0) {
    throw 'Qualified activation must leave the desktop running without retrieving a password unless requested.'
}
Write-Output 'Activation no-change, result qualification and primary/cleanup error tests passed.'
