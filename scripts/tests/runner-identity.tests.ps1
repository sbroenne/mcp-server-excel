$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'AzureRunnerHost.ps1')
$global:ExcelRunnerIdentityCalls = [Collections.Generic.List[string]]::new()
$global:ExcelRunnerIdentityGuestCalls = 0
$global:ExcelRunnerIdentityGuestFailure = $false
function Start-Job {
    param($ScriptBlock, $ArgumentList)
    $command = $ArgumentList[0] -join ' '
    $global:ExcelRunnerIdentityCalls.Add($command)
    return [pscustomobject]@{ State = 'Completed'; Command = $command }
}
function Wait-Job { param($Job, $Timeout) return $Job }
function Remove-Job { param($Job, [switch]$Force) }
function Start-Sleep { param($Seconds) }
function Receive-Job {
    param($Job)
    switch -Wildcard ($Job.Command) {
        'vm get-instance-view *' {
            return '{"identity":{"type":"SystemAssigned","principalId":"synthetic-principal"},"tags":{"project":"mcp-server-excel","purpose":"excel-cloud-agent-devtest"},"instanceView":{"statuses":[{"code":"PowerState/deallocated"}]}}'
        }
        'keyvault list *' { return '[{"name":"synthetic-vault","tags":{"purpose":"excel-cloud-agent-devtest"}}]' }
        'keyvault show *' { return '{"properties":{"accessPolicies":[]}}' }
        'keyvault secret list *' { return '[{"name":"runner-account-password"}]' }
        'role assignment list *' { return '[]' }
        'keyvault set-policy *' { return '{}' }
        'keyvault delete-policy *' { return '{}' }
        'vm auto-shutdown *' { return '{}' }
        'vm start *' { return '{}' }
        'vm restart *' { return '{}' }
        'vm deallocate *' { return '{}' }
        'vm run-command invoke *' {
            if ($global:ExcelRunnerIdentityGuestFailure) {
                return '{"value":[{"message":"Synthetic desktop configuration failure context."}]}'
            }
            $global:ExcelRunnerIdentityGuestCalls++
            $state = if ($global:ExcelRunnerIdentityGuestCalls -eq 1) { 'configured' } else { 'ready' }
            return (@{ value = @(@{
                message = 'EXCELMCP_DESKTOP=' + (@{ state = $state; runnerRegistered = $false } | ConvertTo-Json -Compress)
            }) } | ConvertTo-Json -Depth 5 -Compress)
        }
        default { throw "Unexpected mocked Azure command: $($Job.Command)" }
    }
}
& (Join-Path (Split-Path -Parent $PSScriptRoot) 'Initialize-ExcelAgentDesktop.ps1') -Confirm:$false
if ($global:ExcelRunnerIdentityGuestCalls -ne 2 -or
    @($global:ExcelRunnerIdentityCalls | Where-Object { $_ -like 'vm identity *' }).Count -ne 0 -or
    @($global:ExcelRunnerIdentityCalls | Where-Object { $_ -like 'keyvault delete-policy *' }).Count -ne 1) {
    throw 'Desktop setup must preserve existing identity and remove only its temporary vault grant.'
}
$global:ExcelRunnerIdentityGuestFailure = $true
$failure = $null
try { & (Join-Path (Split-Path -Parent $PSScriptRoot) 'Initialize-ExcelAgentDesktop.ps1') -Confirm:$false }
catch { $failure = $_.Exception }
if (-not $failure -or $failure.ToString() -notmatch 'Synthetic desktop configuration failure context') {
    throw 'Desktop orchestration must preserve guest failure context.'
}
$global:ExcelRunnerIdentityRoles = @()
function Invoke-RunnerAzure {
    param([string[]]$Arguments)
    if (($Arguments[0..2] -join ' ') -ne 'role assignment list' -or
        $Arguments -notcontains '--assignee-object-id' -or $Arguments -notcontains '--all') {
        throw 'Identity checks must inspect all direct Azure role assignments without a directory lookup.'
    }
    return $global:ExcelRunnerIdentityRoles
}
$vm = @{ identity = @{ type = 'SystemAssigned'; principalId = 'synthetic-principal' } }
Assert-RunnerIdentityWithoutAzureAccess -Vm $vm -VaultPolicies @()
Assert-RunnerIdentityWithoutAzureAccess -Vm @{ identity = $null } -VaultPolicies @()
foreach ($case in @('role', 'vault', 'user-assigned', 'missing-principal')) {
    $candidate = @{ identity = @{ type = 'SystemAssigned'; principalId = 'synthetic-principal' } }
    $policies = @()
    $global:ExcelRunnerIdentityRoles = @()
    switch ($case) {
        'role' { $global:ExcelRunnerIdentityRoles = @(@{ roleDefinitionName = 'Reader' }) }
        'vault' { $policies = @(@{ objectId = 'synthetic-principal' }) }
        'user-assigned' { $candidate.identity.type = 'UserAssigned' }
        'missing-principal' { $candidate.identity.principalId = $null }
    }
    $failure = $null
    try { Assert-RunnerIdentityWithoutAzureAccess -Vm $candidate -VaultPolicies $policies }
    catch { $failure = $_.Exception }
    if (-not $failure) { throw "Unsafe identity case was accepted: $case" }
}
Write-Output 'Existing identity preservation and Azure-access rejection tests passed.'
