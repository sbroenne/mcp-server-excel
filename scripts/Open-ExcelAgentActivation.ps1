<#
.SYNOPSIS
Starts the qualified desktop for one-time user activation through free Bastion.
.DESCRIPTION
Does not activate Excel or change coding-agent routing. Leaves the VM running
for the user's sign-in, with a one-hour shutdown backstop. An explicit switch
can replace the local clipboard with the generated Windows login password;
the password is never printed. Do not put Office licence keys in chat.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ResourceGroup = 'rg-excel-copilot-runner',
    [string]$VmName = 'vm-excel-copilot-runner',
    [string]$SubscriptionId,
    [switch]$CopyPasswordToClipboard
)

$ErrorActionPreference = 'Stop'
if (-not $PSCmdlet.ShouldProcess($VmName, 'Open the private desktop for one-time user activation')) { return }
$deadline = [DateTime]::UtcNow.AddMinutes(25)
$subscriptionArguments = if ($SubscriptionId) { @('--subscription', $SubscriptionId) } else { @() }
. (Join-Path $PSScriptRoot 'AzureRunnerHost.ps1')
$vm = Invoke-RunnerAzure @('vm', 'get-instance-view', '--resource-group', $ResourceGroup, '--name', $VmName)
if ($vm.tags.project -ne 'mcp-server-excel' -or $vm.tags.purpose -ne 'excel-cloud-agent-devtest') {
    throw 'The VM is not owned by this Excel deployment.'
}
$states = @($vm.instanceView.statuses | Where-Object { $_.code -like 'PowerState/*' })
if ($states.Count -ne 1 -or $states[0].code -ne 'PowerState/deallocated') {
    throw 'Activation access must start from a deallocated VM.'
}
$vaults = @(Invoke-RunnerAzure @('keyvault', 'list', '--resource-group', $ResourceGroup))
$owned = @($vaults | Where-Object { $_.tags.purpose -eq 'excel-cloud-agent-devtest' })
if ($owned.Count -ne 1) { throw 'Exactly one owned credential vault is required.' }
$vaultName = $owned[0].name
$vault = Invoke-RunnerAzure @('keyvault', 'show', '--name', $vaultName)
Assert-RunnerIdentityWithoutAzureAccess -Vm $vm -VaultPolicies @($vault.properties.accessPolicies)
$bastion = Invoke-RunnerAzure @(
    'network', 'bastion', 'show', '--resource-group', $ResourceGroup, '--name', 'excel-activation-desktop'
)
if ($bastion.sku.name -ne 'Developer' -or $bastion.provisioningState -ne 'Succeeded' -or
    $bastion.tags.purpose -ne 'excel-cloud-agent-devtest') {
    throw 'A successfully provisioned, owned free Developer Bastion is required.'
}
$null = Invoke-RunnerAzure @(
    'vm', 'auto-shutdown', '--resource-group', $ResourceGroup, '--name', $VmName,
    '--time', [DateTime]::UtcNow.AddHours(1).ToString('HHmm')
)
$path = Join-Path ([IO.Path]::GetTempPath()) "excel-activation-$([Guid]::NewGuid().ToString('N')).ps1"
$successful = $false
$failures = [Collections.Generic.List[Exception]]::new()
try {
    $null = Invoke-RunnerAzure @('vm', 'start', '--resource-group', $ResourceGroup, '--name', $VmName) 900
    Start-Sleep -Seconds 60
    @'
$ErrorActionPreference = 'Stop'
$guest = 'C:\ProgramData\ExcelMcp\Provisioning\setup-excel-desktop.ps1'
& $guest -Action Health
& $guest -Action Activation
'@ | Set-Content -LiteralPath $path -Encoding UTF8
    $response = Invoke-RunnerAzure @(
        'vm', 'run-command', 'invoke', '--resource-group', $ResourceGroup,
        '--name', $VmName, '--command-id', 'RunPowerShellScript', '--scripts', "@$path"
    )
    $results = @(
        foreach ($value in $response.value) {
            foreach ($line in ($value.message -split "`r?`n")) {
                if ($line.StartsWith('EXCELMCP_DESKTOP=')) { $line.Substring(17) | ConvertFrom-Json }
            }
        }
    )
    if ($results.Count -ne 2 -or $results[0].state -ne 'ready' -or
        $results[1].state -ne 'activation-access-enabled' -or
        $results[0].runnerRegistered -ne $false -or $results[1].runnerRegistered -ne $false) {
        throw 'The private desktop did not qualify for activation access.'
    }
    if ($CopyPasswordToClipboard) {
        $secret = Invoke-RunnerAzure @(
            'keyvault', 'secret', 'show', '--vault-name', $vaultName, '--name', 'runner-account-password'
        )
        if (-not $secret.value) { throw 'The Windows login password is missing.' }
        Set-Clipboard -Value $secret.value
        $secret = $null
        Write-Output 'Windows login password copied to the local clipboard; it was not printed. Clear it after signing in.'
    }
    $successful = $true
}
catch { $failures.Add($_.Exception) }
finally {
    $secret = $null
    try { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force } }
    catch { $failures.Add($_.Exception) }
    if (-not $successful) {
        $deadline = [DateTime]::UtcNow.AddMinutes(10)
        try { $null = Invoke-RunnerAzure @('vm', 'deallocate', '--resource-group', $ResourceGroup, '--name', $VmName) 600 }
        catch { $failures.Add($_.Exception) }
    }
}
if ($failures.Count -gt 0) { throw [AggregateException]::new('Activation preparation failed.', $failures) }
Write-Output 'Private desktop ready for one-time activation. Windows username: excelrunner.'
Write-Output ("PortalUrl=https://portal.azure.com/#resource" + $vm.id)
Write-Output 'Automatic shutdown was set one hour after this command began. Excel activation remains unverified.'
