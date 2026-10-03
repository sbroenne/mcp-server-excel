<#
.SYNOPSIS
Creates the non-admin Excel desktop using mcp-windows automatic-logon setup.
.DESCRIPTION
Uses temporary managed-identity vault access only during setup, removes that grant,
then restarts and checks the real interactive desktop. Leaves the VM deallocated.
Does not register a GitHub runner or assert that Excel is activated.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ResourceGroup = 'rg-excel-copilot-runner',
    [string]$VmName = 'vm-excel-copilot-runner',
    [string]$SubscriptionId
)

$ErrorActionPreference = 'Stop'
if (-not $PSCmdlet.ShouldProcess($VmName, 'Prepare the dedicated non-admin Excel desktop')) { return }
$deadline = [DateTime]::UtcNow.AddMinutes(40)
$subscriptionArguments = if ($SubscriptionId) { @('--subscription', $SubscriptionId) } else { @() }
. (Join-Path $PSScriptRoot 'AzureRunnerHost.ps1')
$source = Join-Path (Split-Path -Parent $PSScriptRoot) 'infrastructure\azure\setup-excel-desktop.ps1'
$guestPath = 'C:\ProgramData\ExcelMcp\Provisioning\setup-excel-desktop.ps1'

function Invoke-DesktopGuest {
    param([string]$Script)
    $path = Join-Path ([IO.Path]::GetTempPath()) "excel-desktop-command-$([Guid]::NewGuid().ToString('N')).ps1"
    try {
        ("`$ErrorActionPreference = 'Stop'`n" + $Script) | Set-Content -LiteralPath $path -Encoding UTF8
        $response = Invoke-RunnerAzure @(
            'vm', 'run-command', 'invoke', '--resource-group', $ResourceGroup,
            '--name', $VmName, '--command-id', 'RunPowerShellScript', '--scripts', "@$path"
        )
        $messages = @(
            foreach ($value in $response.value) {
                foreach ($line in ($value.message -split "`r?`n")) {
                    if ($line.StartsWith('EXCELMCP_DESKTOP=')) { $line.Substring(17) }
                }
            }
        )
        if ($messages.Count -ne 1) {
            $details = @($response.value | ForEach-Object message) -join "`n"
            throw "Desktop setup did not return exactly one result. Guest output: $details"
        }
        return $messages[0] | ConvertFrom-Json
    }
    finally { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force } }
}

$vm = Invoke-RunnerAzure @('vm', 'get-instance-view', '--resource-group', $ResourceGroup, '--name', $VmName)
if ($vm.tags.project -ne 'mcp-server-excel' -or $vm.tags.purpose -ne 'excel-cloud-agent-devtest') {
    throw 'The VM is not owned by this Excel deployment.'
}
$states = @($vm.instanceView.statuses | Where-Object { $_.code -like 'PowerState/*' })
if ($states.Count -ne 1 -or $states[0].code -ne 'PowerState/deallocated') {
    throw 'Desktop setup must start from a deallocated VM.'
}
$vaults = @(Invoke-RunnerAzure @('keyvault', 'list', '--resource-group', $ResourceGroup))
$ownedVaults = @($vaults | Where-Object { $_.tags.purpose -eq 'excel-cloud-agent-devtest' })
if ($ownedVaults.Count -ne 1) { throw 'Exactly one owned password vault is required.' }
$vaultName = $ownedVaults[0].name
$vault = Invoke-RunnerAzure @('keyvault', 'show', '--name', $vaultName)
Assert-RunnerIdentityWithoutAzureAccess -Vm $vm -VaultPolicies @($vault.properties.accessPolicies)
$existingPrincipalId = if ($vm.identity -and $vm.identity.type -eq 'SystemAssigned') {
    $vm.identity.principalId
} else { $null }
$secrets = @(Invoke-RunnerAzure @('keyvault', 'secret', 'list', '--vault-name', $vaultName))
if (@($secrets | Where-Object name -EQ 'runner-account-password').Count -eq 0) {
    $privateDirectory = Join-Path ([IO.Path]::GetTempPath()) "excel-desktop-secret-$([Guid]::NewGuid().ToString('N'))"
    $passwordPath = Join-Path $privateDirectory 'password.txt'
    try {
        New-Item -ItemType Directory -Path $privateDirectory | Out-Null
        $acl = Get-Acl -LiteralPath $privateDirectory
        $acl.SetAccessRuleProtection($true, $false)
        foreach ($sid in @(
            [Security.Principal.WindowsIdentity]::GetCurrent().User,
            [Security.Principal.SecurityIdentifier]::new('S-1-5-18')
        )) {
            $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new(
                $sid, 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow'
            ))
        }
        Set-Acl -LiteralPath $privateDirectory -AclObject $acl
        $password = 'Aa1!' + [Convert]::ToBase64String([Security.Cryptography.RandomNumberGenerator]::GetBytes(32))
        [IO.File]::WriteAllText($passwordPath, $password, [Text.UTF8Encoding]::new($false))
        $null = Invoke-RunnerAzure @(
            'keyvault', 'secret', 'set', '--vault-name', $vaultName,
            '--name', 'runner-account-password', '--file', $passwordPath
        )
    }
    finally {
        $password = $null
        if (Test-Path -LiteralPath $passwordPath) { Remove-Item -LiteralPath $passwordPath -Force }
        if (Test-Path -LiteralPath $privateDirectory) { Remove-Item -LiteralPath $privateDirectory -Force }
    }
}

$null = Invoke-RunnerAzure @(
    'vm', 'auto-shutdown', '--resource-group', $ResourceGroup, '--name', $VmName,
    '--time', [DateTime]::UtcNow.AddHours(1).ToString('HHmm')
)
$failures = [Collections.Generic.List[Exception]]::new()
$identityAttempted = $false
$principalId = $existingPrincipalId
try {
    $null = Invoke-RunnerAzure @('vm', 'start', '--resource-group', $ResourceGroup, '--name', $VmName) 900
    try {
        if (-not $existingPrincipalId) {
            $identityAttempted = $true
            $identity = Invoke-RunnerAzure @('vm', 'identity', 'assign', '--resource-group', $ResourceGroup, '--name', $VmName)
            $principalId = $identity.principalId
        }
        if (-not $principalId) { throw 'The temporary VM identity has no principal ID.' }
        $null = Invoke-RunnerAzure @(
            'keyvault', 'set-policy', '--name', $vaultName, '--object-id', $principalId,
            '--secret-permissions', 'get'
        )
        $encoded = [Convert]::ToBase64String([IO.File]::ReadAllBytes($source))
        $hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
        $result = Invoke-DesktopGuest @"
`$directory = Split-Path '$guestPath'
New-Item `$directory -ItemType Directory -Force | Out-Null
& icacls.exe `$directory /inheritance:r /grant:r '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' | Out-Null
if (`$LASTEXITCODE -ne 0) { throw 'Could not protect provisioning scripts.' }
[IO.File]::WriteAllBytes('$guestPath', [Convert]::FromBase64String('$encoded'))
if ((Get-FileHash -LiteralPath '$guestPath' -Algorithm SHA256).Hash -ne '$hash') {
    throw 'Desktop setup script hash mismatch.'
}
& '$guestPath' -Action Setup -KeyVaultName '$vaultName'
"@
        if ($result.state -ne 'configured') { throw 'Desktop configuration did not succeed.' }
    }
    finally {
        $workDeadline = $deadline
        $deadline = [DateTime]::UtcNow.AddMinutes(10)
        try {
            if ($principalId) {
                try {
                    $null = Invoke-RunnerAzure @('keyvault', 'delete-policy', '--name', $vaultName, '--object-id', $principalId)
                }
                catch { $failures.Add($_.Exception) }
            }
            if ($identityAttempted) {
                try {
                    $null = Invoke-RunnerAzure @(
                        'vm', 'identity', 'remove', '--resource-group', $ResourceGroup,
                        '--name', $VmName, '--identities', '[system]'
                    )
                }
                catch { $failures.Add($_.Exception) }
            }
        }
        finally { $deadline = $workDeadline }
    }
    if ($failures.Count -gt 0) { throw 'Temporary vault access cleanup failed; do not use this VM for coding.' }
    $currentVm = Invoke-RunnerAzure @('vm', 'get-instance-view', '--resource-group', $ResourceGroup, '--name', $VmName)
    $vault = Invoke-RunnerAzure @('keyvault', 'show', '--name', $vaultName)
    Assert-RunnerIdentityWithoutAzureAccess -Vm $currentVm -VaultPolicies @($vault.properties.accessPolicies)
    $null = Invoke-RunnerAzure @('vm', 'restart', '--resource-group', $ResourceGroup, '--name', $VmName) 900
    Start-Sleep -Seconds 60
    $result = Invoke-DesktopGuest "& '$guestPath' -Action Health"
    if ($result.state -ne 'ready' -or $result.runnerRegistered -ne $false) {
        throw 'The non-admin desktop did not recover after restart.'
    }
    Write-Output 'Dedicated non-admin desktop and automatic logon verified. No coding runner is registered.'
}
catch { $failures.Add($_.Exception) }
finally {
    $deadline = [DateTime]::UtcNow.AddMinutes(10)
    try {
        $null = Invoke-RunnerAzure @('vm', 'deallocate', '--resource-group', $ResourceGroup, '--name', $VmName) 600
    }
    catch { $failures.Add($_.Exception) }
}
if ($failures.Count -gt 0) { throw [AggregateException]::new('Desktop provisioning failed.', $failures) }
Write-Output 'Excel VM is deallocated.'
