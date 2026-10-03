<#
.SYNOPSIS
Provisions the dedicated Excel VM using the mcp-windows resource pattern.
.DESCRIPTION
Creates the VM, isolated network, password vault and shutdown protection.
Does not install Excel, register a runner or change Copilot routing.
The VM is deallocated after provisioning. Existing owned deployments reuse
their administrator password and pinned Windows image.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ResourceGroup = 'rg-excel-copilot-runner',
    [string]$Location = 'eastus2',
    [string]$VmSize = 'Standard_D2s_v7',
    [string]$SubscriptionId,
    [ValidatePattern('^win11-[a-z0-9-]+$')]
    [string]$WindowsImageSku = 'win11-25h2-pro'
)

$ErrorActionPreference = 'Stop'
$vmName = 'vm-excel-copilot-runner'
$template = Join-Path (Split-Path -Parent $PSScriptRoot) 'infrastructure\azure\excel-runner.bicep'

if (-not $PSCmdlet.ShouldProcess($ResourceGroup, 'Provision the dedicated Excel development VM')) {
    return
}

function Invoke-AzureJson {
    param([string[]]$Arguments)
    $response = & az @Arguments --only-show-errors --output json
    if ($LASTEXITCODE -ne 0) {
        throw "Azure command failed: $($Arguments[0]) $($Arguments[1])."
    }
    if (-not $response) { throw "Azure command returned no result: $($Arguments[0]) $($Arguments[1])." }
    $parsed = ($response -join "`n") | ConvertFrom-Json
    return $parsed
}

$accountArgs = @('account', 'show')
if ($SubscriptionId) { $accountArgs += @('--subscription', $SubscriptionId) }
$account = Invoke-AzureJson $accountArgs
if ($account.state -ne 'Enabled' -or -not $account.id) { throw 'An enabled Azure subscription is required.' }
$subscriptionArgs = @('--subscription', $account.id)
$groupExists = Invoke-AzureJson (@('group', 'exists', '--name', $ResourceGroup) + $subscriptionArgs)
if ($groupExists) {
    $group = Invoke-AzureJson (@('group', 'show', '--name', $ResourceGroup) + $subscriptionArgs)
    if ($group.tags.project -ne 'mcp-server-excel' -or
        $group.tags.purpose -ne 'excel-cloud-agent-devtest') {
        throw 'The existing resource group is not owned by this Excel runner deployment.'
    }
    if ($group.location -ne $Location) { throw 'Existing resource group location does not match the requested location.' }
}

$deployer = Invoke-AzureJson @('ad', 'signed-in-user', 'show')
if (-not $deployer.id) { throw 'A signed-in user is required for initial vault ownership.' }
$image = Invoke-AzureJson (@(
    'vm', 'image', 'show', '--location', $Location,
    '--urn', "MicrosoftWindowsDesktop:windows-11:${WindowsImageSku}:latest"
) + $subscriptionArgs)
if ($image.architecture -ne 'x64' -or $image.hyperVGeneration -ne 'V2' -or -not $image.name) {
    throw 'The Windows image must be x64, Generation 2 and resolve to a specific version.'
}

$password = $null
if ($groupExists) {
    $vms = @(Invoke-AzureJson (@('vm', 'list', '--resource-group', $ResourceGroup) + $subscriptionArgs))
    $existingVm = $vms | Where-Object name -EQ $vmName
    if ($existingVm) {
        if ($existingVm.tags.purpose -ne 'excel-cloud-agent-devtest') { throw 'The existing VM is not owned by this deployment.' }
        $instance = Invoke-AzureJson (@(
            'vm', 'get-instance-view', '--resource-group', $ResourceGroup, '--name', $vmName
        ) + $subscriptionArgs)
        $powerStates = @($instance.instanceView.statuses | Where-Object { $_.code -like 'PowerState/*' })
        if ($powerStates.Count -ne 1 -or $powerStates[0].code -ne 'PowerState/deallocated') {
            throw 'The existing VM must be deallocated before provisioning; do not interrupt active work.'
        }
        if ($existingVm.hardwareProfile.vmSize -ne $VmSize -or
            $existingVm.storageProfile.imageReference.sku -ne $WindowsImageSku) {
            throw 'Changing an existing VM size or Windows edition requires a separate migration.'
        }
        $image.name = $existingVm.storageProfile.imageReference.version
        if (-not $image.name -or $image.name -eq 'latest') { throw 'Existing Windows image version is not pinned.' }
    }
    $vaults = @(Invoke-AzureJson (@('keyvault', 'list', '--resource-group', $ResourceGroup) + $subscriptionArgs))
    $ownedVaults = @($vaults | Where-Object { $_.tags.purpose -eq 'excel-cloud-agent-devtest' })
    if ($ownedVaults.Count -gt 1) { throw 'Multiple owned password vaults were found; refusing an ambiguous deployment.' }
    if ($ownedVaults.Count -eq 1) {
        $secret = Invoke-AzureJson (@(
            'keyvault', 'secret', 'show', '--vault-name', $ownedVaults[0].name, '--name', 'vm-admin-password'
        ) + $subscriptionArgs)
        if (-not $secret.value) { throw 'The stored administrator password is empty.' }
        $password = $secret.value
    }
    elseif ($existingVm) { throw 'An existing VM has no owned password vault; do not reset its credentials implicitly.' }
}
if (-not $password) {
    $bytes = New-Object byte[] 32
    $generator = [Security.Cryptography.RandomNumberGenerator]::Create()
    try { $generator.GetBytes($bytes) }
    finally { $generator.Dispose() }
    $password = 'Aa1!' + [Convert]::ToBase64String($bytes)
    [Array]::Clear($bytes, 0, $bytes.Length)
}

$privateDirectory = Join-Path ([IO.Path]::GetTempPath()) "excel-runner-deploy-$([Guid]::NewGuid().ToString('N'))"
$parameterPath = Join-Path $privateDirectory 'parameters.json'
try {
    New-Item -ItemType Directory -Path $privateDirectory | Out-Null
    $acl = Get-Acl -LiteralPath $privateDirectory
    $acl.SetAccessRuleProtection($true, $false)
    foreach ($sid in @(
        [Security.Principal.WindowsIdentity]::GetCurrent().User,
        [Security.Principal.SecurityIdentifier]::new('S-1-5-18')
    )) {
        $rule = [Security.AccessControl.FileSystemAccessRule]::new(
            $sid, 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow'
        )
        $acl.AddAccessRule($rule)
    }
    Set-Acl -LiteralPath $privateDirectory -AclObject $acl
    $parameters = @{
        '$schema' = 'https://schema.management.azure.com/schemas/2019-04-01/deploymentParameters.json#'
        contentVersion = '1.0.0.0'
        parameters = @{
            location = @{ value = $Location }
            vmSize = @{ value = $VmSize }
            adminPassword = @{ value = $password }
            deployerObjectId = @{ value = $deployer.id }
            windowsImageSku = @{ value = $WindowsImageSku }
            windowsImageVersion = @{ value = $image.name }
            autoShutdownTime = @{ value = [DateTime]::UtcNow.AddHours(1).ToString('HHmm') }
        }
    }
    $parameters | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $parameterPath -Encoding utf8
    if (-not $groupExists) {
        $null = Invoke-AzureJson (@(
            'group', 'create', '--name', $ResourceGroup, '--location', $Location,
            '--tags', 'project=mcp-server-excel', 'purpose=excel-cloud-agent-devtest'
        ) + $subscriptionArgs)
    }
    $deploymentArguments = @(
        '--resource-group', $ResourceGroup, '--template-file', $template,
        '--parameters', "@$parameterPath"
    ) + $subscriptionArgs
    $null = Invoke-AzureJson (@('deployment', 'group', 'validate') + $deploymentArguments)
    $deployment = Invoke-AzureJson (@(
        'deployment', 'group', 'create', '--name', 'excel-copilot-runner'
    ) + $deploymentArguments)
    if ($deployment.properties.provisioningState -ne 'Succeeded') { throw 'Runner infrastructure deployment did not succeed.' }
    & az vm deallocate --name $vmName --resource-group $ResourceGroup @subscriptionArgs --only-show-errors --output none
    if ($LASTEXITCODE -ne 0) { throw 'VM deallocation failed; the shutdown backstop remains enabled. Check compute billing.' }
    Write-Output 'Excel VM provisioned and deallocated. Excel installation and runner qualification are still required.'
}
finally {
    if (Test-Path -LiteralPath $parameterPath) { Remove-Item -LiteralPath $parameterPath -Force }
    if (Test-Path -LiteralPath $privateDirectory) { Remove-Item -LiteralPath $privateDirectory -Force }
    $password = $null
}
