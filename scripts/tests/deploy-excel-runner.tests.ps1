$ErrorActionPreference = 'Stop'
$deploymentScript = Join-Path (Split-Path -Parent $PSScriptRoot) 'Deploy-ExcelAgentRunner.ps1'
$global:ExcelRunnerDeploymentTestCalls = [Collections.Generic.List[string]]::new()
$global:ExcelRunnerDeploymentTestScenario = 'what-if'
$global:ExcelRunnerDeploymentTestParameterPath = $null

function az {
    $command = $args -join ' '
    $global:ExcelRunnerDeploymentTestCalls.Add($command)
    $global:LASTEXITCODE = 0
    if ($command.StartsWith('account show ')) {
        if ($global:ExcelRunnerDeploymentTestScenario -eq 'account-failure') {
            $global:LASTEXITCODE = 1
            return
        }
        return '{"id":"test-subscription","state":"Enabled"}'
    }
    if ($command.StartsWith('group exists ')) {
        if ($global:ExcelRunnerDeploymentTestScenario -eq 'fresh') { return 'false' }
        return 'true'
    }
    if ($command.StartsWith('group create ')) { return '{"name":"synthetic-group"}' }
    if ($command.StartsWith('group show ')) {
        if ($global:ExcelRunnerDeploymentTestScenario -in @('rerun', 'deployment-failure', 'running-vm')) {
            return '{"location":"eastus2","tags":{"project":"mcp-server-excel","purpose":"excel-cloud-agent-devtest"}}'
        }
        return '{"location":"eastus2","tags":{"project":"unrelated-project"}}'
    }
    if ($command.StartsWith('ad signed-in-user show ')) {
        return '{"id":"test-deployer"}'
    }
    if ($command.StartsWith('vm image show ')) {
        return '{"name":"new-image-version","architecture":"x64","hyperVGeneration":"V2"}'
    }
    if ($command.StartsWith('vm list ')) {
        return '[{"name":"vm-excel-copilot-runner","tags":{"purpose":"excel-cloud-agent-devtest"},"hardwareProfile":{"vmSize":"Standard_D2s_v7"},"storageProfile":{"imageReference":{"sku":"win11-25h2-pro","version":"original-image-version"}}}]'
    }
    if ($command.StartsWith('vm get-instance-view ')) {
        $powerState = if ($global:ExcelRunnerDeploymentTestScenario -eq 'running-vm') {
            'PowerState/running'
        }
        else { 'PowerState/deallocated' }
        return "{`"instanceView`":{`"statuses`":[{`"code`":`"$powerState`"}]}}"
    }
    if ($command.StartsWith('keyvault list ')) {
        return '[{"name":"test-vault","tags":{"purpose":"excel-cloud-agent-devtest"}}]'
    }
    if ($command.StartsWith('keyvault secret show ')) {
        return '{"value":"synthetic-test-password"}'
    }
    if ($command.StartsWith('deployment group ')) {
        $parameterIndex = [Array]::IndexOf($args, '--parameters')
        if ($parameterIndex -lt 0) { throw 'Deployment must use a parameter file.' }
        $path = $args[$parameterIndex + 1].Substring(1)
        $global:ExcelRunnerDeploymentTestParameterPath = $path
        $parameters = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        if ($global:ExcelRunnerDeploymentTestScenario -eq 'fresh') {
            if ($parameters.parameters.adminPassword.value -notmatch '^Aa1![a-zA-Z0-9+/]{43}=$' -or
                $parameters.parameters.windowsImageVersion.value -ne 'new-image-version') {
                throw 'Fresh deployment must generate a strong password and pin the resolved image.'
            }
        }
        elseif ($parameters.parameters.adminPassword.value -ne 'synthetic-test-password') {
            throw 'A rerun must preserve the stored administrator password.'
        }
        if ($global:ExcelRunnerDeploymentTestScenario -ne 'fresh' -and
            $parameters.parameters.windowsImageVersion.value -ne 'original-image-version') {
            throw 'A rerun must preserve the existing Windows image.'
        }
        $directoryAcl = Get-Acl -LiteralPath (Split-Path -Parent $path)
        if (-not $directoryAcl.AreAccessRulesProtected -or
            @($directoryAcl.Access | Where-Object IsInherited).Count -gt 0) {
            throw 'The secret parameter directory must not inherit permissions.'
        }
        if ($command.StartsWith('deployment group create ') -and
            $global:ExcelRunnerDeploymentTestScenario -eq 'deployment-failure') {
            $global:LASTEXITCODE = 1
            return
        }
        return '{"properties":{"provisioningState":"Succeeded"}}'
    }
    if ($command.StartsWith('vm deallocate ')) { return }
    throw "Unexpected Azure command in deployment test: $command"
}

& $deploymentScript -WhatIf
if ($global:ExcelRunnerDeploymentTestCalls.Count -ne 0) { throw 'WhatIf must not call Azure or create resources.' }

foreach ($scenario in @('wrong-owner', 'account-failure')) {
    $global:ExcelRunnerDeploymentTestScenario = $scenario
    $global:ExcelRunnerDeploymentTestCalls.Clear()
    $failure = $null
    try { & $deploymentScript -Confirm:$false }
    catch { $failure = $_.Exception.Message }
    if (-not $failure) { throw "$scenario must fail." }
    $expected = if ($scenario -eq 'wrong-owner') { 'not owned' } else { 'Azure command failed' }
    if ($failure -notmatch $expected) { throw "Unexpected failure for ${scenario}: $failure" }
    if (@($global:ExcelRunnerDeploymentTestCalls | Where-Object { $_ -match ' create | set | validate | deallocate ' }).Count -gt 0) {
        throw "$scenario must not change Azure resources."
    }
}

foreach ($scenario in @('rerun', 'deployment-failure', 'fresh')) {
    $global:ExcelRunnerDeploymentTestScenario = $scenario
    $global:ExcelRunnerDeploymentTestCalls.Clear()
    $failure = $null
    try { & $deploymentScript -Confirm:$false }
    catch { $failure = $_.Exception.Message }
    if ($scenario -in @('rerun', 'fresh')) {
        if ($failure) { throw "Rerun failed: $failure" }
        if (@($global:ExcelRunnerDeploymentTestCalls | Where-Object { $_ -match '^vm deallocate ' }).Count -ne 1) {
            throw 'A successful provision must deallocate the VM.'
        }
    }
    elseif ($failure -notmatch 'Azure command failed') {
        throw 'A failed deployment must report an explicit error.'
    }
    $path = $global:ExcelRunnerDeploymentTestParameterPath
    if (-not $path -or (Test-Path -LiteralPath $path) -or
        (Test-Path -LiteralPath (Split-Path -Parent $path))) {
        throw 'Secret parameters and their directory must be removed on success and failure.'
    }
}

$global:ExcelRunnerDeploymentTestScenario = 'running-vm'
$global:ExcelRunnerDeploymentTestCalls.Clear()
$failure = $null
try { & $deploymentScript -Confirm:$false }
catch { $failure = $_.Exception.Message }
if ($failure -notmatch 'must be deallocated') {
    throw 'An existing running VM must be refused before deployment can interrupt a task.'
}
if (@($global:ExcelRunnerDeploymentTestCalls | Where-Object { $_ -match '^deployment group |^vm deallocate ' }).Count -gt 0) {
    throw 'A running VM must not be redeployed or deallocated.'
}

Write-Output 'Excel runner deployment safety tests passed.'
