<#
.SYNOPSIS
Installs the repository's development tools on the dedicated Azure VM.
.DESCRIPTION
Uses the existing mcp-windows SYSTEM worker/start/deallocate approach.
Does not register a runner or change cloud-agent routing.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ResourceGroup = 'rg-excel-copilot-runner',
    [string]$VmName = 'vm-excel-copilot-runner',
    [string]$SubscriptionId
)
$ErrorActionPreference = 'Stop'
if (-not $PSCmdlet.ShouldProcess($VmName, 'Install development tools on the Azure VM')) { return }
$root = Split-Path -Parent $PSScriptRoot
$sdkSettings = (Get-Content -LiteralPath (Join-Path $root 'global.json') -Raw | ConvertFrom-Json).sdk
$sdk = $sdkSettings.version
$rollForward = $sdkSettings.rollForward
if ($sdk -notmatch '^\d+\.\d+\.\d+$') { throw 'The repository SDK version is invalid.' }
if ($rollForward -notin @('latestFeature', 'disable')) { throw 'The bootstrap requires an explicitly supported SDK roll-forward policy.' }
$deadline = [DateTime]::UtcNow.AddMinutes(60)
$subscriptionArguments = if ($SubscriptionId) { @('--subscription', $SubscriptionId) } else { @() }
. (Join-Path $PSScriptRoot 'AzureRunnerHost.ps1')

function Invoke-ToolchainGuest {
    param([string]$Script)
    $path = Join-Path ([IO.Path]::GetTempPath()) "excel-toolchain-$([Guid]::NewGuid().ToString('N')).ps1"
    try {
        ("`$ErrorActionPreference = 'Stop'`n" + $Script) | Set-Content -LiteralPath $path -Encoding UTF8
        $response = Invoke-RunnerAzure @(
            'vm', 'run-command', 'invoke', '--resource-group', $ResourceGroup,
            '--name', $VmName, '--command-id', 'RunPowerShellScript', '--scripts', "@$path"
        ) 300
        $prefix = 'EXCELMCP_TOOLCHAIN='
        $messages = @(
            foreach ($value in $response.value) {
                foreach ($line in ($value.message -split "`r?`n")) {
                    if ($line.StartsWith($prefix)) { $line.Substring($prefix.Length) }
                }
            }
        )
        if ($messages.Count -ne 1) {
            throw ('Toolchain provisioning did not return exactly one result. ' + (@($response.value.message) -join "`n"))
        }
        return $messages[0] | ConvertFrom-Json
    }
    finally { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force } }
}

$vm = Invoke-RunnerAzure @('vm', 'get-instance-view', '--resource-group', $ResourceGroup, '--name', $VmName)
if ($vm.tags.project -ne 'mcp-server-excel' -or $vm.tags.purpose -ne 'excel-cloud-agent-devtest') {
    throw 'The VM is not owned by this Excel runner deployment.'
}
$states = @($vm.instanceView.statuses | Where-Object code -Like 'PowerState/*')
if ($states.Count -ne 1 -or $states[0].code -ne 'PowerState/deallocated') {
    throw 'Toolchain provisioning must start from a deallocated VM; do not interrupt active work.'
}
$vaults = @(Invoke-RunnerAzure @('keyvault', 'list', '--resource-group', $ResourceGroup))
$vaults = @($vaults | Where-Object { $_.tags.purpose -eq 'excel-cloud-agent-devtest' })
if ($vaults.Count -ne 1) { throw 'Expected exactly one owned password vault.' }
$vault = Invoke-RunnerAzure @('keyvault', 'show', '--name', $vaults[0].name)
Assert-RunnerIdentityWithoutAzureAccess $vm @($vault.properties.accessPolicies)
$null = Invoke-RunnerAzure @(
    'vm', 'auto-shutdown', '--resource-group', $ResourceGroup, '--name', $VmName,
    '--time', [DateTime]::UtcNow.AddHours(2).ToString('HHmm')
)
$failures = [Collections.Generic.List[Exception]]::new()
try {
    $null = Invoke-RunnerAzure @('vm', 'start', '--resource-group', $ResourceGroup, '--name', $VmName) 900
    Start-Sleep -Seconds 60
    $upload = [Text.StringBuilder]::new()
    $null = $upload.AppendLine(@'
$directory = 'C:\ProgramData\ExcelMcp\Provisioning'
foreach ($name in @('ExcelMcp-Install-Toolchain', 'ExcelMcp-Install-Office')) {
    $task = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
    if ($task -and $task.State -eq 'Running') { throw 'Another provisioning worker is running.' }
}
New-Item $directory -ItemType Directory -Force | Out-Null
& icacls.exe $directory /inheritance:r /grant:r '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'Could not protect the provisioning directory.' }
'@)
    foreach ($name in @('install-excel-office.ps1', 'install-excel-toolchain.ps1')) {
        $source = Join-Path $root "infrastructure\azure\$name"
        $encoded = [Convert]::ToBase64String([IO.File]::ReadAllBytes($source))
        $hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
        $null = $upload.AppendLine(@"
[IO.File]::WriteAllBytes((Join-Path `$directory '$name'), [Convert]::FromBase64String('$encoded'))
if ((Get-FileHash -LiteralPath (Join-Path `$directory '$name') -Algorithm SHA256).Hash -ne '$hash') { throw 'Provisioning script hash mismatch.' }
"@)
    }
    $globalJson = [Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $root 'global.json')))
    $null = $upload.AppendLine("[IO.File]::WriteAllBytes((Join-Path `$directory 'global.json'), [Convert]::FromBase64String('$globalJson'))")
    $null = $upload.AppendLine("& 'C:\ProgramData\ExcelMcp\Provisioning\install-excel-toolchain.ps1' -Action Start -SdkVersion $sdk -RollForward $rollForward")
    $state = Invoke-ToolchainGuest $upload.ToString()
    while ($state.state -eq 'running') {
        Start-Sleep -Seconds 30
        $state = Invoke-ToolchainGuest "& 'C:\ProgramData\ExcelMcp\Provisioning\install-excel-toolchain.ps1' -Action Status -SdkVersion $sdk -RollForward $rollForward"
    }
    if ($state.state -ne 'installed' -or $state.tools.requiredSdk -ne $sdk -or
        $state.tools.rollForward -ne $rollForward -or
        $state.tools.sdk -notmatch '^\d+\.\d+\.\d+$' -or
        $state.tools.git -notmatch '^git version \d+\.\d+\.\d+\.windows\.\d+$' -or
        $state.tools.powershell -notmatch '^7\.\d+\.\d+$' -or
        $state.tools.node -notmatch '^v22\.\d+\.\d+$' -or
        $state.tools.bash -notmatch '^GNU bash, version \d+\.\d+' -or
        $state.tools.jq -ne 'jq-1.8.2' -or $state.runnerRegistered -ne $false) {
        throw "Toolchain installation failed or returned incomplete/mismatched tool versions: $($state.error)"
    }
    if ($state.tools.developmentMode -isnot [bool] -or -not $state.tools.developmentMode) {
        throw 'Windows Developer Mode must be reported as an enabled boolean for limited-user symbolic-link extraction.'
    }
    . (Join-Path $root 'infrastructure\azure\install-excel-toolchain.ps1') -SdkVersion $sdk -RollForward $rollForward
    Assert-RequiredRunnerSdk -Required $sdk -Installed @("$($state.tools.sdk) [resolved]") -RollForward $rollForward
    $state.tools | ConvertTo-Json -Compress
    if ($state.rebootRequired) { Write-Output 'A restart is required before desktop qualification.' }
}
catch { $failures.Add($_.Exception) }
finally {
    $deadline = [DateTime]::UtcNow.AddMinutes(10)
    try { $null = Invoke-RunnerAzure @('vm', 'deallocate', '--resource-group', $ResourceGroup, '--name', $VmName) 600 }
    catch { $failures.Add($_.Exception) }
}
if ($failures.Count -gt 0) { throw [AggregateException]::new('Toolchain provisioning failed.', $failures) }
Write-Output 'Development tools installed; VM deallocated. Coding runner remains unregistered.'
