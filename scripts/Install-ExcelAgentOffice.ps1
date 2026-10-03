<#
.SYNOPSIS
Installs current 64-bit Excel 2024 Retail on the owned Azure development VM.
.DESCRIPTION
Runs a bounded SYSTEM installation task, following mcp-windows maintenance.
No installer, licence key or Azure token is uploaded to a public location.
Starts from a deallocated VM and deallocates it again on completion/failure.
Activation and interactive qualification remain separate required steps.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ResourceGroup = 'rg-excel-copilot-runner',
    [string]$VmName = 'vm-excel-copilot-runner',
    [string]$SubscriptionId
)

$ErrorActionPreference = 'Stop'
if (-not $PSCmdlet.ShouldProcess($VmName, 'Install Excel 2024 Retail on the Azure VM')) { return }
$deadline = [DateTime]::UtcNow.AddMinutes(35)
$subscriptionArguments = if ($SubscriptionId) { @('--subscription', $SubscriptionId) } else { @() }
$source = Join-Path (Split-Path -Parent $PSScriptRoot) 'infrastructure\azure\install-excel-office.ps1'
$guestPath = 'C:\ProgramData\ExcelMcp\Provisioning\install-excel-office.ps1'

. (Join-Path $PSScriptRoot 'AzureRunnerHost.ps1')

function Invoke-OfficeGuest {
    param([string]$Script)
    $path = Join-Path ([IO.Path]::GetTempPath()) "excel-office-command-$([Guid]::NewGuid().ToString('N')).ps1"
    try {
        ("`$ErrorActionPreference = 'Stop'`n" + $Script) | Set-Content -LiteralPath $path -Encoding UTF8
        $result = Invoke-RunnerAzure @(
            'vm', 'run-command', 'invoke', '--resource-group', $ResourceGroup,
            '--name', $VmName, '--command-id', 'RunPowerShellScript', '--scripts', "@$path"
        )
        $messages = @(
            foreach ($value in $result.value) {
                foreach ($line in ($value.message -split "`r?`n")) {
                    if ($line.StartsWith('EXCELMCP_OFFICE=')) { $line.Substring(16) }
                }
            }
        )
        if ($messages.Count -ne 1) { throw 'The Office installation probe did not return exactly one result.' }
        return $messages[0] | ConvertFrom-Json
    }
    finally {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }
}

$vm = Invoke-RunnerAzure @('vm', 'get-instance-view', '--resource-group', $ResourceGroup, '--name', $VmName)
if ($vm.tags.project -ne 'mcp-server-excel' -or $vm.tags.purpose -ne 'excel-cloud-agent-devtest') {
    throw 'The VM is not owned by this Excel runner deployment.'
}
$powerStates = @($vm.instanceView.statuses | Where-Object { $_.code -like 'PowerState/*' })
if ($powerStates.Count -ne 1 -or $powerStates[0].code -ne 'PowerState/deallocated') {
    throw 'Office provisioning must start from a deallocated VM; do not interrupt active work.'
}
$shutdownTime = [DateTime]::UtcNow.AddHours(1).ToString('HHmm')
$null = Invoke-RunnerAzure @(
    'vm', 'auto-shutdown', '--resource-group', $ResourceGroup,
    '--name', $VmName, '--time', $shutdownTime
)
$failures = [Collections.Generic.List[Exception]]::new()
try {
    $null = Invoke-RunnerAzure @('vm', 'start', '--resource-group', $ResourceGroup, '--name', $VmName) 900
    $bootDeadline = [DateTime]::UtcNow.AddMinutes(20)
    do {
        $instance = Invoke-RunnerAzure @('vm', 'get-instance-view', '--resource-group', $ResourceGroup, '--name', $VmName)
        $ready = @($instance.instanceView.vmAgent.statuses | Where-Object displayStatus -EQ 'Ready').Count -gt 0
        if (-not $ready) {
            if ([DateTime]::UtcNow -ge $bootDeadline) { throw 'Azure guest agent did not become ready.' }
            Start-Sleep -Seconds 15
        }
    } while (-not $ready)

    $encoded = [Convert]::ToBase64String([IO.File]::ReadAllBytes($source))
    $hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    $state = Invoke-OfficeGuest @"
`$directory = Split-Path '$guestPath'
`$task = Get-ScheduledTask -TaskName 'ExcelMcp-Install-Office' -ErrorAction SilentlyContinue
if (`$task -and `$task.State -eq 'Running') { throw 'Office installation is already running.' }
New-Item `$directory -ItemType Directory -Force | Out-Null
& icacls.exe `$directory /inheritance:r /grant:r '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' | Out-Null
if (`$LASTEXITCODE -ne 0) { throw 'Could not protect the provisioning directory.' }
[IO.File]::WriteAllBytes('$guestPath', [Convert]::FromBase64String('$encoded'))
if ((Get-FileHash -LiteralPath '$guestPath' -Algorithm SHA256).Hash -ne '$hash') {
    throw 'Office provisioning script hash mismatch.'
}
& '$guestPath' -Action Start
"@
    while ($state.state -eq 'running') {
        Start-Sleep -Seconds 30
        $state = Invoke-OfficeGuest "& '$guestPath' -Action Status"
    }
    if ($state.state -ne 'installed' -or $state.installation.product -ne 'Excel2024Retail' -or
        $state.installation.platform -ne 'x64' -or -not $state.installation.version) {
        throw "Excel installation did not complete successfully: $($state.error)"
    }
    Write-Output "Installed Excel2024Retail x64, build $($state.installation.version). Activation and interactive tests have not been verified."
}
catch { $failures.Add($_.Exception) }
finally {
    # Cleanup has its own budget even after an installation deadline.
    $deadline = [DateTime]::UtcNow.AddMinutes(10)
    try {
        $null = Invoke-RunnerAzure @('vm', 'deallocate', '--resource-group', $ResourceGroup, '--name', $VmName) 600
    }
    catch { $failures.Add($_.Exception) }
}
if ($failures.Count -gt 0) { throw [AggregateException]::new('Excel provisioning failed.', $failures) }
Write-Output 'Excel VM is deallocated.'
