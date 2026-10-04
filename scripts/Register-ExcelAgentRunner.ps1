<#
.SYNOPSIS
Installs and registers the dedicated interactive runner, leaving it offline.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ResourceGroup = 'rg-excel-copilot-runner',
    [string]$VmName = 'vm-excel-copilot-runner',
    [string]$SubscriptionId
)
$ErrorActionPreference = 'Stop'
if (-not $PSCmdlet.ShouldProcess($VmName, 'Register the dedicated repository runner without starting or routing jobs')) { return }
$deadline = [DateTime]::UtcNow.AddMinutes(45)
$subscriptionArguments = if ($SubscriptionId) { @('--subscription', $SubscriptionId) } else { @() }
. (Join-Path $PSScriptRoot 'ExcelRunnerHost.ps1')
$vm = Assert-ExcelRunnerOwnedVm
if (@($vm.instanceView.statuses | Where-Object code -EQ 'PowerState/deallocated').Count -ne 1) {
    throw 'Registration must start from a parked VM.'
}
$runners = Invoke-ExcelRunnerGithub 'repos/sbroenne/mcp-server-excel/actions/runners?per_page=100'
if (@($runners.runners | Where-Object name -EQ 'azure-excel-copilot').Count) {
    throw 'An existing registration must be inspected, not replaced automatically.'
}
$release = Invoke-ExcelRunnerGithub 'repos/actions/runner/releases/latest'
$version = $release.tag_name.TrimStart('v')
$assets = @($release.assets | Where-Object name -EQ "actions-runner-win-x64-$version.zip")
if ($assets.Count -ne 1) { throw 'Expected exactly one Windows x64 runner asset.' }
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'infrastructure\azure\configure-excel-runner.ps1')
Assert-ExcelRunnerPackage $assets[0] $version
$operation = [Guid]::NewGuid().ToString('N')
$failures = [Collections.Generic.List[Exception]]::new()
$rsa = $null
$tokenBytes = $null
$token = $null

function Invoke-RegistrationDesktop {
    param([ValidateSet('Prepare', 'Configure')][string]$Stage)
    $null = Invoke-ExcelRunnerGuest @"
`$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\ProgramData\ExcelMcp\Desktop\configure-excel-runner.ps1" -Action $Stage -OperationId $operation'
`$principal = New-ScheduledTaskPrincipal -UserId "`$env:COMPUTERNAME\excelrunner" -LogonType Interactive -RunLevel Limited
`$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 7)
Register-ScheduledTask -TaskName 'ExcelMcp-Register-Runner' -Action `$action -Principal `$principal -Settings `$settings -Force | Out-Null
Start-ScheduledTask -TaskName 'ExcelMcp-Register-Runner'
Write-Output 'EXCELMCP_CONTROL={"state":"started"}'
"@
    $until = [DateTime]::UtcNow.AddMinutes(8)
    do {
        Start-Sleep -Seconds 10
        $state = Invoke-ExcelRunnerGuest @"
`$task = Get-ScheduledTask -TaskName 'ExcelMcp-Register-Runner'
`$profile = @(Get-CimInstance Win32_UserProfile | Where-Object SID -EQ (Get-LocalUser excelrunner).SID.Value)
`$path = Join-Path `$profile[0].LocalPath 'AppData\Local\ExcelMcp\Registration\$operation.json'
if (`$task.State -eq 'Running') { `$state = @{ state = 'running'; operationId = '$operation' } }
else {
    if ((Get-ScheduledTaskInfo -TaskName 'ExcelMcp-Register-Runner').LastTaskResult -ne 0) { throw 'Non-admin registration task failed; inspect its private result.' }
    `$state = Get-Content -LiteralPath `$path -Raw | ConvertFrom-Json
}
Write-Output ('EXCELMCP_CONTROL=' + (`$state | ConvertTo-Json -Compress))
"@ -OperationId $operation
        if ($state.state -ne 'running') { return $state }
    } while ([DateTime]::UtcNow -lt $until)
    throw 'Registration task exceeded its deadline.'
}

try {
    $null = Invoke-RunnerAzure @(
        'vm', 'auto-shutdown', '--resource-group', $ResourceGroup, '--name', $VmName,
        '--time', [DateTime]::UtcNow.AddHours(2).ToString('HHmm')
    )
    $null = Invoke-RunnerAzure @('vm', 'start', '--resource-group', $ResourceGroup, '--name', $VmName) 900
    Wait-ExcelRunnerGuestAgent
    Assert-ExcelRunnerIdle @() (Get-ExcelRunnerGuestActivity)
    Send-ExcelRunnerFiles Desktop @('configure-excel-runner.ps1', 'start-excel-runner.ps1', 'install-excel-toolchain.ps1', 'test-excel-desktop.ps1', 'excel-job-started.ps1', 'excel-job-completed.ps1', 'recover-excel-jobs.ps1', 'ExcelRunnerPolicy.ps1', 'global.json')
    $null = Invoke-ExcelRunnerDesktopHealth
    $url = $assets[0].browser_download_url
    $hash = $assets[0].digest.Substring(7)
    $installed = Invoke-ExcelRunnerGuest @"
`$directory = 'C:\actions-runner'
`$manifest = 'C:\ProgramData\ExcelMcp\Desktop\package.json'
if (Test-Path -LiteralPath `$directory) {
    if (-not (Test-Path -LiteralPath `$manifest)) { throw 'Existing runner directory has no ownership record.' }
    `$owned = Get-Content `$manifest -Raw | ConvertFrom-Json
    if (`$owned.repository -ne 'sbroenne/mcp-server-excel' -or `$owned.digest -ne '$hash') { throw 'Do not replace an unrelated runner directory.' }
}
else {
    @{ repository = 'sbroenne/mcp-server-excel'; digest = '$hash'; version = '$version' } | ConvertTo-Json -Compress | Set-Content `$manifest -Encoding UTF8
    New-Item -ItemType Directory -Path `$directory | Out-Null
}
`$sid = (Get-LocalUser excelrunner).SID.Value
& icacls.exe `$directory /inheritance:r /grant:r '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' "*`${sid}:(OI)(CI)M" | Out-Null
if (`$LASTEXITCODE -ne 0) { throw 'Runner directory permissions failed.' }
`$archive = 'C:\ProgramData\ExcelMcp\Desktop\actions-runner.zip'
try {
    Invoke-WebRequest '$url' -OutFile `$archive -UseBasicParsing -TimeoutSec 180
    if ((Get-FileHash `$archive -Algorithm SHA256).Hash -ne '$hash') { throw 'Runner download hash mismatch.' }
    Expand-Archive -LiteralPath `$archive -DestinationPath `$directory -Force
}
finally { if (Test-Path `$archive) { Remove-Item -LiteralPath `$archive -Force } }
Write-Output 'EXCELMCP_CONTROL={"state":"installed"}'
"@
    if ($installed.state -ne 'installed') { throw 'Runner package was not installed.' }
    $prepared = Invoke-RegistrationDesktop Prepare
    if ($prepared.state -ne 'prepared') { throw 'Transient guest registration key was not prepared.' }
    $parameters = [Security.Cryptography.RSAParameters]::new()
    $parameters.Modulus = [Convert]::FromBase64String($prepared.modulus)
    $parameters.Exponent = [Convert]::FromBase64String($prepared.exponent)
    $rsa = [Security.Cryptography.RSA]::Create()
    $rsa.ImportParameters($parameters)
    $token = Invoke-ExcelRunnerGithub 'repos/sbroenne/mcp-server-excel/actions/runners/registration-token' -Method POST
    if (-not $token.token) { throw 'No short-lived registration token was returned.' }
    $tokenBytes = [Text.Encoding]::UTF8.GetBytes($token.token)
    $encrypted = [Convert]::ToBase64String($rsa.Encrypt($tokenBytes, [Security.Cryptography.RSAEncryptionPadding]::OaepSHA1))
    $token.token = $null
    $null = Invoke-ExcelRunnerGuest @"
@{ operationId = '$operation'; repository = 'sbroenne/mcp-server-excel'; encryptedToken = '$encrypted' } |
    ConvertTo-Json -Compress | Set-Content 'C:\ProgramData\ExcelMcp\Desktop\registration-$operation.json' -Encoding UTF8
Write-Output 'EXCELMCP_CONTROL={"state":"transported"}'
"@
    $configured = Invoke-RegistrationDesktop Configure
    if ($configured.state -ne 'configured') { throw 'Runner was not registered.' }
    $server = Invoke-ExcelRunnerGithub "repos/sbroenne/mcp-server-excel/actions/runners/$($configured.agentId)"
    if ($server.name -ne 'azure-excel-copilot' -or $server.status -ne 'offline' -or $server.busy -ne $false -or
        @($server.labels).Count -ne 1 -or $server.labels[0].name -ne 'excel-copilot') {
        throw 'GitHub did not establish the expected offline, custom-label-only runner.'
    }
    $null = Invoke-ExcelRunnerGuest @"
`$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\ProgramData\ExcelMcp\Desktop\start-excel-runner.ps1"'
`$principal = New-ScheduledTaskPrincipal -UserId "`$env:COMPUTERNAME\excelrunner" -LogonType Interactive -RunLevel Limited
`$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 6)
Register-ScheduledTask -TaskName 'ExcelMcp-GitHub-Runner' -Action `$action -Principal `$principal -Settings `$settings -Force | Out-Null
@'
ACTIONS_RUNNER_HOOK_JOB_STARTED=C:\ProgramData\ExcelMcp\Desktop\excel-job-started.ps1
ACTIONS_RUNNER_HOOK_JOB_COMPLETED=C:\ProgramData\ExcelMcp\Desktop\excel-job-completed.ps1
'@ | Set-Content 'C:\actions-runner\.env' -Encoding UTF8
@{ repository = 'sbroenne/mcp-server-excel'; agentId = $($configured.agentId) } | ConvertTo-Json -Compress |
    Set-Content 'C:\ProgramData\ExcelMcp\Desktop\registration.json' -Encoding UTF8
Write-Output 'EXCELMCP_CONTROL={"state":"configured-offline"}'
"@
}
catch { $failures.Add($_.Exception) }
finally {
    if ($token) { $token.token = $null }
    if ($tokenBytes) { [Array]::Clear($tokenBytes, 0, $tokenBytes.Length) }
    if ($rsa) { $rsa.Dispose() }
    $deadline = [DateTime]::UtcNow.AddMinutes(15)
    try {
        $null = Invoke-ExcelRunnerGuest @"
`$path = 'C:\ProgramData\ExcelMcp\Desktop\registration-$operation.json'
if (Test-Path -LiteralPath `$path) { Remove-Item -LiteralPath `$path -Force }
`$profile = @(Get-CimInstance Win32_UserProfile | Where-Object SID -EQ (Get-LocalUser excelrunner).SID.Value)
`$key = Join-Path `$profile[0].LocalPath 'AppData\Local\ExcelMcp\Registration\$operation.key'
if (Test-Path -LiteralPath `$key) { Remove-Item -LiteralPath `$key -Force }
Write-Output 'EXCELMCP_CONTROL={"state":"transient-data-removed"}'
"@
    }
    catch { $failures.Add($_.Exception) }
    try {
        Assert-ExcelRunnerIdle @() (Get-ExcelRunnerGuestActivity)
        $null = Invoke-RunnerAzure @('vm', 'deallocate', '--resource-group', $ResourceGroup, '--name', $VmName) 600
    }
    catch { $failures.Add($_.Exception) }
}
if ($failures.Count) { throw [AggregateException]::new('Offline runner registration failed.', $failures) }
Write-Output 'Repository runner registered with only excel-copilot label; listener stopped and VM deallocated.'
