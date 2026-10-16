#Requires -Version 7.4
# Callers supply their overall deadline and explicit subscription arguments.
function ConvertTo-RunnerPowerShell7Command {
    param([Parameter(Mandatory)][string]$Script)
    $payload = @"
`$ErrorActionPreference = 'Stop'
if (`$PSVersionTable.PSVersion -lt [version]'7.4') { throw 'PowerShell 7.4 or later is required on the VM.' }
$Script
"@
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($payload))
    # Azure owns the Windows PowerShell transport; repository code runs only in pwsh.
    return @"
`$ErrorActionPreference = 'Stop'
`$pwsh = Join-Path `$env:ProgramFiles 'PowerShell\7\pwsh.exe'
if (-not (Test-Path -LiteralPath `$pwsh -PathType Leaf)) { throw 'Preinstall PowerShell 7.4 or later on the VM image before provisioning.' }
& `$pwsh -NoProfile -NonInteractive -EncodedCommand '$encoded'
if (`$LASTEXITCODE -ne 0) { throw "PowerShell 7 guest command failed with exit code `$LASTEXITCODE." }
[Console]::WriteLine()
Write-Output 'EXCELMCP_PWSH_COMPLETE'
"@
}

function Assert-RunnerIdentityWithoutAzureAccess {
    param($Vm, [object[]]$VaultPolicies)
    if (-not $Vm.identity -or $Vm.identity.type -eq 'None') { return }
    if ($Vm.identity.type -ne 'SystemAssigned' -or -not $Vm.identity.principalId) {
        throw 'Unexpected VM identity; do not replace unrelated credentials.'
    }
    $principalId = $Vm.identity.principalId
    $roles = @(Invoke-RunnerAzure @(
        'role', 'assignment', 'list', '--assignee-object-id', $principalId, '--all'
    ))
    if ($roles.Count -gt 0 -or
        @($VaultPolicies | Where-Object { $_.objectId -eq $principalId }).Count -gt 0) {
        throw 'The existing VM identity has Azure roles or password-vault access; do not expose it to coding work.'
    }
}

function Invoke-RunnerAzure {
    param([string[]]$Arguments, [int]$TimeoutSeconds = 180)
    $remaining = [int]($deadline - [DateTime]::UtcNow).TotalSeconds
    if ($remaining -le 0) { throw 'Azure runner operation exceeded its overall deadline.' }
    $job = Start-Job -ScriptBlock {
        param($AzureArguments)
        $ErrorActionPreference = 'Stop'
        $output = & az @AzureArguments --only-show-errors --output json
        if ($LASTEXITCODE -ne 0) { throw "Azure command failed: $($AzureArguments[0]) $($AzureArguments[1])." }
        return $output
    } -ArgumentList (,(Get-RunnerAzureCommandArguments $Arguments))
    try {
        if (-not (Wait-Job -Job $job -Timeout ([Math]::Min($TimeoutSeconds, $remaining)))) {
            Stop-Job -Job $job
            throw 'Azure command timed out; its result is uncertain. Do not blindly retry.'
        }
        $output = Receive-Job -Job $job -ErrorAction Stop
        if ($job.State -ne 'Completed') { throw 'Azure command job did not complete successfully.' }
        if ($output) {
            $parsed = ($output -join "`n") | ConvertFrom-Json
            if ($Arguments -contains 'RunPowerShellScript') {
                $completed = @($parsed.value | ForEach-Object { $_.message -split "`r?`n" } |
                    Where-Object { $_ -eq 'EXCELMCP_PWSH_COMPLETE' })
                if ($completed.Count -ne 1) {
                    throw 'PowerShell 7 guest execution did not complete successfully; inspect the Azure Run Command error output.'
                }
            }
            return $parsed
        }
        if ($Arguments -contains 'RunPowerShellScript') { throw 'Azure returned no guest execution result.' }
    }
    finally { Remove-Job -Job $job -Force }
}

function Get-RunnerAzureCommandArguments {
    param([string[]]$Arguments)
    if ($Arguments[0] -in @('ad', 'account')) { return ,$Arguments }
    return ,($Arguments + $subscriptionArguments)
}
