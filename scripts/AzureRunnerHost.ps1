# Callers supply their overall deadline and explicit subscription arguments.
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
            return $parsed
        }
    }
    finally { Remove-Job -Job $job -Force }
}

function Get-RunnerAzureCommandArguments {
    param([string[]]$Arguments)
    if ($Arguments[0] -in @('ad', 'account')) { return ,$Arguments }
    return ,($Arguments + $subscriptionArguments)
}
