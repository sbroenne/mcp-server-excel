$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'AzureRunnerHost.ps1')
$subscriptionArguments = @('--subscription', 'synthetic-subscription')
foreach ($command in @('ad', 'account')) {
    $arguments = Get-RunnerAzureCommandArguments @($command, 'show')
    if ($arguments -contains '--subscription') { throw 'Tenant/account Azure CLI commands must not receive unsupported subscription options.' }
}
$arguments = Get-RunnerAzureCommandArguments @('vm', 'show')
if ($arguments.Count -ne 4 -or $arguments[2] -ne '--subscription' -or $arguments[3] -ne 'synthetic-subscription') {
    throw 'Resource operations must retain the explicitly selected subscription.'
}
Write-Output 'Azure directory commands and subscription-scoped resource arguments passed.'
