<#
.SYNOPSIS
Configures protected, password-free GitHub control of the dedicated Excel VM.
.DESCRIPTION
Reuses the repository's analytics federation pattern. No Azure client secret is
created, no identity is assigned to the coding account, and routing stays disabled.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Repository = 'sbroenne/mcp-server-excel',
    [string]$ResourceGroup = 'rg-excel-copilot-runner',
    [string]$VmName = 'vm-excel-copilot-runner',
    [string]$EnvironmentName = 'excel-runner-control'
)
$ErrorActionPreference = 'Stop'

function Get-ExcelRunnerControlApplicationDefinition {
    return @{
        displayName = 'excelmcp-github-runner-control'
        signInAudience = 'AzureADMyOrg'
        tags = @('ExcelMcpRunnerControl')
    }
}

function Assert-ExcelRunnerControlEnvironment {
    param($Environment, $Policies, [string]$DefaultBranch)
    if (-not $Environment.deployment_branch_policy.custom_branch_policies -or
        $Environment.deployment_branch_policy.protected_branches -or
        @($Policies.branch_policies).Count -ne 1 -or
        $Policies.branch_policies[0].name -ne $DefaultBranch -or
        $Policies.branch_policies[0].type -ne 'branch') {
        throw 'Existing control environment must already allow only the default branch; do not overwrite its protections.'
    }
}

if ($MyInvocation.InvocationName -eq '.') { return }
if (-not $PSCmdlet.ShouldProcess($Repository, 'Configure protected GitHub/Azure workload identity and VM-only permissions')) { return }
if ($Repository -ne 'sbroenne/mcp-server-excel' -or $EnvironmentName -ne 'excel-runner-control') {
    throw 'This bootstrap is restricted to the dedicated Excel repository and control environment.'
}
$deadline = [DateTime]::UtcNow.AddMinutes(30)
$subscriptionArguments = @()
. (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'scripts\ExcelRunnerHost.ps1')
$account = Invoke-RunnerAzure @('account', 'show')
if (-not $account.id -or -not $account.tenantId) { throw 'Authenticate to the approved Azure subscription first.' }
$subscriptionArguments = @('--subscription', $account.id)
$vm = Assert-ExcelRunnerOwnedVm
if (@($vm.instanceView.statuses | Where-Object code -EQ 'PowerState/deallocated').Count -ne 1) {
    throw 'Control setup requires the VM to be parked.'
}
$repo = Invoke-ExcelRunnerGithub "repos/$Repository"
$pages = Invoke-ExcelRunnerGithub "repos/$Repository/environments?per_page=100" -Paginate
$environments = @($pages | ForEach-Object { $_.environments } | Where-Object name -EQ $EnvironmentName)
if ($environments.Count -gt 1) { throw 'Control environment ownership is ambiguous.' }
if ($environments.Count -eq 1) {
    $policies = Invoke-ExcelRunnerGithub "repos/$Repository/environments/$EnvironmentName/deployment-branch-policies"
    Assert-ExcelRunnerControlEnvironment $environments[0] $policies $repo.default_branch
}
$applicationName = 'excelmcp-github-runner-control'
$applications = @(Invoke-RunnerAzure @('ad', 'app', 'list', '--display-name', $applicationName))
if ($applications.Count -gt 1) { throw 'Control application name is ambiguous.' }
if ($applications.Count -eq 0) {
    $applicationPath = Join-Path ([IO.Path]::GetTempPath()) "excel-application-$([Guid]::NewGuid().ToString('N')).json"
    try {
        Get-ExcelRunnerControlApplicationDefinition | ConvertTo-Json -Depth 4 |
            Set-Content -LiteralPath $applicationPath -Encoding UTF8
        $application = Invoke-RunnerAzure @(
            'rest', '--method', 'POST', '--url', 'https://graph.microsoft.com/v1.0/applications',
            '--body', "@$applicationPath"
        )
    }
    finally { if (Test-Path -LiteralPath $applicationPath) { Remove-Item -LiteralPath $applicationPath -Force } }
}
else {
    $application = $applications[0]
    if ($application.tags -notcontains 'ExcelMcpRunnerControl') { throw 'Do not reuse an unrelated application with the same display name.' }
}
if (@($application.passwordCredentials).Count -or @($application.keyCredentials).Count) {
    throw 'The control application must not have permanent client secrets or certificate credentials.'
}
$principals = @(Invoke-RunnerAzure @('ad', 'sp', 'list', '--filter', "appId eq '$($application.appId)'"))
if ($principals.Count -gt 1) { throw 'Control service principal is ambiguous.' }
$principal = if ($principals.Count -eq 1) { $principals[0] } else { Invoke-RunnerAzure @('ad', 'sp', 'create', '--id', $application.appId) }
$credential = @{
    name = 'github-excel-runner-control'
    issuer = 'https://token.actions.githubusercontent.com'
    subject = "repo:$Repository`:environment:$EnvironmentName"
    audiences = @('api://AzureADTokenExchange')
}
$existing = @(Invoke-RunnerAzure @('ad', 'app', 'federated-credential', 'list', '--id', $application.appId))
if ($existing.Count -gt 1 -or ($existing.Count -eq 1 -and (
    $existing[0].name -ne $credential.name -or $existing[0].issuer -ne $credential.issuer -or
    $existing[0].subject -ne $credential.subject -or @($existing[0].audiences).Count -ne 1 -or
    $existing[0].audiences[0] -ne $credential.audiences[0]))) {
    throw 'Do not broaden or overwrite an unexpected federated trust.'
}
$path = Join-Path ([IO.Path]::GetTempPath()) "excel-control-$([Guid]::NewGuid().ToString('N')).json"
try {
    if ($existing.Count -eq 0) {
        $credential | ConvertTo-Json | Set-Content -LiteralPath $path -Encoding UTF8
        $null = Invoke-RunnerAzure @('ad', 'app', 'federated-credential', 'create', '--id', $application.appId, '--parameters', $path)
    }
    $null = Invoke-RunnerAzure @(
        'deployment', 'sub', 'create', '--name', 'excel-runner-control', '--location', $vm.location,
        '--template-file', (Join-Path $PSScriptRoot 'excel-runner-control.bicep'),
        '--parameters', "resourceGroupName=$ResourceGroup", "principalId=$($principal.id)"
    ) 600
    if ($environments.Count -eq 0) {
        $payload = @{
            deployment_branch_policy = @{ protected_branches = $false; custom_branch_policies = $true }
        } | ConvertTo-Json -Depth 4
        $payload | Set-Content -LiteralPath $path -Encoding UTF8
        & gh api "repos/$Repository/environments/$EnvironmentName" --method PUT --input $path | Out-Null
        if ($LASTEXITCODE -ne 0) { throw 'Protected control environment setup failed.' }
        @{ name = $repo.default_branch; type = 'branch' } | ConvertTo-Json | Set-Content -LiteralPath $path -Encoding UTF8
        & gh api "repos/$Repository/environments/$EnvironmentName/deployment-branch-policies" --method POST --input $path | Out-Null
        if ($LASTEXITCODE -ne 0) { throw 'Default-branch control restriction failed.' }
    }
    $environment = Invoke-ExcelRunnerGithub "repos/$Repository/environments/$EnvironmentName"
    $policies = Invoke-ExcelRunnerGithub "repos/$Repository/environments/$EnvironmentName/deployment-branch-policies"
    Assert-ExcelRunnerControlEnvironment $environment $policies $repo.default_branch
    foreach ($entry in @{
        EXCEL_RUNNER_AZURE_CLIENT_ID = $application.appId
        EXCEL_RUNNER_AZURE_TENANT_ID = $account.tenantId
        EXCEL_RUNNER_AZURE_SUBSCRIPTION_ID = $account.id
    }.GetEnumerator()) {
        & gh variable set $entry.Key --repo $Repository --env $EnvironmentName --body $entry.Value
        if ($LASTEXITCODE -ne 0) { throw 'Control environment variable setup failed.' }
    }
}
finally { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force } }
Write-Output 'Protected Azure federation configured. Coding routing and maintenance remain opt-in and disabled.'
