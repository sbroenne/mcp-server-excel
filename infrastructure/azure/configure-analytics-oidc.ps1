<#
.SYNOPSIS
    Configures read-only GitHub Actions OIDC access for usage analytics.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Repository = "sbroenne/mcp-server-excel",
    [string]$SubscriptionId,
    [string]$ResourceGroup = "excelmcp-observability",
    [string]$WorkspaceName = "excelmcp-logs",
    [string]$ApplicationName = "excelmcp-github-usage-analytics"
)

$ErrorActionPreference = "Stop"
function Set-GitHubVariable {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    if (-not $PSCmdlet.ShouldProcess(
            "$Repository variable '$Name'",
            "Set GitHub repository variable")) {
        return
    }

    gh variable set $Name --repo $Repository --body $Value
    if ($LASTEXITCODE -ne 0) {
        throw "Unable to configure GitHub repository variable '$Name'."
    }
}

if ($Repository -notmatch "^[^/]+/[^/]+$") {
    throw "Repository must use owner/name format."
}

$subscriptionArgs = if ($SubscriptionId) { @('--subscription', $SubscriptionId) } else { @() }
$account = az account show @subscriptionArgs --only-show-errors | ConvertFrom-Json
if ($LASTEXITCODE -ne 0 -or -not $account.id -or -not $account.tenantId) {
    throw "Unable to read the requested Azure subscription. Authenticate with az login first."
}
$SubscriptionId = $account.id

$workspace = az monitor log-analytics workspace show `
    --resource-group $ResourceGroup `
    --workspace-name $WorkspaceName `
    --subscription $SubscriptionId `
    --only-show-errors | ConvertFrom-Json
if ($LASTEXITCODE -ne 0 -or -not $workspace.id) {
    throw "Log Analytics workspace '$WorkspaceName' was not found."
}

$applications = @(
    az ad app list --display-name $ApplicationName --only-show-errors |
        ConvertFrom-Json
)
if ($LASTEXITCODE -ne 0) { throw "Unable to list Azure app registrations." }
if ($applications.Count -gt 1) {
    throw "Multiple app registrations named '$ApplicationName' exist."
}

if ($applications.Count -eq 0) {
    if (-not $PSCmdlet.ShouldProcess($ApplicationName, "Create Azure app registration")) {
        return
    }
    $application = az ad app create `
        --display-name $ApplicationName `
        --sign-in-audience AzureADMyOrg `
        --only-show-errors | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0) { throw "Unable to create Azure app registration." }
}
else {
    $application = $applications[0]
}
if (-not $application.appId) { throw "Azure app registration response is missing appId." }

$servicePrincipals = @(
    az ad sp list --filter "appId eq '$($application.appId)'" --only-show-errors |
        ConvertFrom-Json
)
if ($LASTEXITCODE -ne 0 -or $servicePrincipals.Count -gt 1) { throw "Unable to identify a single Azure service principal." }
if ($servicePrincipals.Count -eq 0) {
    if (-not $PSCmdlet.ShouldProcess(
            $ApplicationName,
            "Create Azure service principal")) {
        return
    }
    $servicePrincipal = az ad sp create `
        --id $application.appId `
        --only-show-errors | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0) { throw "Unable to create Azure service principal." }
}
else {
    $servicePrincipal = $servicePrincipals[0]
}
if (-not $servicePrincipal.id) { throw "Azure service principal response is missing id." }

$credentialName = "github-main-usage-analytics"
$credentials = @(
    az ad app federated-credential list `
        --id $application.appId `
        --only-show-errors | ConvertFrom-Json
)
if ($LASTEXITCODE -ne 0) { throw "Unable to list Azure federated credentials." }
if (-not ($credentials | Where-Object name -eq $credentialName)) {
    if (-not $PSCmdlet.ShouldProcess(
            $credentialName,
            "Create Azure federated credential")) {
        return
    }
    $credential = @{
        name = $credentialName
        issuer = "https://token.actions.githubusercontent.com"
        subject = "repo:$Repository`:ref:refs/heads/main"
        audiences = @("api://AzureADTokenExchange")
        description = "Weekly public usage analytics workflow"
    } | ConvertTo-Json
    $credentialPath = Join-Path ([IO.Path]::GetTempPath()) `
        "excelmcp-federated-credential-$([Guid]::NewGuid().ToString('N')).json"
    try {
        [IO.File]::WriteAllText(
            $credentialPath,
            $credential,
            [Text.UTF8Encoding]::new($false))
        az ad app federated-credential create `
            --id $application.appId `
            --parameters $credentialPath `
            --only-show-errors | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "Unable to create the GitHub federated credential."
        }
    }
    finally {
        Remove-Item -LiteralPath $credentialPath -Force -ErrorAction SilentlyContinue
    }
}

$existingAssignments = @(
    az role assignment list `
        --subscription $SubscriptionId `
        --assignee-object-id $servicePrincipal.id `
        --scope $workspace.id `
        --role "Log Analytics Reader" `
        --only-show-errors | ConvertFrom-Json
)
if ($LASTEXITCODE -ne 0) { throw "Unable to list Azure role assignments." }
if ($existingAssignments.Count -eq 0) {
    if (-not $PSCmdlet.ShouldProcess(
            $workspace.id,
            "Grant Log Analytics Reader to '$ApplicationName'")) {
        return
    }
    az role assignment create `
        --subscription $SubscriptionId `
        --assignee-object-id $servicePrincipal.id `
        --assignee-principal-type ServicePrincipal `
        --role "Log Analytics Reader" `
        --scope $workspace.id `
        --only-show-errors | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Unable to grant read-only Log Analytics access."
    }
}

Set-GitHubVariable -Name "AZURE_CLIENT_ID" -Value $application.appId
Set-GitHubVariable -Name "AZURE_TENANT_ID" -Value $account.tenantId
Set-GitHubVariable -Name "AZURE_SUBSCRIPTION_ID" -Value $account.id

if (-not $WhatIfPreference) {
    Write-Host "Configured read-only OIDC analytics access for $Repository."
    Write-Host "Add COPILOT_GITHUB_TOKEN separately as a repository secret."
}
