<#
.SYNOPSIS
    Deploys Application Insights infrastructure for ExcelMcp telemetry.

.DESCRIPTION
    Deploys a resource group with Log Analytics Workspace and Application Insights
    for collecting usage analytics and crash reports from the MCP Server.

.PARAMETER Location
    Azure region for deployment. Default: swedencentral

.PARAMETER ParameterFile
    Path to the parameters JSON file. Default: appinsights.parameters.json

.PARAMETER WhatIf
    Shows what would be deployed without making changes.

.EXAMPLE
    .\deploy-appinsights.ps1

.EXAMPLE
    .\deploy-appinsights.ps1 -Location "eastus" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string]$Location = "swedencentral",

    [Parameter()]
    [string]$ParameterFile = "appinsights.parameters.json",

    [string]$SubscriptionId
)

$ErrorActionPreference = "Stop"

# Ensure we're in the right directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $scriptDir

try {
    Write-Host "`n=== ExcelMcp Application Insights Deployment ===" -ForegroundColor Cyan
    Write-Host "Location: $Location"
    Write-Host "Parameters: $ParameterFile`n"

    # Check prerequisites
    Write-Host "Checking prerequisites..." -ForegroundColor Yellow

    # Check Azure CLI
    $azVersion = az version 2>$null | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0 -or -not $azVersion.'azure-cli') {
        throw "Azure CLI not found. Install from: https://aka.ms/installazurecli"
    }
    Write-Host "  Azure CLI: $($azVersion.'azure-cli')" -ForegroundColor Green

    # Check logged in
    $subscriptionArgs = if ($SubscriptionId) { @('--subscription', $SubscriptionId) } else { @() }
    $account = az account show @subscriptionArgs --only-show-errors | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0 -or -not $account.id) {
        throw "Unable to read the requested Azure subscription. Authenticate with az login first."
    }
    Write-Host "  Subscription: $($account.name) ($($account.id))" -ForegroundColor Green

    # Validate template
    Write-Host "`nValidating Bicep template..." -ForegroundColor Yellow
    $validation = az deployment sub validate `
        --subscription $account.id `
        --location $Location `
        --template-file "appinsights.bicep" `
        --parameters $ParameterFile `
        2>&1

    if ($LASTEXITCODE -ne 0) {
        throw "Template validation failed: $validation"
    }
    Write-Host "  Template is valid" -ForegroundColor Green

    # Deploy
    if ($WhatIfPreference -or $PSCmdlet.ShouldProcess($account.id, "Deploy Application Insights infrastructure")) {

        if ($WhatIfPreference) {
            Write-Host "`nWhat-If deployment (no changes will be made)..." -ForegroundColor Yellow
            az deployment sub what-if `
                --subscription $account.id `
                --location $Location `
                --template-file "appinsights.bicep" `
                --parameters $ParameterFile
            if ($LASTEXITCODE -ne 0) { throw "Azure deployment preview failed." }
        }
        else {
            Write-Host "`nDeploying resources..." -ForegroundColor Yellow
            $deployment = az deployment sub create `
                --subscription $account.id `
                --location $Location `
                --template-file "appinsights.bicep" `
                --parameters $ParameterFile `
                --name "excelmcp-appinsights-$(Get-Date -Format 'yyyyMMdd-HHmmss')" `
                | ConvertFrom-Json

            if ($LASTEXITCODE -ne 0) {
                throw "Deployment failed"
            }

            # Extract outputs
            $outputs = $deployment.properties.outputs
            $connectionString = $outputs.appInsightsConnectionString.value
            $instrumentationKey = $outputs.appInsightsInstrumentationKey.value
            $resourceGroup = $outputs.resourceGroupName.value
            $appInsightsName = $outputs.appInsightsName.value
            if (-not $connectionString -or -not $instrumentationKey -or -not $resourceGroup -or -not $appInsightsName) {
                throw "Deployment response is missing required Application Insights outputs."
            }

            Write-Host "`n=== Deployment Successful ===" -ForegroundColor Green
            Write-Host "Resource Group: $resourceGroup"
            Write-Host "Application Insights: $appInsightsName"
            Write-Host ""
            Write-Host "View telemetry at: https://portal.azure.com/#@/resource/subscriptions/$($account.id)/resourceGroups/$resourceGroup/providers/Microsoft.Insights/components/$appInsightsName/overview"
            Write-Host ""

            # Save connection string to file for reference (gitignored)
            $secretsFile = "appinsights.secrets.local"
            @{
                ConnectionString = $connectionString
                InstrumentationKey = $instrumentationKey
                ResourceGroup = $resourceGroup
                AppInsightsName = $appInsightsName
                DeployedAt = (Get-Date).ToString("o")
            } | ConvertTo-Json | Out-File $secretsFile -Encoding utf8

            Write-Host "Connection details saved to ignored local file: $secretsFile. Do not commit or publish it." -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Host "`nError: $_" -ForegroundColor Red
    exit 1
}
finally {
    Pop-Location
}
