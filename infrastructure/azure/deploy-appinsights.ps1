<#
.SYNOPSIS
    Deploys Application Insights infrastructure for ExcelMcp telemetry.

.DESCRIPTION
    Deploys a resource group with Log Analytics Workspace and Application Insights
    for collecting usage analytics and crash reports from the MCP Server.

.PARAMETER Location
    Azure region for deployment. Default: westeurope

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
    [string]$ParameterFile = "appinsights.parameters.json"
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
    if (-not $azVersion) {
        throw "Azure CLI not found. Install from: https://aka.ms/installazurecli"
    }
    Write-Host "  Azure CLI: $($azVersion.'azure-cli')" -ForegroundColor Green

    # Check logged in
    $account = az account show 2>$null | ConvertFrom-Json
    if (-not $account) {
        Write-Host "  Not logged in. Running 'az login'..." -ForegroundColor Yellow
        az login
        $account = az account show | ConvertFrom-Json
    }
    Write-Host "  Subscription: $($account.name) ($($account.id))" -ForegroundColor Green

    # Validate template
    Write-Host "`nValidating Bicep template..." -ForegroundColor Yellow
    $validation = az deployment sub validate `
        --location $Location `
        --template-file "appinsights.bicep" `
        --parameters $ParameterFile `
        2>&1

    if ($LASTEXITCODE -ne 0) {
        throw "Template validation failed: $validation"
    }
    Write-Host "  Template is valid" -ForegroundColor Green

    # Deploy
    if ($WhatIf -or $PSCmdlet.ShouldProcess("Azure subscription", "Deploy Application Insights infrastructure")) {

        if ($WhatIf) {
            Write-Host "`nWhat-If deployment (no changes will be made)..." -ForegroundColor Yellow
            az deployment sub what-if `
                --location $Location `
                --template-file "appinsights.bicep" `
                --parameters $ParameterFile
        }
        else {
            Write-Host "`nDeploying resources..." -ForegroundColor Yellow
            $deployment = az deployment sub create `
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

            Write-Host "`n=== Deployment Successful ===" -ForegroundColor Green
            Write-Host "Resource Group: $resourceGroup"
            Write-Host "Application Insights: $appInsightsName"
            Write-Host ""
            Write-Host "Connection String (for embedding in code):" -ForegroundColor Cyan
            Write-Host $connectionString -ForegroundColor White
            Write-Host ""
            Write-Host "Instrumentation Key (legacy):" -ForegroundColor Cyan
            Write-Host $instrumentationKey -ForegroundColor White
            Write-Host ""
            Write-Host "=== Next Steps ===" -ForegroundColor Yellow
            Write-Host "1. Copy the connection string above"
            Write-Host "2. Add it to src/ExcelMcp.McpServer/Telemetry/ExcelMcpTelemetry.cs"
            Write-Host "3. Build and test the MCP Server"
            Write-Host "4. View telemetry at: https://portal.azure.com/#@/resource/subscriptions/$($account.id)/resourceGroups/$resourceGroup/providers/Microsoft.Insights/components/$appInsightsName/overview"
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

            Write-Host "Connection string saved to: $secretsFile (add to .gitignore!)" -ForegroundColor Yellow
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
