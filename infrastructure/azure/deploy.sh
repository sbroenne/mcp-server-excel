#!/bin/bash
# Deploy Azure Excel Integration Test Runner
# Usage: ./deploy.sh <resource-group-name> <admin-password> <github-runner-token>

set -e

RESOURCE_GROUP=${1:-"rg-excel-runner"}
ADMIN_PASSWORD=${2}
GITHUB_RUNNER_TOKEN=${3}
LOCATION="swedencentral"
GITHUB_REPO_URL="https://github.com/sbroenne/mcp-server-excel"

if [ -z "$ADMIN_PASSWORD" ] || [ -z "$GITHUB_RUNNER_TOKEN" ]; then
    echo "Usage: ./deploy.sh <resource-group-name> <admin-password> <github-runner-token>"
    echo ""
    echo "To generate GitHub runner token:"
    echo "1. Go to https://github.com/sbroenne/mcp-server-excel/settings/actions/runners/new"
    echo "2. Select Windows"
    echo "3. Copy the token from the configuration command"
    exit 1
fi

echo "🚀 Deploying Excel Integration Test Runner..."
echo "   Resource Group: $RESOURCE_GROUP"
echo "   Location: $LOCATION"
echo "   VM Size: Standard_B2ms (2 vCPUs, 8 GB RAM)"
echo "   Monthly Cost: ~$61 (24/7) or ~$36 (12h/day with auto-shutdown)"

# Create resource group
echo "📦 Creating resource group..."
az group create --name "$RESOURCE_GROUP" --location "$LOCATION"

# Deploy Bicep template
echo "🏗️  Deploying infrastructure..."
az deployment group create \
  --resource-group "$RESOURCE_GROUP" \
  --template-file azure-runner.bicep \
  --parameters \
    location="$LOCATION" \
    adminPassword="$ADMIN_PASSWORD" \
    githubRepoUrl="$GITHUB_REPO_URL" \
    githubRunnerToken="$GITHUB_RUNNER_TOKEN"

# Get VM public IP
VM_FQDN=$(az deployment group show \
  --resource-group "$RESOURCE_GROUP" \
  --name azure-runner \
  --query 'properties.outputs.vmPublicIP.value' \
  --output tsv)

echo ""
echo "✅ Deployment complete!"
echo ""
echo "📋 Next Steps:"
echo "1. RDP to VM: $VM_FQDN"
echo "2. Username: azureuser"
echo "3. Install Office 365 Excel from https://portal.office.com"
echo "4. Activate Excel with your Office 365 account"
echo "5. Runner will auto-start after reboot"
echo ""
echo "⚠️  IMPORTANT: GitHub Actions CANNOT auto-start stopped VMs"
echo "   - Keep VM running 24/7 for immediate workflow execution (~$61/month)"
echo "   - OR manually start VM each day (~$36/month with auto-shutdown)"
echo "   - OR set up Azure Automation for scheduled start (see README)"
echo ""
echo "💰 Monthly Cost: ~$61 (24/7) or ~$36 (12h/day) in Sweden Central"
echo ""
echo "🔍 Verify runner status:"
echo "   https://github.com/sbroenne/mcp-server-excel/settings/actions/runners"
