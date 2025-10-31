# GitHub Actions Automated Deployment Setup

This guide shows how to deploy the Azure VM using GitHub Actions instead of running scripts locally.

## Prerequisites

- Azure subscription with permissions to create service principals
- GitHub repository admin access
- Office 365 license for Excel

## Setup (One-Time)

### Step 1: Create Azure Service Principal

**Using Azure CLI:**

```bash
# Login to Azure
az login

# Get your subscription ID
SUBSCRIPTION_ID=$(az account show --query id --output tsv)

# Create service principal with Contributor role
az ad sp create-for-rbac \
  --name "github-excel-runner-deployer" \
  --role Contributor \
  --scopes /subscriptions/$SUBSCRIPTION_ID \
  --sdk-auth
```

**Output (save this JSON):**
```json
{
  "clientId": "xxx",
  "clientSecret": "xxx",
  "subscriptionId": "xxx",
  "tenantId": "xxx",
  ...
}
```

### Step 2: Add Azure Credentials to GitHub Secrets

1. Go to your repository: `https://github.com/sbroenne/mcp-server-excel`
2. Navigate to **Settings** → **Secrets and variables** → **Actions**
3. Click **New repository secret**
4. Create secret:
   - **Name:** `AZURE_CREDENTIALS`
   - **Value:** Paste the entire JSON from Step 1
5. Click **Add secret**

### Step 3: Generate GitHub Runner Token

1. Go to `https://github.com/sbroenne/mcp-server-excel/settings/actions/runners`
2. Click **New self-hosted runner**
3. Select **Windows**
4. Copy the registration token (valid for 1 hour)

## Deployment

### Deploy via GitHub Actions UI

1. Go to **Actions** tab in your repository
2. Select **Deploy Azure Self-Hosted Runner** workflow
3. Click **Run workflow**
4. Fill in the parameters:
   - **Resource Group:** `rg-excel-runner` (or your preference)
   - **Admin Password:** Strong password for VM (e.g., `MySecurePass123!`)
   - **GitHub Runner Token:** Paste token from Step 3 above
5. Click **Run workflow**

**Deployment takes ~5 minutes**

### After Deployment

1. Check workflow run for VM FQDN (displayed in logs)
2. RDP to the VM using the FQDN and credentials
3. Install Office 365 Excel (30 minutes):
   - Sign in to https://portal.office.com
   - Install Office 365 apps
   - Select Excel only during installation
   - Activate with your Office 365 account
4. Reboot VM
5. Runner auto-starts and registers with GitHub

### Verify Deployment

**Check runner status:**
```
https://github.com/sbroenne/mcp-server-excel/settings/actions/runners
```

Should show:
- Name: `azure-excel-runner`
- Status: Idle (green)
- Labels: `self-hosted`, `windows`, `excel`

## Alternative: OIDC (Federated Credentials)

For better security (no secrets stored), use OIDC authentication:

### Setup OIDC

```bash
# Create app registration
APP_ID=$(az ad app create --display-name "github-excel-runner-oidc" --query appId -o tsv)

# Create service principal
az ad sp create --id $APP_ID

# Add federated credential
az ad app federated-credential create \
  --id $APP_ID \
  --parameters '{
    "name": "github-excel-runner",
    "issuer": "https://token.actions.githubusercontent.com",
    "subject": "repo:sbroenne/mcp-server-excel:ref:refs/heads/main",
    "audiences": ["api://AzureADTokenExchange"]
  }'

# Assign Contributor role
SUBSCRIPTION_ID=$(az account show --query id --output tsv)
az role assignment create \
  --assignee $APP_ID \
  --role Contributor \
  --scope /subscriptions/$SUBSCRIPTION_ID
```

### GitHub Secrets for OIDC

Create these secrets:
- `AZURE_CLIENT_ID`: From `$APP_ID`
- `AZURE_TENANT_ID`: From `az account show --query tenantId -o tsv`
- `AZURE_SUBSCRIPTION_ID`: From `az account show --query id -o tsv`

Update workflow to use OIDC (commented lines in `deploy-azure-runner.yml`)

## Troubleshooting

### "Deployment failed" error

**Check:**
1. Azure credentials are correct
2. Service principal has Contributor role
3. Subscription has quota for B2ms VMs in Sweden Central

**View detailed error:**
- Check workflow logs in Actions tab
- Look for error messages in "Deploy Bicep Template" step

### "Runner registration failed" error

**Causes:**
1. GitHub runner token expired (valid 1 hour)
2. Token was already used

**Solution:**
- Generate new token from repository settings
- Re-run workflow with new token

### Azure Login failed

**Check:**
1. `AZURE_CREDENTIALS` secret contains valid JSON
2. Service principal exists: `az ad sp list --display-name "github-excel-runner-deployer"`
3. Service principal has Contributor role

## Cost

- **Monthly:** ~$61 (24/7 operation in Sweden Central)
- **One-time setup:** Free (GitHub Actions minutes)

## Cleanup

To delete all resources:

```bash
az group delete --name rg-excel-runner --yes --no-wait
```

Or use Azure Portal → Resource Groups → Delete

## Security Best Practices

1. **Use OIDC** instead of service principal credentials (more secure)
2. **Rotate secrets** every 90 days
3. **Limit service principal** to specific resource group
4. **Enable Azure Security Center** for VM monitoring
5. **Review permissions** regularly

## Support

- **Azure Issues:** Check workflow logs in Actions tab
- **Repository Issues:** https://github.com/sbroenne/mcp-server-excel/issues
- **Azure Docs:** https://docs.microsoft.com/azure/developer/github/connect-from-azure

---

**Deployment time:** 5 min (automated) + 30 min (Excel install)  
**Cost:** ~$61/month
