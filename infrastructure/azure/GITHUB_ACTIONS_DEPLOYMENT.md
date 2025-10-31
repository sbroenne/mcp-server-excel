# GitHub Actions Automated Deployment Setup

This guide shows how to deploy the Azure VM using GitHub Actions with **OIDC (OpenID Connect)** - the secure, modern approach with no secrets stored.

## Prerequisites

- Azure subscription with permissions to create app registrations
- GitHub repository admin access
- Office 365 license for Excel

## Setup (One-Time)

### Step 1: Create GitHub Personal Access Token (PAT)

The workflow needs a PAT to generate runner registration tokens. The default `GITHUB_TOKEN` lacks this permission.

**Create a Fine-Grained Personal Access Token:**

1. Go to GitHub: **Settings** → **Developer settings** → **Personal access tokens** → **Fine-grained tokens**
2. Click **Generate new token**
3. Configure the token:
   - **Token name**: `excel-runner-deployment`
   - **Expiration**: 90 days (or longer, you'll need to renew it)
   - **Repository access**: Select "Only select repositories" → Choose `mcp-server-excel`
   - **Permissions**:
     - Repository permissions → **Actions**: Read and write
     - Repository permissions → **Administration**: Read and write (required for self-hosted runners)
4. Click **Generate token**
5. **Copy the token immediately** (you won't see it again)

**Add the token to GitHub Secrets:**

1. Go to your repository: `https://github.com/sbroenne/mcp-server-excel`
2. Navigate to **Settings** → **Secrets and variables** → **Actions**
3. Click **New repository secret**
   - Name: `RUNNER_ADMIN_PAT`
   - Secret: Paste the PAT you just created
4. Click **Add secret**

**Alternative - Classic Token (simpler but less secure):**

If you prefer a classic token:
1. Go to **Settings** → **Developer settings** → **Personal access tokens** → **Tokens (classic)**
2. Generate new token with `repo` scope (includes all repository permissions)
3. Add as `RUNNER_ADMIN_PAT` secret

⚠️ **Security Note**: This PAT grants significant permissions. Treat it like a password:
- Never commit it to code
- Rotate it every 90 days
- Revoke it if compromised
- Only grant to repositories you control

### Step 2: Create Azure App Registration with Federated Credentials (OIDC)

**Using Azure CLI:**

```bash
# Login to Azure
az login

# Get your subscription ID
SUBSCRIPTION_ID=$(az account show --query id --output tsv)

# Create app registration
APP_ID=$(az ad app create \
  --display-name "github-excel-runner-oidc" \
  --query appId \
  --output tsv)

echo "App ID: $APP_ID"

# Create service principal
az ad sp create --id $APP_ID

# Add federated credential for GitHub Actions
az ad app federated-credential create \
  --id $APP_ID \
  --parameters "{
    \"name\": \"github-excel-runner\",
    \"issuer\": \"https://token.actions.githubusercontent.com\",
    \"subject\": \"repo:sbroenne/mcp-server-excel:ref:refs/heads/main\",
    \"audiences\": [\"api://AzureADTokenExchange\"]
  }"

# Assign Contributor role to subscription
az role assignment create \
  --assignee $APP_ID \
  --role Contributor \
  --scope /subscriptions/$SUBSCRIPTION_ID

# Get tenant ID
TENANT_ID=$(az account show --query tenantId --output tsv)

echo ""
echo "✅ Setup complete! Add these to GitHub Secrets:"
echo "AZURE_CLIENT_ID: $APP_ID"
echo "AZURE_TENANT_ID: $TENANT_ID"
echo "AZURE_SUBSCRIPTION_ID: $SUBSCRIPTION_ID"
```

**Using Azure Portal:**

1. Go to **Azure Active Directory** → **App registrations**
2. Click **New registration**
   - Name: `github-excel-runner-oidc`
   - Click **Register**
3. Note the **Application (client) ID** and **Directory (tenant) ID**
4. Go to **Certificates & secrets** → **Federated credentials**
5. Click **Add credential**
   - Federated credential scenario: **GitHub Actions deploying Azure resources**
   - Organization: `sbroenne`
   - Repository: `mcp-server-excel`
   - Entity type: **Branch**
   - GitHub branch name: `main`
   - Name: `github-excel-runner`
   - Click **Add**
6. Go to **Subscriptions** → Select your subscription → **Access control (IAM)**
7. Click **Add role assignment**
   - Role: **Contributor**
   - Assign access to: **User, group, or service principal**
   - Select: `github-excel-runner-oidc`
   - Click **Review + assign**

### Step 3: Add Azure Information to GitHub Secrets

1. Go to your repository: `https://github.com/sbroenne/mcp-server-excel`
2. Navigate to **Settings** → **Secrets and variables** → **Actions**
3. Click **New repository secret** for each:

| Secret Name | Value | Where to Find |
|-------------|-------|---------------|
| `RUNNER_ADMIN_PAT` | Personal Access Token | From Step 1 - the PAT you created |
| `AZURE_CLIENT_ID` | Application (client) ID | From Step 2 or App Registration overview |
| `AZURE_TENANT_ID` | Directory (tenant) ID | From Step 2 or Azure AD overview |
| `AZURE_SUBSCRIPTION_ID` | Subscription ID | From Step 2 or Subscriptions page |

**Important**: The `RUNNER_ADMIN_PAT` must be created **before** running the deployment workflow.

**No Azure client secret needed!** OIDC uses federated credentials instead.

## Deployment

### Deploy via GitHub Actions UI

1. Go to **Actions** tab in your repository
2. Select **Deploy Azure Self-Hosted Runner** workflow
3. Click **Run workflow**
4. Fill in the parameters:
   - **Resource Group:** `rg-excel-runner` (or your preference)
   - **Admin Password:** Strong password for VM (e.g., `MySecurePass123!`)
5. Click **Run workflow**

**Note:** GitHub runner registration token is now **automatically generated** by the workflow - no manual token creation needed!

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

## Why This Approach?

**Security benefits:**
- ✅ **Minimal secrets stored** - Uses OIDC for Azure (no client secret), PAT only for runner registration
- ✅ **No Azure credential rotation** - Federated credentials don't expire (PAT needs 90-day renewal)
- ✅ **Automatic runner token generation** - No manual token handling during deployment
- ✅ **Azure-managed** - Azure AD handles authentication
- ✅ **Audit trail** - Every deployment logged in Azure AD and GitHub Actions
- ✅ **Principle of least privilege** - Scoped to specific repository/branch

**vs. Manual Token Generation:**
- ❌ Manual runner token generation required each time
- ❌ Runner tokens expire after 1 hour
- ❌ Error-prone copy/paste process
- ❌ Cannot be used by coding agents

**Note**: While we do require a PAT for runner registration, this is a one-time setup that GitHub renews automatically during workflow execution. The PAT is more secure than storing runner tokens and allows for automated, repeatable deployments.

## Troubleshooting

### "RUNNER_ADMIN_PAT secret is not set" error

**Cause:** The required Personal Access Token is missing or not configured.

**Solution:**
1. Follow Step 1 in this guide to create a PAT
2. Add it as a repository secret named `RUNNER_ADMIN_PAT`
3. Re-run the workflow

### "Failed to generate runner registration token" error

**Causes:**
1. PAT expired or is invalid
2. PAT lacks required permissions
3. You don't have admin access to the repository

**Solution:**
1. Check the PAT expiration date in GitHub Settings → Developer settings → Personal access tokens
2. Verify PAT has these permissions:
   - **Fine-grained**: Actions (Read and write) + Administration (Read and write)
   - **Classic**: `repo` scope
3. Ensure you're a repository admin
4. If PAT expired, create a new one and update the `RUNNER_ADMIN_PAT` secret
5. Re-run the workflow

**Check workflow logs:**
- Go to Actions tab → Failed workflow run
- Look at "Generate GitHub Runner Registration Token" step for specific error details

### "Deployment failed" error

**Check:**
1. Azure credentials are correct
2. Service principal has Contributor role
3. Subscription has quota for B2ms VMs in Sweden Central

**View detailed error:**
- Check workflow logs in Actions tab
- Look for error messages in "Deploy Bicep Template" step

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

1. **Use OIDC for Azure** instead of service principal credentials (more secure)
2. **PAT Token Management**:
   - Rotate PAT tokens every 90 days
   - Use fine-grained tokens with minimum required permissions
   - Never commit PAT tokens to code
   - Revoke immediately if compromised
   - Monitor token usage in GitHub Settings → Developer settings
3. **Automatic runner token generation** eliminates manual token handling risks
4. **Limit service principal** to specific resource group
5. **Enable Azure Security Center** for VM monitoring
6. **Review permissions** regularly

### PAT Token Rotation Schedule

| Action | Frequency | Steps |
|--------|-----------|-------|
| Create new PAT | Every 90 days | Follow Step 1 in Setup section |
| Update secret | When PAT expires | Update `RUNNER_ADMIN_PAT` in GitHub Secrets |
| Revoke old PAT | After updating secret | GitHub Settings → Developer settings → Revoke token |
| Test deployment | After rotation | Run deployment workflow to verify |

## Support

- **Azure Issues:** Check workflow logs in Actions tab
- **Repository Issues:** https://github.com/sbroenne/mcp-server-excel/issues
- **Azure Docs:** https://docs.microsoft.com/azure/developer/github/connect-from-azure

---

**Deployment time:** 5 min (automated) + 30 min (Excel install)  
**Cost:** ~$61/month
