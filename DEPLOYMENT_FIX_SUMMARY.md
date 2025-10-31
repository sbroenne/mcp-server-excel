# GitHub Runner Deployment Fix - Summary

## Problem Identified

The GitHub Actions workflow `deploy-azure-runner.yml` was failing with a **403 Forbidden** error when attempting to create a runner registration token:

```
❌ Failed to generate runner registration token
Response: {
  "message": "Resource not accessible by integration",
  "documentation_url": "https://docs.github.com/rest/actions/self-hosted-runners#create-a-registration-token-for-a-repository",
  "status": "403"
}
```

## Root Cause

The default `GITHUB_TOKEN` provided by GitHub Actions **does not have permission** to create runner registration tokens via the API endpoint `/repos/{owner}/{repo}/actions/runners/registration-token`.

According to GitHub's documentation, this endpoint requires either:
- A Personal Access Token (PAT) with appropriate permissions
- A GitHub App access token with runner management permissions

The `actions: write` permission in the workflow file only grants permission for workflow-related actions, **not** for managing self-hosted runners.

## Solution Implemented

### 1. Workflow Changes (`.github/workflows/deploy-azure-runner.yml`)

**Changed:**
- Replaced `secrets.GITHUB_TOKEN` with `secrets.RUNNER_ADMIN_PAT` (a user-created PAT)
- Removed `actions: write` permission (not needed with PAT)
- Added validation to check if PAT is configured before attempting token generation
- Enhanced error messages with clear troubleshooting steps

**Before:**
```yaml
env:
  GH_TOKEN: ${{ secrets.GITHUB_TOKEN }}
```

**After:**
```yaml
env:
  GH_TOKEN: ${{ secrets.RUNNER_ADMIN_PAT }}
```

### 2. Documentation Updates

#### `infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md`

**Added Step 1: Create GitHub Personal Access Token**
- Detailed instructions for creating a fine-grained PAT
- Required permissions: `Actions: Read and write` + `Administration: Read and write`
- Alternative classic token option with `repo` scope
- Security best practices for PAT management

**Updated Step 3: Add Secrets**
- Added `RUNNER_ADMIN_PAT` to the secrets table
- Clarified PAT is required before running deployment

**Enhanced Troubleshooting**
- New section for missing PAT error
- New section for invalid/expired PAT error
- Updated with specific error scenarios and solutions

**Added Security Section**
- PAT rotation schedule (90 days)
- Token management best practices
- Monitoring and revocation procedures

#### `infrastructure/azure/README.md`

**Updated Quick Start**
- Added PAT creation as first step
- Clarified security model (OIDC for Azure, PAT for runner registration)

## What You Need to Do

### Step 1: Create Personal Access Token

1. Go to GitHub: **Settings** → **Developer settings** → **Personal access tokens** → **Fine-grained tokens**
2. Click **Generate new token**
3. Configure:
   - **Token name**: `excel-runner-deployment`
   - **Expiration**: 90 days (recommended)
   - **Repository access**: Only `sbroenne/mcp-server-excel`
   - **Permissions**:
     - Repository permissions → **Actions**: Read and write
     - Repository permissions → **Administration**: Read and write
4. Click **Generate token** and **copy it immediately**

**Alternative - Classic Token (simpler):**
- Generate classic token with `repo` scope
- Less secure but easier to configure

### Step 2: Add Token to GitHub Secrets

1. Go to repository: `https://github.com/sbroenne/mcp-server-excel`
2. Navigate to **Settings** → **Secrets and variables** → **Actions**
3. Click **New repository secret**
   - **Name**: `RUNNER_ADMIN_PAT`
   - **Secret**: Paste the PAT you created
4. Click **Add secret**

### Step 3: Test the Deployment

1. Go to **Actions** tab
2. Select **Deploy Azure Self-Hosted Runner** workflow
3. Click **Run workflow**
4. Fill in parameters:
   - **Resource Group**: `rg-excel-runner` (or your preference)
   - **Admin Password**: Strong password for VM
5. Click **Run workflow**

The workflow should now:
- ✅ Successfully generate runner registration token using your PAT
- ✅ Deploy Azure resources via OIDC (no client secret needed)
- ✅ Configure the VM with GitHub runner
- ✅ Provide next steps for Excel installation

## Verification

After deployment completes:

1. **Check workflow logs** - Should see "✅ Runner registration token generated successfully"
2. **Verify runner registration** - Go to `https://github.com/sbroenne/mcp-server-excel/settings/actions/runners`
3. **Should see**: Runner named `azure-excel-runner` with status "Offline" (until you install Excel and reboot)

## Security Considerations

### PAT Token Management

- **Expiration**: Tokens expire after 90 days (or your chosen period)
- **Rotation**: Set calendar reminder to rotate before expiration
- **Storage**: Never commit to code, only store in GitHub Secrets
- **Monitoring**: Check token usage in GitHub Settings → Developer settings
- **Revocation**: Revoke immediately if compromised

### Why This Approach?

**Pros:**
- ✅ Automated runner token generation (no manual copy/paste)
- ✅ OIDC for Azure (no client secret stored)
- ✅ Repeatable deployments via GitHub UI
- ✅ Works with GitHub Coding Agents
- ✅ Audit trail in GitHub Actions logs

**Cons:**
- ⚠️ Requires PAT creation (one-time setup)
- ⚠️ PAT needs rotation every 90 days (can set longer)

**Alternative Considered and Rejected:**
- ❌ Manual runner token generation - expires in 1 hour, not automatable
- ❌ Organization-level GitHub App - requires org admin, more complex setup
- ❌ Enterprise runner controller - overkill for single repository

## Files Changed

1. `.github/workflows/deploy-azure-runner.yml` - Updated to use PAT
2. `infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md` - Added PAT setup guide
3. `infrastructure/azure/README.md` - Updated quick start with PAT step

## References

- [GitHub REST API - Self-hosted runners](https://docs.github.com/en/rest/actions/self-hosted-runners)
- [GitHub Fine-grained PATs](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/managing-your-personal-access-tokens#creating-a-fine-grained-personal-access-token)
- [GitHub Community - PAT for runner registration](https://github.com/orgs/community/discussions/120232)
- [Azure Login Action (OIDC)](https://github.com/Azure/login#login-with-openid-connect-oidc-recommended)

## Next Steps After Deployment

Once the VM is deployed and runner is registered:

1. **RDP to VM** using the FQDN from deployment output
2. **Install Office 365 Excel** (30 minutes)
   - Sign in to https://portal.office.com
   - Install Office apps (Excel only)
   - Activate with your Office 365 account
3. **Reboot VM** - Runner service starts automatically
4. **Verify runner is online** - Check GitHub repository settings
5. **Run integration tests** - Trigger workflow to test Excel COM automation

---

**Questions or Issues?**

Check the troubleshooting section in `infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md` or create an issue in the repository.
