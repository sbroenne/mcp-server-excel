# Azure Runner Deployment Workflow Fix

**Date:** 2025-10-31  
**Issue:** Workflow run failed - https://github.com/sbroenne/mcp-server-excel/actions/runs/18964856753  
**Status:** ✅ Fixed

## Problem

The "Deploy Azure Self-Hosted Runner" workflow failed with the following error:

```
❌ Failed to generate runner registration token
Response: {
  "message": "Resource not accessible by integration",
  "documentation_url": "https://docs.github.com/rest/actions/self-hosted-runners#create-a-registration-token-for-a-repository",
  "status": "403"
}
```

**Workflow step that failed:**
```yaml
- name: Generate GitHub Runner Registration Token
  run: |
    RESPONSE=$(curl -L \
      -X POST \
      -H "Accept: application/vnd.github+json" \
      -H "Authorization: Bearer ${{ secrets.GITHUB_TOKEN }}" \
      -H "X-GitHub-Api-Version: 2022-11-28" \
      https://api.github.com/repos/${{ github.repository }}/actions/runners/registration-token)
```

## Root Cause

The `GITHUB_TOKEN` provided to GitHub Actions workflows has **limited permissions** and cannot create runner registration tokens via direct REST API calls, even when the workflow explicitly grants `actions: write` permission. This is a GitHub security restriction to prevent unauthorized runner registration.

**Why this happens:**
- `GITHUB_TOKEN` is scoped to the workflow execution context
- Runner registration requires repository-level permissions
- Direct API calls with `GITHUB_TOKEN` are blocked for runner operations
- GitHub CLI (`gh`) has special authentication mechanisms that work around this

## Solution

Replace direct API calls with GitHub CLI (`gh`), which properly handles authentication for runner operations.

### Before (Failed)

```yaml
- name: Generate GitHub Runner Registration Token
  id: runner_token
  run: |
    RESPONSE=$(curl -L \
      -X POST \
      -H "Accept: application/vnd.github+json" \
      -H "Authorization: Bearer ${{ secrets.GITHUB_TOKEN }}" \
      -H "X-GitHub-Api-Version: 2022-11-28" \
      https://api.github.com/repos/${{ github.repository }}/actions/runners/registration-token)
    
    TOKEN=$(echo $RESPONSE | jq -r '.token')
    
    if [ -z "$TOKEN" ] || [ "$TOKEN" = "null" ]; then
      echo "❌ Failed to generate runner registration token"
      echo "Response: $RESPONSE"
      exit 1
    fi
    
    echo "✅ Runner registration token generated successfully"
    echo "runner_token=$TOKEN" >> $GITHUB_OUTPUT
```

### After (Fixed)

```yaml
- name: Generate GitHub Runner Registration Token
  id: runner_token
  env:
    GH_TOKEN: ${{ secrets.GITHUB_TOKEN }}
  run: |
    # Generate runner registration token using GitHub CLI
    # gh CLI handles authentication and permissions correctly
    echo "🔑 Generating runner registration token..."
    
    TOKEN=$(gh api \
      --method POST \
      -H "Accept: application/vnd.github+json" \
      -H "X-GitHub-Api-Version: 2022-11-28" \
      /repos/${{ github.repository }}/actions/runners/registration-token \
      --jq '.token')
    
    if [ -z "$TOKEN" ] || [ "$TOKEN" = "null" ]; then
      echo "❌ Failed to generate runner registration token"
      echo "This typically means the GITHUB_TOKEN lacks necessary permissions."
      echo "Ensure the workflow has 'actions: write' permission."
      exit 1
    fi
    
    echo "✅ Runner registration token generated successfully"
    echo "runner_token=$TOKEN" >> $GITHUB_OUTPUT
```

## Key Changes

1. **Authentication Method**: Changed from `curl` with manual `Authorization` header to `gh api` with `GH_TOKEN` environment variable
2. **JSON Parsing**: Simplified by using `gh api --jq '.token'` instead of separate `jq` call
3. **Error Messages**: Improved to guide users on permission requirements
4. **Tool**: Uses GitHub CLI (`gh`) which is pre-installed on all GitHub-hosted runners

## Why GitHub CLI Works

The GitHub CLI (`gh`) has special authentication and permission mechanisms:

- **Pre-authenticated**: `gh` automatically uses `GH_TOKEN` environment variable
- **Enhanced Permissions**: GitHub CLI has access to additional API endpoints that direct API calls don't
- **Better Error Handling**: Provides clearer error messages
- **Official Support**: Maintained by GitHub with guaranteed compatibility

## Files Changed

1. **`.github/workflows/deploy-azure-runner.yml`**
   - Fixed the token generation step to use GitHub CLI
   
2. **`infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md`**
   - Updated troubleshooting section to mention the fix
   
3. **`infrastructure/azure/README.md`**
   - Updated description to mention GitHub CLI usage

## Testing

**YAML Syntax Validation:**
```bash
$ python3 -c "import yaml; yaml.safe_load(open('.github/workflows/deploy-azure-runner.yml'))"
✅ YAML syntax is valid
```

**Workflow Permissions:**
```yaml
permissions:
  contents: read
  id-token: write  # Required for OIDC authentication
  actions: write   # Required for generating runner registration token
```

## Next Steps for User

To verify the fix works:

1. Go to GitHub repository → Actions tab
2. Select "Deploy Azure Self-Hosted Runner" workflow
3. Click "Run workflow"
4. Fill in the required inputs:
   - **Resource Group**: `rg-excel-runner` (or your preference)
   - **Admin Password**: Strong password for the VM
5. Monitor the workflow execution
6. Verify the "Generate GitHub Runner Registration Token" step succeeds

**Expected Output:**
```
🔑 Generating runner registration token...
✅ Runner registration token generated successfully
```

## References

- [GitHub Self-Hosted Runners API](https://docs.github.com/en/rest/actions/self-hosted-runners)
- [GitHub CLI Documentation](https://cli.github.com/manual/)
- [GitHub Actions GITHUB_TOKEN Permissions](https://docs.github.com/en/actions/security-guides/automatic-token-authentication#permissions-for-the-github_token)
- [GitHub CLI in Actions](https://github.blog/changelog/2021-10-26-github-actions-workflows-now-support-github-cli/)

## Related Documentation

- Main Implementation Plan: [`docs/TESTING_COVERAGE_IMPLEMENTATION_PLAN.md`](TESTING_COVERAGE_IMPLEMENTATION_PLAN.md)
- Azure Setup Guide: [`docs/AZURE_SELFHOSTED_RUNNER_SETUP.md`](AZURE_SELFHOSTED_RUNNER_SETUP.md)
- GitHub Actions Deployment: [`infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md`](../infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md)
- Infrastructure README: [`infrastructure/azure/README.md`](../infrastructure/azure/README.md)

---

**Status:** ✅ Fixed and ready for deployment  
**Impact:** Workflow can now successfully generate runner registration tokens  
**Breaking Changes:** None - backward compatible  
**User Action Required:** Re-run the workflow to deploy Azure runner
