# NuGet Trusted Publishing Setup Checklist

Quick reference for configuring NuGet Trusted Publishing for the ExcelMcp MCP Server.

## Prerequisites ✅

- [ ] Package exists on NuGet.org (must publish v1.0.0+ manually first)
- [ ] You have owner/admin access to the package on NuGet.org
- [ ] You have admin access to the GitHub repository
- [ ] You know your NuGet.org username (profile name, NOT email)

## GitHub Repository Configuration ✅

### Step 1: Add NUGET_USER Secret

- [ ] Go to <https://github.com/sbroenne/mcp-server-excel/settings/secrets/actions>
- [ ] Click "New repository secret"
- [ ] Name: `NUGET_USER`
- [ ] Value: Your NuGet.org username (profile name, NOT email address)
- [ ] Click "Add secret"

### Step 2: Verify Workflow Configuration

The workflow is **already configured** in `.github/workflows/release-mcp-server.yml`:

- [x] `id-token: write` permission added
- [x] `NuGet/login@v1` action for OIDC authentication
- [x] Short-lived API key from login action output
- [x] Correct source URL (`https://api.nuget.org/v3/index.json`)
- [x] Documentation comments added

## NuGet.org Configuration (Required)

Follow these steps on NuGet.org:

### Step 1: Navigate to Package Management

- [ ] Go to <https://www.nuget.org>
- [ ] Sign in with your Microsoft account
- [ ] Go to <https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer/manage>

### Step 2: Add Trusted Publisher

- [ ] Click on the "Trusted Publishers" tab
- [ ] Click "Add Trusted Publisher" button
- [ ] Select "GitHub Actions" as Publisher Type

### Step 3: Configure GitHub Actions Publisher

Enter these exact values:

| Field | Value | Status |
|-------|-------|--------|
| **Publisher Type** | GitHub Actions | ⬜ |
| **Owner** | `sbroenne` | ⬜ |
| **Repository** | `mcp-server-excel` | ⬜ |
| **Workflow** | `release-mcp-server.yml` | ⬜ |
| **Environment** | *(leave empty)* | ⬜ |

- [ ] Click "Add" to save configuration
- [ ] Verify the trusted publisher appears in the list

## Testing

### Step 1: Create Test Release

- [ ] Create a new tag (e.g., `mcp-v1.0.5`)
- [ ] Push the tag to trigger the workflow
- [ ] Monitor GitHub Actions workflow run

### Step 2: Verify Publish Success

- [ ] Workflow completes successfully
- [ ] No authentication errors in logs
- [ ] Package appears on NuGet.org
- [ ] Version number matches the release tag

## Troubleshooting

If authentication fails:

- [ ] Verify trusted publisher configuration matches exactly
- [ ] Check workflow has `id-token: write` permission
- [ ] Ensure no `--api-key` parameter in publish command
- [ ] Confirm package exists on NuGet.org before setup
- [ ] Review detailed logs in GitHub Actions

## Security Cleanup

After Trusted Publishing is working:

- [ ] Old `NUGET_API_KEY` secret can remain deleted (no longer used with Trusted Publishing)
- [ ] ✅ `NUGET_USER` secret configured (required for NuGet/login action)
- [ ] ✅ Revoke old long-lived API keys on NuGet.org (if any)
- [ ] ✅ Document the setup in team documentation

## Benefits Achieved

Once configured, you have:

- ✅ Short-lived API keys (generated per workflow run, expire immediately)
- ✅ No long-lived API keys to manage or rotate
- ✅ Automatic authentication via OIDC
- ✅ Enhanced security with token-based authentication
- ✅ Full audit trail of all publishes
- ✅ Minimal maintenance required (only `NUGET_USER` secret)

## Reference

- **Full Documentation**: `docs/NUGET_TRUSTED_PUBLISHING.md`
- **Workflow File**: `.github/workflows/release-mcp-server.yml`
- **Package URL**: <https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer>
- **Microsoft Docs**: <https://learn.microsoft.com/en-us/nuget/nuget-org/publish-a-package#trust-based-publishing>

---

**Last Updated**: October 20, 2025  
**Status**: Workflow configured ✅ | NuGet.org configuration needed ⬜
