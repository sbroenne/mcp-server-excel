# NuGet Trusted Publishing Setup Guide

## Overview

The ExcelMcp MCP Server uses **NuGet Trusted Publishing** via OpenID Connect (OIDC) for secure, automated package publishing. This eliminates the need to manage API keys as GitHub secrets.

## What is Trusted Publishing?

Trusted Publishing is a secure method for publishing packages to NuGet.org that uses OpenID Connect (OIDC) tokens instead of long-lived API keys. When configured, GitHub Actions can authenticate directly with NuGet.org using short-lived tokens that are automatically generated and validated.

### Benefits

✅ **More Secure**: No long-lived API keys to manage or store  
✅ **Zero Maintenance**: No API key rotation needed  
✅ **Auditable**: All publishes tied to specific GitHub workflows  
✅ **Best Practice**: Recommended by NuGet.org and Microsoft  

## How It Works

```
┌─────────────────────────────────────────────────────────────────┐
│ 1. Release Published (GitHub)                                   │
│    └─> Tag: v1.0.4                                              │
└─────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│ 2. GitHub Actions Workflow Triggered                            │
│    └─> Generates OIDC token with claims:                        │
│        • Repository: sbroenne/mcp-server-excel                           │
│        • Workflow: publish-nuget.yml                             │
│        • Actor: (whoever triggered)                              │
└─────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│ 3. .NET CLI Publishes Package                                   │
│    └─> Sends OIDC token to NuGet.org                            │
└─────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│ 4. NuGet.org Validates Token                                    │
│    └─> Checks against trusted publisher configuration           │
└─────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│ 5. Package Published ✅                                          │
│    └─> Available at nuget.org/packages/ExcelMcp.McpServer      │
└─────────────────────────────────────────────────────────────────┘
```

## Initial Setup (Required)

### Step 1: First Package Publish

Trusted publishing requires the package to exist on NuGet.org before configuration. You need to publish version 1.0.0 (or any initial version) using an API key.

**Option A: Publish Manually**

1. Build the package locally:

   ```bash
   dotnet pack src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj -c Release -o ./nupkg
   ```

2. Publish using your existing NuGet API key:

   ```bash
   dotnet nuget push ./nupkg/ExcelMcp.McpServer.*.nupkg \
     --api-key YOUR_API_KEY \
     --source https://api.nuget.org/v3/index.json
   ```

**Option B: Use GitHub Actions Temporarily**

1. Add `NUGET_API_KEY` as a repository secret temporarily
2. Modify the workflow to use the API key for the first release:

   ```yaml
   - name: Publish to NuGet.org
     run: |
       dotnet nuget push $packagePath \
         --api-key ${{ secrets.NUGET_API_KEY }} \
         --source https://api.nuget.org/v3/index.json
   ```

3. Create and publish a release
4. After successful publish, remove the `--api-key` parameter and delete the secret

### Step 2: Configure Trusted Publisher on NuGet.org

Once the package exists on NuGet.org:

1. **Sign in to NuGet.org**
   - Go to <https://www.nuget.org>
   - Sign in with your Microsoft account

2. **Navigate to Package Management**
   - Go to <https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer/manage>
   - Or: Find your package → Click "Manage Package"

3. **Add Trusted Publisher**
   - Click on the "Trusted Publishers" tab
   - Click "Add Trusted Publisher" button

4. **Configure GitHub Actions Publisher**

   Enter the following values:

   | Field | Value |
   |-------|-------|
   | **Publisher Type** | GitHub Actions |
   | **Owner** | `sbroenne` |
   | **Repository** | `mcp-server-excel` |
   | **Workflow** | `publish-nuget.yml` |
   | **Environment** | *(leave empty)* |

5. **Save Configuration**
   - Click "Add" to save the trusted publisher
   - You should see the configuration listed

### Step 3: Verify Configuration

After configuration:

1. Create a new release (e.g., v1.0.4)
2. Watch the GitHub Actions workflow run
3. Verify the package publishes successfully without API keys
4. Check the package appears on NuGet.org

## Workflow Configuration

The `.github/workflows/publish-nuget.yml` file is already configured for trusted publishing:

```yaml
jobs:
  publish:
    runs-on: windows-latest
    permissions:
      contents: read
      id-token: write  # Required for OIDC token generation
    
    steps:
    # ... build steps ...
    
    - name: Publish to NuGet.org
      run: |
        dotnet nuget push $packagePath \
          --source https://api.nuget.org/v3/index.json \
          --skip-duplicate
        # No --api-key parameter needed!
```

### Key Configuration Elements

1. **Permission**: `id-token: write` - Required for GitHub to generate OIDC tokens
2. **No API Key**: The `dotnet nuget push` command doesn't need `--api-key` parameter
3. **Automatic**: The .NET CLI automatically uses OIDC authentication when available

## Troubleshooting

### Error: "Authentication failed"

**Cause**: Trusted publisher not configured or misconfigured on NuGet.org

**Solution**:

1. Verify the package exists on NuGet.org
2. Check trusted publisher configuration matches exactly:
   - Owner: `sbroenne`
   - Repository: `mcp-server-excel`
   - Workflow: `publish-nuget.yml`
3. Ensure `id-token: write` permission is set in workflow

### Error: "Package 'ExcelMcp.McpServer' does not exist"

**Cause**: Package not yet published to NuGet.org

**Solution**: Complete Step 1 (First Package Publish) using an API key

### Error: "The workflow 'publish-nuget.yml' is not trusted"

**Cause**: Workflow filename in trusted publisher config doesn't match

**Solution**:

1. Check the exact workflow filename in `.github/workflows/`
2. Update trusted publisher configuration if needed
3. Configuration is case-sensitive

### Workflow Succeeds but Package Not Updated

**Cause**: The `--skip-duplicate` flag prevents republishing existing versions

**Solution**: This is expected behavior. Create a new release with a new version tag.

## Maintenance

### No Ongoing Maintenance Required

Once configured, trusted publishing requires zero maintenance:

- ✅ No API keys to rotate
- ✅ No secrets to update
- ✅ No expiration dates to track
- ✅ Automatic authentication on every release

### Updating Configuration

If you need to change the workflow filename or repository structure:

1. Update the workflow file in the repository
2. Go to NuGet.org package management
3. Remove old trusted publisher
4. Add new trusted publisher with updated values

## Security Considerations

### Why Trusted Publishing is More Secure

**Traditional API Key Approach**:

- Long-lived secrets (6-12 months or never expire)
- Stored in GitHub secrets (potential for exposure)
- Requires manual rotation
- If leaked, valid until revoked

**Trusted Publishing Approach**:

- Short-lived OIDC tokens (minutes)
- Generated on-demand per workflow run
- Automatically validated against configuration
- No storage of secrets
- Cannot be reused or leaked effectively

### OIDC Token Claims

The OIDC token includes these claims that NuGet.org validates:

- `repository`: Must match configured repository
- `workflow`: Must match configured workflow file
- `actor`: GitHub user who triggered the workflow
- `ref`: Git reference (branch/tag)
- `repository_owner`: Must match configured owner

If any claim doesn't match the trusted publisher configuration, authentication fails.

## Alternative: Using API Key (Not Recommended)

If you prefer using traditional API keys instead of trusted publishing, here are the instructions:

### Step 1: Generate NuGet API Key

1. **Sign in to NuGet.org**
   - Go to <https://www.nuget.org>
   - Sign in with your Microsoft account

2. **Create API Key**
   - Go to <https://www.nuget.org/account/apikeys>
   - Click "Create" button
   - Configure the API key:
     - **Key Name**: `ExcelMcp GitHub Actions` (or similar descriptive name)
     - **Expiration**: 365 days (maximum, but requires rotation)
     - **Scopes**: Select "Push new packages and package versions"
     - **Glob Pattern**: `Sbroenne.ExcelMcp.*` (to limit to your packages)

3. **Copy API Key**
   - Copy the generated API key immediately (you won't see it again)
   - Store it securely for the next step

### Step 2: Configure GitHub Repository Secret

1. **Go to Repository Settings**
   - Navigate to <https://github.com/sbroenne/mcp-server-excel/settings/secrets/actions>
   - Or: Go to your repository → Settings → Secrets and variables → Actions

2. **Add Repository Secret**
   - Click "New repository secret"
   - **Name**: `NUGET_API_KEY`
   - **Secret**: Paste your API key from Step 1
   - Click "Add secret"

### Step 3: Update Workflow

Modify `.github/workflows/publish-nuget.yml` to use the API key:

```yaml
jobs:
  publish:
    runs-on: windows-latest
    permissions:
      contents: read
      # Remove: id-token: write  # Not needed for API key method
    
    steps:
    # ... other steps remain the same ...
    
    - name: Publish to NuGet.org
      run: |
        $version = "${{ steps.version.outputs.version }}"
        $packagePath = "nupkg/Sbroenne.ExcelMcp.McpServer.$version.nupkg"
        
        dotnet nuget push $packagePath `
          --api-key ${{ secrets.NUGET_API_KEY }} `
          --source https://api.nuget.org/v3/index.json `
          --skip-duplicate
      shell: pwsh
```

### Step 4: Test the Configuration

1. Create a new release to trigger the workflow
2. Verify the package publishes successfully
3. Check that the secret is not exposed in workflow logs

### Maintenance Requirements for API Key Method

❌ **Regular Maintenance Required**:
- API keys expire (maximum 365 days)
- Need to regenerate and update secret annually
- Monitor for key exposure or compromise
- Rotate keys if repository access changes

⚠️ **Security Considerations**:
- API keys are long-lived secrets
- If leaked, they remain valid until revoked
- Stored in GitHub secrets (potential attack vector)
- Manual rotation required

### Why Trusted Publishing is Preferred

| Aspect | Trusted Publishing | API Key |
|--------|-------------------|---------|
| **Security** | ✅ Short-lived OIDC tokens | ❌ Long-lived secrets |
| **Maintenance** | ✅ Zero maintenance | ❌ Annual rotation required |
| **Setup Complexity** | ⚠️ Requires initial package | ✅ Works immediately |
| **Audit Trail** | ✅ Full workflow traceability | ⚠️ Limited to API key usage |
| **Best Practice** | ✅ Microsoft/NuGet recommended | ❌ Legacy approach |

## References

- [NuGet Trusted Publishing Documentation](https://learn.microsoft.com/en-us/nuget/nuget-org/publish-a-package#trust-based-publishing)
- [GitHub OIDC Documentation](https://docs.github.com/en/actions/deployment/security-hardening-your-deployments/about-security-hardening-with-openid-connect)
- [.NET CLI dotnet nuget push](https://learn.microsoft.com/en-us/dotnet/core/tools/dotnet-nuget-push)

## Support

If you encounter issues:

1. Check the [Troubleshooting](#troubleshooting) section above
2. Review GitHub Actions workflow logs for detailed error messages
3. Verify trusted publisher configuration on NuGet.org
4. Open an issue at <https://github.com/sbroenne/mcp-server-excel/issues>

---

**Status**: ✅ Configured for trusted publishing  
**Package**: <https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer>  
**Workflow**: `.github/workflows/publish-nuget.yml`
