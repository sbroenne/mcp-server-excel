# NuGet Publishing Guide for ExcelMcp

Complete guide for publishing and managing all ExcelMcp NuGet packages using OIDC Trusted Publishing.

> **Distribution Channels:** NuGet is the **secondary** distribution channel. The **primary** channel is standalone self-contained executables distributed via GitHub Releases — no .NET runtime required. See [INSTALLATION.md](INSTALLATION.md) for the recommended installation methods.

## Table of Contents

- [Published Packages](#published-packages)
- [NuGet Trusted Publishing Overview](#nuget-trusted-publishing-overview)
- [Initial Setup](#initial-setup)
- [Release Process](#release-process)
- [Version Numbering Strategy](#version-numbering-strategy)
- [Package Testing](#package-testing)
- [Troubleshooting](#troubleshooting)
- [Security & Maintenance](#security--maintenance)

---

## Published Packages

ExcelMcp publishes two NuGet packages (unified release):

### 1. Sbroenne.ExcelMcp.McpServer (.NET Tool)
- **Package Type**: .NET Global Tool (executable)
- **Purpose**: MCP server for AI assistant integration
- **Tag Pattern**: `v*` (e.g., `v1.2.0`) - **unified with CLI**
- **Workflow**: `.github/workflows/release.yml` (handles both packages)
- **NuGet Page**: https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer
- **Installation**: `dotnet tool install --global Sbroenne.ExcelMcp.McpServer`

### 2. Sbroenne.ExcelMcp.CLI (.NET Tool)
- **Package Type**: .NET Global Tool (executable)
- **Purpose**: Command-line interface for Excel automation
- **Tag Pattern**: `v*` (e.g., `v1.2.0`) - **unified with MCP Server**
- **Workflow**: `.github/workflows/release.yml` (handles both packages)
- **NuGet Page**: https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI
- **Installation**: `dotnet tool install --global Sbroenne.ExcelMcp.CLI`

**Note**: MCP Server and CLI are always released together with the same version number, both from the single unified `release.yml` workflow. Core and ComInterop are internal library dependencies — they are **not** published to NuGet.

---

## NuGet Trusted Publishing Overview

All packages use **NuGet Trusted Publishing** via OpenID Connect (OIDC) for secure, automated package publishing without API keys.

### What is Trusted Publishing?

Trusted Publishing uses short-lived OIDC tokens instead of long-lived API keys for authentication with NuGet.org.

### Benefits

✅ **More Secure**: No long-lived API keys to manage or store  
✅ **Zero Maintenance**: No API key rotation needed  
✅ **Auditable**: All publishes tied to specific GitHub workflows  
✅ **Best Practice**: Recommended by NuGet.org and Microsoft  

### How It Works

```
1. Git Tag Pushed (e.g., v1.2.2)
   ↓
2. GitHub Actions Workflow Triggered (release.yml)
   └─> Generates OIDC token with claims:
       • Repository: sbroenne/mcp-server-excel
       • Workflow: release.yml
       • Actor: (whoever triggered)
   ↓
3. NuGet Login Action Exchanges OIDC Token
   └─> Receives short-lived API key
   ↓
4. .NET CLI Publishes Both Packages
   └─> Uses short-lived API key
   └─> Publishes MCP Server
   └─> Publishes CLI
   ↓
5. NuGet.org Validates Token
   └─> Checks against trusted publisher configuration
   ↓
6. Packages Published ✅
   └─> Available at nuget.org/packages/[PackageId]
```

---

## Initial Setup

### Step 1: Configure GitHub Secret

Add your NuGet.org username as a repository secret (one-time setup):

1. **Go to Repository Settings**
   - Navigate to: https://github.com/sbroenne/mcp-server-excel/settings/secrets/actions
   - Or: Repository → Settings → Secrets and variables → Actions

2. **Add Repository Secret**
   - Click "New repository secret"
   - **Name**: `NUGET_USER`
   - **Secret**: Your NuGet.org username (profile name, **NOT email**)
   - Click "Add secret"

### Step 2: First-Time Package Publishing

Trusted publishing requires packages to exist on NuGet.org before configuration.

**Option A: Manual Publishing (Recommended)**

```powershell
# Build the package
dotnet pack src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -c Release -o ./nupkg

# Publish using your NuGet API key (first time only)
dotnet nuget push ./nupkg/Sbroenne.ExcelMcp.CLI.1.0.0.nupkg \
  --api-key YOUR_API_KEY \
  --source https://api.nuget.org/v3/index.json
```

**Option B: Temporary Workflow API Key**

1. Add `NUGET_API_KEY` as repository secret temporarily
2. Modify workflow to use API key for first release
3. Create and publish release
4. Configure trusted publisher (Step 3)
5. Remove API key and restore OIDC authentication

### Step 3: Configure Trusted Publishers on NuGet.org

For **each package**, configure a trusted publisher:

#### MCP Server

1. Go to: https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer/manage
2. Click "Trusted Publishers" tab → "Add Trusted Publisher"
3. Select "GitHub Actions"
4. Enter:
   - **Owner**: `sbroenne`
   - **Repository**: `mcp-server-excel`
   - **Workflow**: `release.yml`
   - **Environment**: *(leave empty)*
5. Click "Add"

#### CLI Tool

1. Go to: https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI/manage
2. Same steps, use workflow: `release.yml`

### Step 4: Verify Configuration

After configuration:

1. Create a test release tag
2. Watch GitHub Actions workflow run
3. Verify package publishes without API keys
4. Check package appears on NuGet.org

---

## Release Process

### Publishing Order

**MCP Server and CLI are released together** via the unified release workflow (`release.yml`). Core and ComInterop are internal dependencies not separately published to NuGet.

```
MCP Server + CLI (released together via release.yml workflow_dispatch)
```

### Standard Release Commands

Releases are triggered via the GitHub Actions UI, not by pushing package-specific tags:

1. Go to **Actions → Release All Components → Run workflow**
2. Select a version bump type (patch/minor/major) or enter a custom version
3. Run the workflow — it builds, packs, and publishes both NuGet packages, creates the GitHub release, and tags the commit (e.g. `v1.2.2`)
4. Monitor the run at https://github.com/sbroenne/mcp-server-excel/actions
5. Verify packages on NuGet.org (wait 5-10 minutes for indexing) and test installation

### Quick Release (All Components with Single Tag)

```powershell
# Create and push unified tag - releases ALL components (MCP Server, CLI, VS Code Extension, MCPB)
git tag v1.2.2 -m "Release v1.2.2"
git push origin v1.2.2
```

---

## Version Numbering Strategy

All packages follow **Semantic Versioning (SemVer)**:

- **MAJOR** (1.x.x): Breaking API changes
- **MINOR** (x.1.x): New features, backward compatible
- **PATCH** (x.x.1): Bug fixes, backward compatible

### Version Alignment Strategy

**MCP Server and CLI always share the same version number** — they are released together from the unified `release.yml` workflow and both are stamped with the tag's version (e.g. `v1.2.0`). Core and ComInterop are internal library dependencies built from the same commit; they are not independently versioned or published, so there is no cross-package version drift to manage.

---

## Package Testing

### Build Packages Locally

```powershell
# Build the published packages
dotnet pack src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj -c Release -o ./nupkg
dotnet pack src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -c Release -o ./nupkg
```

### Test Local Installation

```powershell
# Install .NET tool from local package
dotnet tool install --global Sbroenne.ExcelMcp.CLI --add-source ./nupkg --version 1.0.0

# Test the tool
excelcli --help

# Uninstall
dotnet tool uninstall --global Sbroenne.ExcelMcp.CLI
```

### Validate Package Contents

```powershell
# Extract package (NuGet packages are ZIP files)
unzip -l ./nupkg/Sbroenne.ExcelMcp.CLI.1.0.0.nupkg

# Verify:
# - README.md is included
# - LICENSE is included
# - DLLs are present
# - Dependencies are correct in .nuspec
```

---

## Troubleshooting

### Package Not Appearing on NuGet.org

**Wait Time**: NuGet.org indexing takes 5-10 minutes after publishing.

**Check Workflow Logs**:
1. Go to: https://github.com/sbroenne/mcp-server-excel/actions
2. Find the release workflow run
3. Check "Publish to NuGet.org" step for errors

### Trusted Publishing Authentication Failed

**Cause**: Trusted publisher not configured or misconfigured

**Solution**:
1. Verify package exists on NuGet.org
2. Check trusted publisher configuration matches exactly:
   - Owner: `sbroenne`
   - Repository: `mcp-server-excel`
   - Workflow: `release-[package].yml` (exact filename)
3. Ensure `NUGET_USER` secret is set correctly
4. Verify workflow has `id-token: write` permission

### Error: "Package does not exist"

**Cause**: Package not yet published to NuGet.org

**Solution**: Complete Step 2 (First-Time Package Publishing) using an API key

### Error: "Workflow is not trusted"

**Cause**: Workflow filename in trusted publisher config doesn't match

**Solution**:
1. Check exact workflow filename in `.github/workflows/`
2. Update trusted publisher configuration if needed
3. Configuration is case-sensitive

### Package Validation Errors

All packages have `EnablePackageValidation=true` which catches issues during build.

**Common Issues**:
- Missing README or LICENSE
- Incorrect package metadata
- Missing dependencies

**Solution**: Check build output for validation warnings/errors

---

## Security & Maintenance

### Security Benefits of Trusted Publishing

**vs. Traditional API Keys:**

| Aspect | Trusted Publishing | API Key |
|--------|-------------------|---------|
| **Security** | ✅ Short-lived tokens (minutes) | ❌ Long-lived secrets (up to 1 year) |
| **Maintenance** | ✅ Zero maintenance | ❌ Annual rotation required |
| **Setup** | ⚠️ Requires initial package | ✅ Works immediately |
| **Audit** | ✅ Full workflow traceability | ⚠️ Limited to API key usage |
| **Best Practice** | ✅ Microsoft/NuGet recommended | ❌ Legacy approach |
| **Storage** | ✅ No stored secrets | ❌ Stored in GitHub secrets |
| **Leak Risk** | ✅ Expires in minutes | ❌ Valid until revoked |

### OIDC Token Claims

The OIDC token includes validated claims:

- `repository`: Must match configured repository
- `workflow`: Must match configured workflow file
- `actor`: GitHub user who triggered workflow
- `ref`: Git reference (branch/tag)
- `repository_owner`: Must match configured owner

If any claim doesn't match trusted publisher configuration, authentication fails.

### Zero Maintenance Required

Once configured:
- ✅ No API keys to rotate
- ✅ No secrets to update
- ✅ No expiration dates to track
- ✅ Automatic authentication on every release

### Updating Configuration

If you rename a workflow file:
1. Update workflow file in repository
2. Go to NuGet.org package management
3. Remove old trusted publisher
4. Add new trusted publisher with updated workflow name

---

## Monitoring Releases

### GitHub Actions

- **Workflow Runs**: https://github.com/sbroenne/mcp-server-excel/actions
- Each release workflow creates:
  - NuGet package upload
  - GitHub release with notes
  - Binary assets (for MCP Server and CLI)

### NuGet.org Package Pages

- **MCP Server**: https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer
- **CLI**: https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI

### Download Statistics

Monitor package adoption via NuGet.org statistics pages (linked above).

---

## References

- [NuGet Trusted Publishing Documentation](https://learn.microsoft.com/en-us/nuget/nuget-org/publish-a-package#trust-based-publishing)
- [GitHub OIDC Documentation](https://docs.github.com/en/actions/deployment/security-hardening-your-deployments/about-security-hardening-with-openid-connect)
- [Semantic Versioning](https://semver.org/)
- [.NET Global Tools](https://learn.microsoft.com/en-us/dotnet/core/tools/global-tools)
- [.NET CLI dotnet nuget push](https://learn.microsoft.com/en-us/dotnet/core/tools/dotnet-nuget-push)

---

## Support

For issues with NuGet publishing:

1. Check this guide's troubleshooting section
2. Review GitHub Actions workflow logs
3. Verify NuGet.org trusted publisher configuration
4. Open an issue at: https://github.com/sbroenne/mcp-server-excel/issues

---

**Status**: ✅ Both packages configured for trusted publishing  
**Workflows**: Unified `release.yml` publishes both packages together  
**Security**: OIDC trusted publishing eliminates API key management
