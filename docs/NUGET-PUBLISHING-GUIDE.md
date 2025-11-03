# NuGet Package Publishing Guide

This guide explains how to publish all ExcelMcp packages to NuGet.org.

## Published Packages

ExcelMcp consists of four NuGet packages:

### 1. Sbroenne.ExcelMcp.ComInterop (Library)
- **Package Type**: Library (.dll)
- **Purpose**: Low-level COM interop utilities for Excel automation
- **Tag Pattern**: `cominterop-v*` (e.g., `cominterop-v1.0.0`)
- **Workflow**: `.github/workflows/release-cominterop.yml`
- **NuGet Page**: https://www.nuget.org/packages/Sbroenne.ExcelMcp.ComInterop
- **Installation**: `dotnet add package Sbroenne.ExcelMcp.ComInterop`

### 2. Sbroenne.ExcelMcp.Core (Library)
- **Package Type**: Library (.dll)
- **Purpose**: High-level Excel automation commands
- **Dependencies**: Sbroenne.ExcelMcp.ComInterop
- **Tag Pattern**: `core-v*` (e.g., `core-v1.0.0`)
- **Workflow**: `.github/workflows/release-core.yml`
- **NuGet Page**: https://www.nuget.org/packages/Sbroenne.ExcelMcp.Core
- **Installation**: `dotnet add package Sbroenne.ExcelMcp.Core`

### 3. Sbroenne.ExcelMcp.McpServer (.NET Tool)
- **Package Type**: .NET Global Tool (executable)
- **Purpose**: MCP server for AI assistant integration
- **Dependencies**: Sbroenne.ExcelMcp.Core
- **Tag Pattern**: `mcp-v*` (e.g., `mcp-v1.2.0`)
- **Workflow**: `.github/workflows/release-mcp-server.yml`
- **NuGet Page**: https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer
- **Installation**: `dotnet tool install --global Sbroenne.ExcelMcp.McpServer`

### 4. Sbroenne.ExcelMcp.CLI (.NET Tool)
- **Package Type**: .NET Global Tool (executable)
- **Purpose**: Command-line interface for Excel automation
- **Dependencies**: Sbroenne.ExcelMcp.Core
- **Tag Pattern**: `cli-v*` (e.g., `cli-v2.1.0`)
- **Workflow**: `.github/workflows/release-cli.yml`
- **NuGet Page**: https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI
- **Installation**: `dotnet tool install --global Sbroenne.ExcelMcp.CLI`

## Publishing Order

When releasing multiple packages with breaking changes, follow this order to ensure dependencies are available:

```
1. ComInterop (foundation layer)
   ↓
2. Core (depends on ComInterop)
   ↓
3. MCP Server and/or CLI (both depend on Core)
```

**Example:**
```bash
# 1. Release ComInterop first
git tag cominterop-v1.1.0
git push origin cominterop-v1.1.0

# 2. Wait for NuGet publishing to complete, then release Core
git tag core-v1.1.0
git push origin core-v1.1.0

# 3. Wait for Core publishing, then release MCP Server and CLI
git tag mcp-v1.3.0
git push origin mcp-v1.3.0

git tag cli-v2.2.0
git push origin cli-v2.2.0
```

## NuGet Trusted Publishing Setup

All packages use **NuGet Trusted Publishing** via OpenID Connect (OIDC) for secure, keyless authentication.

### Required Configuration

#### 1. GitHub Repository Secret

Add your NuGet.org username as a repository secret:

1. Go to: https://github.com/sbroenne/mcp-server-excel/settings/secrets/actions
2. Click "New repository secret"
3. **Name**: `NUGET_USER`
4. **Secret**: Your NuGet.org profile name (NOT your email)
5. Click "Add secret"

#### 2. NuGet.org Trusted Publishers

For each package, configure a trusted publisher on NuGet.org:

**Sbroenne.ExcelMcp.ComInterop:**
- Package URL: https://www.nuget.org/packages/Sbroenne.ExcelMcp.ComInterop/manage
- Trusted Publisher → Add → GitHub Actions
  - Owner: `sbroenne`
  - Repository: `mcp-server-excel`
  - Workflow: `release-cominterop.yml`
  - Environment: *(leave empty)*

**Sbroenne.ExcelMcp.Core:**
- Package URL: https://www.nuget.org/packages/Sbroenne.ExcelMcp.Core/manage
- Trusted Publisher → Add → GitHub Actions
  - Owner: `sbroenne`
  - Repository: `mcp-server-excel`
  - Workflow: `release-core.yml`
  - Environment: *(leave empty)*

**Sbroenne.ExcelMcp.McpServer:**
- Package URL: https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer/manage
- Trusted Publisher → Add → GitHub Actions
  - Owner: `sbroenne`
  - Repository: `mcp-server-excel`
  - Workflow: `release-mcp-server.yml`
  - Environment: *(leave empty)*

**Sbroenne.ExcelMcp.CLI:**
- Package URL: https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI/manage
- Trusted Publisher → Add → GitHub Actions
  - Owner: `sbroenne`
  - Repository: `mcp-server-excel`
  - Workflow: `release-cli.yml`
  - Environment: *(leave empty)*

### Initial Package Publishing

Trusted publishing requires the package to exist on NuGet.org before configuration. For first-time publishing:

**Option 1: Manual Publishing (Recommended for first release)**

```bash
# Build the package
dotnet pack src/ExcelMcp.ComInterop/ExcelMcp.ComInterop.csproj -c Release -o ./nupkg

# Publish using your NuGet API key (first time only)
dotnet nuget push ./nupkg/Sbroenne.ExcelMcp.ComInterop.1.0.0.nupkg \
  --api-key YOUR_API_KEY \
  --source https://api.nuget.org/v3/index.json
```

**Option 2: Temporary Workflow API Key**

1. Add `NUGET_API_KEY` as a repository secret temporarily
2. Modify workflow to use API key for first release
3. Create and publish a release
4. Configure trusted publisher on NuGet.org
5. Remove API key and restore OIDC authentication

After the first release, all subsequent releases use OIDC trusted publishing automatically.

## Release Process

### Standard Release (Individual Package)

```bash
# 1. Ensure main branch is up to date
git checkout main
git pull

# 2. Create a release tag
git tag [package-prefix]-v[version]
# Examples:
# git tag cominterop-v1.0.0
# git tag core-v1.0.0
# git tag mcp-v1.2.0
# git tag cli-v2.1.0

# 3. Push the tag to trigger workflow
git push origin [tag-name]

# 4. Monitor workflow
# - Go to: https://github.com/sbroenne/mcp-server-excel/actions
# - Watch the release workflow run
# - Verify NuGet publishing succeeds
# - Verify GitHub release is created

# 5. Verify package on NuGet.org
# - Wait 5-10 minutes for NuGet indexing
# - Check package page for new version
# - Test installation: dotnet tool install --global [PackageId] --version [version]
```

### Coordinated Multi-Package Release

When releasing multiple packages with interdependencies:

```bash
# 1. Release ComInterop first (if updated)
git tag cominterop-v1.1.0
git push origin cominterop-v1.1.0
# Wait for workflow to complete and NuGet to index

# 2. Release Core (if updated)
git tag core-v1.1.0
git push origin core-v1.1.0
# Wait for workflow to complete and NuGet to index

# 3. Release MCP Server and/or CLI (if updated)
git tag mcp-v1.3.0
git push origin mcp-v1.3.0

git tag cli-v2.2.0
git push origin cli-v2.2.0
```

## Version Numbering

Each package has independent versioning following Semantic Versioning (SemVer):

- **MAJOR** version (1.x.x): Breaking API changes
- **MINOR** version (x.1.x): New features, backward compatible
- **PATCH** version (x.x.1): Bug fixes, backward compatible

### Version Alignment Guidance

While packages have independent versions, consider aligning major versions for clarity:

- **ComInterop v1.x.x** → **Core v1.x.x** → **MCP Server v1.x.x**, **CLI v1.x.x**

Breaking changes in ComInterop or Core should trigger major version bumps in dependent packages.

## Package Testing

Before releasing to NuGet.org, test packages locally:

### Build Packages

```bash
# Build all packages
dotnet pack src/ExcelMcp.ComInterop/ExcelMcp.ComInterop.csproj -c Release -o ./nupkg
dotnet pack src/ExcelMcp.Core/ExcelMcp.Core.csproj -c Release -o ./nupkg
dotnet pack src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj -c Release -o ./nupkg
dotnet pack src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -c Release -o ./nupkg
```

### Test Local Installation

```bash
# Install .NET tool from local package
dotnet tool install --global Sbroenne.ExcelMcp.CLI --add-source ./nupkg --version 1.0.0

# Test the tool
excelcli --help

# Uninstall
dotnet tool uninstall --global Sbroenne.ExcelMcp.CLI
```

### Validate Package Contents

```bash
# Extract package (NuGet packages are ZIP files)
unzip -l ./nupkg/Sbroenne.ExcelMcp.Core.1.0.0.nupkg

# Verify:
# - README.md is included
# - LICENSE is included
# - DLLs are present
# - Dependencies are correct in .nuspec
```

## Troubleshooting

### Package Not Appearing on NuGet.org

**Wait Time**: NuGet.org indexing takes 5-10 minutes after publishing.

**Check Workflow Logs**:
1. Go to GitHub Actions
2. Find the release workflow run
3. Check "Publish to NuGet.org" step for errors

### Trusted Publishing Authentication Failed

**Cause**: Trusted publisher not configured or misconfigured

**Solution**:
1. Verify package exists on NuGet.org
2. Check trusted publisher configuration matches exactly
3. Ensure `NUGET_USER` secret is set correctly
4. Verify workflow has `id-token: write` permission

### Dependency Version Mismatch

**Cause**: Dependent package references outdated version

**Solution**:
1. Release ComInterop first
2. Wait for NuGet indexing
3. Release Core (references latest ComInterop)
4. Wait for indexing
5. Release MCP Server/CLI (reference latest Core)

### Package Validation Errors

All packages have `EnablePackageValidation=true` which catches issues during build.

**Common Issues**:
- Missing README or LICENSE
- Incorrect package metadata
- Missing dependencies

**Solution**: Check build output for validation warnings/errors

## Monitoring Releases

### GitHub Actions

- **Workflow Runs**: https://github.com/sbroenne/mcp-server-excel/actions
- Each release workflow creates:
  - NuGet package upload
  - GitHub release with notes
  - Binary assets (for MCP Server and CLI)

### NuGet.org

- **ComInterop**: https://www.nuget.org/packages/Sbroenne.ExcelMcp.ComInterop
- **Core**: https://www.nuget.org/packages/Sbroenne.ExcelMcp.Core
- **MCP Server**: https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer
- **CLI**: https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI

### Download Statistics

Monitor package adoption via NuGet.org statistics pages (linked above).

## Maintenance

### Zero Maintenance with Trusted Publishing

Once configured, trusted publishing requires:
- ✅ No API key rotation
- ✅ No secret updates
- ✅ No expiration management
- ✅ Automatic authentication on every release

### Updating Workflows

If you rename a workflow file:
1. Update the workflow file in repository
2. Go to NuGet.org package management
3. Remove old trusted publisher
4. Add new trusted publisher with updated workflow name

## Security Best Practices

### Why Trusted Publishing is Secure

- **Short-lived tokens**: Generated per-workflow run, expire in minutes
- **No stored secrets**: OIDC tokens are not stored anywhere
- **Automatic validation**: NuGet.org validates workflow identity
- **Audit trail**: All publishes tied to specific workflow runs

### Traditional API Key Comparison

| Aspect | Trusted Publishing | API Key |
|--------|-------------------|---------|
| Security | ✅ Short-lived tokens | ❌ Long-lived secrets |
| Maintenance | ✅ Zero maintenance | ❌ Annual rotation |
| Setup | ⚠️ Requires initial package | ✅ Works immediately |
| Audit | ✅ Full workflow trace | ⚠️ Limited to API key usage |
| Best Practice | ✅ Microsoft/NuGet recommended | ❌ Legacy approach |

## References

- [NuGet Trusted Publishing Documentation](https://learn.microsoft.com/en-us/nuget/nuget-org/publish-a-package#trust-based-publishing)
- [GitHub OIDC Documentation](https://docs.github.com/en/actions/deployment/security-hardening-your-deployments/about-security-hardening-with-openid-connect)
- [Semantic Versioning](https://semver.org/)
- [.NET Global Tools](https://learn.microsoft.com/en-us/dotnet/core/tools/global-tools)

## Support

For issues with NuGet publishing:

1. Check this guide's troubleshooting section
2. Review GitHub Actions workflow logs
3. Verify NuGet.org trusted publisher configuration
4. Open an issue at: https://github.com/sbroenne/mcp-server-excel/issues
