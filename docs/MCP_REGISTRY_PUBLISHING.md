# MCP Registry Publishing Guide

This document describes how the ExcelMcp server is published to the [Model Context Protocol (MCP) Registry](https://registry.modelcontextprotocol.io/).

## Overview

The ExcelMcp server is automatically published to the MCP Registry whenever a new release is tagged with the format `mcp-v*` (e.g., `mcp-v1.0.10`). This is handled by the GitHub Actions workflow `.github/workflows/release-mcp-server.yml`.

## Configuration Files

### server.json

Location: `src/ExcelMcp.McpServer/.mcp/server.json`

This is the MCP registry metadata file that describes the server:

```json
{
  "$schema": "https://static.modelcontextprotocol.io/schemas/2025-10-17/server.schema.json",
  "name": "io.github.sbroenne/mcp-server-excel",
  "title": "Excel COM Automation",
  "description": "Excel COM automation - Power Query, DAX measures, VBA, Tables, ranges, connections",
  "version": "1.0.0",
  "packages": [
    {
      "registryType": "nuget",
      "identifier": "Sbroenne.ExcelMcp.McpServer",
      "version": "1.0.0",
      "transport": {
        "type": "stdio"
      }
    }
  ],
  "repository": {
    "url": "https://github.com/sbroenne/mcp-server-excel",
    "source": "github"
  }
}
```

Key fields:
- `name`: Registry namespace (uses GitHub namespace `io.github.sbroenne/*`)
- `title`: Human-readable name
- `description`: Brief description of capabilities
- `version`: Server version (automatically updated by release workflow)
- `packages`: Array of deployment options (NuGet package in this case)

### Package Validation

For NuGet packages, the MCP Registry validates ownership by checking for `mcp-name:` in the package README.

Location: `src/ExcelMcp.McpServer/README.md`

The README includes this validation metadata:
```markdown
<!-- mcp-name: io.github.sbroenne/mcp-server-excel -->
```

This HTML comment is invisible to users but allows the registry to verify the package belongs to this server.

## Publishing Workflow

The publishing process is automated in `.github/workflows/release-mcp-server.yml`:

### 1. Version Update
When a tag like `mcp-v1.0.10` is pushed, the workflow:
- Extracts version `1.0.10` from the tag
- Updates `server.json` with the new version
- Updates `.csproj` file with the new version

### 2. Build and Test
- Restores dependencies
- Builds the MCP server in Release configuration
- Skips tests (they require Excel)

### 3. NuGet Publishing
- Packs the NuGet package
- Authenticates using NuGet Trusted Publishing (OIDC)
- Publishes to NuGet.org
- Waits for package to be available

### 4. MCP Registry Publishing (Non-Blocking)
- Downloads the MCP Publisher CLI tool
- Authenticates using GitHub OIDC (no secrets required)
- Publishes `server.json` to the MCP Registry
- **Note**: These steps use `continue-on-error: true` to ensure release completes even if MCP Registry publishing fails
- Check the "MCP Registry Status" step output for publishing results

### 5. GitHub Release
- Creates a ZIP file with binaries
- Creates a GitHub release with release notes
- Attaches the ZIP file

## Authentication

### NuGet Authentication
Uses **NuGet Trusted Publishing** via OIDC:
- No API keys stored in GitHub secrets
- Automatic token exchange via GitHub Actions
- Configured in NuGet.org package settings

**Required Secret:**
- `NUGET_USER`: Your NuGet.org username (profile name)

**NuGet.org Configuration:**
- Package: `Sbroenne.ExcelMcp.McpServer`
- Trusted Publisher: GitHub Actions
- Owner: `sbroenne`
- Repository: `mcp-server-excel`
- Workflow: `release-mcp-server.yml`

### MCP Registry Authentication
Uses **GitHub OIDC**:
- No secrets required
- Automatic authentication via `mcp-publisher login github-oidc`
- Works for `io.github.*` namespaces

**Required Permissions:**
The workflow has `id-token: write` permission enabled for OIDC authentication.

## Release Process

### Creating a New Release

1. **Ensure all changes are merged to main**
   ```bash
   git checkout main
   git pull origin main
   ```

2. **Create and push a version tag**
   ```bash
   git tag mcp-v1.0.10
   git push origin mcp-v1.0.10
   ```

3. **Monitor the workflow**
   - Go to GitHub Actions
   - Watch the "Release MCP Server" workflow
   - Verify all steps complete successfully

4. **Verify publication**
   - **NuGet**: https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer
   - **MCP Registry**: https://registry.modelcontextprotocol.io/servers/io.github.sbroenne/mcp-server-excel
   - **GitHub Release**: https://github.com/sbroenne/mcp-server-excel/releases

### Version Numbering

- Use semantic versioning: `MAJOR.MINOR.PATCH`
- Tag format: `mcp-v{version}` (e.g., `mcp-v1.0.10`)
- The workflow automatically updates all version references

## Troubleshooting

### NuGet Publishing Fails

**Issue**: "Authentication failed" or "API key expired"

**Solution**: 
- Verify `NUGET_USER` secret is set correctly (must be profile name, not email)
- Check NuGet.org Trusted Publishers configuration
- Ensure package exists and you have ownership

### MCP Registry Publishing Fails

**Issue**: "Package validation failed"

**Solution**:
- Verify the NuGet package has been published successfully
- Check that `mcp-name:` is present in the package README
- Wait a few minutes for NuGet indexing
- Verify the mcp-name matches the server.json name exactly

**Note**: As of the latest workflow update, MCP Registry publishing failures do not block the release process. The NuGet package will still be published successfully, and you can manually publish to the MCP Registry later if needed.

**Issue**: "Authentication failed"

**Solution**:
- Verify workflow has `id-token: write` permission
- Ensure you're using the correct namespace format
- For `io.github.*` namespaces, GitHub OIDC should work automatically

**Note**: The workflow uses `continue-on-error: true` for MCP Registry steps, so authentication failures will not prevent the release from completing.

### Workflow Doesn't Trigger

**Issue**: Tag pushed but workflow doesn't run

**Solution**:
- Verify tag format is `mcp-v*` (prefix is required)
- Check that the workflow file is on the main branch
- Look for workflow errors in GitHub Actions

## Registry Features

Once published, the server will be:

1. **Discoverable**: Users can search for it in MCP-compatible clients
2. **Auto-installable**: `dnx Sbroenne.ExcelMcp.McpServer --yes`
3. **Version-managed**: Users can install specific versions
4. **Documented**: Description and README visible in registry

## Manual Publishing (Emergency)

If the automated workflow fails, you can publish manually:

### 1. Install MCP Publisher
```bash
# Windows PowerShell
$arch = if ([System.Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture -eq "Arm64") { "arm64" } else { "amd64" }
Invoke-WebRequest -Uri "https://github.com/modelcontextprotocol/registry/releases/latest/download/mcp-publisher_windows_$arch.tar.gz" -OutFile "mcp-publisher.tar.gz"
tar xf mcp-publisher.tar.gz
```

### 2. Authenticate
```bash
./mcp-publisher login github
```

### 3. Update server.json Version
Edit `src/ExcelMcp.McpServer/.mcp/server.json` and update the version fields.

### 4. Publish
```bash
cd src/ExcelMcp.McpServer
../../mcp-publisher publish
```

## References

- [MCP Registry Publishing Guide](https://github.com/modelcontextprotocol/registry/blob/main/docs/guides/publishing/publish-server.md)
- [GitHub Actions Automation](https://github.com/modelcontextprotocol/registry/blob/main/docs/guides/publishing/github-actions.md)
- [NuGet Trusted Publishing](https://learn.microsoft.com/en-us/nuget/nuget-org/publish-a-package)
- [MCP Registry](https://registry.modelcontextprotocol.io/)
