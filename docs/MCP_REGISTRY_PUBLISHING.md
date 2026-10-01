# MCP Registry Publishing Guide

This document describes how the ExcelMcp server is published to the [Model Context Protocol (MCP) Registry](https://registry.modelcontextprotocol.io/).

## Overview

The ExcelMcp server is automatically published to the MCP Registry as part of the unified release workflow in `.github/workflows/release.yml`.

## Configuration Files

### server.json

Location: `src/ExcelMcp.McpServer/.mcp/server.json`

This is the MCP registry metadata file that describes the server:

```json
{
  "$schema": "https://static.modelcontextprotocol.io/schemas/2025-12-11/server.schema.json",
  "name": "io.github.sbroenne/mcp-server-excel",
  "title": "MCP Server for Excel",
  "description": "Real Excel automation for AI. Complete on Windows; experimental beta on Apple Silicon macOS.",
  "version": "1.0.0",
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
- `repository`: Source repository reference

### Package Validation

The registry offers both NuGet and npm installations. It validates NuGet
ownership through `mcp-name:` in the package README and npm ownership through
the `mcpName` property in the launcher package.

Registry presence is not platform acceptance. Apple Silicon macOS support is
experimental beta; prefer the npm deployment for Mac. NuGet installation checks
run on Windows, and neither channel bypasses
[Mac feature gates](../specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta).

Location: `src/ExcelMcp.McpServer/README.md`

The README includes this validation metadata:
```markdown
<!-- mcp-name: io.github.sbroenne/mcp-server-excel -->
```

This HTML comment is invisible to users but allows the registry to verify the package belongs to this server.

The npm package at `npm-packages/mcp-server-excel/package.json` contains:

```json
{
  "mcpName": "io.github.sbroenne/mcp-server-excel"
}
```

## Publishing Workflow

The publishing process is automated by `.github/workflows/publish-mcp-registry.yml`,
which the unified release workflow calls after NuGet and npm publication:

### 1. Version Update
The workflow:
- Parses `server.json`, updates the top-level version plus the NuGet and npm
  package versions, then validates that all values match the release version.

### 2. Wait for NuGet and npm Propagation
- The MCP Registry offers both NuGet and npm deployment mechanisms
- `scripts/Test-McpRegistryPublication.ps1` verifies the source `server.json`
  identity and release version, including its NuGet and npm package entries
- The job waits for the NuGet README's `mcp-name:` marker, NuGet package identity
  and version, and the npm launcher's identity, version, and `mcpName` metadata
- It also checks each required npm runtime's published identity and version,
  and requires the launcher's corresponding `optionalDependencies` entry to
  match the release version
- Required runtimes come from the exact released source's launcher manifest,
  passed through `-NpmLauncherManifestPath`: x64 is required, and ARM64 is
  required when declared. Historical x64-only releases can therefore still be
  repaired without requiring an ARM64 package that did not exist
- Polls up to 3 times with 10-minute intervals
- Decodes the NuGet README response as UTF-8 when NuGet returns
  `application/octet-stream`

### 3. MCP Registry Publishing
- Downloads the MCP Publisher CLI tool
- Authenticates using GitHub OIDC (no secrets required)
- Publishes `server.json` to the MCP Registry
- Fails the registry job if publication fails so the release workflow cannot report
  the registry update as successful when it was not published

## Authentication

### MCP Registry Authentication
Uses **GitHub OIDC**:
- No secrets required
- Automatic authentication via `mcp-publisher login github-oidc`
- Works for `io.github.*` namespaces
- The publish job uses the protected `mcp-registry` environment. Repository
  settings must keep its custom deployment branch policy restricted to `main`
  and require approval from the repository owner before the OIDC token is issued.

**Required Permissions:**
The workflow has `id-token: write` permission enabled for OIDC authentication.

## Release Process

See [RELEASE-STRATEGY.md](RELEASE-STRATEGY.md) for the full release process.

After release, verify publication:
- **MCP Registry**: https://registry.modelcontextprotocol.io/v0/servers?search=io.github.sbroenne/mcp-server-excel
- **GitHub Release**: https://github.com/sbroenne/mcp-server-excel/releases

## Troubleshooting

### Package Metadata Is Not Ready

**Issue**: The validation gate reports `Published package metadata is not ready`
or `Published x64/arm64 npm runtime metadata is not ready`.

**Solution**:
- Check that the NuGet package and npm launcher have the requested release
  version and ownership metadata described above
- Check that each runtime required by that release's launcher manifest is
  published at the same version, and that the published launcher's matching
  `optionalDependencies` entries reference that version
- Allow package metadata to propagate before retrying the repair workflow with
  the exact existing release tag. Do not substitute the current branch's
  launcher manifest when repairing an older release

### MCP Registry Publishing Fails

**Issue**: "Authentication failed" or OIDC error

**Solution**: 
- Verify `id-token: write` permission is set in the workflow job
- Ensure repository is configured for GitHub OIDC
- Resolve the failure, then manually run **Publish MCP Registry** with the exact
  existing release tag. The repair workflow validates that tag's immutable
  source metadata, requires its commit to be reachable from protected `main`,
  and validates the existing NuGet and npm packages before publishing
  only the MCP Registry entry. It does not rebuild or republish any package,
  GitHub release asset, extension, or plugin.

### Version Not Updated

**Issue**: Registry shows old version

**Solution**: 
- Check the `publish-mcp-registry` job logs
- Confirm the top-level, NuGet, and npm `server.json` versions were stamped with the release version
  before rerunning the workflow or publishing manually
