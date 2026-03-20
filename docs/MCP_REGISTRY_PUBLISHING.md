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
  "description": "Excel automation for AI - Sheets, Power Query, DAX, VBA, Tables, Ranges and more. Windows only.",
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

The MCP Registry validates ownership by checking for `mcp-name:` in the package README.

Location: `src/ExcelMcp.McpServer/README.md`

The README includes this validation metadata:
```markdown
<!-- mcp-name: io.github.sbroenne/mcp-server-excel -->
```

This HTML comment is invisible to users but allows the registry to verify the package belongs to this server.

## Publishing Workflow

The publishing process is automated as `publish-mcp-registry` job in `.github/workflows/release.yml`:

### 1. Version Update
The workflow:
- Updates `server.json` with the new version number

### 2. MCP Registry Publishing (Non-Blocking)
- Downloads the MCP Publisher CLI tool
- Authenticates using GitHub OIDC (no secrets required)
- Publishes `server.json` to the MCP Registry
- Uses `continue-on-error: true` to ensure release completes even if MCP Registry publishing fails

## Authentication

### MCP Registry Authentication
Uses **GitHub OIDC**:
- No secrets required
- Automatic authentication via `mcp-publisher login github-oidc`
- Works for `io.github.*` namespaces

**Required Permissions:**
The workflow has `id-token: write` permission enabled for OIDC authentication.

## Release Process

See [RELEASE-STRATEGY.md](RELEASE-STRATEGY.md) for the full release process.

After release, verify publication:
- **MCP Registry**: https://registry.modelcontextprotocol.io/servers/io.github.sbroenne/mcp-server-excel
- **GitHub Release**: https://github.com/sbroenne/mcp-server-excel/releases

## Troubleshooting

### MCP Registry Publishing Fails

**Issue**: "Authentication failed" or OIDC error

**Solution**: 
- Verify `id-token: write` permission is set in the workflow job
- Ensure repository is configured for GitHub OIDC
- MCP Registry publishing failures don't block the release (`continue-on-error: true`)

### Version Not Updated

**Issue**: Registry shows old version

**Solution**: 
- Check the `publish-mcp-registry` job logs
- Re-run the workflow or manually update `server.json` and run mcp-publisher
