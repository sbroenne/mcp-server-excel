# Glama.ai Docker Support

This document explains the Docker support added for [Glama.ai](https://glama.ai) MCP registry integration.

## Overview

The `Dockerfile` in the repository root enables Glama.ai to:
- **Inspect** the MCP server and discover available tools
- **Display** tool schemas, parameters, and descriptions in their registry
- **Allow users** to browse capabilities before installing locally

**Registry URL:** https://glama.ai/mcp/servers/@sbroenne/mcp-server-excel

## Important Limitations

> **The Docker container CANNOT perform actual Excel operations.**

Excel MCP requires:
- **Windows OS** - COM interop is Windows-only
- **Microsoft Excel** - The actual Excel application must be installed
- **Local installation** - Cannot run Excel automation in a container

The Docker image is purely for **tool discovery and inspection** by Glama.ai's registry crawler.

## Files

| File | Purpose |
|------|---------|
| `Dockerfile` | Multi-stage build that compiles the MCP server for Linux |
| `.dockerignore` | Excludes unnecessary files from Docker build context |
| `glama.json` | Glama registry metadata (maintainer info) |

## Building Locally (Optional)

If you want to test the Docker build locally:

```powershell
# Build the image
docker build -t mcp-server-excel:test .

# Test MCP protocol response
echo '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"1.0","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}}' | docker run -i --rm mcp-server-excel:test
```

## How It Works

1. **Glama.ai crawler** pulls the Docker image from the registry
2. **Sends MCP protocol messages** (initialize, tools/list) to the container
3. **Extracts tool metadata** - names, descriptions, parameters, schemas
4. **Displays in registry** - Users can browse tools before installing

## For Actual Excel Automation

To use Excel MCP for real automation, install locally on Windows:

```powershell
# Via NPX (recommended)
npx @anthropic-ai/excel-mcp-server

# Or download from releases
# https://github.com/sbroenne/mcp-server-excel/releases
```

See the main [README.md](../README.md) for complete installation instructions.

## Related

- [MCP Registry Publishing](MCP_REGISTRY_PUBLISHING.md) - NPM registry publishing
- [Installation Guide](INSTALLATION.md) - Local installation options
