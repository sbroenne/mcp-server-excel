# ExcelMcp Release Strategy

This document outlines the unified release process for all ExcelMcp components.

## Overview

All four ExcelMcp components are released together with a single version tag:

| Component | Distribution | Description |
|-----------|--------------|-------------|
| **MCP Server** | NuGet + ZIP | Model Context Protocol server for AI assistants |
| **CLI** | NuGet + ZIP | Command-line interface for scripting |
| **VS Code Extension** | VSIX + Marketplace | One-click installation with bundled MCP Server |
| **MCPB** | Claude Desktop bundle | One-click installation for Claude Desktop |

## Unified Release Workflow

**Workflow**: `.github/workflows/release.yml`  
**Trigger**: Tags matching `v*` (e.g., `v1.5.6`)

### What Gets Released

When you push a `v*` tag:

1. **MCP Server** → NuGet (`Sbroenne.ExcelMcp.McpServer`) + ZIP
2. **CLI** → NuGet (`Sbroenne.ExcelMcp.CLI`) + ZIP
3. **VS Code Extension** → VS Code Marketplace + VSIX
4. **MCPB** → Claude Desktop bundle (`.mcpb` file)
5. **MCP Registry** → Updated after NuGet propagation
6. **GitHub Release** → Created with all artifacts

### Release Artifacts

| Artifact | Format | Distribution |
|----------|--------|--------------|
| `ExcelMcp-MCP-Server-{version}-windows.zip` | ZIP | GitHub Release |
| `ExcelMcp-CLI-{version}-windows.zip` | ZIP | GitHub Release |
| `excelmcp-{version}.vsix` | VSIX | GitHub Release + VS Code Marketplace |
| `excel-mcp-{version}.mcpb` | MCPB | GitHub Release |
| `Sbroenne.ExcelMcp.McpServer.{version}.nupkg` | NuGet | NuGet.org |
| `Sbroenne.ExcelMcp.CLI.{version}.nupkg` | NuGet | NuGet.org |

## Release Process

### 1. Update Changelog

Before creating a release tag, update `CHANGELOG.md`:

```markdown
## [Unreleased]

## [1.5.7] - 2025-01-21

### Added
- New feature description

### Changed
- Changed feature description

### Fixed
- Bug fix description
```

### 2. Create Release Tag

```bash
# Ensure you're on main with latest changes
git checkout main
git pull origin main

# Create and push tag
git tag v1.5.7
git push origin v1.5.7
```

### 3. Monitor Workflow

The release workflow runs automatically:

1. **build-mcp-server** (3-5 min) → Builds and publishes to NuGet
2. **build-cli** (3-5 min) → Builds and publishes to NuGet
3. **build-vscode** (3-5 min) → Builds and publishes to VS Code Marketplace
4. **build-mcpb** (3-5 min) → Builds Claude Desktop bundle
5. **publish-mcp-registry** (10-30 min) → Waits for NuGet propagation, updates MCP Registry
6. **create-release** → Creates GitHub Release with all artifacts

### 4. Verify Release

After workflow completes:

- [ ] GitHub Release created with all 4 artifacts
- [ ] NuGet packages available (may take 10-30 min for full propagation)
- [ ] VS Code Marketplace updated
- [ ] MCP Registry updated

## Version Management

### Single Version Number

All components use the same version number extracted from the tag:

```
Tag: v1.5.7
↓
MCP Server: 1.5.7
CLI: 1.5.7
VS Code Extension: 1.5.7
MCPB: 1.5.7
```

### Version Sources

| Component | Version Source |
|-----------|----------------|
| MCP Server | `.csproj` (updated at build time from tag) |
| CLI | `.csproj` (updated at build time from tag) |
| VS Code Extension | `package.json` (updated at build time from tag) |
| MCPB | `manifest.json` (updated at build time from tag) |

### Development Version

During development, use placeholder version `1.0.0` in:
- `Directory.Build.props`
- `package.json`
- `manifest.json`

The release workflow injects the correct version from the tag.

## Changelog Format

The root `CHANGELOG.md` follows [Keep a Changelog](https://keepachangelog.com/) format:

```markdown
# Changelog

## [Unreleased]

## [1.5.7] - 2025-01-21

### Added
- Feature description

### Changed
- Change description

### Fixed
- Bug fix description
```

The release workflow extracts the section for the current version and includes it in GitHub Release notes.

## Required Secrets

Configure these GitHub repository secrets:

| Secret | Purpose |
|--------|---------|
| `NUGET_USER` | NuGet.org username (for OIDC trusted publishing) |
| `VSCE_TOKEN` | VS Code Marketplace PAT |
| `APPINSIGHTS_CONNECTION_STRING` | Application Insights (optional telemetry) |

## Troubleshooting

### NuGet Publishing Fails

- Verify `NUGET_USER` secret is set to your NuGet.org profile name (not email)
- Check NuGet.org trusted publishers are configured for OIDC

### VS Code Marketplace Fails

- Verify `VSCE_TOKEN` is valid and not expired
- Check extension ID matches marketplace listing

### MCPB Build Fails

- Ensure `mcpb/manifest.json` is valid JSON
- Verify `mcpb/icon-512.png` exists (512x512 PNG)

### MCP Registry Update Fails

- MCP Registry update uses GitHub OIDC
- Failures don't block the release (marked continue-on-error)
- Can be retried manually via MCP publisher tool

## Legacy Workflows

The following workflows have been deprecated:

- `.github/workflows/release-mcp-server.yml.deprecated` - Replaced by unified workflow
- `.github/workflows/release-vscode-extension.yml.deprecated` - Replaced by unified workflow

These files are kept for reference but are not triggered.

## Benefits of Unified Releases

1. **Single version** across all components ensures compatibility
2. **One tag** triggers all releases - simpler process
3. **Synchronized updates** - users always get matching versions
4. **Reduced coordination** - no need to remember multiple tag patterns
5. **Complete changelog** - all changes documented in one place
6. **Faster releases** - parallel builds for all components
