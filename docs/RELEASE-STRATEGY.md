# ExcelMcp Release Strategy

This document outlines the unified release process for all ExcelMcp components.

## Overview

All ExcelMcp components are released together with a single version tag:

| Component | Distribution | Description |
|-----------|--------------|-------------|
| **MCP Server** | NuGet (unified) + ZIP | Model Context Protocol server + CLI (single NuGet package) |
| **CLI** | Bundled in MCP Server | Command-line interface for scripting (no longer a separate NuGet) |
| **VS Code Extension** | VSIX + Marketplace | Self-contained — bundles MCP Server + CLI + agent skills (no .NET required) |
| **MCPB** | Claude Desktop bundle | Self-contained one-click installation for Claude Desktop |
| **Agent Skills** | ZIP | Reusable skill packages for AI coding assistants |

## Unified Release Workflow

**Workflow**: `.github/workflows/release.yml`  
**Trigger**: Tags matching `v*` (e.g., `v1.5.6`)

### What Gets Released

When you push a `v*` tag:

1. **CLI** → Built as dependency (artifact shared with MCP Server job)
2. **MCP Server** → Unified NuGet (`Sbroenne.ExcelMcp.McpServer` — includes CLI) + ZIP
3. **VS Code Extension** → Self-contained VSIX (bundles both exes + skills) → VS Code Marketplace
4. **MCPB** → Claude Desktop bundle (`.mcpb` file)
5. **Agent Skills** → ZIP package for AI coding assistants
6. **MCP Registry** → Updated after NuGet propagation
7. **GitHub Release** → Created with all artifacts + auto-PR to update CHANGELOG

### Release Artifacts

| Artifact | Format | Distribution |
|----------|--------|--------------|
| `ExcelMcp-MCP-Server-{version}-windows.zip` | ZIP | GitHub Release |
| `excelmcp-{version}.vsix` | VSIX | GitHub Release + VS Code Marketplace (~68-70 MB, self-contained) |
| `excel-mcp-{version}.mcpb` | MCPB | GitHub Release |
| `excel-skills-v{version}.zip` | ZIP | GitHub Release |
| `Sbroenne.ExcelMcp.McpServer.{version}.nupkg` | NuGet | NuGet.org (unified — includes CLI) |

> **Note:** Separate CLI ZIP and CLI NuGet package are no longer produced. The CLI is bundled in the unified MCP Server NuGet package and in the VS Code extension.

## Release Process

### 1. Update Changelog

Before creating a release tag, ensure all changes are documented under `## [Unreleased]` in `CHANGELOG.md`:

```markdown
## [Unreleased]

### Added
- New feature description

### Changed
- Changed feature description

### Fixed
- Bug fix description

## [1.5.6] - 2025-01-15
...
```

> **Important:** Do NOT rename `[Unreleased]` to a version number manually. The release workflow extracts content from `[Unreleased]` for release notes, then creates an auto-PR to rename it to `[X.Y.Z] - date` and add a fresh `[Unreleased]` section.

### 2. Create Release Tag

```powershell
# Ensure you're on main with latest changes
git checkout main
git pull origin main

# Create and push tag
git tag v1.5.7
git push origin v1.5.7
```

### 3. Monitor Workflow

The release workflow runs automatically (7 jobs):

1. **build-cli** (2-3 min) → Builds CLI as dependency artifact (not published separately)
2. **build-mcp-server** (3-5 min) → Downloads CLI artifact, builds unified NuGet package, publishes to NuGet.org
3. **build-vscode** (5-8 min) → Builds self-contained MCP Server + CLI exes, copies skills, packages VSIX, publishes to VS Code Marketplace
4. **build-mcpb** (3-5 min) → Builds Claude Desktop bundle
5. **build-agent-skills** (1-2 min) → Builds agent skills ZIP package
6. **publish-mcp-registry** (10-30 min) → Waits for NuGet propagation, updates MCP Registry
7. **create-release** → Creates GitHub Release with all artifacts, then creates auto-PR to update CHANGELOG

**Job dependencies:**
- `build-mcp-server` depends on `build-cli` (needs CLI build output)
- `publish-mcp-registry` depends on `build-mcp-server` (needs NuGet published)
- `create-release` depends on `build-mcp-server`, `build-vscode`, `build-mcpb`, `build-agent-skills`
- `build-vscode`, `build-mcpb`, `build-agent-skills` run independently (no dependencies)

### 4. Verify Release

After workflow completes:

- [ ] GitHub Release created with all artifacts (ZIP, VSIX, MCPB, skills ZIP)
- [ ] NuGet unified package available (may take 10-30 min for full propagation)
- [ ] VS Code Marketplace updated (verify self-contained extension works without .NET)
- [ ] MCP Registry updated
- [ ] Auto-PR created for CHANGELOG rename (merge it to update `[Unreleased]` → `[X.Y.Z]`)

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

The release workflow extracts content from `## [Unreleased]` for GitHub Release notes. After the release is created, an auto-PR renames `[Unreleased]` to `[X.Y.Z] - date` and adds a fresh `[Unreleased]` section.

> **Why auto-PR instead of direct push?** Branch protection requires pull requests for all changes to `main`. The `github-actions[bot]` cannot be added to the bypass list in GitHub Rulesets, so the workflow creates a PR with `continue-on-error: true` to handle this gracefully.

## Required Secrets

Configure these GitHub repository secrets:

| Secret | Purpose |
|--------|---------|
| `NUGET_USER` | NuGet.org username (for OIDC trusted publishing) |
| `VSCE_TOKEN` | VS Code Marketplace PAT |
| `APPINSIGHTS_CONNECTION_STRING` | Application Insights (optional telemetry) |

> **Note:** NuGet uses OIDC trusted publishing (no API key needed). The `NUGET_USER` is just the NuGet.org profile name for OIDC token exchange.

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
2. **One tag** triggers all releases — simpler process
3. **Synchronized updates** — users always get matching versions
4. **Reduced coordination** — no need to remember multiple tag patterns
5. **Complete changelog** — all changes documented in one place, auto-updated via PR
6. **Faster releases** — parallel builds for independent components
7. **Self-contained distributions** — VS Code and MCPB bundle .NET runtime, no external dependencies
8. **Unified NuGet** — single package installs both MCP Server and CLI
