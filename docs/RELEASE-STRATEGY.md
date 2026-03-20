# ExcelMcp Release Strategy

This document outlines the unified release process for all ExcelMcp components.

## Overview

All ExcelMcp components are released together with a single version tag:

| Component | Distribution | Description |
|-----------|--------------|-------------|
| **MCP Server** | Standalone exe ZIP | Self-contained `mcp-excel.exe` — no .NET runtime required |
| **CLI** | Standalone exe ZIP | Self-contained `excelcli.exe` — no .NET runtime required |
| **VS Code Extension** | VSIX + Marketplace | Self-contained — bundles MCP Server + CLI + agent skills (no .NET required) |
| **MCPB** | Claude Desktop bundle | Self-contained one-click installation for Claude Desktop |
| **Agent Skills** | ZIP | Reusable skill packages for AI coding assistants |

## Unified Release Workflow

**Workflow**: `.github/workflows/release.yml`  
**Trigger**: `workflow_dispatch` with version bump (major/minor/patch) or custom version

### What Gets Released

When you run the release workflow:

1. **CLI** → Standalone self-contained exe (`excelcli.exe`) → ZIP artifact
2. **MCP Server** → Standalone self-contained exe (`mcp-excel.exe`) → ZIP artifact
3. **VS Code Extension** → Self-contained VSIX (bundles both exes + skills) → VS Code Marketplace
4. **MCPB** → Claude Desktop bundle (`.mcpb` file)
5. **Agent Skills** → ZIP package for AI coding assistants
6. **MCP Registry** → Updated via GitHub OIDC
7. **GitHub Release** → Created with all artifacts + auto-PR to update CHANGELOG

### Release Artifacts

| Artifact | Format | Distribution |
|----------|--------|--------------|
| `ExcelMcp-MCP-Server-{version}-windows.zip` | ZIP | GitHub Release (contains `mcp-excel.exe`) |
| `ExcelMcp-CLI-{version}-windows.zip` | ZIP | GitHub Release (contains `excelcli.exe`) |
| `excelmcp-{version}.vsix` | VSIX | GitHub Release + VS Code Marketplace (~68-70 MB, self-contained) |
| `excel-mcp-{version}.mcpb` | MCPB | GitHub Release |
| `excel-skills-v{version}.zip` | ZIP | GitHub Release |

> **No .NET required:** All standalone distributions are self-contained and bundle the .NET 10 runtime.

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

### 2. Run the Release Workflow

1. Go to **Actions** → **Release All Components** → **Run workflow**
2. Select the version bump type:
   - **patch** (default): `1.5.6` → `1.5.7`
   - **minor**: `1.5.6` → `1.6.0`
   - **major**: `1.5.6` → `2.0.0`
3. Or enter a **custom version** (e.g., `1.5.7`) to override the bump

The workflow will:
1. Calculate the next version from the latest git tag
2. Build all components with the new version as standalone self-contained executables
3. Create and push the git tag (`v1.5.7`)
4. Publish to VS Code Marketplace and npm (agent skills)
5. Update MCP Registry
6. Create GitHub Release with all artifacts
7. Auto-PR to update `CHANGELOG.md`

### 3. Monitor Workflow

The release workflow runs automatically (8 jobs):

1. **build-cli** (3-5 min) → Builds standalone `excelcli.exe` (win-x64, self-contained), packages ZIP
2. **build-mcp-server** (4-6 min) → Builds standalone `mcp-excel.exe` (win-x64, self-contained), packages ZIP
3. **build-vscode** (5-8 min) → Builds self-contained VSIX, publishes to VS Code Marketplace
4. **build-mcpb** (3-5 min) → Builds Claude Desktop bundle
5. **build-agent-skills** (1-2 min) → Builds agent skills ZIP package
6. **create-tag** → Creates git tag (waits for all builds)
7. **publish-mcp-registry** → Updates MCP Registry via GitHub OIDC
8. **publish** → Publishes to VS Code Marketplace and npm
9. **create-release** → Creates GitHub Release with all artifacts, then creates auto-PR to update CHANGELOG

**Job dependencies:**
- `create-tag` depends on all build jobs (cli, mcp-server, vscode, mcpb, agent-skills)
- `publish-mcp-registry` depends on `create-tag` and `build-mcp-server`
- `publish` depends on `create-tag`
- `create-release` depends on `create-tag` and `publish`
- `build-cli`, `build-mcp-server`, `build-vscode`, `build-mcpb`, `build-agent-skills` run in parallel

### 4. Verify Release

After workflow completes:

- [ ] GitHub Release created with all artifacts (MCP Server ZIP, CLI ZIP, VSIX, MCPB, skills ZIP)
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
| `VSCE_TOKEN` | VS Code Marketplace PAT |
| `APPINSIGHTS_CONNECTION_STRING` | Application Insights (optional telemetry) |

> **Note:** No NuGet secrets needed — packages are no longer published to NuGet.org.

## Troubleshooting

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
7. **Self-contained distributions** — all exe distributions bundle .NET runtime, no external dependencies
8. **No NuGet dependency** — users just download and run; no package manager needed
