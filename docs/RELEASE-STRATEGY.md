# ExcelMcp Release Strategy

This document outlines the unified release process for all ExcelMcp components.

## Overview

All ExcelMcp components are released together with a single version tag:

| Component | Primary Distribution | Secondary Distribution | Description |
|-----------|---------------------|----------------------|-------------|
| **MCP Server** | Standalone exe ZIP | NuGet (.NET tool) | `mcp-excel.exe` — no .NET runtime required |
| **CLI** | Standalone exe ZIP | NuGet (.NET tool) | `excelcli.exe` — no .NET runtime required |
| **VS Code Extension** | VSIX + Marketplace | — | Self-contained — bundles MCP Server + CLI + skills |
| **MCPB** | Claude Desktop bundle | — | Self-contained one-click installation |
| **Agent Skills** | ZIP | npm | Reusable skill packages for AI coding assistants |

## Unified Release Workflow

**Workflow**: `.github/workflows/release.yml`
**Trigger**: `workflow_dispatch` with version bump (major/minor/patch) or custom version

### What Gets Released

When you run the release workflow, all components are released together:

1. **CLI** → Standalone self-contained exe (`excelcli.exe`) + ZIP [primary] + NuGet pack [secondary]
2. **MCP Server** → Standalone self-contained exe (`mcp-excel.exe`) + ZIP [primary] + NuGet pack [secondary]
3. **VS Code Extension** → Self-contained VSIX (bundles both exes + skills) → VS Code Marketplace
4. **MCPB** → Claude Desktop bundle (`.mcpb` file)
5. **Agent Skills** → ZIP package for AI coding assistants
6. **GitHub Copilot Plugins** → Republished to the GitHub Copilot plugin marketplace repo via `publish-plugins.yml` (see [Phase 3 Plugin Publishing](../.github/workflows/docs/publish-plugins-setup.md))
7. **NuGet** → Both packages published to NuGet.org (secondary channel)
8. **MCP Registry** → Updated after NuGet propagation
9. **GitHub Release** → Created with all artifacts + auto-PR to update CHANGELOG

### Release Artifacts

| Artifact | Format | Distribution |
|----------|--------|--------------|
| `ExcelMcp-MCP-Server-{version}-windows.zip` | ZIP | GitHub Release (primary — contains `mcp-excel.exe`) |
| `ExcelMcp-CLI-{version}-windows.zip` | ZIP | GitHub Release (primary — contains `excelcli.exe`) |
| `excelmcp-{version}.vsix` | VSIX | GitHub Release + VS Code Marketplace (~68-70 MB, self-contained) |
| `excel-mcp-{version}.mcpb` | MCPB | GitHub Release |
| `excel-skills-v{version}.zip` | ZIP | GitHub Release |
| `Sbroenne.ExcelMcp.McpServer.{version}.nupkg` | NuGet | NuGet.org (secondary — requires .NET 10 runtime) |
| `Sbroenne.ExcelMcp.CLI.{version}.nupkg` | NuGet | NuGet.org (secondary — requires .NET 10 runtime) |

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
2. Build all components (standalone exes + NuGet packages)
3. Create and push the git tag (`v1.5.7`)
4. Publish to NuGet.org, VS Code Marketplace, MCP Registry
5. Create GitHub Release with all artifacts
6. Auto-PR to update `CHANGELOG.md`

### 3. Monitor Workflow

The main release workflow runs automatically (9 jobs), then the plugin publish workflow runs automatically if the release succeeds:

1. **build-cli** (3-5 min) → Builds standalone `excelcli.exe` (win-x64, self-contained), creates ZIP + NuGet pack
2. **build-mcp-server** (4-6 min) → Builds standalone `mcp-excel.exe` (win-x64, self-contained), creates ZIP + NuGet pack
3. **build-vscode** (5-8 min) → Builds self-contained VSIX (bundles exes), publishes to VS Code Marketplace
4. **build-mcpb** (3-5 min) → Builds Claude Desktop bundle
5. **build-agent-skills** (1-2 min) → Builds agent skills ZIP package
6. **create-tag** → Creates git tag (waits for all builds)
7. **publish-mcp-registry** (10-30 min) → Waits for NuGet propagation, updates MCP Registry
8. **publish** → Publishes to NuGet.org, VS Code Marketplace, and npm
9. **create-release** → Creates GitHub Release with all artifacts, then creates auto-PR to update CHANGELOG
10. **publish-plugins.yml** (follow-on workflow) → Sync-gated republish of `excel-mcp` and `excel-cli` to `sbroenne/mcp-server-excel-plugins` when plugin-facing install artifacts changed

### 4. Verify Release

After workflow completes:

- [ ] GitHub Release created with all artifacts (MCP Server ZIP, CLI ZIP, VSIX, MCPB, skills ZIP)
- [ ] NuGet packages available on NuGet.org (may take 10-30 min for full propagation)
- [ ] VS Code Marketplace updated (verify self-contained extension works without .NET)
- [ ] MCP Registry updated
- [ ] `publish-plugins.yml` completed; if the sync gate detected plugin-facing changes, `sbroenne/mcp-server-excel-plugins` was updated
- [ ] Auto-PR created for CHANGELOG rename (merge it to update `[Unreleased]` → `[X.Y.Z]`)

### 5. Agent Plugin Publishing (Automatic)

**Workflow**: `.github/workflows/publish-plugins.yml`
**Trigger**: Runs automatically after `release.yml` completes successfully, with a manual `workflow_dispatch` re-sync path for existing source release tags
**Published Repo**: `sbroenne/mcp-server-excel-plugins` (published plugin artifact repo)

The `publish-plugins.yml` workflow automatically publishes updated plugins when the release workflow completes:

1. **Extracts version** from the release tag created by `release.yml`
2. **Runs a source-side sync gate** and skips the downstream publish when no plugin-published source files changed since the previous release tag
3. **Builds plugins** via `scripts/Build-Plugins.ps1`:
    - Copies validated plugin structure from the published repo
    - Updates version in plugin.json and version.txt
    - Refreshes skill content (always uses latest source)
4. **Checks published-repo guards** before mutation:
   - Rejects explicit tag/version mismatches
   - Rejects downgrade publishes
   - Skips automatic duplicate publishes when the published repo already has the same version and tag
5. **Publishes plugin artifacts** by committing and tagging the published repo when needed

Maintainers can also replay plugin publication for an existing release tag without cutting a new release:

```powershell
gh workflow run publish-plugins.yml -f release_tag=v1.2.3
```

**Key Points:**
- ✅ **Automatic** — No manual intervention required
- ✅ **Idempotent** — Safe to re-run on the same version
- ✅ **Version-aligned** — Uses the exact version from the release
- ✅ **Sync-gated** — skips downstream plugin republish when plugin install-surface inputs did not change since the prior release tag
- ✅ **Guarded replay** — downgrade syncs are rejected, automatic duplicates are skipped, and manual repair/replay runs must keep the requested tag aligned with the incoming plugin manifest/version
- ✅ **Manual repair path** — maintainers keep a `workflow_dispatch` re-sync entry point for repair/replay scenarios
- ⚠️ **Requires cross-repo token** — First-time setup needs a repository secret `PLUGINS_REPO_TOKEN` (PAT with `public_repo` scope) in the source repo (see [Phase 3 Plugin Publishing docs](../.github/workflows/docs/publish-plugins-setup.md))
- ℹ️ **Setup command** — After creating the PAT: `gh secret set PLUGINS_REPO_TOKEN --repo sbroenne/mcp-server-excel --body "<token-value>"`

**Surface note:**
- The release automation publishes plugin bundles (manifest, skills, agents, hooks, MCP config, helper scripts) to the published repo.
- Those artifacts can be relevant across multiple plugin-capable clients, but marketplace registration, discovery, and installation UX remain client-specific.
- The current workflow and docs only claim a verified GitHub Copilot install flow; they do **not** claim automatic publication into VS Code or Claude-specific plugin marketplaces.

**Hardening note:**
- Automatic publication now passes through a source-side sync gate so unchanged plugin install surfaces do not produce redundant downstream publishes.
- The published-side sync path rejects downgrade attempts and keeps explicit repair/replay runs honest by requiring tag/version alignment.
- Maintainers still have a manual `workflow_dispatch` re-sync entry point when a repair or replay is needed.

For detailed setup instructions and troubleshooting, see [Phase 3 Plugin Publishing Setup](../.github/workflows/docs/publish-plugins-setup.md).

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

## Required Secrets and Variables

Configure these GitHub repository secrets and variables:

| Type | Name | Purpose |
|------|------|---------|
| Secret | `NUGET_USER` | NuGet.org username (for OIDC trusted publishing) |
| Secret | `VSCE_TOKEN` | VS Code Marketplace PAT |
| Secret | `APPINSIGHTS_CONNECTION_STRING` | Application Insights (optional telemetry) |
| Secret | `PLUGINS_REPO_TOKEN` | Cross-repo PAT with `public_repo` scope for publishing plugins to `sbroenne/mcp-server-excel-plugins` |

> **Notes:**
> - NuGet uses OIDC trusted publishing (no API key needed). The `NUGET_USER` is just the NuGet.org profile name for OIDC token exchange.
> - npm skill packages also use trusted publishing via GitHub OIDC (`id-token: write` in `release.yml`), so no npm token is required.
> - The follow-on plugin publish workflow uses a stored cross-repo PAT (`PLUGINS_REPO_TOKEN`) with `contents:write` on the published plugin repo.

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

### Publish Plugins Fails

- Confirm the repository secret `PLUGINS_REPO_TOKEN` exists in `sbroenne/mcp-server-excel`
- Confirm the token has `public_repo` scope (or `repo` if the plugin repo were private)
- Verify the token hasn't expired and has push access to `sbroenne/mcp-server-excel-plugins`
- If the main release succeeded but plugins did not update, inspect the separate follow-on `publish-plugins.yml` run
- If you need to replay the publish without cutting a new release, dispatch `publish-plugins.yml` manually with `release_tag=vX.Y.Z`
- If the workflow reports a downgrade, duplicate, or tag/version mismatch, fix the published repo state first instead of forcing a lower or inconsistent version through

## Benefits of Unified Releases

1. **Single version** across all components ensures compatibility
2. **One tag** triggers all releases — simpler process
3. **Synchronized updates** — users always get matching versions
4. **Reduced coordination** — no need to remember multiple tag patterns
5. **Complete changelog** — all changes documented in one place, auto-updated via PR
6. **Faster releases** — parallel builds for independent components
7. **Dual distribution** — standalone exe (primary, no .NET needed) + NuGet (secondary, for .NET users)
8. **Self-contained VS Code** — extension bundles everything, no external dependencies
