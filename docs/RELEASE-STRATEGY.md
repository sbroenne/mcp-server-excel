# ExcelMcp Release Strategy

This document outlines the unified release process for all ExcelMcp components.

## Overview

All ExcelMcp components are released together with a single version tag:

| Component | Primary Distribution | Secondary Distribution | Description |
|-----------|---------------------|----------------------|-------------|
| **MCP Server** | npm + standalone exe ZIP | NuGet (.NET tool) | `npx -y @sbroenne/mcp-server-excel@latest` or `mcp-excel.exe` — no .NET runtime required |
| **CLI** | npm + standalone exe ZIP | NuGet (.NET tool) | `npx -y @sbroenne/excelcli@latest` or `excelcli.exe` — no .NET runtime required |
| **VS Code Extension** | VSIX + Marketplace | — | Self-contained — bundles MCP Server and its skill |
| **MCPB** | Claude Desktop bundle | — | Direct npx configuration with `@latest`; requires Node.js/npm on PATH |
| **GitHub Copilot Plugins** | Published plugin marketplace | — | `excel-mcp` and `excel-cli` plugins with npx launch configuration, an argument-safe CLI wrapper, and skills |
| **Agent Skills** | GitHub Release ZIP | Direct skill extraction | Reusable skill packages for AI coding assistants (`npx skills add`) |

## Unified Release Workflow

**Workflow**: `.github/workflows/release.yml`
**Trigger**: `workflow_dispatch` with version bump (major/minor/patch) or custom version

### What Gets Released

When you run the release workflow, all components are released together:

1. **CLI** → Standalone self-contained exe (`excelcli.exe`) shipped as:
   - npm launcher and Windows runtime packages (primary distribution)
   - ZIP file (primary distribution)
   - NuGet package (secondary distribution)
2. **MCP Server** → npm launcher and Windows runtime packages + standalone self-contained exe ZIP [primary] + NuGet pack [secondary]
3. **VS Code Extension** → Self-contained Windows x64 and ARM64 VSIX packages (bundle the MCP executable and skill) → VS Code Marketplace
4. **MCPB** → Claude Desktop bundle (`.mcpb` file)
5. **Agent Skills** → ZIP package for AI coding assistants
6. **GitHub Copilot Plugins** → `publish-plugins.yml` compares prepared output and
   publishes only real content changes; unchanged plugins retain their prior
   version/tag while npx launchers use the latest npm runtime, with npm-managed
   resolution and caching (see
   [Plugin Publishing](../.github/workflows/docs/publish-plugins-setup.md))
7. **NuGet** → Both packages published to NuGet.org (secondary channel)
8. **MCP Registry** → Updated after NuGet and npm propagation
9. **GitHub Release** → Created with all artifacts and the prepared changelog notes

### Release Artifacts

| Artifact | Format | Distribution |
|----------|--------|--------------|
| `@sbroenne/mcp-server-excel@{version}` | npm | npm registry (primary launcher package) |
| `@sbroenne/mcp-server-excel-win32-x64@{version}` | npm | npm registry (self-contained Windows runtime) |
| `@sbroenne/mcp-server-excel-win32-arm64@{version}` | npm | npm registry (self-contained native Windows ARM64 runtime) |
| `@sbroenne/excelcli@{version}` | npm | npm registry (primary CLI launcher package) |
| `@sbroenne/excelcli-win32-x64@{version}` | npm | npm registry (self-contained Windows CLI runtime) |
| `@sbroenne/excelcli-win32-arm64@{version}` | npm | npm registry (self-contained native Windows ARM64 CLI runtime) |
| `ExcelMcp-MCP-Server-{version}-windows.zip` | ZIP | GitHub Release (primary — contains `mcp-excel.exe`) |
| `ExcelMcp-CLI-{version}-windows.zip` | ZIP | GitHub Release (primary — contains `excelcli.exe`) |
| `SHA256SUMS` | GNU-style SHA-256 manifest (`<hash>  <filename>`) | GitHub Release (covers both Windows runtime ZIPs and the prepared plugin ZIP) |
| `Sbroenne.ExcelMcp.CLI.{version}.nupkg` | NuGet | NuGet.org (secondary — contains `excelcli.exe`, requires .NET 10 runtime) |
| `Sbroenne.ExcelMcp.McpServer.{version}.nupkg` | NuGet | NuGet.org (secondary — contains `mcp-excel.exe`, requires .NET 10 runtime) |
| `excel-skills-v{version}.zip` | ZIP | GitHub Release (contains `excel-cli` + `excel-mcp` skills for direct extraction) |
| `excel-mcp-{version}.vsix` | VSIX | GitHub Release + VS Code Marketplace (Windows x64; self-contained MCP executable and skill) |
| `excel-mcp-{version}-win32-arm64.vsix` | VSIX | GitHub Release + VS Code Marketplace (Windows ARM64; same x64 MCP executable through Windows emulation, plus skill) |
| `excel-mcp-{version}.mcpb` | MCPB | GitHub Release (Claude Desktop metadata bundle; server fetched through npx with `@latest`) |
| `excel-plugins-v{version}.zip` | ZIP | GitHub Release (prepared plugin payload for exact-release repairs) |

## Release Process

### 1. Add a Changeset (every PR, not just before release)

`CHANGELOG.md` is no longer hand-edited. Every PR that changes user-visible
behavior adds a small **changeset fragment** describing the change, and a CI
check (`.github/workflows/changeset-check.yml`) fails the PR if one is missing
(unless the PR is labeled `skip-changelog`). See [`.changeset/README.md`](../.changeset/README.md)
for the full guide and examples of good vs. too-technical entries.

```powershell
npx changeset
# → pick a bump type (metadata only, doesn't drive the real release version)
# → write a short, end-user-facing summary
# → commit the generated .changeset/<name>.md with your PR
```

Nothing needs to happen "before creating a release tag" — fragments accumulate
in `.changeset/` across PRs and are compiled automatically when the release
workflow runs (see [Changelog Generation](#changelog-generation) below).

### 2. Run the Release Workflow

1. Go to **Actions** → **Release All Components** → **Run workflow**
2. Select the version bump type:
   - **patch** (default): `1.5.6` → `1.5.7`
   - **minor**: `1.5.6` → `1.6.0`
   - **major**: `1.5.6` → `2.0.0`
3. Or enter a **custom version** (e.g., `1.5.7`) to override the bump

The workflow will:
1. Calculate the next version from the latest git tag
2. Consume the counts reviewed in the source PR and compile pending changesets into the current release changelog
3. Build all components (npm packages, standalone exes, and NuGet packages), packaging the generated documentation where applicable
4. Commit the generated release metadata and create the git tag (`v1.5.7`) at that commit
5. Publish to npm, NuGet.org, VS Code Marketplace, and MCP Registry
6. Create the GitHub Release with all artifacts and the prepared release notes

### 3. Monitor Workflow

The release shares prepared inputs instead of repeating builds in each package job:

1. **version** → Calculates the version from the latest tag and dispatch input
2. **prepare-release** → Compiles changesets once and uploads the exact metadata patch and notes
3. **build-packages** → Applies that patch and calls `Build-ReleasePackages.ps1` for all NuGet, npm, runtime ZIP, VSIX, MCPB, skill and plugin outputs. ARM64 archives are checked here; `verify-arm64` then installs and executes both prepared npm distributions on native Windows ARM64 before tag creation.
4. **create-tag** → Applies the same patch, checks that `main` has not advanced, commits only allowed release metadata, then tags that commit
5. **create-release** → Publishes GitHub assets, checksums and prepared notes
6. **publish** → Publishes npm and NuGet packages
7. **publish-vscode** → Publishes the already verified VSIX independently
8. **publish-mcp-registry** → Waits for matching npm/NuGet metadata and registers the release
9. **publish-plugins** → Calls the reusable publisher after GitHub assets exist, passing exact release identity and prepared plugins

Registry propagation failures do not suppress plugin publication or GitHub assets.
Each distribution reports its own result; repair only the failed destination.

### 4. Verify Release

After workflow completes:

- [ ] GitHub Release created with all artifacts (MCP Server ZIP, CLI ZIP, `SHA256SUMS`, VSIX, MCPB, skills ZIP)
- [ ] All six npm packages (MCP Server and CLI launchers plus their x64 and ARM64 Windows runtimes) are available at the release version
- [ ] NuGet packages available on NuGet.org (may take 10-30 min for full propagation)
- [ ] VS Code Marketplace updated (verify self-contained extension works without .NET)
- [ ] MCP Registry updated
- [ ] `publish-plugins.yml` completed; if the sync gate detected plugin-facing changes, `sbroenne/mcp-server-excel-plugins` was updated
- [ ] Release tag contains the current version in `CHANGELOG.md` and points to the release metadata commit on `main`

### 5. Agent Plugin Publishing (Automatic)

**Workflow**: `.github/workflows/publish-plugins.yml`
**Trigger**: Called by `release.yml` after GitHub assets exist, with a manual `workflow_dispatch` repair path for existing source release tags
**Published Repo**: `sbroenne/mcp-server-excel-plugins` (the actual Copilot CLI marketplace repo)

The `publish-plugins.yml` workflow consumes prepared release plugins:

1. **Verifies tag, commit and version** explicitly supplied by `release.yml`
2. **Consumes prepared output**, or the exact requested release's payload for manual repair
3. **Uses plugins assembled earlier** via `scripts/Build-Plugins.ps1`:
    - Copies canonical plugin structure from this source repository
    - Copies npx launch configuration and the argument-safe CLI wrapper from `.github/plugins/`
    - Strips committed runtime payloads from plugin bundles so the published repo contains no bundled runtimes
    - Stamps plugin.json, version.txt and skill VERSION for candidate validation,
      while npx launchers still target the latest npm runtime
    - Consumes complete generated skills from the released source and stamps the release version
4. **Checks published-repo guards** before mutation, reading the current published plugin version from the canonical marketplace manifest when present (or the legacy root manifest before migration):
    - Rejects explicit tag/version mismatches
    - Rejects downgrade publishes
    - Compares the complete prepared tree using the
      [exact normalization rules](../.github/workflows/docs/publish-plugins-setup.md#publish-only-changed-output)
      before destination writes
5. **Publishes plugin artifacts** only for real changes. Equivalent output skips
   commit/push/tag entirely, retaining the actual earlier plugin version/tag.
   Root-overlay changes are included; root-only publication does not update
   Awesome Copilot listings.
6. **Optional listing update** calls the guarded
   [Awesome Copilot updater](../.github/workflows/docs/awesome-copilot-update-setup.md)
   only after real changed-plugin publication and explicit opt-in.

Maintainers can also replay plugin publication for an existing release tag without cutting a new release:

```powershell
gh workflow run publish-plugins.yml -f release_tag=v1.2.3
```

**Key Points:**
- ✅ **Automatic** — No manual intervention required
- ✅ **Idempotent** — Safe to re-run on the same version
- ✅ **Exact candidate version** — A real publication uses the source release's
  version; unchanged output retains its earlier plugin version/tag
- ✅ **No stale fallback** — Distributable builds require an explicit version; canonical skill sources contain no `VERSION` file
- ✅ **Output-compared** — version-only prepared output creates no commit, push
  or tag; product/npm publication still proceeds
- ✅ **Guarded replay** — downgrade syncs are rejected, automatic duplicates are skipped, and manual repair/replay runs must keep the requested tag aligned with the incoming plugin manifest/version
- ✅ **Manual repair path** — maintainers keep a `workflow_dispatch` re-sync entry point for repair/replay scenarios
- ⚠️ **Requires cross-repo token** — First-time setup needs a repository secret `PLUGINS_REPO_TOKEN` in the source repo. Use either a PAT with `public_repo` scope or an app token with `contents:write` on `sbroenne/mcp-server-excel-plugins` (see [Phase 3 Plugin Publishing docs](../.github/workflows/docs/publish-plugins-setup.md))
- ℹ️ **Setup command** — Enter it through the interactive prompt:
  `gh secret set PLUGINS_REPO_TOKEN --repo sbroenne/mcp-server-excel`

**Surface note:**
- The release automation publishes plugin bundles (manifests, skills, agents, MCP config, and the CLI wrapper) to the published repo.
- Those bundles intentionally exclude self-contained runtime binaries. They use `npx -y @sbroenne/mcp-server-excel@latest` or `npx -y @sbroenne/excelcli@latest`; npm manages package resolution and caching. Plugins do not download GitHub release ZIPs or install global helpers.
- Plugin versions can intentionally lag product/npm versions.
- The published repo is the marketplace; this source repo only owns inputs, overlays, and automation.
- Those artifacts can be relevant across multiple plugin-capable clients, but marketplace registration, discovery, and installation UX remain client-specific.
- The current workflow and docs only claim a verified GitHub Copilot install flow; they do **not** claim automatic publication into VS Code or Claude-specific plugin marketplaces.

**Hardening note:**
- Publication uses explicit release identity and prepared output, not an incomplete source-path comparison.
- The published-side sync path rejects downgrade attempts, rewrites the published repo to the canonical `.github/plugin/marketplace.json` layout, and keeps explicit repair/replay runs honest by requiring tag/version alignment.
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

## Changelog Generation

`CHANGELOG.md` is **end-user facing** and generated from [changesets](https://github.com/changesets/changesets), not hand-edited. This replaced a manual `## [Unreleased]` section that repeatedly went stale — the auto-rename step below relied on a PR merge that didn't always happen, so several released versions (v1.8.64–v1.9.0) sat mislabeled as "Unreleased" for weeks. See the changeset fragments' own guide at [`.changeset/README.md`](../.changeset/README.md) for day-to-day usage.

**How it works:**

1. **Every PR** that changes user-visible behavior adds a small markdown fragment via `npx changeset` (from repo root), describing the change in 1-3 end-user-facing sentences. This is committed to `.changeset/<random-name>.md` alongside the PR.
2. **CI enforces this** via `.github/workflows/changeset-check.yml`, which fails the PR if no fragment was added — unless the PR carries the `skip-changelog` label (docs/tests/CI/dependency-only changes).
3. **At release time**, the `prepare-release` job in `release.yml` runs `scripts/Build-Changelog.ps1 -Version <version> -Date <date>`, which:
   - Runs `npx changeset version` to consume every pending fragment, deleting them.
   - Normalizes the tool's generated version header to this repo's `## [X.Y.Z] - YYYY-MM-DD` (Keep a Changelog) style.
   - Extracts the new section verbatim into `release_notes_body.md`, used directly as the "What's New" body of the GitHub Release.
4. **Artifact builds** consume that prepared changelog, so packaged VS Code and MCPB changelog files include the version being released.
5. **After all builds pass**, the `create-tag` job regenerates the metadata using the same release date and verifies it byte-for-byte against the prepared artifact. It commits `CHANGELOG.md`, synchronized version metadata, and consumed `.changeset/*.md` deletions to `main` through the Git Data API, then points the release tag at that exact commit.

Advertised tool and operation totals are generated from code by running
`scripts/check-doc-counts.ps1 -Update` and reviewed in the same PR as the code.
Strict CI rejects stale counts before merge. The command writes the canonical include file
`doc-counts.json` (repo root) plus every managed headline claim across the
repository. Release automation and the website never derive or restate these
numbers themselves — they read `doc-counts.json` or the headlines it already
wrote, both of which are always current on `main` by the time a release or a
site build runs. Feature PRs update operation tables and per-category counts,
but do not manually update repeated headline totals.

Root `package.json` and `.changeset/config.json` (using `@changesets/changelog-github` for PR-linked entries) exist solely to drive this tooling — they have no bearing on the actual MCP Server / CLI / VS Code Extension / MCPB version, which remains fully controlled by the `version_bump` / `custom_version` workflow inputs as described above.

```markdown
# Changelog

## [Unreleased]

_No pending changes. New entries are added automatically from changesets when a release is cut._

## [1.9.1] - 2025-01-21

### Minor Changes
- **Feature description** (#123): what changed and why it matters to users.

### Patch Changes
- **Bug fix description** (#124): what was broken and what's fixed now.
```

> **Why the Git Data API?** Branch protection prevents the default workflow token from pushing the bookkeeping commit directly. `RELEASE_PAT` belongs to the repository administrator and has `contents: write`, allowing the workflow to create a fast-forward metadata commit before tagging without opening a release-time pull request.

## Required Secrets and Variables

Configure these GitHub repository secrets and variables:

| Type | Name | Purpose |
|------|------|---------|
| Secret | `NUGET_USER` | NuGet.org username (for OIDC trusted publishing) |
| Secret | `NPM_TOKEN` | Granular npm publish token for the first package publish and fallback authentication |
| Secret | `VSCE_TOKEN` | VS Code Marketplace PAT |
| Secret | `APPINSIGHTS_CONNECTION_STRING` | Application Insights (optional telemetry) |
| Secret | `RELEASE_PAT` | Repository-admin token with `contents: write` for the pre-tag release metadata commit |
| Secret | `PLUGINS_REPO_TOKEN` | Cross-repo PAT (`public_repo`) or app token (`contents:write`) for publishing plugins to `sbroenne/mcp-server-excel-plugins` |

> **Notes:**
> - NuGet uses OIDC trusted publishing (no API key needed). The `NUGET_USER` is just the NuGet.org profile name for OIDC token exchange.
> - npm trusted publishing cannot bootstrap a new package. Use `NPM_TOKEN` for the first release of each new package, including the two ARM64 runtimes. Configure `release.yml` as the trusted publisher for all six npm packages (MCP Server and CLI launchers and their x64/ARM64 runtimes), then remove the token; npm automatically prefers OIDC and generates provenance.
> - The follow-on plugin publish workflow uses a stored cross-repo token (`PLUGINS_REPO_TOKEN`) with write access to the published plugin repo. A PAT needs `public_repo`; an app token needs `contents:write`.

## Troubleshooting

### NuGet Publishing Fails

- Verify `NUGET_USER` secret is set to your NuGet.org profile name (not email)
- Check NuGet.org trusted publishers are configured for OIDC

### npm Publishing Fails

- For the first release, verify `NPM_TOKEN` can publish public packages under the `@sbroenne` scope
- After the first release, configure all six npm packages to trust the `release.yml` GitHub Actions workflow
- Confirm the workflow has `id-token: write` and uses npm 11 or later

### npm Packaging Development

`npm-packages/shared/launcher.js` is copied into each staged launcher's `lib/`
by `scripts/Build-NpmPackages.ps1`; do not maintain separate launcher
implementations. Package directories are build inputs, not directly runnable
source installations. The shared tests cover both launchers and inspect real
tarballs built with fixture payloads (no Excel required):

```powershell
npm ci --prefix npm-packages/shared --ignore-scripts
npm test --prefix npm-packages/shared
```

On Windows, build and smoke-test each real runtime with
`scripts/Build-NpmPackages.ps1` and `scripts/Test-NpmPackages.ps1`.
Pass `-Component Cli` for `excelcli`; the default remains `McpServer`.
Pass `-Architecture x64` (the default) or `-Architecture arm64` to both scripts.
The executable's PE machine type must match the package architecture.
Both launcher dependencies are stamped to the same release version.

`Build-ReleasePackages.ps1` builds both npm architectures while preserving x64
payloads for standalone ZIPs and other bundles. All four runtime packages are
published before either launcher. ARM64 Node.js selects ARM64; x64 Node.js
selects x64, even on ARM64 Windows. Missing matching runtimes fail explicitly.

`Test-NpmPackages.ps1` inspects both archives, but installs and executes a runtime
only when Node.js matches its architecture. On the x64 hosted release runner,
ARM64 archive validation runs and ARM64 execution is reported as **not run**.
Validate native ARM64 execution and real Excel automation locally on Windows
ARM64 before release. Package installation, help, and MCP discovery alone do
not establish Excel compatibility.

The CLI smoke test checks help, version, subcommand arguments, output, and
failure exit codes; the MCP smoke test checks initialization and tool discovery.
These smoke tests do not exercise Excel automation.

### VS Code Marketplace Fails

- Verify `VSCE_TOKEN` is valid and not expired
- Check extension ID matches marketplace listing

### MCPB Build Fails

- Ensure `mcpb/manifest.json` is valid JSON
- Verify `mcpb/icon-512.png` exists (512x512 PNG)

### MCP Registry Update Fails

- MCP Registry update uses GitHub OIDC
- Manually run the **Publish MCP Registry** workflow with the exact existing
  release tag
- The repair requires owner approval through the protected `mcp-registry`
  environment, rejects tag commits not reachable from protected `main`, and
  validates the source manifest plus published NuGet and npm metadata
- Repository settings for `mcp-registry` must retain a custom deployment branch
  policy of exactly `main` and the repository owner as a required reviewer
- The workflow publishes only the MCP Registry entry
- Do not rerun the unified release to repair a registry-only failure

### Publish Plugins Fails

- Confirm the repository secret `PLUGINS_REPO_TOKEN` exists in `sbroenne/mcp-server-excel`
- Confirm the token is valid for the published repo (PAT with `public_repo`, or app token with `contents:write`)
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
7. **Multiple distribution choices** — npm and standalone exe (primary, no .NET needed) + NuGet (secondary, for .NET users)
8. **Self-contained VS Code** — extension bundles everything, no external dependencies
