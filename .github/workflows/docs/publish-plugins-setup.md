# Phase 3: GitHub Copilot Plugin Publishing

## Overview

ExcelMcp plugins are published to the official GitHub Copilot CLI marketplace via a separate published marketplace repository. This document explains the maintainer-side workflow for automatic plugin republishing.

**Architecture:**
- **Source repo** (`sbroenne/mcp-server-excel`) — Development, releases, skill inputs, and plugin overlays
- **Published repo** (`sbroenne/mcp-server-excel-plugins`) — Official marketplace artifacts
- **Two plugins:** `excel-mcp` (MCP Server + skill) and `excel-cli` (CLI + skill)
- **Auto-sync:** `.github/workflows/publish-plugins.yml` copies templates after each release

**Trigger:** After "Release All Components" workflow completes successfully, the publish workflow automatically syncs plugin artifacts to the marketplace.

**User Impact:** GitHub Copilot CLI users can install both plugins via `copilot plugin install`.

**Key Design:**
- Validated plugin structure lives in the published repo, not the source repo
- Source repo stores overlay content in `.github/plugins/` (not installable plugin roots)
- Version and skills are injected during the publish workflow, not maintained separately

See [GitHub Copilot Plugin Distribution](../../docs/AWESOME-COPILOT-PROPOSAL.md) for the user-facing documentation.

## What can be automated from this environment?

- **Token creation:** **No** — you must create a PAT or obtain an app token outside this workflow and store it as a repository secret.
- **Source-repo wiring with the token:** **Yes** — store the secret with `gh`.
- **Workflow readiness checks:** **Yes** — this repo already contains a preflight gate in `publish-plugins.yml` that fails fast if the secret is missing or the published repo is unreachable.

### CLI command to store the token

```powershell
# Store the PAT or app token as a repository secret in the source repo
gh secret set PLUGINS_REPO_TOKEN --repo sbroenne/mcp-server-excel --body "<token-value>"
```

### Validate the repo-side wiring

```powershell
# Confirm the secret name exists (GitHub never returns the secret value)
gh secret list -R sbroenne/mcp-server-excel
```

---

## Required Repository Secret

The workflow needs write access to the published repository. Store a token in the source repo:

### Token Setup (Required)

Choose **one** of these options:

#### Option A: Personal Access Token (PAT)

1. Go to [GitHub Settings → Developer settings → Personal access tokens → Tokens (classic)](https://github.com/settings/tokens)
2. Click **Generate new token (classic)**
3. **Token name:** `ExcelMcp Plugin Publisher`
4. **Expiration:** 90 days (recommended; rotate every 90 days or manually when workflow fails)
5. **Scopes:** Select `public_repo` (minimum scope for publishing to a public repo)
6. Click **Generate token** and copy the token value
7. Store it in the source repo:
   ```powershell
   gh secret set PLUGINS_REPO_TOKEN --repo sbroenne/mcp-server-excel --body "<token-value>"
   ```

#### Option B: GitHub App Token

If you've already created a GitHub App for other purposes:
1. Generate a temporary app installation token from the app's settings
2. Store it as `PLUGINS_REPO_TOKEN` (same as above)

### Why Stored Token?

- ✅ Simple setup — one secret, no extra variables
- ✅ Works immediately — no browser-based app creation or installation flow
- ✅ Easy to rotate — update the secret when needed
- ✅ Same behavior as the legacy PAT approach

---

## Workflow Behavior

### Trigger Conditions
- ✅ Runs ONLY when "Release All Components" workflow completes successfully
- ✅ Runs ONLY on `main` branch releases
- ✅ Maintainers also get a manual re-sync entry point for repair/replay scenarios
- ❌ Does NOT run on failed releases
- ❌ Does NOT run on PR builds or test runs

### What It Does

1. **Resolve Tag + Version** — Extracts version from the triggering workflow's HEAD commit tag, or validates the manually supplied source release tag
2. **Source-Side Sync Gate** — Skips downstream publish when the plugin-published source surface did not change since the previous release tag
3. **Clone Repos** — Clones BOTH source and published repos
4. **Build Plugins** — Runs `scripts/Build-Plugins.ps1` which:
    - Copies validated plugin structure from `../mcp-server-excel-plugins/plugins/`
    - Applies source-owned overlays from `.github/plugins/` (overlay content only)
    - Updates `plugin.json` version and `version.txt`
    - Bundles `excelcli.exe` and companion publish output into `plugins/excel-cli/bin/`
    - Refreshes skills content from source repo (`skills/excel-mcp`, `skills/excel-cli`)
    - Refreshes shared references from source repo (`skills/shared/*.md`)
5. **Migrate Marketplace Layout** — Rewrites the published repo into the canonical marketplace layout by applying the source-owned root overlay, writing `.github/plugin/marketplace.json`, and removing any legacy root `marketplace.json`
6. **Published-Repo Guards** — Rejects downgrade or tag/version mismatch publishes before mutating the published repo
7. **Sync to Published Repo** — Only commits and pushes when the guarded sync path says publication is needed
8. **Create or Repair Tag** — Tags the published repo with the same version (for example `v1.2.3`) when the tag is missing
9. **Summary** — Generates workflow summary with the publish/skip decision and GitHub Copilot CLI install examples for changed published artifacts

### Version Extraction Strategy

**Corrected:** Uses `workflow_run.head_sha` plus the checked-out git tag graph to find the annotated source release tag created by the release workflow.

- ✅ Avoids race condition: Uses the exact commit that was just released
- ✅ No drift: If multiple releases happen close together, each publish uses the correct version
- ❌ Old (incorrect) approach: "latest release" could grab the wrong version in rapid succession

### Sync Gate

- The hardened source-side flow skips downstream plugin publication when the install-surface inputs have not changed since the prior release tag.
- Result: normal releases still publish all core artifacts, but plugin republishing only happens when plugin-facing content actually changed.

### Version and Tag Guards

- Published-side sync rejects downgrade attempts.
- Manual repair/replay runs must keep the requested tag/version aligned with the incoming plugin manifest/version metadata.
- The sync step now rewrites the published repo to the canonical marketplace layout on every needed publish, so legacy root-manifest state is repaired automatically.
- Result: maintainers can re-sync safely without accidentally stamping the wrong release tag onto plugin artifacts.

### Concurrency Control
- Only one publish workflow runs at a time
- Does NOT cancel in-progress runs (waits for completion)
- Prevents race conditions during concurrent releases

### Idempotency
- Automatic release-follow-on runs skip entirely when no plugin-published source files changed since the previous release tag
- Automatic duplicate publishes are skipped when the published repo already has the same version and tag
- Manual re-sync runs can replay an existing release tag without cutting a new source release
- If the published repo is already in sync, the workflow exits with a clear summary instead of making an empty commit

---

## Testing the Workflow

### Test After Token Setup

1. **Trigger a test release** (or wait for next real release):
   ```powershell
   # From source repo, trigger a release manually
   gh workflow run release.yml -f version_bump=patch
   ```

2. **Monitor the publish workflow**:
   ```powershell
   # Watch for publish-plugins workflow to start
   gh run watch

   # Or list recent runs
   gh run list --workflow=publish-plugins.yml
   ```

3. **Verify published repo updated**:
     ```powershell
     cd ../mcp-server-excel-plugins
     git pull
     git log -1  # Should see the latest publish commit
     git tag     # Should see new version tag
     Test-Path .github\plugin\marketplace.json  # Should be True after migration
     ```

### Manual Re-Sync

If the automatic follow-on publish needs to be replayed after a transient failure, use the workflow's manual `workflow_dispatch` entry point with an existing source release tag:

```powershell
gh workflow run publish-plugins.yml -f release_tag=v1.2.3
```

Keep the requested release tag aligned with the plugin manifest/version the workflow is syncing; the published-side guards reject mismatched or downgrade attempts.

### Troubleshooting

**Workflow fails with "Resource not accessible":**
- Token does not have write access to `sbroenne/mcp-server-excel-plugins`
- Token is expired or revoked
- Check that `PLUGINS_REPO_TOKEN` secret exists in the source repo

**Workflow fails immediately with a missing configuration message:**
- Add repository secret `PLUGINS_REPO_TOKEN` in `sbroenne/mcp-server-excel`
- Verify the secret contains a valid PAT or app token with write access to the published repo
- Rotate PAT if it has expired or been compromised

**Workflow completes but no commit in published repo:**
- The source-side sync gate detected no plugin-published source changes since the prior release tag
- OR: The published repo already had the same version and tag, so the automatic duplicate publish was skipped
- OR: Build-Plugins.ps1 generated identical content and the published repo stayed in sync without a new commit

**Workflow fails with a version/tag guard message:**
- Confirm the requested `release_tag` exists in the source repo and matches the plugin manifest/version being synced
- Check `.github/plugin/marketplace.json` (or the legacy root `marketplace.json` if the published repo has not been migrated yet) and existing tags in `sbroenne/mcp-server-excel-plugins`
- Downgrade attempts and inconsistent "tag exists but version differs" states are intentionally blocked

**Build step fails:**
- Build-Plugins.ps1 requires .NET 10.0 SDK
- VERSION file missing in `skills/excel-mcp/VERSION`

---

## File Locations

| File | Purpose | Location |
|------|---------|----------|
| `publish-plugins.yml` | Workflow definition | `.github/workflows/` (source repo) |
| `Build-Plugins.ps1` | Plugin build script | `scripts/` (source repo) |
| `Sync-PublishedPluginRepo.ps1` | Canonical published-repo sync script | `scripts/` (source repo) |
| This document | Setup instructions | `.github/workflows/docs/` (source repo) |

---

## Maintenance

### Updating Plugin Structure

If you change the plugin structure (add/remove files, change manifest schema):
1. Update `Build-Plugins.ps1` to reflect new structure
2. If the change affects plugin-owned helper files or published-repo root content, update the matching source overlay under `.github/plugins/`
3. Test locally: `./scripts/Build-Plugins.ps1 -Version 0.0.0`
4. Commit changes to source repo
5. Workflow will use updated script on next release

### Changing Published Repo Name

If you rename `mcp-server-excel-plugins`:
1. Update `PUBLISHED_REPO` env var in `publish-plugins.yml`
2. Rotate or update `PLUGINS_REPO_TOKEN` secret if needed
3. Update the published-repo metadata owned here (for example `.github/plugins/marketplace-repo/README.md`) so the next sync rewrites the target repo correctly

### Debugging Build Issues

Run the build + sync scripts locally to test:
```powershell
# From source repo root
./scripts/Build-Plugins.ps1 -Version 1.2.3
./scripts/Sync-PublishedPluginRepo.ps1 -PublishedRepoDir ..\mcp-server-excel-plugins -BuiltPluginsDir .\plugins -Version 1.2.3

# Verify output
ls plugins/
ls plugins/excel-mcp/
ls plugins/excel-cli/
Test-Path ..\mcp-server-excel-plugins\.github\plugin\marketplace.json
```

---

## Architecture Notes

### Why workflow_run?

The workflow uses `workflow_run` trigger with `head_sha` version extraction:
- ✅ **Avoids binary race condition** — Waits for release workflow to complete, ensuring GitHub Release artifacts exist
- ✅ **Atomic trigger** — One publish per release, no manual intervention
- ✅ **Version alignment** — Extracts tag from the exact commit that was just released (not "latest release")
- ✅ **No drift** — If multiple releases happen close together, each gets the correct version

**Version extraction logic:**
```yaml
HEAD_SHA="${{ github.event.workflow_run.head_sha }}"
TAG=$(gh api repos/.../git/matching-refs/tags/v --jq ".[] | select(.object.sha == \"$HEAD_SHA\") | .ref" ...)
```

### Why copy from validated templates?

**Build-Plugins.ps1 strategy:** COPY Phase 1/2 validated plugin structure, don't regenerate.

- ✅ **Preserves validated implementations** — Phase 1/2 created working bin scripts, READMEs, configs
- ✅ **Prevents regression** — Regenerating would introduce drift and stale content
- ✅ **Build script's job** — Version injection + skill refresh, NOT plugin authoring
- ❌ Old (incorrect) approach: Hand-authoring plugin content in build script

**What gets copied:**
- Plugin structure → From `../mcp-server-excel-plugins/plugins/` (validated Phase 1/2 implementations)
- Skills → From source repo `skills/excel-mcp`, `skills/excel-cli` (always fresh)
- Shared refs → From source repo `skills/shared/*.md` (always fresh)
- Marketplace ownership → Stays in the published repo; the source repo contributes overlays and automation, not a local marketplace manifest

**What gets updated:**
- `plugin.json` version field
- `version.txt` (for MCP plugin download script)

### Why two repos?

**Source repo** (`mcp-server-excel`):
- Development, testing, releases
- CI/CD, integration tests, documentation
- Binary build outputs (MCP Server, CLI)

**Published repo** (`mcp-server-excel-plugins`):
- Distribution only (lightweight marketplace)
- Canonical Copilot CLI marketplace manifest lives at `.github/plugin/marketplace.json`
- No build dependencies (just JSON, Markdown, PowerShell scripts)
- Clean separation: users don't clone 200MB source repo to get plugins

### Why not git submodules?

The published repo is NOT a submodule of the source repo. Instead:
- Workflow pushes built artifacts directly to published repo
- Published repo is standalone (easier for users to clone/fork)
- No submodule complexity for plugin consumers
- Allows published repo to have different README, docs, structure

---

## Success Criteria

✅ **Workflow created:** `.github/workflows/publish-plugins.yml`
✅ **Build script created:** `scripts/Build-Plugins.ps1`
✅ **Documentation created:** This file
⚠️ **Token configuration required:** User must add repository secret `PLUGINS_REPO_TOKEN`
⚠️ **First run test required:** Validate after next release

---

**Status:** Implementation complete, pending token setup and first-run validation
**Next Steps:** User must configure repository secret `PLUGINS_REPO_TOKEN` in `sbroenne/mcp-server-excel`, then validate both the automatic release-follow-on path and the manual `workflow_dispatch` re-sync path. The next successful sync will also migrate the published repo to the canonical marketplace manifest path/layout.
