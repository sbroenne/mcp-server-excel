# Phase 3: Plugin Publishing Workflow Setup

## Overview

The `publish-plugins.yml` workflow automates synchronization of plugin artifacts from the source repository (`sbroenne/mcp-server-excel`) to the published repository (`sbroenne/mcp-server-excel-plugins`).

**Key Design:** Copies validated Phase 1/2 plugin structure, injects version, refreshes skills content.

**Trigger:** After successful completion of the "Release All Components" workflow
**Actions:** Builds plugins from validated templates, updates versions, commits to published repo, creates tags
**Authentication:** Requires `PLUGINS_REPO_TOKEN` secret

---

## Required Repository Secret

The workflow needs write access to the published repository. Configure the following:

### Fine-Grained Personal Access Token (Required)

1. Go to [GitHub Settings → Developer settings → Personal access tokens → Fine-grained tokens](https://github.com/settings/tokens?type=beta)
2. Click "Generate new token"
3. **Token name:** `plugins-repo-publish` (or similar)
4. **Resource owner:** `sbroenne` (your account)
5. **Repository access:** Select "Only select repositories" → Choose `mcp-server-excel-plugins`
6. **Permissions:**
   - Repository permissions:
     - Contents: **Read and write**
     - Metadata: **Read-only** (automatic)
7. Click "Generate token"
8. **Copy the token immediately** (you won't see it again!)

Then add to source repo:
1. Go to [Source Repo Settings → Secrets and variables → Actions](https://github.com/sbroenne/mcp-server-excel/settings/secrets/actions)
2. Click "New repository secret"
3. **Name:** `PLUGINS_REPO_TOKEN`
4. **Secret:** Paste the token
5. Click "Add secret"

**⚠️ Critical:** This token is used BOTH for:
- Cloning the published repo (to get validated plugin templates)
- Pushing updated plugins back to the published repo

---

## Workflow Behavior

### Trigger Conditions
- ✅ Runs ONLY when "Release All Components" workflow completes successfully
- ✅ Runs ONLY on `main` branch releases
- ❌ Does NOT run on failed releases
- ❌ Does NOT run on PR builds or test runs

### What It Does

1. **Get Version** — Extracts version from the triggering workflow's HEAD commit tag
2. **Clone Repos** — Clones BOTH source and published repos
3. **Build Plugins** — Runs `scripts/Build-Plugins.ps1` which:
   - Copies validated plugin structure from `../mcp-server-excel-plugins/plugins/`
   - Updates `plugin.json` version and `version.txt`
   - Refreshes skills content from source repo (`skills/excel-mcp`, `skills/excel-cli`)
   - Refreshes shared references from source repo (`skills/shared/*.md`)
4. **Sync to Published Repo** — Commits and pushes plugin updates
5. **Create Tag** — Tags the published repo with the same version (e.g., `v1.2.3`)
6. **Summary** — Generates workflow summary with install instructions

### Version Extraction Strategy

**Corrected:** Uses `workflow_run.head_sha` to find the tag created by the release workflow.

- ✅ Avoids race condition: Uses the exact commit that was just released
- ✅ No drift: If multiple releases happen close together, each publish uses the correct version
- ❌ Old (incorrect) approach: "latest release" could grab the wrong version in rapid succession

### Concurrency Control
- Only one publish workflow runs at a time
- Does NOT cancel in-progress runs (waits for completion)
- Prevents race conditions during concurrent releases

### Idempotency
- If plugins are already up to date, workflow completes with "No changes to commit"
- Safe to re-run on the same version (no duplicate commits)

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
   git log -1  # Should see commit from github-actions[bot]
   git tag     # Should see new version tag
   ```

### Troubleshooting

**Workflow fails with "Resource not accessible by integration":**
- Secret is missing or has wrong name
- Token lacks required permissions (needs `contents: write`)
- Token expired or was revoked

**Workflow completes but no commit in published repo:**
- Plugins were already up to date (check "No changes to commit" in workflow logs)
- OR: Build-Plugins.ps1 generated identical content

**Build step fails:**
- Build-Plugins.ps1 requires .NET 10.0 SDK
- VERSION file missing in `skills/excel-mcp/VERSION`

---

## File Locations

| File | Purpose | Location |
|------|---------|----------|
| `publish-plugins.yml` | Workflow definition | `.github/workflows/` (source repo) |
| `Build-Plugins.ps1` | Plugin build script | `scripts/` (source repo) |
| This document | Setup instructions | `.github/workflows/docs/` (source repo) |

---

## Maintenance

### Updating Plugin Structure

If you change the plugin structure (add/remove files, change manifest schema):
1. Update `Build-Plugins.ps1` to reflect new structure
2. Test locally: `./scripts/Build-Plugins.ps1 -Version 0.0.0`
3. Commit changes to source repo
4. Workflow will use updated script on next release

### Changing Published Repo Name

If you rename `mcp-server-excel-plugins`:
1. Update `PUBLISHED_REPO` env var in `publish-plugins.yml`
2. Update `PLUGINS_REPO_TOKEN` secret to grant access to new repo
3. Update `repository.url` in plugin.json templates in `Build-Plugins.ps1`

### Debugging Build Issues

Run the build script locally to test:
```powershell
# From source repo root
./scripts/Build-Plugins.ps1 -Version 1.2.3

# Verify output
ls plugins/
ls plugins/excel-mcp/
ls plugins/excel-cli/
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
⚠️ **Secret configuration required:** User must add `PLUGINS_REPO_TOKEN`
⚠️ **First run test required:** Validate after next release

---

**Status:** Implementation complete, pending token configuration and first-run validation
**Next Steps:** User must configure `PLUGINS_REPO_TOKEN` secret, then test on next release
