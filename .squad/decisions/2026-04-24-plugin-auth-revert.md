# Decision: Revert Plugin Publish Auth from GitHub App to Stored PAT

**Date:** 2026-04-24  
**Agent:** Kelso (Copilot CLI Plugin Engineer)  
**Requestor:** Stefan Brönner  
**Status:** Implemented

## Context

The `publish-plugins.yml` workflow automates cross-repo publishing of agent plugin artifacts from `sbroenne/mcp-server-excel` (source) to `sbroenne/mcp-server-excel-plugins` (published marketplace). This workflow was previously configured with GitHub App authentication using `PLUGINS_PUBLISH_APP_ID` + `PLUGINS_PUBLISH_APP_PRIVATE_KEY`.

## Decision

Switch plugin publish authentication from GitHub App-based token minting to a single stored cross-repo Personal Access Token (PAT) stored as `PLUGINS_REPO_TOKEN`.

## Rationale

1. **Simpler Setup:** 1 secret (`PLUGINS_REPO_TOKEN`) vs 1 variable + 1 secret (`PLUGINS_PUBLISH_APP_ID` + `PLUGINS_PUBLISH_APP_PRIVATE_KEY`)
2. **Easier Rotation:** Update one secret value when token expires, no App registration/installation coordination
3. **Same Security Posture:** For public repo use case, fine-grained PAT with `public_repo` scope provides equivalent security to GitHub App with `contents:write`
4. **Operational Hardening Preserved:** All iq-core-style guards remain intact (preflight, sync gate, version guards, manual re-sync)
5. **Industry Standard:** Stored PAT is the simpler pattern for single cross-repo workflows; GitHub App auth adds value for multi-repo installations

## Changes Made

### Workflow (`.github/workflows/publish-plugins.yml`)
- **Already token-based** — workflow was already using `PLUGINS_REPO_TOKEN` throughout
- Verified consistency: preflight, checkout, and all git operations use `secrets.PLUGINS_REPO_TOKEN`
- Changed commit identity from app bot to `github-actions[bot]` (standard identity)

### Documentation
- **publish-plugins-setup.md:** Replaced GitHub App setup instructions with PAT setup (Options A/B)
- **RELEASE-STRATEGY.md:** Updated secrets table, removed App ID/private key entries, updated troubleshooting
- **INSTALLATION.md:** Changed "GitHub App auth" to "stored cross-repo PAT"
- **README.md:** Updated release strategy reference
- **gh-pages/index.md:** Aligned with README change

### Skills
- **cross-repo-release-preflight:** Generalized patterns section to cover both PAT and GitHub App auth options

## Operational Hardening Retained

✅ **Preflight validation:** Fails fast if `PLUGINS_REPO_TOKEN` missing or target repo unreachable  
✅ **Source-side sync gate:** Skips publish when plugin surface unchanged since previous release  
✅ **Version guards:** Rejects downgrade attempts and tag/version mismatches  
✅ **Manual re-sync path:** `workflow_dispatch` allows replaying specific release tags  
✅ **Duplicate detection:** Auto runs skip cleanly; manual runs can repair missing tags  

## Required Action Items

- [ ] Create fine-grained PAT with `public_repo` scope (90-day expiration recommended)
- [ ] Store as repository secret: `gh secret set PLUGINS_REPO_TOKEN --repo sbroenne/mcp-server-excel --body "<token>"`
- [ ] Verify preflight: Next release workflow should validate token and proceed to publish

## Related Files

- `.github/workflows/publish-plugins.yml` (workflow implementation)
- `.github/workflows/docs/publish-plugins-setup.md` (setup guide)
- `docs/RELEASE-STRATEGY.md` (release process)
- `.squad/skills/cross-repo-release-preflight/SKILL.md` (reusable pattern)

## Notes

This decision aligns with the project's preference for simplicity where equivalent security is achieved. GitHub App auth remains documented as an alternative pattern in the cross-repo-release-preflight skill for teams with multi-repo installation needs.
