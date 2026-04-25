---
name: "cross-repo-release-preflight"
description: "Fail fast when a release workflow depends on credentials or access to a second repository."
domain: "release-management"
confidence: "high"
source: "earned"
---

## Context
Use this when a source repo release triggers a second workflow that clones, updates, or tags a separate published/distribution repo.

## Patterns
- Add a dedicated preflight job before the first cross-repo checkout.
- Validate the required secret exists with an explicit error message naming the exact secret the maintainer must add.
- Validate target repo reachability with the same credential the workflow will use for checkout/push.
- Add a source-side sync gate so downstream publication runs only when the published/install surface changed since the previous release tag.
- Add a manual `workflow_dispatch` path that can replay publication for an existing release tag without forcing a brand new source release.
- For cross-repo PAT auth, store the token as a repository secret and use it consistently for all checkout/push operations in the workflow.
- Alternatively, for GitHub App auth: keep the App ID in a repository variable, keep the private key in a repository secret, and mint a fresh installation token per job instead of sharing one long-lived token across the workflow.
- If the target repo owns the final commit, add defensive version guards there: reject explicit tag/version mismatches and reject downgrade syncs before overwriting the published surface.
- Let the duplicate guard distinguish automatic versus manual runs: automatic duplicates should skip cleanly, while manual runs should still be able to repair missing tags or re-sync the same version.
- Reflect the follow-on workflow in release docs and release verification checklists so maintainers confirm both workflows, not just the primary release.

## Examples
- `publish-plugins.yml` checks `PLUGINS_REPO_TOKEN` secret exists, then verifies access to `sbroenne/mcp-server-excel-plugins` before cloning it.
- `docs/RELEASE-STRATEGY.md` includes the follow-on plugin publish workflow in the release monitoring and verification steps.
- `mcaps-microsoft/iq-core-dev` uses a release-time sync gate so workflow-only/dev-only releases do not trigger downstream published-repo sync.
- `mcaps-microsoft/iq-core` rejects downgrade syncs and explicit tag/version mismatches before committing the copied plugin surface.
- `publish-plugins.yml` can expose `workflow_dispatch` with `release_tag=vX.Y.Z` so maintainers can replay the same published bundle path after a partial failure.

## Anti-Patterns
- Letting the first `actions/checkout` failure be the only signal that credentials are missing.
- Reusing a long-lived PAT or one shared app token across multiple jobs when each job can mint the minimal installation token it needs.
- Triggering cross-repo publication on every release even when nothing in the install surface changed.
- Accepting an explicit version/tag input without asserting it matches the incoming manifest version.
- Documenting only the primary release workflow when a second workflow performs the actual publication.
