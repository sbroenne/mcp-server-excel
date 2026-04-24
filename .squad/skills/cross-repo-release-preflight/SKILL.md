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
- Reflect the follow-on workflow in release docs and release verification checklists so maintainers confirm both workflows, not just the primary release.

## Examples
- `publish-plugins.yml` checks `PLUGINS_REPO_TOKEN` before cloning `sbroenne/mcp-server-excel-plugins`.
- `docs/RELEASE-STRATEGY.md` includes the follow-on plugin publish workflow in the release monitoring and verification steps.

## Anti-Patterns
- Letting the first `actions/checkout` failure be the only signal that credentials are missing.
- Documenting only the primary release workflow when a second workflow performs the actual publication.
