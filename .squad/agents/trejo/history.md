# Trejo — History

## Core Context

- **Project:** A Windows COM interop MCP server and CLI for programmatic Excel automation with equal MCP Server and CLI entry points.
- **Role:** Docs Lead
- **Joined:** 2026-03-15T10:42:22.625Z

## Current Working Posture

- Keep user-facing documentation honest, concise, and aligned with the implemented release/install experience.
- Treat maintainer docs as the home for workflow mechanics, release gates, and recovery procedures.
- Preserve the distinction between plugins, skills, and MCP so install guidance stays surface-accurate.

## Cross-Agent Impact Notes

- **2026-04-24:** Kelso owns Copilot plugin packaging and publish automation; Trejo owns the maintainer/user documentation layer that explains those flows without overclaiming client support.

## Recent Work

### 2026-04-24: Publish Workflow Hardening Docs Sync
- Aligned maintainer docs with the hardened `publish-plugins.yml` flow: source-side sync gate, published-repo downgrade/tag-version guards, and manual `workflow_dispatch` replay via an existing `release_tag`.
- Kept user-facing wording to one accurate promise: plugin republishing is automatic but guarded, and install instructions remain client-specific.
- Recorded the docs-layering decision for Scribe merge work.

### 2026-04-24: GitHub App Publish Docs Sync
- Updated workflow setup, release strategy, and public wording together when cross-repo publication moved from PAT auth to GitHub App auth.
- Standardized the maintainer setup details around `PLUGINS_PUBLISH_APP_ID` plus `PLUGINS_PUBLISH_APP_PRIVATE_KEY`.

### 2026-04-24: Release Docs Cleanup
- Linked the main README release story to `docs\RELEASE-STRATEGY.md`.
- Added explicit Copilot plugin release coverage so the main release workflow and follow-on publish workflow are discoverable together.

### 2026-04-23: Plugin Distribution Documentation
- Updated source and published-repo plugin docs to reflect the validated two-plugin distribution story and the honest local-testing blockers.
- Kept counts aligned to the authoritative feature inventory and documented the release-asset dependency for first-time binary download.

## Learnings

- **Docs layering:** Put sync-gate, version/tag guard, and manual replay details in maintainer docs first; keep user-facing docs to concise, accurate statements.
- **GitHub App auth changes:** When publication auth changes, update workflow setup notes, release docs, and any “published automatically” wording together.
- **Plugin surface wording:** Describe published artifacts as GitHub Copilot plugins, but keep install commands scoped to the client flows we have actually validated.
- **Release discoverability:** If release mechanics change, make the canonical release doc discoverable from README instead of expecting contributors to infer workflow relationships.
- **Skills architecture:** Skills remain single-source guidance; plugin packaging wraps them, but should not fork or restate their behavioral content unnecessarily.

## Archive

- Detailed session history was moved to `.squad\agents\trejo\history-archive-2026-04-24.md` on 2026-04-24.
