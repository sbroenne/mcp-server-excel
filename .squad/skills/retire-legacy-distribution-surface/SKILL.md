---
name: "retire-legacy-distribution-surface"
description: "Pattern for safely removing an obsolete package/distribution surface without breaking the surviving release path."
---

## Context

Use this when a repo has replaced one publish/install surface with another, but old package directories or workflow steps still linger.

## Pattern

1. Prove the old surface is no longer the supported path.
2. Remove the legacy source artifacts/directories.
3. Remove any build script logic that still populates those artifacts.
4. Remove release workflow upload/publish steps for the retired surface.
5. Remove local validation/hooks that still expect the retired outputs.
6. Preserve and re-validate the replacement distribution path.
7. Search active operational files for stale references; leave historical archives/history alone.

## Why

- Prevents “half-retired” packaging where deleted directories are still expected by CI or pre-commit hooks.
- Keeps the supported distribution path clear and reduces release noise.
- Avoids accidentally reviving a dead package line just because automation still mentions it.

## Example

- Delete legacy `packages/excel-*-skill` directories.
- Remove the matching `release.yml` populate/upload/publish steps.
- Simplify `Build-AgentSkills.ps1` and `scripts/pre-commit.ps1` so they only validate the surviving ZIP artifact.
- Keep `publish-plugins.yml` and the `sbroenne/mcp-server-excel-plugins` marketplace flow intact.
