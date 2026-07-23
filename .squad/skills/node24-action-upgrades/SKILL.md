---
name: "node24-action-upgrades"
description: "Upgrade checked-in GitHub Actions references off deprecated Node 20 runtimes with the smallest safe major bumps."
domain: "github-actions"
confidence: "high"
source: "earned"
---

## Context

Use this when GitHub Actions logs report `Node.js 20 actions are deprecated` and the repo needs a focused cleanup without rewriting release logic.

## Patterns

1. Start from real run warnings, not just static grep, so you know which checked-in refs actually emitted the deprecation message.
2. For each warned action, inspect the candidate major tag's `action.yml` and confirm `runs.using: node24` before changing the workflow.
3. Prefer the first Node24-capable major rather than the newest possible major if that avoids unrelated workflow churn.
4. For composite actions, inspect nested `uses:` lines too (for example `upload-pages-artifact` can hide an internal `upload-artifact` pin).
5. If the latest available upstream tag still declares `node20`, leave it explicit in the report rather than forcing `FORCE_JAVASCRIPT_ACTIONS_TO_NODE24` without validation.

## Example mapping

- `actions/checkout@v4` -> `actions/checkout@v5`
- `actions/setup-dotnet@v4` -> `actions/setup-dotnet@v5`
- `actions/setup-node@v4` -> `actions/setup-node@v5`
- `actions/upload-artifact@v4` -> `actions/upload-artifact@v6`
- `actions/download-artifact@v4` -> `actions/download-artifact@v7`
- `actions/configure-pages@v4` -> `actions/configure-pages@v6`
- `actions/upload-pages-artifact@v3` -> `actions/upload-pages-artifact@v5`
- `actions/deploy-pages@v4` -> `actions/deploy-pages@v5`

## Anti-Patterns

- Blanket-search-replacing to the latest major without checking runtime metadata.
- Treating composite actions as safe without checking their internal `uses:` pins.
- Hiding unresolved upstream Node20 actions instead of reporting them plainly.
