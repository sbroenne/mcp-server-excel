---
name: "precommit-release-gates"
description: "How to harden pre-commit against release-only failures, especially packaging mismatches"
domain: "release-management"
confidence: "high"
source: "earned"
---

## Context
Use this when a release workflow fails even though local builds or lighter smoke checks were green.

## Patterns
- Reproduce the failure with the exact release command first (for VS Code extensions, `npm run package`, not just `npm install` or `npm run compile`).
- Promote that exact release command into `scripts/pre-commit.ps1` if the failure is deterministic and commit-blocking.
- Keep hook documentation synchronized with the real script; stale check lists are a docs bug.
- If the failure is a manifest or metadata mismatch, align the authoritative config files together (for example `package.json` and `package-lock.json`).

## Examples
- `vscode-extension/package.json` and `vscode-extension/package-lock.json` both moved `engines.vscode` to `^1.110.0` to match `@types/vscode ^1.110.0`.
- `scripts/pre-commit.ps1` now runs the same `npm run package` path that release uses.

## Anti-Patterns
- Assuming install or compile coverage proves the package is releasable.
- Updating the hook without updating docs that enumerate its gates.
- Fixing the symptom in one manifest file while leaving the lockfile or related metadata stale.
