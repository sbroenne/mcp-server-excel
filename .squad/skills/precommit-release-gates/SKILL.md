---
name: "precommit-release-gates"
description: "Diagnose and fix release-only build failures by promoting exact release commands into pre-commit hooks and aligning manifest metadata. Use when CI releases fail despite green local builds, pre-commit hooks miss packaging errors, version pins diverge between package.json and lockfiles, or VS Code extension packaging breaks. Triggers: release failure, pre-commit hook, packaging mismatch, CI pipeline error, works locally fails in CI, version conflict, npm run package, manifest sync."
---

## Context

Use this when a release workflow fails even though local builds or lighter smoke checks were green.

## Workflow

1. **Reproduce** the failure with the exact release command first (for VS Code extensions, `npm run package`, not just `npm install` or `npm run compile`).
2. **Align** authoritative config files if the failure is a manifest or metadata mismatch (for example `package.json` and `package-lock.json`).
3. **Promote** the exact release command into `scripts/pre-commit.ps1` if the failure is deterministic and commit-blocking.
4. **Update docs** that enumerate pre-commit gates — stale check lists are a docs bug.
5. **Verify** the fix locally:
   ```powershell
   ./scripts/pre-commit.ps1  # Must exit 0
   ```

## Examples

- `vscode-extension/package.json` and `vscode-extension/package-lock.json` both moved `engines.vscode` to `^1.110.0` to match `@types/vscode ^1.110.0`.
- `scripts/pre-commit.ps1` now runs the same `npm run package` path that release uses.

## Anti-Patterns

- Assuming install or compile coverage proves the package is releasable.
- Updating the hook without updating docs that enumerate its gates.
- Fixing the symptom in one manifest file while leaving the lockfile or related metadata stale.
