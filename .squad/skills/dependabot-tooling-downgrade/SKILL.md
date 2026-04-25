---
name: "dependabot-tooling-downgrade"
description: "Use a validated tooling downgrade when Dependabot flags an unpatchable transitive vulnerability in build-only dependencies."
domain: "dependency-management"
confidence: "medium"
source: "earned"
---

## Context

Use this when a Dependabot or `npm audit` failure comes from a dev-only packaging/build tool and the current major line has no viable patched transitive path.

## Patterns

1. Confirm the failing dependency chain from the Dependabot or audit logs before changing anything.
2. Check whether Dependabot's suggested replacement version actually removes the vulnerable transitive stack.
3. Prefer a real dependency change over ignoring the alert when the downgraded tool still supports the repo's required workflow.
4. If you later need to roll the downgrade back, restore the exact historical lockfile and validate from a clean `npm ci`; a generic `npm install` can drift transitive packages and create fake packaging regressions.
5. Validate the exact release command after the change, not just `npm install` or `npm audit`.
6. Keep the fix surgical: update the dependency, refresh or restore the lockfile intentionally, and verify the packaging path.

## Examples

- `vscode-extension` moved from `@vscode/vsce` `^3.7.1` to `^2.25.0` after Dependabot showed the 3.x line was stuck on `@azure/msal-node -> uuid@^8.3.0`; `npm audit` and `npm run package` both passed afterward.

## Anti-Patterns

- Ignoring the alert before proving there is no workable package-level fix.
- Assuming the latest version is always the safest path.
- Validating only with `npm audit` while skipping the real release/package command.
