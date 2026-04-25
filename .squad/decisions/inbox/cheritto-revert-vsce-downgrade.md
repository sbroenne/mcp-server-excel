# Cheritto Decision Inbox — Revert `@vscode/vsce` Downgrade

## Context

Main carried the merged `@vscode/vsce` downgrade from `^3.7.1` to `^2.25.0` in `vscode-extension/package.json` and `package-lock.json`. Stefan explicitly asked for an immediate corrective revert while keeping the later workflow and release fixes intact.

## Decision

Revert only the VS Code extension packaging dependency downgrade by restoring the pre-PR-609 `package.json` and `package-lock.json` state (`@vscode/vsce` back to `^3.7.1`), with no additional workflow or release-file churn in the corrective branch.

## Why

- The regression call was specifically about the downgrade itself, not the broader workflow fixes that landed around it.
- A narrow revert keeps PR #613 and the other release/workflow fixes intact.
- Validation proved the packaging surface can still be exercised successfully after the revert when the exact historical lockfile is replayed with a clean `npm ci` and `npm run package`.

## Validation

- `npm ci --ignore-scripts` in `vscode-extension` using the restored historical lockfile
- `npm run package` in `vscode-extension` → produced `excel-mcp-1.6.9.vsix`
- `npm audit --json` still reports the known moderate advisory chain reintroduced by `@vscode/vsce` 3.7.1; that remains intentionally out of scope for this emergency revert
