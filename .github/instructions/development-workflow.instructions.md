---
applyTo: ".github/workflows/**/*.yml,**/*.csproj,global.json,Directory.Build.*,scripts/**/*.ps1"
excludeAgent: "code-review"
---

# Build and release constraints

- Keep workflow SDK setup compatible with `global.json`. Preserve analyzer and
  warning-as-error settings in `Directory.Build.props` and `.editorconfig`.
- `ci.yml` has Excel-free runtime and documentation gates. Local pre-commit
  selects checks by changed paths; preserve runtime/non-runtime and merge-parent
  handling so imported changes do not trigger unrelated Excel E2E.
- Local CLI builds invoke `scripts\Stop-ExcelMcpProcesses.ps1`. Preserve its
  pipe-scoped process ownership; builds in one worktree must not stop another's
  Excel sessions.
- `release.yml` owns versions and changelog generation. Do not dispatch it as a
  test. Authorized merges use squash.

Procedures: `docs/RELEASE-STRATEGY.md`, `vscode-extension/DEVELOPMENT.md`, and
`.github/workflows/docs/publish-plugins-setup.md#maintenance-and-updates`.
