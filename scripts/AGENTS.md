# Build and release constraints

These rules also apply to workflows, project/SDK configuration and changesets;
the root task map requires them outside this directory.

- MSBuild owns .NET dependency evaluation, restore and incremental compilation.
  Do not add a project XML interpreter or source-hash freshness cache.
- `Get-ValidationPlan.ps1` owns explicit changed-path rules. Standard .NET and
  xUnit commands own compilation, test discovery and execution; shared
  PowerShell coordinates groups, deadlines and required-case evidence.
  Do not recreate source-reference inference or a custom build executable.
  Existing PowerShell package
  commands own distribution preparation with standard SDK/npm/VSIX tools;
  Actions owns authorized publication. Unknown paths and failed/cancelled checks must fail
  explicitly, never silently broaden selection or satisfy a completion gate.
- Keep SDK setup aligned with `global.json` and preserve compiler safeguards.
  Compiler inputs, including EXL030 documentation, must reach selected builds.
- Commit checks cover selected builds and quick CLI/MCP/Core/Service tests only.
  Keep slower process/tooling/distribution coverage in CI; run local process
  checks when their boundary changes, not on every push. Desktop Excel checks
  still need local Excel when affected; hosted runners cannot replace them.
- Preserve merge-parent selection, per-owner test filters and runtime/package
  separation. PR CI checks affected package components and npm inputs. Keep
  complete main/manual/release checks and configured CodeQL coverage.
- Cache downloads, not writable build outputs. Isolate hosted test partitions
  so rebuilding tests cannot race packaging.
- Local CLI builds stop only this worktree's background CLI service, once,
  even on self-hosted desktops with `CI=true`. Validate PID/start-time identity
  and use the retained process handle. Never save or stop Excel, MCP,
  foreground CLI or another worktree's service as build cleanup.
- `release.yml` owns versions/changelogs. Do not dispatch publication as a test.
  Follow the publication guides before authorized publication changes.

Procedures and build ownership: [Development](../docs/DEVELOPMENT.md#build-and-check-ownership),
[test commands](../tests/README.md), [release](../docs/RELEASE-STRATEGY.md),
and [plugin publication](../.github/workflows/docs/publish-plugins-setup.md#maintenance-and-updates).
