# Build and release constraints

These rules also apply to workflows, project/SDK configuration, changesets, and
`tools/ExcelMcp.Build`; the root task map requires them outside this directory.

- MSBuild owns .NET dependency evaluation, restore and incremental compilation.
  Do not add a project XML interpreter or source-hash freshness cache.
- Typed build components own check selection, execution and results; scripts
  translate arguments. Unknown paths and failed/cancelled checks must fail
  explicitly, never silently broaden selection or satisfy a completion gate.
- Keep SDK setup aligned with `global.json` and preserve compiler safeguards.
  Compiler inputs, including EXL030 documentation, must reach selected builds.
- Commit checks cover selected builds and quick CLI/MCP/Core/Service tests only.
  Keep slower process/tooling/distribution coverage in CI; run local process
  checks when their boundary changes, not on every push. Desktop Excel checks
  still need local Excel when affected; hosted runners cannot replace them.
- Preserve merge-parent selection, per-owner test filters and runtime/package
  separation. Normal CI checks every package; only explicit manual runs narrow
  components. Keep complete main/manual CI and configured CodeQL coverage.
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
