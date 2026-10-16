# Testing strategy

Follow the [repository rules](../AGENTS.md). These are implementation
instructions; review tasks use the root [Code Review Rules](../AGENTS.md#code-review-rules).

## Commands

Use `build.ps1 plan` and `build.ps1 validate --plan <path>` for affected
behavior and dependencies, and retain the selected results. Use
`scripts\Invoke-ExcelTests.ps1` for explicit Excel group runs after a Release
build; `scripts\Test-E2E.ps1` runs complete CLI/MCP acceptance. These runners
set hard execution deadlines; returning control while a test keeps running
is not a timeout.
PR and commit checks share the typed changed-area policy. The commit hook runs
the selected Release build and quick CLI/MCP/Core/Service tests (`Commit`
projection of `Fast`) only. CI owns hosted slower checks; run local process
checks when changing their boundary, not on every push. Run selected Excel
checks locally because hosted runners lack Excel. A test-only change
selects its class and actual helper consumers; shared changes include affected
dependencies. Empty or unknown selections must not be presented as passed tests.
Session/batch infrastructure changes also require relevant ComInterop OnDemand
tests. Core OnDemand tests are optional diagnostics, not mandatory CI gates.
VBA needs Trust Center access; run screenshots separately because they use
desktop/clipboard resources.

Commands and prerequisites: [tests/README.md](README.md#quick-start).

## Fixtures and assertions

- Public workbook behavior requires real Excel and defaults to
  `ExcelMcpService.ProcessAsync`, not mocked `IExcelBatch` or direct Core.
  Direct Core needs a named internal contract Service cannot expose. Pure
  parsing, mapping, serialization, and generator tests need no Excel.
- Data Model/OLAP PivotTable behavior tests should reuse
  `PersistentServiceDataModelFixture` and create a test-owned PivotTable from
  its model. Verify the resulting Excel state through Service before cleanup;
  do not rely only on command success.
- Independent Facts/Theories must isolate their state, even with a reviewed
  class fixture.
  Lifecycle, transport, global-state, and raw-COM capability tests retain the
  actual boundary and isolation their subject requires.
- Cleanup always runs, preserves primary and cleanup failures, and never
  silently recreates a failed shared session.
- Follow nearby trait conventions: `Category`, `Feature`, `Layer`,
  `Speed`, and `RunType` where applicable. Compiler diagnostics own supported
  test classification and fixture-ownership checks; fix their source.
- Other global-state tests also need explicitly non-parallel collections;
  separate Excel testhosts must not overlap. `RequiresExcel=false`
  alone does not establish parallel safety or low cost.
- Positive tests assert successful prerequisites and verification responses,
  then exact values for every result field or workbook state the change adds
  or advertises. Tests must fail if those fields are unexpectedly empty,
  default, or stuck at one value. Cover each value of status flags (for example,
  complete and incomplete) and promised data shapes (for example, multi-level
  hierarchies, captions differing from names, and collections larger than a
  requested bound). Setup failure or an expected error is not a successful
  positive test.
- Verify inside the test after the operation and before cleanup. No exception,
  success alone, object existence, or a loose count is not an outcome check.
  Replacements prove old/new state; refresh changes the source first;
  persistence checks exact reopened contents.
- Negative tests check the intended error and preserved state, rollback, or
  unusable-session/recovery behavior promised by that operation. Ask whether
  the test would fail for a no-op, wrong target, or partial change before error.
- Do not assume failed operations automatically roll back or clean up workbook
  changes. Assert the actual partial state and error; leave cleanup decisions
  to the caller unless rollback is an explicit operation contract.
- Cleanup may terminate only exact owned identities using PID plus start time
  and retained handles. Never kill unrelated Excel processes. Fixture COM
  access follows [COM safety](../src/ExcelMcp.ComInterop/AGENTS.md).

Fixture, template, raw-COM/OOXML, and parallel-safety procedures:
[tests/README.md](README.md#saved-workbook-templates).

## Entry-point and migration coverage

CLI and MCP tests cover adapter responsibilities, not a second full workbook
matrix. Real-Excel CLI/MCP tests are limited to one save-and-reopen smoke per
entry point and process boundaries (background service, crash, deadline,
protected file, and Excel ownership). Test other adapter behavior without
Excel against a recorded Service request/response; workbook outcomes belong in
Service tests. Preserve adapter regressions as Excel-free tests; do not
shorten production waits for test speed.

Adapter test design:
[CLI and MCP coverage](README.md#cli-and-mcp-coverage).
Migration case mapping and final-source evidence:
[Complete normal-suite verification](README.md#complete-normal-suite-verification).

## Save and round-trip behavior

Do not call `batch.Save()` for in-memory assertions. When testing persistence,
save/close and reopen in a new batch before asserting. Use `.xlsm` for VBA.

Test design and failure investigation: [tests/README.md](README.md).
