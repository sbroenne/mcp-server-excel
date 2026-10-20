# Testing strategy

Follow the [repository rules](../AGENTS.md). These are implementation
instructions; review tasks use the root [Code Review Rules](../AGENTS.md#code-review-rules).

## Commands

Follow the [local validation scope rule](../AGENTS.md#build-and-validation):
use `build.ps1 plan` to identify candidate checks, not as an instruction to run
every selected feature. Use focused existing filters or runner groups for the
changed behavior and its affected callers; use `build.ps1 validate --plan <path>`
when the complete selection is relevant. Retain results and explain any narrower
selection in the PR. Use
`scripts\Invoke-ExcelTests.ps1` for explicit Excel group runs after a Release
build; `scripts\Test-E2E.ps1` runs complete CLI/MCP acceptance. These runners
set hard execution deadlines; returning control while a test keeps running
is not a timeout.
PR and commit checks share the PowerShell changed-area policy. The commit hook runs
the selected Release build and deterministic CLI/MCP/Core/Service tests
(`Commit` projection of `Fast`), excluding `AdapterTestKind=System` and
`Speed=Slow`. CI retains those system and slow tests; run local process checks
when changing their boundary, not on every push. Run relevant Excel
checks locally because hosted runners lack Excel. Test-only changes select their
owning project group; unit and recording-only protocol tests remain Excel-free.
Real-Excel MCP transport fixtures and their shared helpers select Excel coverage.
Shared changes include affected dependencies, not every unrelated Excel feature.
Empty or unknown selections must not be presented as passed tests.
Run ComInterop OnDemand diagnostics only when they cover a specific changed
ownership, threading, or cleanup behavior not already established by focused tests.
Core OnDemand tests are optional diagnostics, not mandatory CI gates.
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
- Tests never stop Excel or Power Query processes, including ones they started.
  Close test sessions through the product (`CloseSession`, `close`, dispose).
  If Excel is still running afterwards, the test or run fails and names the
  process; do not add code that stops it. Never close every listed session to
  establish a "clean" starting state. The compiler bans machine-wide process
  search and the product's process-stop functions in test code; do not bypass it.
  Fixture COM
  access follows [COM safety](../src/ExcelMcp.ComInterop/AGENTS.md).

Fixture, template, raw-COM/OOXML, and parallel-safety procedures:
[tests/README.md](README.md#saved-workbook-templates).

## Diagnosing stalled Excel tests

When an Excel test hangs, times out, or leaves Excel running, inspect Excel's
windows yourself with the available computer-use tools before guessing at the
cause or asking the user for screenshots. Read the window list and
accessibility text first; capture an image only when the text is not
available. Look for modal prompts such as Printer Setup or a security notice.
Do not click, close, or answer them, and do not stop Excel: report the dialog
text to the user. Keep dialog text out
of public artifacts.

PDF/XPS export tests first check that Windows can open the default printer and
fail with a plain message if it cannot, because Excel would otherwise show
Printer Setup during export. Fix that by repairing the printer or choosing a
working default printer; do not change the tests to skip export.

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
