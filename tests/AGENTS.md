# Testing strategy

Follow the [repository rules](../AGENTS.md). These are implementation
instructions; review tasks use the root [Code Review Rules](../AGENTS.md#code-review-rules).

## Commands

During development, use `dotnet test` with a filter for the affected project,
class, or feature. Retain TRX results and use `--blame-hang-timeout` for Excel
tests. Run each Excel-dependent command sequentially. Use
`scripts\Invoke-ExcelTests.ps1` only for explicit group or complete-suite runs;
its shared runner sets hard execution deadlines. Returning control while a
test keeps running is not a timeout. Final runtime acceptance uses
`scripts\Test-E2E.ps1` as required by the root rules.
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
- Independent Facts/Theories must isolate their state, even with a reviewed
  class fixture. Do not combine class and collection fixtures on one class.
  Lifecycle, transport, global-state, and raw-COM capability tests retain the
  actual boundary and isolation their subject requires.
- Cleanup always runs, preserves primary and cleanup failures, and never
  silently recreates a failed shared session.
- Use a unique workbook per isolated test or per reviewed persistent Service
  class. Do not combine `IClassFixture<T>` with a collection fixture on the same
  class: it can create competing Excel sessions.
- For a saved empty or populated baseline, create one immutable workbook
  template and copy it to a unique destination for each test. Copying must not
  start Excel. Keep workbook-creation, format, lifecycle, and persistence tests
  on their explicit creation/open/save/reopen paths.
- Keep independent scenarios as separate Facts/Theories and isolate their
  sheets, tables, maps, and other objects. Fixture cleanup must always run and
  preserve both the primary failure and every cleanup failure. A failed shared
  session fails explicitly; do not silently recreate it.
- Lifecycle/ownership/crash/timeout, transport, global-state, locale-rendering,
  clipboard, protected-file, concurrency, and raw-COM capability tests retain
  their actual boundary when that boundary is the test subject. Public VBA
  behavior may use a persistent macro-enabled Service class when modules are
  isolated and removed. Public screenshot behavior stays at Service, but cases
  that mutate window, selection, protection, or rendering state use an isolated
  visible Service session per test. Persistence requires explicit
  save/close/reopen and assertions against the reopened session.
- Follow nearby trait conventions: `Category`, `Feature`, `Layer`,
  `RequiresExcel`, `Speed`, and `RunType` where applicable. Every discovered
  test must resolve to exactly one `RequiresExcel=true` or `false` value.
- Excel and other global-state tests belong to an explicitly non-parallel
  collection; separate Excel testhosts must not overlap. `RequiresExcel=false`
  alone does not establish parallel safety or low cost.
- Positive tests assert successful prerequisites and verification responses,
  then concrete workbook state and returned fields. Setup failure or an
  expected error is not a successful positive test.
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
  access follows [COM safety](../docs/agents/rules/excel-com-interop.md).

Fixture, template, raw-COM/OOXML, and parallel-safety procedures:
[tests/README.md](README.md#saved-workbook-templates).

## Entry-point and migration coverage

CLI and MCP tests cover adapter responsibilities, not a second full workbook
matrix. Preserve adapter regressions and real Excel smokes; do not shorten
production waits for test speed.

Adapter test design:
[CLI and MCP coverage](README.md#cli-and-mcp-coverage).
Migration case mapping and final-source evidence:
[Complete normal-suite verification](README.md#complete-normal-suite-verification).

## Save and round-trip behavior

Do not call `batch.Save()` for in-memory assertions. When testing persistence,
save/close and reopen in a new batch before asserting. Use `.xlsm` for VBA.

Test design and failure investigation: [tests/README.md](README.md).
