---
applyTo: "tests/**/*.cs"
excludeAgent: "code-review"
---

# Testing strategy

## Commands

Select one project and feature/name filter, not the full Excel suite. Set a hard
execution timeout; returning control while a test keeps running is not a timeout.
Session/batch infrastructure changes also require relevant ComInterop OnDemand
tests. Core OnDemand tests are optional diagnostics, not mandatory CI gates.
VBA needs Trust Center access; run screenshots separately because they use
desktop/clipboard resources.

Commands and prerequisites: [tests/README.md](../../tests/README.md#quick-start).

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
- Treat workbook files as opaque in tests and fixtures. Do not construct,
  inspect, parse, or mutate ZIP/OOXML workbook parts. Copy intact
  Excel-authored templates when a saved fixture is required, and verify
  behavior only through supported Excel APIs.
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
- Cleanup may terminate only exact owned identities using PID plus start time
  and retained handles. Never kill unrelated Excel processes. Fixture COM
  access follows `excel-com-interop.instructions.md`.

Fixture, template, raw-COM/OOXML, and parallel-safety procedures:
[tests/README.md](../../tests/README.md#saved-workbook-templates).

## Entry-point and migration coverage

CLI and MCP tests cover adapter responsibilities, not a second full workbook
matrix. Preserve adapter regressions and real Excel smokes; do not shorten
production waits for test speed.

Adapter test design:
[CLI and MCP coverage](../../tests/README.md#cli-and-mcp-coverage).
Migration case mapping and final-source evidence:
[Complete normal-suite verification](../../tests/README.md#complete-normal-suite-verification).

## Save and round-trip behavior

Do not call `batch.Save()` for in-memory assertions. When testing persistence,
save/close and reopen in a new batch before asserting. Use `.xlsm` for VBA.

Test design and failure investigation: [tests/README.md](../../tests/README.md).
