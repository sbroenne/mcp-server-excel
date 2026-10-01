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
