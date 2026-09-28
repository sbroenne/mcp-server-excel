---
applyTo: "tests/**/*.cs"
excludeAgent: "code-review"
---

# Testing strategy

## Commands

Select one project and feature/name filter, not the full Excel suite. Set a hard
execution timeout; returning control while a test keeps running is not a timeout.

```powershell
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "Feature=PowerQuery&RunType!=OnDemand"
dotnet test tests\ExcelMcp.ComInterop.Tests\ExcelMcp.ComInterop.Tests.csproj --filter "RunType=OnDemand"
```

The second command is required for session/batch infrastructure changes; narrow
by test name where appropriate. Core OnDemand tests are optional diagnostics,
not mandatory CI gates. VBA needs Trust Center access; run screenshots
separately because they use desktop/clipboard resources.

## Fixtures and assertions

- Public workbook behavior requires real Excel and defaults to
  `ExcelMcpService.ProcessAsync`, not mocked `IExcelBatch` or direct Core.
  Independent Facts/Theories may share a reviewed class fixture when their
  per-test state is isolated. Fresh workbook, process, desktop, or registry
  isolation does not by itself change the test boundary. Use direct Core only
  for a named internal contract that the public boundary cannot expose.
- Raw COM is acceptable for required setup or concrete verification, but the
  operation under test must still cross its intended Service boundary. Pure
  parsing, mapping, serialization, and generator tests need no Excel.
- Tests may construct or inspect ZIP/OOXML workbook parts for fixtures and
  concrete verification. Never move that package access into production code or
  use it as a substitute for exercising Excel COM behavior.
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
  collection. Only proven Excel-free, stateless collections may use the four
  configured workers. xUnit collection serialization applies only inside one
  testhost; orchestration must prevent Excel lanes in different testhosts from
  overlapping.
- `RequiresExcel=false` means only that Excel is unnecessary. Processes, waits,
  repository mutation, and console/current-directory/culture/environment state
  still require explicit cost and parallel-safety review. Keep pure logic tests
  fast and preserve separate real lifecycle checks.
- Positive tests assert every prerequisite and verification response succeeded,
  then assert concrete workbook state and returned fields. Never return green
  after setup fails or accept an expected error as a substitute for the
  intended success path. Put invalid input in a separate negative test and
  report unavailable prerequisites explicitly. One rejected input or formula
  does not prove that the feature is unsupported.
- Cleanup may terminate only exact owned identities using PID plus start time
  and retained handles. Never kill unrelated Excel processes. Fixture COM
  access follows `excel-com-interop.instructions.md`.

## Entry-point and migration coverage

- CLI and MCP tests cover their entry-point responsibilities: argument and name
  mapping, defaults, protocol/serialization, exit/error behavior, and lifecycle.
  Do not repeat the full workbook-domain matrix at each adapter. Protocol-only
  MCP tests may use the real Program transport with an explicit recording
  Service backend to assert the exact command, session, arguments, and
  serialized result or error. This does not prove workbook behavior; map that
  behavior to concrete Service coverage. Reuse a real host/session for safe
  independent adapter cases, while retaining small real adapter-to-Excel smokes
  and intentional fresh-process, deadline, crash, and ownership coverage. Do
  not remove adapter regressions or shorten production waits for test speed.
- Map every migrated method and every theory row to equivalent concrete inputs,
  assertions, and an individual discovered result. Record an explicit
  replacement or a narrow retained/removal reason; “specialized” and aggregate
  counts are not sufficient completion evidence.
- Evidence must match the final source. Report passed, failed, skipped,
  not-applicable, and not-run separately. A clean rerun does not erase the
  original failure. Do not infer broad suite completion or measured whole-suite
  speedup from a subgroup, operation count, or theoretical launch count.

## Save and round-trip behavior

Do not call `batch.Save()` for in-memory assertions. When testing persistence,
save/close and reopen in a new batch before asserting. Use `.xlsm` for VBA.

Test design and failure investigation: `tests/README.md`.
