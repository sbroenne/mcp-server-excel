# ExcelMcp Tests

Excel-dependent behavior uses real Excel integration tests. Parsing, mapping,
serialization, and generation can use focused tests without Excel. The former
blanket ban on unit tests is [superseded](../docs/ADR-001-NO-UNIT-TESTS.md).

## Quick Start

```powershell
# One ordinary workbook feature through Service
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "Feature=PowerQuery&RunType!=OnDemand"

# Excel-independent parsing
dotnet test tests\ExcelMcp.Core.Tests\ExcelMcp.Core.Tests.csproj --filter "FullyQualifiedName~ServiceRegistryJsonParsingTests"

# Session/batch changes: narrow by test name when appropriate
dotnet test tests\ExcelMcp.ComInterop.Tests\ExcelMcp.ComInterop.Tests.csproj --filter "RunType=OnDemand"

# VBA behavior (requires VBA trust enabled)
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "Feature=VBA&RunType!=OnDemand"
```

Set a hard execution timeout for every Excel-dependent run. Run only the
relevant project and filter, not the full integration suite during iteration.

### Saved workbook templates

Fixtures that only need an already-saved blank or populated workbook create one
immutable template per baseline and extension, then copy it to a unique path for
each test. The copy is ordinary file I/O and does not launch Excel. Opening,
mutating, saving, and reopening each copy still uses the normal owned Excel
process path.

Do not use a template when workbook creation, file format, lifecycle, save, or
reopen behavior is the subject of the test. Do not pool a live Excel
application, batch, or workbook across unrelated tests or classes. A reviewed
class-scoped Service fixture may share one session as described below.

### Persistent Service class fixtures

Ordinary workbook behavior tests may share one in-process `ExcelMcpService`
session per test class. Each behavior remains a separate Fact or Theory and
calls the public `ProcessAsync` boundary through the production command
contracts. Every test creates and cleans its own sheets, tables, maps, or other
objects, so test order does not matter. Fixture shutdown closes the session,
checks the Service session count, verifies exact owned process identities, and
uses the shared assembly exit gate. Cleanup still runs after a primary failure,
and reports both the primary and cleanup failures.

Fresh workbook/process, desktop, registry, or per-test Service isolation is
independent from the operation boundary and is not by itself a reason to call
Core directly. Direct Core coverage needs a specific internal contract that
Service cannot expose. Raw COM may prepare or verify state, but must not replace
the Service call under test.

Tests whose subject is process/session lifecycle, ownership, PID reuse,
crash/timeout/abort behavior, transport, application-wide state, locale text,
clipboard, protected files, concurrency, or raw COM capability keep that actual
boundary and the isolation it requires. Public VBA behavior may share one
macro-enabled Service workbook when each test owns and removes its modules.
Screenshots that change window, selection, protection, or rendering state use
one isolated visible Service session per test. Persistence uses explicit
save/close/reopen, reacquires the session token, and verifies reopened state.
A Service validation or serialization difference is a boundary mismatch, not
permission to weaken an assertion.

Positive tests must assert successful setup and verification responses before
checking concrete state. Never catch a setup failure and return green, treat an
expected error as a substitute for intended success, or infer that a whole
feature is unsupported from one rejected input. Keep invalid-input behavior in
separate negative tests and identify unavailable prerequisites explicitly.

### CLI and MCP coverage

CLI and MCP tests concentrate on argument/name/default mapping,
protocol/serialization, exit and error behavior, and lifecycle. They do not
repeat the full workbook-domain matrix already exercised through Service. MCP
protocol-only tests use the real Program transport with an explicit recording
Service backend and assert the exact dispatched command, session, arguments,
and serialized result or error. This proves the adapter contract, not workbook
behavior; the matching workbook inputs and state assertions stay in Service
tests. Independent safe adapter cases may reuse a real host/session. Retain
focused real adapter-to-Excel smokes plus intentional fresh-process, deadline,
crash, and ownership coverage. Do not remove adapter regressions or shorten
production waits merely to speed tests.

CLI parser contracts use the production Spectre command app in-process with an
explicit request client and captured input/output. They assert the exact
Service request and public output envelope without starting a daemon or Excel.
Batch validation may use a real in-process Service when validation belongs to
that public boundary. Keep a small executable-and-pipe smoke set plus the
distinct mutex, startup, deadline, crash, persistence, and forced-stop cases;
do not launch a process for every argument or alias row.

Use these Excel-free groups for quick adapter feedback:

```powershell
dotnet test tests\ExcelMcp.CLI.Tests\ExcelMcp.CLI.Tests.csproj -c Release --no-build --filter 'RunType!=OnDemand&RequiresExcel=false&AdapterTestKind!=System'
dotnet test tests\ExcelMcp.McpServer.Tests\ExcelMcp.McpServer.Tests.csproj -c Release --no-build --filter 'RunType!=OnDemand&RequiresExcel=false&AdapterTestKind!=System'
```

The quick groups are not acceptance gates. Complete normal validation still
uses `RunType!=OnDemand`, including the separately classified real Excel,
process, deadline, crash, rebuild, and ownership cases below.

### Parallel collections

Each project allows up to four xUnit collection workers, but only
`RequiresExcel=false` tests that are proven stateless may run in parallel.
Excel/COM, Data Model, CLI service and daemon, MCP transport, console,
current-directory, culture/environment, VBA, screenshot/desktop/clipboard,
repository mutation, timeout/stress, and intentional-concurrency tests remain
in collection definitions with `DisableParallelization=true`.

`RequiresExcel=false` does not imply cheap, unit-level, stateless, or safe to
parallelize. Process launches, waits, repository mutation, and global state
still require review. Collection settings serialize only one testhost;
orchestration must keep separate Excel testhosts from overlapping.

The classification architecture test reflects over the built test assembly. It
fails when a discovered test is missing `RequiresExcel=true/false`, has both
values, or an Excel test is outside an exclusive collection.

### Fair speed comparisons

Compare equivalent concrete behavior on the same machine and build settings.
Run the selected old cases together so their shared fixture is counted once,
and run the equivalent new checks together. Include fixture setup, test
execution, cleanup, outer wall time, and actual exact-owned Excel starts.
Begin with one run of each; repeat only to investigate variability or provide
requested confidence.

Do not use summed test durations as suite wall time, or a theoretical
one-start-per-case count as a measured old launch count. Make no numerical
speed promise before collecting comparable evidence. Keep temporary ledgers
and result artifacts outside committed instructions.

## Complete normal-suite verification

For a full validation pass, run all seven test projects with
`RunType!=OnDemand`. This filter includes normal Service VBA and screenshot
tests; their prerequisites must be satisfied, not silently skipped. Run
screenshots separately from other Excel tests when investigating them because
they use the interactive desktop. Research diagnostics and external-service
LLM tests remain separate from the normal behavior suite.

Use `--disable-build-servers` for preparatory builds and test runs so their
compiler/build hosts do not retain locks on the shared build-task assembly.
If an earlier ordinary build left a cached host behind, finish active builds
before using the repository's `dotnet build-server shutdown` preparation step.
Do not terminate unrelated processes to make a rebuild test pass.

Retain TRX output with `--logger "trx;LogFileName=<stage>.trx"` and use
`--list-tests` with the same filters to reconcile discovered and executed cases.
Set a per-test hang deadline (for example `--blame-hang-timeout 3m`) and a hard
deadline for each complete stage. A shell wait limit alone is not an execution
deadline. Preserve intentional concurrent-workbook tests when ordering stages.

Excel cleanup must use identities captured from the owning session or request,
including both process ID and start time. Never infer ownership from all Excel
processes started since a test began, or close existing Excel instances to
prepare a fixture. An independently opened workbook must survive successful and
failed teardown. Leak assertions must still fail if a confirmed owned process
survives shutdown.

Lifecycle probes announce readiness only after the owned Excel process has been
protected against testhost termination. Their cleanup reads every complete,
verified Excel identity in the child journals, including failed job assignments
and reopened sessions. Malformed records and read failures must be reported
without preventing cleanup of valid records elsewhere in those journals.

Core, ComInterop, and in-process MCP test hosts initialize this protection at
module startup. Set `EXCELMCP_TEST_OWNERSHIP_DIRECTORY` to retain their journals
with the run's other evidence; otherwise they are written beneath
`%TEMP%\ExcelMcpTestOwnership`. The shared xUnit framework reports surviving owned
Excel processes as assembly cleanup failures before announcing completed results;
`ProcessExit` remains a fallback. Both paths check before closing the job, and
disposal is idempotent. Exit checks use the original verified process handle,
retained until after job disposal, rather than reopening historical PIDs.
This prevents a later inaccessible process reusing a PID from being reported
as an Excel leak. This does not protect a separate
CLI daemon merely because its client test is running in a protected host.
Reused test hosts rearm protection at each execution start after the previous
execution disposed it, with a distinct journal for each execution.

The lifecycle probes require successful workbook Close, Quit, and STA completion,
plus Save where requested. Reopening also checks persisted data and a localized
table name. Each normal phase allows only the existing exact-owned post-Quit
fallback within the original 15-second process-cleanup budget, and confirms exit
through its retained handle before job disposal. Office can intermittently wait
on its own background services after our STA exits, even for simple creation.
This fallback is recorded explicitly; join-timeout, operation-timeout, quit
errors, disconnected proxies, and other forced paths are not accepted.
An intentionally undisposed child proves that cleanup failure makes the direct
test runner fail, even when its individual test result was successful.

Only change current-user VBA project access with explicit permission. Capture
the original registry value's presence, value, and type outside the repository
before changing it; restore and verify them in `finally`. Do not override managed
policy or change other macro security settings. Coordinate with other test
sessions before using the shared desktop or VBA setting.

### Ordered acceptance commands

The complete gate is an **ordered, reconciled set of runs**, not a claim that
an unpartitioned concurrent `dotnet test Sbroenne.ExcelMcp.sln` is green.
Use these filters after a successful Release solution build. Ordinary groups
that require Excel run sequentially, including across projects. Only proven
Excel-free, stateless selections may overlap. Rebuild tests must not compete
with another build.

```powershell
$testArgs = @('-c', 'Release', '--no-build', '--disable-build-servers',
    '--blame-hang-timeout', '10m', '--results-directory', 'TestResults\Normal')

# Ordinary groups. Run sequentially for the simplest reproducible ordering.
dotnet test tests\ExcelMcp.Core.Tests\ExcelMcp.Core.Tests.csproj @testArgs --filter 'RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust&Feature!=Screenshot' --logger 'trx;LogFileName=Core-main.trx'
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj @testArgs --filter 'RunType!=OnDemand&Feature!=VBA&Feature!=Screenshot' --logger 'trx;LogFileName=Service-main.trx'
dotnet test tests\ExcelMcp.CLI.Tests\ExcelMcp.CLI.Tests.csproj @testArgs --filter 'RunType!=OnDemand&FullyQualifiedName!~VbaRun_OnMacroWorkbook' --logger 'trx;LogFileName=CLI-main.trx'
dotnet test tests\ExcelMcp.McpServer.Tests\ExcelMcp.McpServer.Tests.csproj @testArgs --filter 'RunType!=OnDemand&FullyQualifiedName!~VbaRun_OnMacroWorkbook' --logger 'trx;LogFileName=MCP-main.trx'
dotnet test tests\ExcelMcp.ComInterop.Tests\ExcelMcp.ComInterop.Tests.csproj @testArgs --filter 'RunType!=OnDemand' --logger 'trx;LogFileName=ComInterop-normal.trx'
dotnet test tests\ExcelMcp.SkillGeneration.Tests\ExcelMcp.SkillGeneration.Tests.csproj @testArgs --filter 'RunType!=OnDemand' --logger 'trx;LogFileName=Skills-normal.trx'
dotnet test tests\ExcelMcp.Diagnostics.Tests\ExcelMcp.Diagnostics.Tests.csproj @testArgs --filter 'RunType!=OnDemand' --logger 'trx;LogFileName=Diagnostics-normal.trx'
dotnet test tests\ExcelMcp.ComInterop.Tests\ExcelMcp.ComInterop.Tests.csproj @testArgs --filter 'RunType=OnDemand&FullyQualifiedName!~BeginBatch_RealIrmWorkbook&Locale!=ja-JP' --logger 'trx;LogFileName=ComInterop-infrastructure.trx'

# Exclusive VBA groups: obtain permission, capture the original setting, enable
# access only for the run, then restore/verify it as described above.
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj @testArgs --filter 'RunType!=OnDemand&Feature=VBA' --logger 'trx;LogFileName=Service-vba.trx'
dotnet test tests\ExcelMcp.CLI.Tests\ExcelMcp.CLI.Tests.csproj @testArgs --filter 'RunType!=OnDemand&FullyQualifiedName~VbaRun_OnMacroWorkbook' --logger 'trx;LogFileName=CLI-vba.trx'
dotnet test tests\ExcelMcp.McpServer.Tests\ExcelMcp.McpServer.Tests.csproj @testArgs --filter 'RunType!=OnDemand&FullyQualifiedName~VbaRun_OnMacroWorkbook' --logger 'trx;LogFileName=MCP-vba.trx'

# Exclusive desktop group: no other Excel test hosts.
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj @testArgs --filter 'RunType!=OnDemand&Feature=Screenshot' --logger 'trx;LogFileName=Service-screenshot.trx'
```

The Service normal group includes the four-workbook isolation workflow in its
exclusive Excel collection. CLI acceptance keeps representative executable,
pipe, persistence, and shutdown wiring without repeating that workbook matrix.
Use an outer hard deadline in addition to the per-test hang limit: up to eight
hours for ordinary Core, two hours for ordinary CLI/MCP, 90 minutes for
ComInterop, 30 minutes for skills/screenshots, 45 minutes for Core VBA, and
10 minutes for each transport-VBA run. Stop on nonzero
exit codes or zero matching cases.

Refresh `--list-tests` for every final binary and reconcile case counts,
including repeated display names. Verify the partitions do not overlap.
Use serializable theory arguments so discovery enumerates every row rather than
reporting one deferred method that expands into several cases during execution.
For migrations, map each method and theory row to equivalent inputs, concrete
assertions, and an individual result, or record a narrow retained/removal
reason. A broad “specialized” exception or count-only comparison is not enough.
The configured real-IRM probe and Japanese-format probe are separate,
prerequisite-dependent on-demand runs. The latter requires a ja-JP Windows and
Excel installation; do not change host locale or policies to manufacture a pass.

Evidence must describe the exact final source and classify passed, failed,
skipped, not-applicable, and not-run results separately. Preserve the original
failure even when a later run is clean. Do not claim broad completion or a
measured whole-suite speedup from a subgroup, operation totals, or theoretical
process counts.

## Documentation

**For complete testing guidance, see:**

- **[Testing Strategy](../.github/instructions/testing-strategy.instructions.md)** - Quick reference, templates, common mistakes
- **[Repository Rules](../.github/copilot-instructions.md)** - Build, E2E, and contribution requirements

## Test Architecture

```
tests/
├── ExcelMcp.Core.Tests/           # Internal contracts and pure parsing tests
├── ExcelMcp.Service.Tests/        # Public workbook behavior through Service
├── ExcelMcp.Diagnostics.Tests/    # Excel COM behavior research (OnDemand, Manual)
├── ExcelMcp.McpServer.Tests/      # MCP protocol layer (Integration)
├── ExcelMcp.CLI.Tests/            # CLI wrapper (Integration)
├── ExcelMcp.ComInterop.Tests/     # COM utilities and session infrastructure
└── ExcelMcp.SkillGeneration.Tests/ # Generated skill and plugin checks

llm-tests/                          # LLM tool behavior validation (Manual)
```

## Test Categories

| Category | Speed | Requirements | Run By Default |
|----------|-------|--------------|----------------|
| **Unit** | Fast | No Excel for pure logic | Select the relevant tests |
| **Integration** | Medium (10-20 min) | Excel + Windows | ✅ Yes (local) |
| **OnDemand** | Slow (3-5 min) | Excel + Windows | ❌ No (explicit only) |
| **Diagnostics** | Slow (varies) | Excel + Windows | ❌ No (manual, excluded from CI) |
| **LLM Tests** | Slow (varies) | Excel + Azure OpenAI | ❌ No (manual only) |

## Diagnostics Tests

Diagnostics tests are research/exploratory tests in `ExcelMcp.Diagnostics.Tests` that document the actual behavior of Excel's COM APIs without our abstraction layer. These tests are **excluded from CI** to keep automation focused on core functionality.

**Purpose:**
- Understand Excel COM API behavior for Power Query, Data Model, PivotTables, etc.
- Document findings and edge cases for future implementation decisions
- Test alternative approaches to complex Excel operations

**Trait markers:**
- `Layer=Diagnostics`  
- `RunType=OnDemand`

**Run diagnostics tests locally:**
```powershell
# All diagnostics tests
dotnet test tests/ExcelMcp.Diagnostics.Tests/ --filter "RunType=OnDemand&Layer=Diagnostics"

# Specific diagnostic tests
dotnet test tests/ExcelMcp.Diagnostics.Tests/ --filter "Feature=PowerQuery&RunType=OnDemand"
```

**CI Behavior:**
- Diagnostics tests are **NOT** run in CI workflows (GitHub Actions)
- Path filter includes folder to trigger builds when tests change
- Test execution uses `RunType!=OnDemand` filter to exclude them

## Feature-Specific Tests

```powershell
# Test specific feature only
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "Feature=PowerQuery&RunType!=OnDemand"
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "Feature=DataModel&RunType!=OnDemand"
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "Feature=Tables&RunType!=OnDemand"
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "Feature=PivotTables&RunType!=OnDemand"
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "Feature=Ranges&RunType!=OnDemand"
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "Feature=Connections&RunType!=OnDemand"
```

## When to Run Which Tests

| Scenario | Command |
|----------|---------|
| **Daily development** | Run the smallest project and feature/name filter covering the change. |
| **Before commit** | Rerun affected tests and applicable checks; follow the [runtime E2E requirements](../.github/copilot-instructions.md#build-and-validation). |
| **Modified session/batch code** | Run relevant OnDemand tests in `ExcelMcp.ComInterop.Tests`; see [Testing Strategy](../.github/instructions/testing-strategy.instructions.md#commands). |
| **VBA development** | `dotnet test --filter "(Feature=VBA\|Feature=VBATrust)&RunType!=OnDemand"` |
| **LLM behavior validation** | See [LLM Tests](#llm-tests) section below |

## LLM Tests

The `llm-tests/` project validates that LLMs correctly use Excel MCP Server and CLI tools using [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering).

### When to Run LLM Tests

- **Manual/on-demand only** - Not part of CI/CD
- After changing tool descriptions or adding new tools
- To validate LLM behavior patterns (e.g., incremental updates vs rebuild)

### Running LLM Tests

```powershell
# From llm-tests/
uv sync
uv run pytest -m aitest -v
```

### Prerequisites

- `AZURE_OPENAI_ENDPOINT` environment variable
- Windows desktop with Excel installed
- GitHub auth via `gh auth login` or `GITHUB_TOKEN`

**See [LLM Tests README](../llm-tests/README.md) for complete documentation.**

## VBA Testing

### Normal-suite coverage

`RunType!=OnDemand` includes normal VBA tests. VBA project operations require
Excel's "Trust access to the VBA project object model" setting. A targeted
non-VBA development run is not a complete normal-suite validation pass.

### When to Run VBA Tests

Run the focused VBA group when:
- Modifying VBA-related code (ScriptCommands, VbaTrustDetection)
- Adding new VBA features
- Before releasing VBA-related changes
- Troubleshooting VBA-specific issues

Include these tests in every explicitly requested complete normal-suite run.

### How to Run VBA Tests

```powershell
# Run ONLY VBA tests
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand"
```

For a complete normal-suite run, use the serialized project commands in
[Ordered acceptance commands](#ordered-acceptance-commands).

### VBA Test Files

VBA tests use the `VBA` or `VBATrust` feature trait:

```
tests/ExcelMcp.Service.Tests/
  - PersistentServiceVbaCommandsTests.cs
  - PersistentServiceVbaFailureTests.cs
  - PersistentServiceVbaFixture.cs

tests/ExcelMcp.Core.Tests/Integration/Commands/Vba/
  - VbaCommandsTimeoutTests.cs

tests/ExcelMcp.CLI.Tests/Integration/
  - VbaRunValidationCliTests.cs
  - VbaRunCliTransportProofTests.cs
```

### VBA Trust Setup

VBA project operations require VBA trust enabled in Excel. Inspect the setting:

```powershell
Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "AccessVBOM"
```

An absent value is not enabled. Do not change this setting without permission.
For an authorized temporary test run, capture its original presence, value, and
type first, restore them in `finally`, and verify restoration. Do not change
managed policies or other macro security settings.

## Key Principles

### Designing a regression test

Reproduce the reported failure before changing the implementation. Cover
meaningful boundary and error cases, plus both entry points when the contract
crosses CLI and MCP. There is no fixed test quota.

Assert resulting workbook state and relevant returned fields rather than only
`Success`. For update/replace behavior, assert both that old content is absent
and that new content is exact. Error assertions should distinguish the intended
failure from other exceptions instead of accepting incompatible outcomes.

Use a unique workbook with the established feature fixture. Combining
`IClassFixture<T>` and a collection fixture on the same class can create competing
Excel sessions. Follow neighboring trait conventions and the COM cleanup rules
for any references acquired by the test itself.

Reviewed ordinary workbook behavior uses a class-scoped Service fixture: one
live Service/workbook session per class, separate Facts/Theories, unique object
names, and fail-closed per-test cleanup. Tests still use their actual boundary when they cover fresh files or processes,
transport adapters, application-wide state, raw COM details not returned
publicly, or lifecycle behavior. Persistence behavior may use the Service
fixture only when it performs an explicit save/close/reopen and verifies the
reopened workbook.

For in-memory changes, inspect the same batch without saving. For persistence,
save and close, reopen in a new batch, then assert the state; do not open the
same workbook in two live batches. Use `.xlsm` for VBA persistence.

### Diagnosing a failing test

Run the failure alone before broadening the run. Check workbook isolation,
fixture selection, actual Excel state, cleanup, and whether the assertion needs
a save/reopen cycle. Inspect fallback/retry paths when a primary-path fix is
insufficient. Do not hide a deterministic failure with skip/xfail or loosen an
assertion merely to make it pass.

- ✅ **State Isolation** - Each test owns unique objects, or a fresh workbook
  when its behavior requires one
- ✅ **Binary Assertions** - Pass OR fail, never "accept both"
- ✅ **Verify Excel State** - Always verify actual Excel state after operations
- **Explicit persistence** - Call `batch.Save()` only when testing save/close/reopen behavior (see [Testing Strategy](../.github/instructions/testing-strategy.instructions.md#save-and-round-trip-behavior)).

## Getting Help

- **Test failures**: Check test output for detailed error messages
- **Excel issues**: Ensure Excel 2016+ installed and activated
- **Session/batch issues**: Run OnDemand tests to verify cleanup
- **Writing tests**: See [Testing Strategy](../.github/instructions/testing-strategy.instructions.md)
