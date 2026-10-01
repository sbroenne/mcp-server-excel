# ExcelMcp Tests

Excel-dependent behavior uses real Excel integration tests. Parsing, mapping,
serialization, and generation can use focused tests without Excel.
[ADR-001](../docs/ADR-001-TESTING-STRATEGY.md) explains this split.

**Platform scope:** the COM fixtures, owned-process checks, VBA/registry setup,
and ordered normal-suite commands below require Windows desktop Excel. Apple
Silicon macOS is experimental beta; its shared-Excel workbook ownership must
be tested separately. A portable contract test or cross-target build is not
proof of Excel behavior on either platform.

After a successful Release build, use these Mac lanes sequentially:

```powershell
pwsh -NoProfile -File scripts/Invoke-ExcelFreeTests.ps1 -Local -Contracts
pwsh -NoProfile -File scripts/Test-MacE2E.ps1 -SkipBuild
```

The E2E lane requires interactive Excel for Mac and accepted Automation
permission. Office.js candidates additionally require user-controlled
certificate trust and exact-workbook activation; they remain unavailable until
accepted. See [Mac support and limitations](../specs/MACOS-SUPPORT.md).
Never overlap Excel testhosts, fixtures, or CLI/MCP E2E runs.

## Quick Start

```powershell
# One ordinary workbook feature through Service
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj -c Release --filter 'RequiresExcel=true&Feature=PowerQuery&RunType!=OnDemand' --blame-hang-timeout 5m --logger trx

# Excel-independent parsing
dotnet test tests\ExcelMcp.Core.Tests\ExcelMcp.Core.Tests.csproj --filter "FullyQualifiedName~ServiceRegistryJsonParsingTests"

# Session/batch changes: narrow by test name when appropriate
dotnet test tests\ExcelMcp.ComInterop.Tests\ExcelMcp.ComInterop.Tests.csproj --filter "RunType=OnDemand"

# VBA behavior (requires VBA trust enabled)
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "Feature=VBA&RunType!=OnDemand"
```

### Excel integration tests and saved results

Excel integration tests are ordinary C# tests tagged `RequiresExcel=true`.
They exercise ExcelMcp against desktop Excel and verify workbook or session
outcomes. They are not a separate investigation suite.

During development, run `dotnet test` for the affected project, class, or feature,
as above. Use `--logger trx` to retain results, `--blame-hang-timeout 5m` for hang
protection, and `--results-directory` when a named results directory is useful.
Check that the filter actually executed the intended tests: `dotnet test` can
return zero when no tests match. Never overlap Excel-dependent commands.

For an explicit group or complete-suite run, build Release first and use
`scripts\Invoke-ExcelTests.ps1`. It keeps class fixtures together, runs groups
and projects sequentially, and uses the same `Invoke-TestStage.ps1` helper as
Excel-free checks and E2E. The helper retains stdout/stderr logs, TRX reports,
ownership journals, reports wall times to the console, enforces a hard deadline,
and rejects empty, skipped, failed, or contradictory reports. Group and E2E
runs also compare discovered test names with executed results, including repeated
theory rows; missing, extra, or duplicated cases fail the run.
`-ListTests` lists the selected tests
without running workbook operations; it is not passing test evidence.

Complete-suite runs are not routine development steps. For runtime changes,
run `scripts\Test-E2E.ps1` once on final PR source. Investigation diagnostics
marked `RunType=OnDemand` stay separate; the group runner includes infrastructure
diagnostics only with explicit `-IncludeInfrastructureDiagnostics`.

Windows/Azure runner setup and administration scripts are not part of the
automated test suite. Product checks remain, including COM-reference safety,
owned pre-build cleanup, test-result reporting, and real Excel acceptance.
The retained PowerShell script tests run with PowerShell 7.

Missing Excel, VBA trust, desktop, or other prerequisites mean incomplete
validation; do not manufacture a pass by skipping tests or changing host
settings. Explicit on-demand locale/IRM probes need their own focused run.

### Verify the outcome before cleanup

Establish the relevant initial state, check the operation's response, then
compare the actual Excel outcome with independently determined expected state.
Check untouched cells/objects when preservation is promised. Use public reads
where sufficient and controlled raw COM for state not exposed by them.
Verification belongs inside each test before its objects are cleaned up; a
generic after-test callback cannot determine what that test intended.

Examples of meaningful checks:

| Behavior | Required evidence |
| --- | --- |
| Write or replace | Exact destination values; a replacement differs from the original |
| Formula | Formula text and calculated meaning when both are part of the contract |
| Refresh | Baseline data, changed source while destination is still old, then exact new destination |
| Filter and clear | Exact visible records after filtering, complete restored records after clearing |
| Settings | Requested setting on the target and preserved untargeted state |
| Save/reopen | One changed marker survives reopening; a filename alone does not prove saving |
| Rejected input | Intended error and unchanged state where validation promises no mutation |
| Real timeout | Reached the blocked operation, rejected poisoned session, safe recovery |

Select negatives by risk: missing objects, bad dimensions/addresses, bounds,
name conflicts, merged/protected destinations, invalid options, cancellation,
partial failures, and documented unsupported operations. Not every operation
needs every category. Do not promise rollback where it is not supported.
Use known pre-failure state and inspect it after the failure.

Check rejected destinations before clearing a guard or retrying. Otherwise a
partial write can be hidden by the test's own cleanup. For calculation tests,
establish a calculated baseline, change its input in manual mode, and prove
the old result is still present before requesting recalculation. For native
dimensions that Excel rounds, compare the effective size before and after
the operation instead of assuming the requested size is stored exactly.
Expected records and dependency edges must come from the seeded scenario,
not agreement between a returned count and the operation's own output.

Test behavior owned by our code, not Excel's general reliability. Save checks
need one persisted marker and the requested save/discard behavior, not repeated
checks of every cell, formula, and format after reopening. Check those features
in their own operation tests.

Keep one test for each distinct owned scenario. When tests repeat the same
setup, operation and expected outcome, merge any unique assertions into the
retained test instead of running the scenario again under another name.
Preserve distinct inputs, regressions, native controls, entry-point behavior
and cleanup responsibilities. Similar-looking bodies or shared helpers alone
do not establish duplication; inspect their data and actual call targets.
Check theory rows and acceptance selections as well as test definitions.

For Power Query, check complete records in every requested destination, not
just agreement between `list`, `view`, and `get-load-config`. Loading to both
a worksheet and the Data Model creates separate connections; identify them
by exact mashup `Location` and check the model table's source connection.
Give model seed columns explicit M types before testing numeric DAX measures.
A stored measure is not proof that it can calculate. Check its result before
and after source or schema changes; removing a referenced column must establish
the resulting calculation failure and recovery, rather than allow any outcome.
For worksheet schema changes, verify the complete native table shape, removed
cells and any calculated columns or neighbors, not only the expected rectangle
within the result. Reading two requested columns can miss an unwanted third
column. Native refresh controls established that `PreserveColumnInfo=false`
leaves synthetic extra headers when source columns disappear; the corrected
setting must also work on existing tables and retain calculated columns.
For DAX-backed worksheet Tables, compare the complete independently expected
records, native start cell and dimensions, exact connection command and retained
model source. A rejected query update can retain the new command and old rows;
check that concrete partial state and successful recovery instead of promising
rollback. Unsupported model renames must preserve query sources, connection
settings, complete model records, relationships and actual measure calculations.
XML rejection tests should retain an existing mapped value, its native XPath,
exported structure and neighboring cells, not only an empty map list.

Setting updates must start with a different effective setting, verified before
the update; setting Percentage to Percentage cannot catch a missing update.
Assertions about a required field must fail when that field is absent, not sit
inside an optional branch. Cancellation tests must establish where execution
reached before cancellation; a timer alone cannot prove the intended boundary.
Seed guards on isolated test sheets, and normalize existing line endings before
comparing complete VBA source so the test does not introduce double carriage
returns. The VBA editor can change identifier casing using project symbols;
allow that native change without ignoring changes to string literals or comments.
Restoration checks must run after disposal, outside another automatic guard that
would hide the restored state.

During review, ask: **would this test fail if the operation did nothing, changed
the wrong target, or made a partial change before reporting an error?**
No exception, `Success`, existence, a loose count, or an unchecked read does
not answer that question. Do not require screenshots for ordinary verification.

Set a hard execution timeout for every Excel-dependent run. Run only the
relevant project and filter, not the full integration suite during iteration.
Returning control while tests keep running is not a timeout. Session/batch
infrastructure changes require relevant ComInterop OnDemand tests; narrow by
test name where appropriate. Core OnDemand tests are optional diagnostics,
not mandatory CI gates.

### Saved workbook templates

Treat all workbook files as opaque. Templates must be authored by Excel and
copied intact; never construct, inspect, or mutate ZIP/OOXML workbook contents,
including through libraries or fixture scripts. Verify content through
supported Excel APIs only.

Fixtures that only need an already-saved blank or populated workbook create one
immutable template per baseline and extension, then copy it to a unique path for
each test. The copy is ordinary file I/O and does not launch Excel. Opening,
mutating, saving, and reopening each copy still uses the normal owned Excel
process path.

Do not use a template when workbook creation, file format, lifecycle, save, or
reopen behavior is the subject of the test. Do not pool a live Excel
application, batch, or workbook across unrelated tests or classes. A reviewed
class-scoped Service fixture may share one session as described below.

Tests may construct or inspect ZIP/OOXML workbook parts for fixtures and
concrete verification. This is a test-only exception: never move that package
access into production code or use it instead of exercising Excel COM behavior.

### Persistent Service class fixtures

Ordinary workbook behavior tests may share one in-process `ExcelMcpService`
session per test class. Each behavior remains a separate Fact or Theory and
calls the public `ProcessAsync` boundary through the production command
contracts. Every test creates and cleans its own sheets, tables, maps, or other
objects, so test order does not matter. Fixture shutdown closes the session,
checks the Service session count, verifies exact owned process identities, and
uses the shared assembly exit gate. Cleanup still runs after a primary failure,
and reports both the primary and cleanup failures. A failed shared session
fails explicitly; do not silently recreate it.

Remove dependent measures before deleting their query-backed model tables.
Deleting the table first can remove the measure implicitly and make later
cleanup fail, hiding the original test outcome.

Fresh workbook/process, desktop, registry, or per-test Service isolation is
independent from the operation boundary and is not by itself a reason to call
Core directly. Direct Core coverage needs a specific internal contract that
Service cannot expose. Raw COM may prepare or verify state, but must not replace
the Service call under test.

Locale-sensitive native comparisons must use the equivalent Excel API.
For example, `WorksheetFunction.Text` and a worksheet `TEXT` formula can
interpret the same quoted format differently. Verify formula calculations
against native worksheet formulas using independently calculated inputs, and
check the complete stored formula separately. Do not change host locale to
make a comparison pass.

Avoid starting Excel repeatedly to verify different properties of the same
saved workbook at one checkpoint. Reopen it once and check sheet order, marker
values, calculated results and formulas within that owned batch, then dispose
it. Combine unrelated seed writes within one setup batch as well. Keep separate
reopens when an intervening operation or a save/reopen lifecycle is the subject;
do not cache snapshots across operations or pool Excel between independent tests.

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

### Checking exact feature outcomes

For new workbook operations, assert the exact changed state and the unchanged
surrounding state, not only counts or matching writer/getter responses. Filter
tests should identify the visible records; layout tests should check expected
coordinates, identities, sizes, and unselected objects. Partial formatting
updates should check every supplied field and every omitted setting using
distinct initial values. Use raw COM or saved-file inspection when a shared
writer/getter mistake could otherwise pass.

Keep success, invalid-input, cancellation, output/cleanup failure, and
save/reopen cases separate where each applies. Adapter tests check exact
requests, defaults, and failure output; real Excel tests establish workbook
behavior. A cancellation timer alone does not prove which execution phase was
interrupted. Establish the claimed boundary first; for example, cancel after
Excel requests a controlled local M data source but before it receives the
response. Do not reuse a forcibly aborted COM context to prove recovery: test
the actual session's rejection and disposal contract instead. For high-risk
changes, temporarily introduce a specific wrong
mapping, selection, omitted-field reset, or missing cleanup and confirm the
intended test fails. Restore the change and rerun the final source before
delivery; retain these fault-check results with the run evidence.

Before PR delivery, include affected existing callers as well as new feature
tests. Run the full existing Excel-free selection with
`scripts\Invoke-ExcelFreeTests.ps1` (without `-Local`), including packaged-plugin
validation, and `npx --no-install changeset status --since=origin/main`.
These complement focused native tests and final-source runtime E2E; they do
not replace either.
Preserve documented response shapes. For example, MCP `file` action `test`
returns a file assessment: `success=false` can mean a missing or protected
file, not a failed tool request. Check its complete diagnostic fields and the
protocol error flag separately; do not treat that assessment as successful
workbook opening or apply its exception to ordinary operation envelopes.

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

### Changed-path CI selection

`scripts\Get-ValidationPlan.ps1` owns the selections used by CI and the local
hook. Pull requests compare their head with the base branch's merge base.
Ordinary documentation and developer instructions, including nested `AGENTS.md`
and `CLAUDE.md`, avoid unrelated .NET tests and package jobs. Shipped documentation
is different: root, product, and npm-package READMEs, `LICENSE`, `CHANGELOG.md`, and
`docs\AGENT-SKILLS.md` select their consuming packages and owning tooling checks,
not runtime tests. Authored skill and plugin template trees remain package inputs,
including instruction files that their copy steps have not excluded. The extension
excludes developer instructions from its VSIX. Runtime and shared inputs select conservatively; multiple
inputs form a union. Main and manual CI runs select complete validation.

Hosted tests run in separate checkouts: `Fast` contains normal Excel-free
tests except `AdapterTestKind=System`; `Process` contains the CLI system
regressions; `Tooling` contains the selected skill-generation, packaging, and
script-safety checks. Each Tooling project has its own filter, so a script-safety
change cannot broaden a documentation-count or publication selection in another
project. Runtime source changes retain runtime/contract and process coverage,
but do not automatically select publication tests, metadata-only MCPB/plugins,
or authored skill ZIPs. Binary changes still select their consuming packages,
including the extension when its bundled server changes. Standalone tooling
test-project changes select their owning project, not every runtime project.
The full partitions cover the complete normal Excel-free selection without overlap.
Package, npm launcher, and lockfile checks have their own selections.
Changes to `doc-counts.json` or `scripts\check-doc-counts.ps1` select the
`Tooling` documentation-count regressions and source correctness checks,
without unrelated runtime tests or packages. Source checks run once in `Fast`
when selected, otherwise in `Tooling` for these count inputs. Preparatory CI
builds restore and build only selected test projects and their dependencies.
The group owning source/count checks still builds the full Release solution,
as required by `check-doc-counts.ps1 -SkipBuild`. All preparatory builds disable
build servers so rebuild regressions do not inherit assembly locks. Package
commands build their own required binaries without a preceding solution build;
metadata-only packages do not require .NET setup. NuGet and npm caches hold
dependency downloads, not shared compiled outputs.
`Docs Site` always runs. The required `CI Gate` always reports and rejects
failed detection, cancelled or failed work, and unexpectedly skipped jobs.
Hosted runners do not run real-Excel tests.

CodeQL uses the same planner, selecting actual C#, JavaScript/TypeScript,
Python, and GitHub Actions source or dependency/configuration inputs rather
than every file in a language directory. Publication `.mjs` scripts and tests
are included; Markdown alone does not select a language. Each selected analysis
still scans the complete language, with a traced manual C# build to include
generated code. Main, merge-group, scheduled, and manual CodeQL runs analyze
all languages. The required `CodeQL Completion` check reports even when no
analysis is selected, and rejects failed detection or incomplete selected work.

After a Release build, the hosted test partitions can also run locally:

```powershell
& .\scripts\Invoke-ExcelFreeTests.ps1 -Group Fast
& .\scripts\Invoke-ExcelFreeTests.ps1 -Group Process
& .\scripts\Invoke-ExcelFreeTests.ps1 -Group Tooling
```

Use `-PlanFile <plan.json>` with an explicit `-Group` to reproduce a selected
CI partition. To prepare just its build inputs:

```powershell
& .\scripts\Get-CiValidationPlan.ps1 -BaseRef origin/main -OutputPath artifacts\ci-plan.json
& .\scripts\Build-CiInputs.ps1 -PlanFile artifacts\ci-plan.json -Group Tooling
& .\scripts\Invoke-ExcelFreeTests.ps1 -PlanFile artifacts\ci-plan.json -Group Tooling
```

Choose a group listed in the saved plan; an unselected or empty group is an error.
Omitting the group retains the complete Excel-free run; existing
`-Local`, `-Contracts`, `-HookTests`, `-SkillTests`, and `-PackagingTests`
selections remain supported.

Generated MCP parameter tests inspect our emitted method declarations directly.
Protocol checks cover our names, descriptions, selected output fields, and
request handling, not the SDK's primitive JSON Schema type encoding or
provider-specific schema restrictions.

Keep representative tests for rules our generators implement, such as optional
enum strings, whole-second timeouts, file aliases, mixed cell values, and
injected parameters. One declaration-to-generated-contract completeness check
replaces repeated per-category inventories; compilation alone does not verify
that every declared action was emitted. Keep real parser, request, result, and
custom validation regressions at their owning entry point.

Guidance tests do not establish assistant understanding or performance. Do not
freeze sentences, editorial phrases, emoji rules, or example-count quotas.
Retain mechanical checks for metadata, packaged links, native example-block
selection, package integrity, and failed preparation preserving existing output.
Check native CLI help through the CLI parser, not by launching it from a skill
documentation scanner. Authored prose still needs review; API tests do not
validate its examples automatically.

Shared skill Markdown selects the existing skill-generation checks locally,
not Excel E2E. Program source, mixed runtime changes, and unknown inputs still
select Excel validation. Run the focused skill selection after a Release build:

```powershell
& .\scripts\Invoke-ExcelFreeTests.ps1 -Local -SkillTests
```

Excel-free distribution and script checks have separate projects:
`ExcelMcp.Packaging.Tests` owns release metadata, plugin publication, package
contents, and launch wrappers; `ExcelMcp.ScriptSafety.Tests` owns commit hooks,
validation selection, and script safety. `ExcelMcp.SkillGeneration.Tests` owns
only skill preparation, references, native examples, and standalone skill ZIPs.
Changed paths select the owning checks locally and in CI. Full validation runs
all three projects. Changes to the shared packaging helpers select both skill
and packaging checks, without unrelated runtime tests.

```powershell
& .\scripts\Invoke-ExcelFreeTests.ps1 -Local -PackagingTests
& .\scripts\Invoke-ExcelFreeTests.ps1 -Local -HookTests
```

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
Keep pure logic tests fast while preserving separate real lifecycle checks.

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

### Focused real-Excel groups and acceptance

Build the Release solution first. `Invoke-ExcelTests.ps1` obtains its inventory
from the built test assemblies and runs groups and projects sequentially.
Class fixtures stay together; multi-feature classes use one group rather than
overlapping feature filters.

```powershell
& .\scripts\Invoke-ExcelTests.ps1 -Groups Editing,Data -ListTests
& .\scripts\Invoke-ExcelTests.ps1 -Groups Editing,Data
& .\scripts\Invoke-ExcelTests.ps1 -Groups Infrastructure -IncludeInfrastructureDiagnostics
& .\scripts\Test-E2E.ps1 -SkipBuild
```

The available groups are `Editing`, `Reporting`, `Data`, `Lifecycle`,
`Infrastructure`, `Acceptance`, `VBA`, and `Desktop`. No `-Groups` selects all
normal groups. VBA and desktop tests require their actual prerequisites.
`-IncludeInfrastructureDiagnostics` adds ComInterop OnDemand probes except
configured IRM and Japanese-locale probes; those require separate configured
runs. Other OnDemand diagnostics and external-service evaluations remain separate.

`Acceptance` runs the complete required E2E stages, then the remaining normal
adapter acceptance cases without repeating required cases. Local commit hooks
run changed-area Excel-free checks and remind contributors to run complete
Excel E2E once against the final PR source; they do not repeat full E2E on each
commit. During development, use focused real-Excel groups or a focused E2E
stage as needed. Focused runs are not final acceptance. Run complete three-stage
E2E after the last runtime-affecting change, and rerun it if subsequent commits
change runtime behavior. Run affected Excel tests separately, including when
changing Excel-dependent tests.

`Test-E2E.ps1` defaults to three sequential stages: independent executable CLI
scenarios, the linked stale-build save/rebuild/reopen regression, and independent
real-protocol MCP scenarios. Each stage has a separate TRX report and a hard
execution deadline. Empty selections, skipped tests, failures, and assembly
cleanup failures fail the run. `-Stages Cli`, `-Stages Rebuild`, or `-Stages Mcp`
is a focused run, not complete runtime acceptance. `Test-CliWorkflow.ps1` is a
compatible wrapper for the CLI stage, including `-PipeName` and `-KeepFile`.
The CLI stage also retains the expanded native API workflow in
`Test-CliApiCoverage.ps1`, hosted by its own acceptance case with a private pipe
and a hard deadline. MCP native formatting/style and report-depth assertions
remain part of their independently reported acceptance scenarios.

Reports and ownership journals go into a new `TestResults` directory by default.
`-ResultsDirectory` can select another new directory. Reusing an existing
stage report is rejected so stale evidence cannot turn a failed run green.
`-ListTests` discovers cases without starting workbook operations.

For a full validation pass, run all solution test projects with
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
dotnet test tests\ExcelMcp.Packaging.Tests\ExcelMcp.Packaging.Tests.csproj @testArgs --filter 'RunType!=OnDemand' --logger 'trx;LogFileName=Packaging-normal.trx'
dotnet test tests\ExcelMcp.ScriptSafety.Tests\ExcelMcp.ScriptSafety.Tests.csproj @testArgs --filter 'RunType!=OnDemand' --logger 'trx;LogFileName=ScriptSafety-normal.trx'
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

**Repository requirements and the short instruction checklist:**

- **[Testing Strategy](AGENTS.md)** - Required safeguards; detailed procedures live in this guide
- **[Repository Rules](../AGENTS.md)** - Build, E2E, and contribution requirements

## Test Architecture

```
tests/
├── ExcelMcp.Portable.Tests/      # Mac adapter contracts and explicitly selected desktop probes
├── ExcelMcp.Core.Tests/           # Internal contracts and pure parsing tests
├── ExcelMcp.Service.Tests/        # Public workbook behavior through Service
├── ExcelMcp.Diagnostics.Tests/    # Excel COM behavior research (OnDemand, Manual)
├── ExcelMcp.McpServer.Tests/      # MCP protocol layer (Integration)
├── ExcelMcp.CLI.Tests/            # CLI wrapper (Integration)
├── ExcelMcp.ComInterop.Tests/     # COM utilities and session infrastructure
├── ExcelMcp.SkillGeneration.Tests/ # Skill preparation and standalone skill ZIPs
├── ExcelMcp.Packaging.Tests/        # Packaging, releases, and plugin publication
└── ExcelMcp.ScriptSafety.Tests/     # Commit hooks and script safety

llm-tests/                          # LLM tool behavior validation (Manual)
```

## Test Categories

| Category | Speed | Requirements | Run By Default |
|----------|-------|--------------|----------------|
| **Unit** | Fast | No Excel for pure logic | Select the relevant tests |
| **Integration** | Medium (10-20 min) | Excel + Windows | ✅ Yes (local) |
| **OnDemand** | Slow (3-5 min) | Excel + Windows | ❌ No (explicit only) |
| **Diagnostics** | Slow (varies) | Excel + Windows | ❌ No (manual, excluded from CI) |
| **LLM Tests** | Slow (varies) | Excel + GitHub Copilot | ❌ No (manual only) |

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
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "(Feature=Table|Feature=Tables)&RunType!=OnDemand"
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "Feature=PivotTables&RunType!=OnDemand"
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "(Feature=Range|Feature=Ranges)&RunType!=OnDemand"
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj --filter "(Feature=Connection|Feature=Connections)&RunType!=OnDemand"
```

## When to Run Which Tests

| Scenario | Command |
|----------|---------|
| **Daily development** | Run the smallest project and feature/name filter covering the change. |
| **Before commit** | Rerun affected tests and applicable checks; follow the [runtime E2E requirements](../AGENTS.md#build-and-validation). |
| **Modified session/batch code** | Run relevant OnDemand tests in `ExcelMcp.ComInterop.Tests`; see [Testing Strategy](AGENTS.md#commands). |
| **VBA development** | `dotnet test --filter "(Feature=VBA\|Feature=VBATrust)&RunType!=OnDemand"` |
| **LLM behavior validation** | See [LLM Tests](#llm-tests) section below |

## LLM Tests

The `llm-tests/` project checks selected real agent workflows through the MCP
Server and CLI using [pytest-skill-engineering 1.x](https://github.com/sbroenne/pytest-skill-engineering).
It complements deterministic behavior tests; it does not prove that agents
correctly use every operation. Independent workbook checks retain useful
chart, slicer, and permission coverage. Weak prose-only/API-success scenarios
have been retired or replaced.

### When to Run LLM Tests

- **Manual/on-demand only** - Not part of CI/CD
- After changing tool descriptions or adding new tools
- To check agent behavior, permission handling, and preservation of existing state
- To compare skill availability against no skill, with fixed requests/model/budgets

### Running LLM Tests

```powershell
# From llm-tests/
uv sync
uv run python -m unittest test_eval_harness.py test_cli_mcp_server.py test_consent_scenarios.py test_skill_value_checks.py -v
uv run pytest mcp_tests\test_mcp_chart_positioning.py --aitest-json TestResults\chart-new-run.json -v
```

### Prerequisites

- Windows desktop with Excel installed
- GitHub Copilot access and auth via `gh auth login`, `GITHUB_TOKEN`, or `GH_TOKEN`
- Release binaries and freshly generated skills; see the LLM README for setup

The offline command above needs neither Excel nor authentication. Live runs
must be sequential and may incur costs. The default model is `gpt-6.1-sol`,
not `auto`; skill-value runs never fall back to another model.
The default 60-execution comparison matrix requires explicit `--run-skill-value` and
a new `--skill-value-output` directory. Native JSON replaces removed AI/HTML
reports; Azure OpenAI is not a prerequisite. Missing usage is not zero cost.
The completed single-model comparison found equal verified correctness with
higher recorded token usage for both skills; it does not justify a general
reliability claim. See the linked results for the corrected matched comparison,
all-attempt overhead, and limits of the small sample.

The optional `--skill-value-suite real-world` suite combines four pinned,
licensed SpreadsheetBench workbook cases with Power Query recovery and Data
Model/PowerPivot refresh. It defaults to 48 executions; no new paid comparison
has been run. Dataset setup and no-model checker proofs are documented in the
[real-world pilot guide](../llm-tests/README.md#real-world-cases-and-native-excel-workflows).

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
- **Explicit persistence** - Call `batch.Save()` only when testing save/close/reopen behavior (see [Testing Strategy](AGENTS.md#save-and-round-trip-behavior)).

## Getting Help

- **Test failures**: Check test output for detailed error messages
- **Excel issues**: Ensure Excel 2016+ installed and activated
- **Session/batch issues**: Run OnDemand tests to verify cleanup
- **Writing tests**: See [Testing Strategy](AGENTS.md)
