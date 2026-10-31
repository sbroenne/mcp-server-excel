# ExcelMcp Tests

Excel-dependent behavior uses real Excel integration tests. Parsing, mapping,
serialization, and generation can use focused tests without Excel.
[ADR-001](../docs/ADR-001-TESTING-STRATEGY.md) explains this split.
Common SDK properties and xUnit/analyzer dependencies live in
`tests\Directory.Build.props`; each project keeps its platform overrides,
references, linked fixtures, assets, and test scheduling configuration.

Service worker and controller tests launch the Release `excelcli` executable.
The Service test project includes a build-only CLI project reference so selected
builds produce it without adding the CLI assembly to the test references.
Use `-c Release` for these tests; `--no-build` requires that build first.

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

### Behavior validation and saved evidence

Use the [changed-area plan](#changed-path-ci-selection) for PR validation.
It builds affected projects and dependencies, runs selected checks, and includes
selected real-Excel cases. Excel behavior tests are tagged `RequiresExcel=true`
and verify actual workbook or Excel-session outcomes against desktop Excel.
Excel-free parsing, adapter, packaging, and tooling checks remain separate.

During development, use a focused `dotnet test` filter as above and retain TRX
results. Use `--settings tests\strict.runsettings` to make empty selections,
skips and runner warnings fail. Direct `dotnet test` has per-test hang protection
with `--blame-hang-timeout`, but not the shared PowerShell runner's owned-tree
deadline, which also covers child processes and output drain.

For explicit Excel group runs, build Release first and use
`scripts\Invoke-ExcelTests.ps1`. It keeps class fixtures together and runs groups
and projects sequentially. `scripts\Test-E2E.ps1` runs complete CLI/MCP acceptance.
These runners retain stdout/stderr logs, TRX reports, and discovery output in a
fresh results directory. They enforce hard deadlines
and reject empty, skipped, failed, or contradictory reports. Excel selections
compare discovered and executed test names as multisets, including theory rows;
missing, extra, or duplicated cases fail the run.

Each stage prints its live stdout/stderr paths and an elapsed-time message
every 30 seconds. The stdout log includes individual test results as they
finish; a quiet log does not mean the current test has stopped. Read progress
from another PowerShell window without waiting for the runner to finish:

```powershell
Get-Content -LiteralPath '<printed path to stage.stdout.log>' -Tail 20 -Wait
```

Logs remain available after failure or timeout. Final TRX reports and the
runner's exit status determine success, not the progress messages.
MCP transport checks also allow the product's sequential auto-save, close,
STA-thread shutdown, and natural Excel exit waits before declaring a host
shutdown timeout; they do not stop Excel.
For GitHub check status, query the current PR directly with
`gh pr checks <number>`; this returns the current state without waiting for
all checks to finish.

GitHub-hosted machines do not have desktop Excel. Run Excel-dependent checks
locally on Windows with Excel installed. Product checks include COM-reference
safety, owned pre-build cleanup, test-result reporting, and real Excel acceptance.
The retained PowerShell script tests run with PowerShell 7.4 or later (`pwsh`).

Tests must leave pre-existing Excel instances, workbooks, sessions, and Power
Query workers untouched. Tests never stop Excel processes, including ones they
started: they close their own sessions through the product, and Excel still
running afterwards fails the test or run. Test compilation rejects machine-wide
process search and the product's process-stop functions. A timeout or failed
test is not permission for test code to stop Excel. The coding agent running
the tests may close or stop any Excel copy afterwards and must report each one
by process number and start time.

Use `-ListTests` with `scripts\Invoke-ExcelTests.ps1` or `--list-tests` with
`build.ps1 test --plan <path> --no-build` to discover without executing tests.
Discovery-only output is not passing test evidence. Neither
discovery nor a passing subgroup proves that every test contains strong
assertions. Keep the case-by-case source review separate and record unreviewed
cases honestly. Missing Excel, VBA trust, desktop, or other prerequisites mean
incomplete validation; do not manufacture a pass by skipping tests or changing
host settings. Explicit on-demand locale/IRM probes need their own focused run.

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
| Real timeout | Reached the blocked operation, preserved unsaved work, refused overlapping work, reused the same session after completion and readiness |

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

Pure validation-selection matrices run their input rows in one PowerShell
process per behavior, rather than paying for process startup for every path.
Keep each row's exact expectations and print its input before checking it;
run the row in its own script scope so local variables cannot leak to the next
row. Tests whose subject is process isolation, exit/failure behavior, Git state,
or environment changes still use separate processes. A lower runner case count
after this consolidation does not mean fewer input scenarios were checked.

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
Returning control while tests keep running is not a timeout. Follow the
[local validation scope rule](../AGENTS.md#build-and-validation) when choosing
coverage. Use ComInterop OnDemand diagnostics for a specific changed ownership,
threading, or cleanup risk not already covered by focused tests; narrow by test
name. Core OnDemand tests are optional diagnostics, not mandatory CI gates.

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
tests. Use the shared changed-area plan rather than requiring every Excel-free
project or complete acceptance after a feature-local edit. PR CI validates the
complete packages, site, and launchers through their owning checks. Complete normal
runs remain explicit commands: `scripts\Invoke-ExcelFreeTests.ps1` (without
`-Local`) includes packaged-plugin validation and complements, rather than
replaces, focused native tests and normal hooked runtime E2E. Run
`npx --no-install changeset status --since=origin/main` when applicable;
internal tooling, tests, and CI work uses `skip-changelog`.
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
tests. Independent safe adapter cases may reuse a real host/session. Real
Excel is used by CLI/MCP tests only for one save-and-reopen smoke per entry
point (`CliWorkflowAcceptanceTests` and `McpServerSmokeTests`) and for
fresh-process, deadline, crash, protected-file, and ownership coverage. When an
adapter regression involves workbook behavior, keep the workbook assertion in
Service and the adapter assertion as an Excel-free recorded-request test. Do
not remove adapter regressions or shorten production waits merely to speed tests.

CLI parser contracts use the production Spectre command app in-process with an
explicit request client and captured input/output. They assert the exact
Service request and public output envelope without starting a daemon or Excel.
Batch validation may use a real in-process Service when validation belongs to
that public boundary. Keep a small executable-and-pipe smoke set plus the
distinct mutex, startup, deadline, crash, persistence, and forced-stop cases;
do not launch a process for every argument or alias row.

The commit hook keeps only deterministic in-process adapter tests: its `Commit`
projection excludes `AdapterTestKind=System` and `Speed=Slow`. CI still runs
both classes in its regular `Fast` partition and the separately classified
`Process` partition. The `Fast` partition is a validation group, not a promise
that every included case is cheap.

MCP still has a small set of real SDK calls in the hook. These catch request
binding and serialization mistakes that source or generated-metadata checks
cannot prove. Exhaustive all-action schema and protocol sweeps remain in CI.

Use these Excel-free filters for quick adapter feedback:

```powershell
dotnet test tests\ExcelMcp.CLI.Tests\ExcelMcp.CLI.Tests.csproj -c Release --no-build --filter 'RunType!=OnDemand&RequiresExcel=false&AdapterTestKind!=System&Speed!=Slow'
dotnet test tests\ExcelMcp.McpServer.Tests\ExcelMcp.McpServer.Tests.csproj -c Release --no-build --filter 'RunType!=OnDemand&RequiresExcel=false&AdapterTestKind!=System&Speed!=Slow'
```

The quick groups are not acceptance gates. Complete normal validation still
uses `RunType!=OnDemand`, including the separately classified real Excel,
process, deadline, crash, rebuild, and ownership cases below.

### Changed-path CI selection

`scripts\Get-ValidationPlan.ps1` owns the explicit changed-path and feature-label
rules used by CI, CodeQL and the local hook. Standard .NET and xUnit commands
perform builds and tests; there is no custom C# tool to compile or load.
Test-only changes select their owning project group; unit and recording-only
protocol tests remain Excel-free. Real-Excel MCP transport fixtures and their
shared helpers select Excel coverage. Shared runtime helpers select broader coverage
instead of rebuilding a source-reference catalogue. Pull requests compare
their head with the base branch's merge base. The hook uses staged inputs and
retains merge-parent and unstaged-input safeguards.
Edits to partial `PersistentServiceWindowTests` files select the complete
owning test class, keeping its fixture together without selecting unrelated
Service Excel classes.

MSBuild owns .NET project evaluation, imports, dependencies, restore, compilation,
and incremental build decisions. The wrappers call the normal SDK commands;
they do not interpret project XML, reconstruct package consumers from Git
snapshots, or maintain a separate up-to-date cache. Multiple selected entry
projects use a native solution filter, whose dependencies MSBuild resolves.
Root `Directory.*`, SDK, NuGet, and solution changes select the complete Release
solution build and shared checks. This intentionally trades narrower checks on
central package updates for removing a second, incomplete dependency evaluator.
Malformed build inputs fail in MSBuild, not in a custom XML interpreter.
Changed-path routing still selects test environments;
it is not a substitute build system. Unknown paths fail explicitly.
The hook reuses its saved staged-input plan for quick CLI, MCP Server, Core,
and Service tests (`Commit`, a projection of `Fast`). ComInterop-specific
tests remain in CI; their real-Excel cases require a Windows/Excel environment.
It does not run process, tooling, packaging, Excel, or acceptance checks on each
commit. CI owns supported slower checks; run local process checks when changing
their boundary, not on every push. Selected Excel checks still require desktop
Excel locally. CI uses its
selected base, head, and `MergeBase`/`Direct` comparison mode for changed paths.
When no selected test group builds an explicitly selected test project, CI also
runs a separate build-only job for the saved inputs.
The required CI gate checks its exact result, without presenting an empty
test selection as passed tests.

```powershell
.\build.ps1 plan --base-ref origin/main --head-ref HEAD --output artifacts\validation-plan.json
.\build.ps1 validate --plan artifacts\validation-plan.json
.\build.ps1 test --area PowerQuery
```

`validate` builds selected .NET inputs with compiler safety checks
and runs selected .NET tests. Package preparation and website/npm checks remain
separate CI jobs; `validate` does not publish or prepare release artifacts.
Test-only edits select their classes and actual shared-helper consumers.
Feature implementations select their Core and Service behavior without an
unchanged CLI/MCP matrix or npm source suite. Contract edits also select generated
consumers. Unknown inputs fail instead of selecting everything. Every selected
test report must contain executed passing cases; unexpected skips, cleanup
failures, and hard-deadline expiry fail the command.
Compiler safety rules live in `src\ExcelMcp.Analyzers` and banned-API build inputs,
with focused compiler regressions in `SafetyAnalyzerTests`. The old standalone
source-pattern scanner commands have been retired. These checks identify
supported risky patterns, not complete COM lifetime or
success/error control-flow correctness.
Ordinary documentation and developer instructions, including nested `AGENTS.md`
and `CLAUDE.md`, avoid unrelated .NET test selections. Shipped documentation
is different: root, product, and npm-package READMEs, `LICENSE`, `CHANGELOG.md`, and
`docs\AGENT-SKILLS.md` select their owning tooling checks,
not runtime tests. Authored skill and plugin template trees remain package inputs,
including instruction files that their copy steps have not excluded. The extension
excludes developer instructions from its VSIX. Shared inputs include affected
dependencies; multiple inputs form a union. Main and manual CI runs select
complete validation.

Hosted tests run in separate checkouts: `Fast` contains normal Excel-free
tests except `AdapterTestKind=System`; `Process` contains the CLI system
regressions; `Tooling` contains the selected skill-generation, packaging, and
script-safety checks. Each Tooling project has its own filter, so a script-safety
change cannot broaden a documentation-count or publication selection in another
project. Skill sources, the shared report-formatting guide, and skill preparation
also select plugin bootstrap tests, which check the assembled plugins and their
published layout. Plugin validation checks current files, versions, launch
settings, local links, and script behavior rather than a list of historical
command names.
Runtime source changes select owning behavior and necessary boundaries,
but do not automatically select publication tests, metadata-only MCPB/plugins,
or standalone skill ZIPs (which are no longer published).
PR CI uses `PackageComponents` and `NpmTests` from the saved plan to run affected
distribution checks, including the extension source suite when Extension is
selected. Shared packaging inputs select all dependent components; test-only
npm changes run npm checks without real runtime packages. Main/manual CI and
release preparation remain complete. Runtime-only PR changes retain product
checks but can defer distribution-only defects to main.
Explicit local runs may select `-Components` or opt out of extension tests with
`-SkipExtensionTests`; CI never skips selected extension tests.
Standalone tooling test-project changes select their owning project, not every
runtime project.
The full partitions cover the complete normal Excel-free selection without overlap.
Changes to `infrastructure\azure\configure-analytics-oidc.ps1` or
`infrastructure\azure\deploy-appinsights.ps1` also record the Azure preview and
subscription-restoration checks in `ToolingOnDemandFilters`. Both local and
saved-plan Tooling runs include these selected `RunType=OnDemand` cases alongside
the normal safety checks; unrelated OnDemand diagnostics remain excluded.
EXL030 validates count inputs during MCP compilation using the compiled Core
manifest; there is no standalone count scanner or updater. Use the diagnostic's
exact expected values to update affected source counts, review the changes,
then rebuild. Retain `McpToolSurfaceTests` for compiled read-tool identities and
help-banner behavior, and `SkillPreparationTests` for exact prepared-skill
identities; these are stronger runtime contracts, not source regex checks.
Preparatory CI builds follow the saved plan for selected projects and dependencies.
When a full Release solution build is required, the first selected group owns it.
All preparatory builds disable
build servers so rebuild regressions do not inherit assembly locks. Package
commands build their own required binaries without a preceding solution build;
metadata-only packages avoid product runtime builds, but still require the
repository SDK from `global.json` for the internal build tool. NuGet and npm caches hold
dependency downloads, not shared compiled outputs.
`Docs Site` always reports, but builds the site only when its inputs changed.
The required `CI Gate` always reports and rejects
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
selections remain supported. `-Local -Contracts` includes generated routing,
MCP generation and registration, both CLI and MCP usage-analytics mappings,
the advertised MCP tool surface, and authored calculation examples checked
against the published schema. Contract and generator changes select these same
suites. Tests retain their existing feature traits as well as
`Feature=GeneratedContracts`.

These commands call shared PowerShell coordination. `Invoke-TestStage.ps1` uses
the same hard deadlines, owned-process environment and strict report checks as
saved-plan execution. `Get-ExcelTestGroups.ps1` exports the actual built inventory,
not a guessed list of source classes. MSBuild evaluates selected build inputs.
CI completion checks use PowerShell without installing an SDK; detection failure,
missing cases and unexpected skips are errors.

Package preparation uses the same PowerShell command as Actions, without publishing anything:

```powershell
.\scripts\Build-ReleasePackages.ps1 -Components Cli,Mcp -OutputDirectory artifacts\local-packages
```

Use a new output directory. Existing `Build-ReleasePackages.ps1`,
`Build-NpmPackages.ps1`, `Test-NpmPackages.ps1`, `Build-AgentSkills.ps1`,
`Build-Plugins.ps1` and `mcpb\Build-McpBundle.ps1` keep their parameters and
own their named operations directly. Shared helpers cover process deadlines and
filesystem/archive safety; no C# packaging bootstrap is involved.
This is optional local troubleshooting, not an added commit check.
Runtime packages retain architecture and version
checks, npm archive inspection and native launcher checks. ARM64 archive
validation is not ARM64 execution on an x64 machine. Extension preparation still
compiles, lints and inspects both VSIX targets when unchanged source tests are
explicitly skipped.

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
not Excel E2E. Program source and mixed runtime changes can still select Excel
validation; unknown operational inputs fail selection. Run the focused skill
selection after a Release build:

```powershell
& .\scripts\Invoke-ExcelFreeTests.ps1 -Local -SkillTests
```

Excel-free distribution and script checks have separate projects:
`ExcelMcp.Packaging.Tests` owns release metadata, plugin publication, package
contents, and launch wrappers; `ExcelMcp.ScriptSafety.Tests` owns commit hooks,
validation selection, and script safety. `ExcelMcp.SkillGeneration.Tests` owns
only skill preparation, references, and native examples.
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

The [compiler classification analyzer](../src/ExcelMcp.Analyzers/TestClassificationAnalyzer.cs)
owns the supported trait, collection, fixture, and system-test ownership checks.
The remaining built-assembly inventory and feature-group tests validate runtime
selection semantics; compilation does not replace those checks or execute tests.

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
adapter acceptance cases without repeating required cases. Local commit checks
use the shared changed-area policy, not complete E2E after every runtime edit.
Run affected real-Excel
groups separately, including when changing Excel-dependent tests.

`Test-E2E.ps1` defaults to three sequential stages: independent executable CLI
scenarios, real-protocol MCP scenarios, and external OLAP PivotTable tests
through the Service boundary. The OLAP stage covers field placement, measures,
calculated members, slicers, charts, and unsupported cube operations. It also
checks saved-and-reopened connection/layout/selection/chart state, multi-level
Calendar rows with RelativePeriod columns and filters, compatible connection
changes, rejected changes without altering either PivotTable, and failed refresh
against a nonexistent cube followed by explicit repair. These are real Excel
checks against the synthetic provider, not coverage of every server's behavior.
Configure all four `EXCELMCP_TEST_OLAP_CONNECTION_STRING`,
`EXCELMCP_TEST_OLAP_CUBE`, `EXCELMCP_TEST_OLAP_HIERARCHY`, and
`EXCELMCP_TEST_OLAP_LEVEL` settings to use a compatible external fixture.
Partial settings fail rather than silently testing a different cube.
Configured fixtures must provide the same synthetic rows, hierarchy captions,
relative periods, and measures described below. Member identities and the
All-member identity are read independently from Excel, so a compatible SSAS
fixture need not use Atoti's member names. Selection snapshots still require
the documented ordinal ordering of those identities. Exact Atoti identity
assertions remain enabled for the runner-owned synthetic provider.
With none set, `scripts\Start-OlapTestCube.ps1` starts the synthetic cube in
`scripts\olap_test_cube.py` on a free localhost port and the runner stops its
owned Python/Java process tree afterwards. The first run needs Python 3.12 or
newer and internet access and installs Atoti 6.2.2 into `%LOCALAPPDATA%\emo`.
Environment creation and package installation share a 300-second setup deadline;
cube startup has a separate 300-second deadline before the OLAP test deadline.
Timed-out setup stops its owned process tree and does not mark installation complete.
Atoti Community Edition is free for non-commercial use, requires online license
activation, and sends usage telemetry; see its
[license](https://docs.activeviam.com/engine/python-sdk/latest/eula).
The cube has 24 monthly records across 2024-2025, Amount values 1-24, a
Year/Month Calendar hierarchy, and fixed Prior Month/Current Month/Next Month
members for October/November/December 2025. Older records use Earlier Months.
It contains no customer data.
Excel keeps the cube connection open while the fixture PivotTable uses it.
Each stage has a separate TRX report and a hard execution deadline. Empty
selections, skipped tests, failures, and assembly cleanup failures fail the run.
`-Stages Cli`, `-Stages Mcp`, or an OLAP-only selection is focused, not complete
runtime acceptance. `Test-CliWorkflow.ps1` is a
compatible wrapper for the CLI stage, including `-PipeName` and `-KeepFile`.
Use `Test-E2E.ps1 -Stages Cli` for the direct CLI acceptance command.

### OLAP server coverage

SSAS checks require a separately configured SQL Server Analysis Services
instance and a local Windows desktop with Excel. With no
`EXCELMCP_TEST_OLAP_*` configuration, the OLAP stage uses Atoti; it does not
discover or select an installed SSAS instance automatically.

The All-member regression introduced for issue #1135 was verified against
Atoti Community Edition 6.2.2, Excel's built-in Data Model (Power Pivot), and
SSAS 17.0.25.223 in multidimensional mode using the AdventureWorks sample cube.
That focused SSAS Service regression passed; this does not
mean the complete external suite was run against SSAS. Azure Analysis Services
and Power BI XMLA have not been tested. SSAS tabular date-filter coverage is
documented separately below; it does not cover this All-member regression.

The 20 configured external OLAP acceptance cases have also passed individually
against SSAS 17.0.25.218 in tabular mode, using a separate synthetic fixture with
the 24 rows described above. This covers schema paging, provider-specific member
identities, report-filter All/selected/cleared states, multi-level snapshots,
and the existing external PivotTable operations and persistence checks. It does
not establish coverage of AdventureWorks, other SSAS configurations, Azure
Analysis Services, or Power BI XMLA.

The focused SSAS regression is separate from the Atoti fixture. Configure
`EXCELMCP_TEST_SSAS_CONNECTION_STRING` in the test process for the sample database;
the test uses `[Product].[Category]`, the Bikes member, and
`[Measures].[Internet Sales Amount]`. Without this setting the test is skipped,
which is not a successful SSAS validation. Keep credentials and connection
strings out of committed files, public logs, and PR descriptions.

```powershell
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj -c Release --no-restore --disable-build-servers --filter 'FullyQualifiedName~SsasReportFilter_AdventureWorks_NonstandardAllRoundTripsThroughService' --logger 'trx;LogFileName=ssas.trx' --results-directory TestResults\ssas-report-filter --blame-hang-timeout 5m --verbosity quiet
```

It reads the actual All-member identity through Excel's ADO connection and
requires a nonstandard name. It checks filter reads and semantic snapshots,
explicit All assignment, a real selected member, clearing, unchanged
single-selection mode, data totals, and save/reopen followed by explicit refresh.

#### SSAS tabular date filters

The separate date-filter regression uses a small synthetic SSAS tabular model,
not the multidimensional Adventure Works cube or the generic Atoti fixture.
Execute this TMSL in an SSAS management query window connected to a tabular
test instance. The create command refuses to overwrite an existing database:

```json
{
  "create": {
    "database": {
      "name": "ExcelMcpDateFilterTest",
      "compatibilityLevel": 1200,
      "model": {
        "culture": "en-US",
        "tables": [{
          "name": "Dates",
          "dataCategory": "Time",
          "columns": [
            { "name": "Date", "type": "calculatedTableColumn", "dataType": "dateTime", "sourceColumn": "[Date]", "isKey": true, "formatString": "yyyy-mm-dd" },
            { "name": "Amount", "type": "calculatedTableColumn", "dataType": "int64", "sourceColumn": "[Amount]" }
          ],
          "partitions": [{
            "name": "Dates",
            "source": {
              "type": "calculated",
              "expression": "UNION(ROW(\"Date\",DATE(2024,1,15),\"Amount\",100),ROW(\"Date\",DATE(2024,2,10),\"Amount\",200),ROW(\"Date\",DATE(2024,3,5),\"Amount\",400),ROW(\"Date\",DATE(2024,4,12),\"Amount\",800))"
            }
          }],
          "measures": [{ "name": "Total Amount", "expression": "SUM(Dates[Amount])", "formatString": "0" }]
        }]
      }
    }
  }
}
```

Process the newly created `ExcelMcpDateFilterTest` database:

```json
{
  "refresh": {
    "type": "full",
    "objects": [{ "database": "ExcelMcpDateFilterTest" }]
  }
}
```

The model has one time table, `Dates`, with a date-time key column `Date`,
an `Amount` column, and the `Total Amount` measure. Its four dates are
2024-01-15, 2024-02-10, 2024-03-05, and 2024-04-12, with amounts 100, 200,
400, and 800. It has no external data source or credentials.

Set `EXCELMCP_TEST_SSAS_DATE_CONNECTION_STRING` in the test process to an
MSOLAP connection to this database. This setting is independent of the
Adventure Works connection above. Without it, the on-demand tests skip;
a skipped test is not SSAS validation.

```powershell
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj -c Release --no-restore --filter 'FullyQualifiedName~SsasDateFilters_TypedDateLevel' --logger 'trx;LogFileName=ssas-date.trx' --results-directory TestResults\ssas-date --blame-hang-timeout 3m --verbosity quiet
```

All eight supported date-filter types were verified against SSAS 17.0.25.218
in tabular mode through desktop Excel and the Service entry point. Each case
checks exact returned criteria, actual date rows and totals, semantic snapshot
completeness, and clear/readback restoring all four dates. Excel reports this
valid SSAS date level's `PivotField.DataType` as `xlText`, so that property
cannot be used to reject OLAP date filters. Native OLAP calls omit
`WholeDayFilter`; regular PivotTables retain their date-type check and
whole-day behavior. Source metadata still matters: the existing Data Model
`SalesTable.Date` fixture rejects native date filtering and remains a negative
regression, not evidence that every OLAP date field is unsupported.

This is SSAS tabular coverage only; it does not establish date-filter support
on SSAS multidimensional, Atoti, Power BI XMLA, or every Data Model date field.

To use the general external suite against another SSAS fixture, configure all
four `EXCELMCP_TEST_OLAP_*` settings listed above in the test process.
The existing external suite assumes the fixture's Calendar/RelativePeriod
hierarchies, Amount measure, member unique names, and monthly values. Pointing
the settings at an arbitrary SSAS cube is not sufficient; use a compatible
fixture or add focused SSAS tests that assert its actual schema and results.
To cover this regression, the test must establish the server's actual All-member
identity and check both All and a genuine selected member through Excel.

Run locally against the exact feature-branch commit being reviewed, not just
`main`. Record the SSAS version and cube mode, tested commit, exact commands,
and TRX outcomes in the PR before claiming SSAS coverage. Run Excel-dependent
commands sequentially.

Development builds directly stop only CLI services from the current worktree
without saving workbooks. They do not kill Excel or stop other worktrees' services.
The old graceful-save rebuild gate is removed; normal CLI/MCP persistence and
safe-close tests remain. Test-run service stopping also specifies its private
pipe so it cannot stop another local fixture's CLI service.

MCP native formatting/style and report-depth assertions
remain part of their independently reported acceptance scenarios.

Reports go into a new `TestResults` directory by default.
`-ResultsDirectory` can select another new directory. Reusing an existing
stage report is rejected so stale evidence cannot turn a failed run green.
`-ListTests` discovers cases without starting workbook operations.

For a full validation pass, run all solution test projects with
`RunType!=OnDemand`, then run complete three-stage E2E. The OLAP E2E stage
explicitly selects the OnDemand external-cube tests rather than skipping them.
The normal filter includes Service VBA and screenshot
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

Tests never stop Excel. Each test closes the sessions it created through the
product; per-test checks then wait for those exact Excel processes (process ID
plus start time) to exit and fail the test if any are still running. Never
infer ownership from all Excel processes started since a test began, or close
existing Excel instances to prepare a fixture. An independently opened workbook
must survive successful and failed teardown.

Core, ComInterop, Service, and in-process MCP test hosts also run a run-level
check (`TestRunExcelLeftoverCheck`). It keeps the original process handle of every
Excel process the product starts during the run and, before results are
reported, waits up to 15 seconds for each to exit. Any that remain fail the run
and are named in the error; they are marked so the product's own exit handling
also leaves them running for inspection. This does not cover a separate CLI
background service started by a client test.

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
dotnet test tests\ExcelMcp.CLI.Tests\ExcelMcp.CLI.Tests.csproj @testArgs --filter 'RunType!=OnDemand' --logger 'trx;LogFileName=CLI-main.trx'
dotnet test tests\ExcelMcp.McpServer.Tests\ExcelMcp.McpServer.Tests.csproj @testArgs --filter 'RunType!=OnDemand' --logger 'trx;LogFileName=MCP-main.trx'
dotnet test tests\ExcelMcp.ComInterop.Tests\ExcelMcp.ComInterop.Tests.csproj @testArgs --filter 'RunType!=OnDemand' --logger 'trx;LogFileName=ComInterop-normal.trx'
dotnet test tests\ExcelMcp.SkillGeneration.Tests\ExcelMcp.SkillGeneration.Tests.csproj @testArgs --filter 'RunType!=OnDemand' --logger 'trx;LogFileName=Skills-normal.trx'
dotnet test tests\ExcelMcp.Packaging.Tests\ExcelMcp.Packaging.Tests.csproj @testArgs --filter 'RunType!=OnDemand' --logger 'trx;LogFileName=Packaging-normal.trx'
dotnet test tests\ExcelMcp.ScriptSafety.Tests\ExcelMcp.ScriptSafety.Tests.csproj @testArgs --filter 'RunType!=OnDemand' --logger 'trx;LogFileName=ScriptSafety-normal.trx'
dotnet test tests\ExcelMcp.Diagnostics.Tests\ExcelMcp.Diagnostics.Tests.csproj @testArgs --filter 'RunType!=OnDemand' --logger 'trx;LogFileName=Diagnostics-normal.trx'
dotnet test tests\ExcelMcp.ComInterop.Tests\ExcelMcp.ComInterop.Tests.csproj @testArgs --filter 'RunType=OnDemand&FullyQualifiedName!~BeginBatch_RealIrmWorkbook&Locale!=ja-JP' --logger 'trx;LogFileName=ComInterop-infrastructure.trx'

# Exclusive VBA groups: obtain permission, capture the original setting, enable
# access only for the run, then restore/verify it as described above.
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj @testArgs --filter 'RunType!=OnDemand&Feature=VBA' --logger 'trx;LogFileName=Service-vba.trx'

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
├── ExcelMcp.Core.Tests/           # Internal contracts and pure parsing tests
├── ExcelMcp.Service.Tests/        # Public workbook behavior through Service
├── ExcelMcp.Diagnostics.Tests/    # Excel COM behavior research (OnDemand, Manual)
├── ExcelMcp.McpServer.Tests/      # MCP protocol layer (Integration)
├── ExcelMcp.CLI.Tests/            # CLI wrapper (Integration)
├── ExcelMcp.ComInterop.Tests/     # COM utilities and session infrastructure
├── ExcelMcp.SkillGeneration.Tests/ # Skill preparation and native examples
├── ExcelMcp.Packaging.Tests/        # Packaging, releases, and plugin publication
└── ExcelMcp.ScriptSafety.Tests/     # Commit hooks and script safety

llm-tests/                          # LLM tool behavior validation (Manual)
```

## Test Categories

| Category | Speed | Requirements | Run By Default |
|----------|-------|--------------|----------------|
| **Unit** | Fast | No Excel for pure logic | Select the relevant tests |
| **Integration** | Medium (10-20 min) | Excel + Windows | ✅ Yes (local) |
| **OnDemand** | Varies | Excel + Windows; test-specific prerequisites | ❌ No (explicit only) |
| **Diagnostics** | Slow (varies) | Excel + Windows | ❌ No (manual, excluded from CI) |
| **LLM Tests** | Slow (varies) | Excel + GitHub Copilot | ❌ No (manual only) |

## Optional Legacy `.xls` Checks

The legacy `.xls` save and validation checks use `RunType=OnDemand` because
their availability depends on the local Office policy for Excel 97-2003 files.
They still verify the saved format, session path, workbook contents, and file
preservation. The `.xlsx`, `.xlsm`, and `.xlsb` checks remain in the normal
suite.

Run them only on a machine whose existing Office policy permits legacy workbook
files. These tests do not change Trust Center File Block settings:

```powershell
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj -c Release --filter 'RequiresExcel=true&RunType=OnDemand&FullyQualifiedName~LegacyXls' --blame-hang-timeout 5m --logger trx
```

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
| **Modified session/batch code** | Cover the changed lifetime and failure paths; add targeted ComInterop OnDemand diagnostics where focused tests leave a specific risk uncovered. See [Testing Strategy](AGENTS.md#commands). |
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

Use a unique workbook with the established feature fixture.
Follow neighboring trait conventions and the COM cleanup rules
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
- **Session/batch issues**: Start with focused lifetime and failure tests; add a targeted OnDemand diagnostic for a cleanup risk those tests do not cover.
- **Writing tests**: See [Testing Strategy](AGENTS.md)
