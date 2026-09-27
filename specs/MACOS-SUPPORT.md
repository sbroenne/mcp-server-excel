# macOS support: initial release and parity plan

**Status:** The capability-gated macOS version ships the MCP Server and
`excelcli` on Apple Silicon. Windows retains the complete COM backend. macOS
now combines native Apple Events operations with secure saved-package Power
Query reads. Additional capability tiers are introduced gradually; unavailable
operations fail explicitly rather than returning success-shaped fallbacks.

## Outcome

Native Mac Excel automation supports the initial workbook/range operation set
through Apple Events. Real-Excel validation passed creation, explicit workbook targeting
across separate processes, values, formulas, formatting, bulk calculation,
save/reopen, discard, and owned-workbook cleanup.

**A prompt-free test path is now demonstrated:** seed blank synthetic `.xlsx`
files in an ordinary temporary directory and open them through macOS
LaunchServices. Four runs passed, including one after a user-managed Excel
restart, with no dialogs confirmed by the user. All editing, calculation and
save/reopen assertions used real Excel. The same handoff now works through the
actual CLI daemon and MCP stdio server, with both workflows passing and no
dialogs confirmed by the user. This proves the tested existing-file workflows,
not Excel-native workbook creation or all file-opening options.

Both foreign-container approaches failed the no-prompt requirement. Access from
Copilot prompted again after explicit one-time approval; a separately launched,
ad-hoc-signed native app also prompted again. Successful workbook operations
did not make those runs prompt-free. No broad privacy grants, macro-security
changes, or UI auto-clicking were used.

The product goal is **maximum parity with Windows ExcelMcp, especially Power
Query and VBA**, not a permanently reduced workbook/range product. The initial
release is the first delivery milestone, not the target feature set. Backend evolution
must account for Power Query and VBA.
Resolve permission/onboarding and shared-Excel ownership alongside those
feasibility probes. Do not promise unverified parity or unattended arbitrary-file
access.

A repository-owned `[MS-QDEFF]` package reader/writer now parses exact M query
definitions and can update an exact query transactionally while resetting
Power Query permissions. Production `list`, `view`, and `get-load-config` use
this reader for clean, saved workbooks. Package-only `update` is available when
`refresh=false`; it uses exact-workbook close, same-directory backup, validated
replacement, reopen, and rollback. Worksheet load state is derived from OOXML
worksheet, table, QueryTable, and connection relationships because the live
Apple Events object model does not expose that graph reliably. Refreshing
updates and destination mutations remain gated pending unattended real-Excel
fixtures that can prove completion, errors, and cleanup.

## Implementation status

The first implementation increment now exists behind runtime platform selection:

- MCP Server and `excelcli` compile as `net10.0` hosts on macOS. Windows retains
  `net10.0-windows`, WinForms tray integration, SID-secured pipes, COM routing
  and owned-process cleanup.
- The shared Service selects a serialized JXA/Apple Events backend on macOS.
  CLI IPC uses a stable hashed per-user identity and current-user-only Unix
  named pipes. The Mac daemon has no tray and never force-kills shared Excel.
- Logical Mac sessions track workbook paths and serialize operations per
  workbook. Timeout handling attempts workbook cleanup without killing Excel;
  complete invalidation and recovery guarantees still need hardening.
- Implemented native bridge actions are session create/open/close, worksheet
  create/list/rename/delete, range get/set values and formulas (including
  existing JSON/CSV file transforms), number-format read/write, explicit
  row/column sizing, clear operations, calculation, Goal Seek, and one- or
  two-variable Data Tables.
- Power Query `list`, `view`, and `get-load-config` are implemented for clean,
  saved workbooks using package inspection. Package-only `update` requires
  `refresh=false`; the contract default remains `refresh=true`, so Mac callers
  must opt out explicitly until refresh parity is proven. Dirty workbooks fail
  with save-or-discard guidance rather than returning stale package data.
  Workbooks without a DataMashup return an empty query list. Workbooks
  containing a Data Model fail explicitly until package-level query-to-model
  identity can be established; they are not misreported as connection-only.
- Other production Mac Power Query actions and VBA source CRUD return an
  explicit capability-tier failure. They are not reported as success and no
  unvalidated rewrite is substituted.
- Existing-file open uses a non-prompting native Automation preflight, rejects
  an already-open target, hands the exact path to LaunchServices, then attaches
  through JXA under a shared deadline.
- Portable regression tests cover platform-neutral path validation and errors,
  service startup/status without Excel, stable user-scoped IPC naming, native
  permission status/layout, handoff ordering, failures and deadlines.
- Actual CLI/MCP tests cover inline/file-backed values, formulas/calculation,
  number formats, dimensions, worksheet creation, missing-sheet failure and
  recovery, save/reopen, discard, sentinel preservation, clean empty Power
  Query lists, dirty-workbook rejection, and cleanup.

This is **not completion of the parity plan**. Worksheet creation uses the
working nested AppleScript form after JXA returned Excel parameter errors.
Power Query create/rename/delete/load/unload/evaluate and all refresh variants
remain gated until their completion, error, destination, and cleanup semantics
have repository-safe unattended real-Excel evidence. VBA
list/view/import/update/delete remain gated because Excel's installed Apple
Events dictionary exposes macro execution but not the VB project/code-module
object model. The approved design now permits gradual, optional helpers:
Office.js for broad workbook features, a macro helper for execution, and a VBA
project-model tier only after explicit user trust. Base macOS support must not
require macros or VBA project trust.

## Required parity and priorities

The existing Windows source contracts are the baseline. Match operation names,
parameters, defaults, validation, results, errors, persistence and observable
behavior wherever Mac Excel permits. MCP and CLI must remain equal entry points
on both platforms. Implementation details may differ; a replacement workflow
that merely produces similar cells does not establish contract parity.

**Power Query and VBA are primary acceptance requirements**, ahead of expanding
secondary features such as chart/window customization. A range-only preview may
be useful internally, but does not satisfy the planned macOS feature scope.
There is no accepted Power Query or VBA scope reduction in this plan.

Build an action-level compatibility inventory from the current contracts,
starting with [IPowerQueryCommands](../src/ExcelMcp.Core/Commands/PowerQuery/IPowerQueryCommands.cs)
and [IVbaCommands](../src/ExcelMcp.Core/Commands/Vba/IVbaCommands.cs), then covering
every other category. For each action and relevant parameter variant, record:
Windows semantics, candidate Mac mechanism, fixture/evidence, platform/version
requirements, and status (verified parity, partial, blocked, or not yet tested).
Do not classify an untested action as a platform limitation. Every proposed
exception needs evidence, its user impact, alternatives considered, and an
explicit scope decision before release.

### Power Query parity workstream

Target the complete existing lifecycle, not just refreshing queries authored on
Windows:

| Contract area | Required investigation and acceptance evidence |
| --- | --- |
| `list`, `view`, `get-load-config` | Enumerate queries, return exact M code and actual load state; preserve compact list metadata and the bounded preview; surface inspection failures |
| `create`, `update`, `rename`, `delete` | Author and maintain M through Mac Excel; preserve inline/file input rules, exact code by default, refresh defaults, name validation, exact query identity and no implicit save |
| `evaluate` | Execute M through Excel's engine, return matching columns/rows and useful engine errors, and verify removal of temporary query/sheet/table/connection artifacts on success and failure |
| `load-to`, `unload` | Exercise worksheet and connection-only destinations, target sheet/cell behavior, existing-data preservation, and destination transitions; investigate Data Model-dependent variants separately |
| `refresh`, `refresh-all` | Observe actual refresh completion and source errors rather than just command acceptance; match caller/default timeouts, failure reporting and recovery behavior |

Fixtures must include literal M tables, dependent queries, prefix/case-sensitive
identity edge cases according to the Windows contract, syntax/runtime errors,
typed dates/nulls, worksheet-source queries, local file sources and approved
authenticated-source scenarios. Verify saved query definitions and load state
after reopening, with no unrelated workbook data or connections changed.
Keep credentials out of fixtures and logs. External M formatting remains
opt-in, as on Windows.

First investigate native Mac `Workbook.Queries`/`WorkbookQuery` access through a
VBA bridge and the actual worksheet-loading/refresh mechanism available on Mac.
Do not assume Windows Mashup OLE DB connection strings/providers exist on Mac.
Query-definition access alone is not sufficient: loading, evaluation cleanup,
refresh completion, authentication and errors must also work. Record connector
and Excel-version limitations individually instead of disabling the category.

Microsoft's [Mac Power Query documentation](https://support.microsoft.com/en-us/excel/import-and-shape-data-in-excel-for-mac-power-query),
under "Author and transfer Power Query VBA code", explicitly documents Mac
support for `Queries`, `WorkbookQuery` and `Workbook.Queries`. It also supplies
a worksheet-loading example using the Mashup connection syntax and synchronous
`QueryTable.Refresh`. This is a concrete candidate for the repository-owned
package writer, not local execution evidence or proof of every Windows
destination. The page also
contains an outdated statement about editor availability; rely on its specific
API claims as hypotheses to verify against the installed Excel version.

### VBA parity workstream

Target **all six current actions**: `list`, `view`, `import`, `update`, `run` and
`delete`. Running an already-installed macro alone does not establish VBA parity.

Investigate access to the Mac VBA project model, component/procedure enumeration,
source reading, creation of standard modules, editing/deleting existing
components, and qualified `Module.Procedure` execution with parameters. Preserve
the existing `.xlsm` prerequisites, inline/file code inputs, timeout behavior,
result shape and explicit save semantics. Distinguish access to project source
from permission to execute macros; neither implies the other.

Use synthetic macro-enabled fixtures to verify import/view/run/update/run/delete,
module and procedure discovery, workbook-qualified execution with multiple
workbooks open, save/reopen persistence, missing modules/procedures, compile and
runtime errors, protected projects, unavailable project access, and timeout
handling that does not terminate shared Excel. Include parameter-sensitive
macros and observable worksheet changes so execution is verified, not inferred
from a successful dispatch.

If a helper/add-in is necessary, prove its user-approved installation, trust,
versioning and removal, plus operation on target workbooks without injecting
permanent helper code into them. No automatic macro-security changes or UI
clicking to bypass trust. Document any required user-managed Mac trust settings
and unsupported Windows-only VBA dependencies; report failures explicitly rather
than rewriting user code or claiming it ran successfully.

### Native investigation and progressive helper policy

The initial helper-free investigation established the secure native baseline
and the limits of Excel's Apple Events dictionary. The product direction now
permits optional helpers as later capability tiers. Installation and activation
must be explicit, independently detectable, versioned, removable, and reported
consistently by MCP and CLI. ExcelMcp must never alter macro security or VBA
project-model trust, click trust dialogs, or make a helper a prerequisite for
the native and saved-package tiers.

Read-only inspection of the installed Excel 16.112.3 scripting dictionary found:

- `run VB Macro` (`smXL2620`) accepts a macro/function reference and up to 30
  arguments. It does not accept arbitrary VBA source as a code-execution API.
- `has vb project` is a workbook boolean, not project/component/source access.
  No `VBProject`, `VBComponents`, `CodeModule`, or `do Visual Basic` surface was
  declared.
- `query table` exposes `connection`, `sql`, `destination`, `result range`,
  `refreshing`, and background-query settings. `refresh query table`
  (`smXLXrQT`) accepts `background query` and returns a boolean.
- `refresh all` (`smXL1831`) targets a workbook and also refreshes non-Power-Query
  external ranges and PivotTables. Its signature does not establish completion
  or per-query errors.
- No workbook `Queries` collection or `WorkbookQuery` M-formula object was
  declared. Legacy `QueryTable.sql` is not the query's M source.
- `evaluate` (`smXL2435`) evaluates Excel names/formulas. It is not an exposed
  general VBA interpreter or a Power Query M evaluator.

The dictionary was read using `/usr/bin/sdef '/Applications/Microsoft Excel.app'`
and XML inspection of command/class/property declarations. No workbook was
opened or modified, no macro executed, and no consent requested for this
inspection. Absence from the dictionary means no declared route was found, not
proof that every undocumented mechanism is impossible. Replacing JXA with
Swift/ScriptingBridge or direct .NET Apple Events does not itself add missing
Excel object-model endpoints.

The action-level native inventory is below. "No declared route" is a blocker
for the current helper-free backend, not a claim that Mac Excel lacks the
underlying feature.

| Existing action | Helper-free native candidate and evidence | Current status |
| --- | --- | --- |
| Power Query `list` | `[MS-QDEFF]` DataMashup parsing plus OOXML relationship traversal | Implemented for clean, saved workbooks without a Data Model; worksheet and connection-only state verified |
| Power Query `view` | `[MS-QDEFF]` DataMashup parsing with exact case-insensitive identity | Implemented for clean, saved workbooks without a Data Model |
| Power Query `get-load-config` | OOXML worksheet/table/QueryTable/connection traversal | Implemented for clean, saved workbooks without a Data Model; ambiguous worksheet destinations fail |
| Power Query `create` | No declared query-definition creation API | Blocked |
| Power Query `update` | Transactional saved-package exact formula replacement | Implemented only with `refresh=false`, for clean, saved workbooks without a Data Model; explicit save commits and discard restores the session baseline |
| Power Query `rename` | Renaming a QueryTable is not proven equivalent to renaming the query and preserving dependencies | Blocked |
| Power Query `delete` | Removing a destination is not equivalent to deleting the query definition and its exact destinations | Blocked |
| Power Query `evaluate` | No declared M execution API or temporary-query lifecycle | Blocked |
| Power Query `load-to` | QueryTable construction is a candidate only after an existing query can be identified; destination transitions and Data Model variants are unproven | Partial candidate, not executed |
| Power Query `unload` | Destination removal must preserve the query definition and remove only exact targets | Partial candidate, not executed |
| Power Query `refresh` | Exact connection identity plus synchronous `refresh query table` may work for an existing worksheet-loaded query | Promising candidate; requires real query success/failure fixtures |
| Power Query `refresh-all` | Native command has broader scope; connection-only coverage, completion and errors are unproven | Gated, not parity |
| VBA `list` | `has vb project` cannot enumerate components or procedures | Blocked |
| VBA `view` | No declared code-module read API | Blocked |
| VBA `import` | No declared component creation/source import API | Blocked |
| VBA `update` | No declared code-module write API | Blocked |
| VBA `run` | `run VB Macro` exists, but Excel displays a macro-trust prompt on every tested open | Capability-gated until execution is prompt-free and covered by real-entry-point tests |
| VBA `delete` | No declared component removal API | Blocked |

Other documented alternatives were checked, without installing or running them:

| Alternative | Finding against the required contracts |
| --- | --- |
| [Office.js Query](https://learn.microsoft.com/en-us/javascript/api/excel/excel.query?view=excel-js-preview) / [QueryCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.querycollection?view=excel-js-preview) | Metadata is available; the reviewed preview also has query deletion and collection refresh. No M source authoring or VBA project CRUD route is exposed by these query APIs. Preview deletion leaves associated tables disconnected rather than satisfying our full delete contract. The approved Office.js tier targets broader workbook surfaces, not replacement of the package-backed M lifecycle. |
| [Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/resources/vba-differences) / [ExcelScript.Query](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.query) | Office Scripts do support Mac, but the reviewed query surface provides metadata getters, not M authoring or VBA source operations. Documented invocation is user-started or through Power Automate, not a local external replacement for the current native backend. |
| [Power Automate refresh](https://learn.microsoft.com/en-us/office/dev/scripts/testing/power-automate-troubleshooting#refresh-not-fully-supported-in-power-automate) | Microsoft documents that `refreshAllDataConnections` refreshes only Power BI sources in flows and otherwise can return successfully without doing anything. This fails the required refresh success contract and is not local desktop Excel execution. |
| Saved OOXML/VBA package editing | May enable offline inspection or editing, but does not expose unsaved live query/project state, preserve current session semantics, or establish real-Excel evaluation/refresh. Not substituted for these contracts. |
| UI scripting or an in-process injected native library | No supported unattended object-model route established; UI automation needs additional permission and is fragile, and private injection would not be a reliable supported backend. Neither was attempted. |

### Helper-free saved-package investigation

The native inventory above is still accurate, but it is no longer the only
credible helper-free route. A clean-workbook transaction can combine published
file formats with real Excel execution:

1. Require the workbook to be saved and clean; never save user changes
   implicitly.
2. Close that exact workbook without saving while holding its session lock.
3. Patch a same-directory temporary copy and validate its ZIP/OPC structures.
4. Atomically replace the original, reopen the exact path through
   LaunchServices, and attach it to the existing logical session.
5. Restore the original if patching or reopening fails.

Dirty workbooks must fail with explicit save-or-discard guidance. The visible
close/reopen restriction is less capable than Windows live COM, but preserves
user data and avoids an in-Excel helper.

Power Query feasibility is now demonstrated beyond parser round-tripping:

- Microsoft's published `[MS-QDEFF]` specification describes the DataMashup
  root, Package Parts OPC archive, permissions, metadata and permission
  bindings. A repository-owned implementation can therefore be written from
  the specification.
- `Vladinator/excel-datamashup` successfully extracted and rewrote
  `Formulas/Section1.m`, but is GPL-3.0 and must remain research evidence unless
  an explicit licensing decision is made. Its code must not be copied into the
  repository implementation.
- An Excel-generated workbook was copied to a globally unique basename,
  patched from its file-source query to
  `#table(1, {{"ExcelMcp macOS parity probe"}})`, and reopened by current Mac
  Excel without a dialog.
- Native `Refresh All` then changed the existing worksheet-loaded table from
  the original vehicle dataset to a `Column1` table containing exactly
  `ExcelMcp macOS parity probe`. This proves that Excel accepted the rewritten
  M and executed it through the real Mashup engine.
- The probe preserved an existing query identity and load destination. It does
  not yet prove query creation, rename/delete dependency rewrites,
  connection-only queries, Data Model loads, temporary-query cleanup, source
  errors, or refresh completion across all connectors.

VBA package work reached a narrower result:

- `[MS-OVBA]` documents the project storage format. The MIT-licensed
  `Beakerboy/MS-OVBA` and `MS-Pcode-Assembler` projects passed all 106 tests
  after installing the assembler dependency omitted by MS-OVBA's declared
  dependencies.
- A generated `vbaProject.bin` preserved the exact
  `ParityProbe.WriteParityMarker` source when independently extracted with
  `olevba`. Mac Excel opened the resulting `.xlsm` and reported
  `has vb project = true`.
- The generated procedure did not execute: the workbook-qualified
  `run VB Macro` call returned a parameter error and left the marker unchanged.
  A known-good VBA fixture returned `42` through the same qualified invocation,
  so invocation syntax is not the cause. Source preservation and structural
  recognition are therefore insufficient; executable-cache/recompilation
  compatibility remains unresolved.
- Excel displayed its macro warning on every open, including reopening the same
  exact file after the user had enabled macros. Automated tests must never click
  that dialog or change macro security. Unattended VBA execution requires a
  separately approved user-managed trust prerequisite or another non-prompting
  trust design.

**Conclusion:** helper-free Power Query authoring is technically viable through
transactional saved-package editing and real Excel refresh, with the complete
contract still to implement and validate. VBA source parsing is viable, but
source mutation is not yet executable in Excel and remains blocked on a valid
recompilation/cache strategy plus unattended trust. No production capability
is enabled solely from these probes.

## Experiment

Environment: Apple Silicon/Arm64, Excel for Mac **16.112.3**, PowerShell 7,
AppleScript via `/usr/bin/osascript`. The repository SDK resolved to **10.0.401**
under `global.json`'s `10.0.302` / `latestFeature` policy. This is one machine and
one Excel version; Intel Macs and other Excel versions were not exercised.

The original feasibility experiments are isolated in:

- [PowerShell runner](../scripts/spikes/macos/Test-MacOsExcel.ps1)
- [Ordinary-file handoff](../scripts/spikes/macos/Test-MacFileHandoff.ps1)
- [AppleScript operations](../scripts/spikes/macos/ExcelSpike.applescript)

It does not invoke the existing MCP server, CLI, Service, or COM runtime. The
runner uses .NET `ProcessStartInfo.ArgumentList`, not shell/source interpolation.
AppleScriptObjC/Foundation serializes structured values to JSON. Workbook paths
are arguments and each command runs in a fresh host process.

### Reproduction and consent

Requires an interactive Mac desktop, PowerShell 7, licensed Excel already
running, and previously granted Automation access for the desktop host. The
default spike now uses ordinary temporary files and LaunchServices:

```powershell
pwsh -NoProfile -File scripts/spikes/macos/Test-MacOsExcel.ps1
```

This delegates to `Test-MacFileHandoff.ps1`. It does not read, create, inspect or
delete anything in Excel's container. Blank OOXML files are seeded solely as
synthetic input fixtures; all subsequent edits, calculation and persistence
use Excel. Files are opened using `/usr/bin/open -b com.microsoft.Excel`, not
the Excel `open workbook` command's bare text-path parameter.

A bounded native Automation permission check uses
`AEDeterminePermissionToAutomateTarget` with `askUserIfNeeded = false`; denied,
undecided, unavailable, and unexpected statuses fail explicitly before the test
creates fixtures. Excel-not-running was observed to fail with OSStatus -600
without dispatching any workbook operation. Tests do not request consent or
change privacy settings. Initial Automation onboarding on a fresh host remains
a separate manual prerequisite.

**Limits:** the evidence covers these synthetic files, this host, and this
Excel version. It does not establish behavior for protected directories,
untrusted downloaded files, external links, macros, a fresh host without
Automation permission, or permission revocation during a run. A preflight
from PowerShell is not proof of permissions for every production process.

`-AllowExcelContainerAccess` retains the older container experiment only for
explicit investigation. It is known to prompt repeatedly and must not be used
for unattended tests. The failed container setup receipt was invalidated and
the receipt-based initializer removed; a successful setup run is not evidence
that future container access is authorized.

### Managed versus native automation boundary

Keep .NET for MCP, CLI and shared contracts. A native language does not remove
sandboxing, Automation consent, or missing Excel APIs.

The bounded comparison can run without Excel commands or protected-directory
access:

```powershell
pwsh -NoProfile -File scripts/spikes/macos/Test-MacAutomationBoundary.ps1
```

It compiles a disposable Swift permission probe, compares it with direct .NET
interop and the JXA host, emits structured results, and removes its temporary
binary. It returns nonzero when a candidate cannot establish permission. Each
probe has a 15-second host deadline; the Swift compiler has a separate deadline.
This is a development spike, not a shipped signed/notarized helper.

On the current ARM64 Mac, .NET and Swift both returned OSStatus 0 without
requesting consent. The managed descriptor layout matched the SDK's 12-byte
packed `AEDesc`. The JXA probe could not complete the native check: passing the
Foundation descriptor produced `Ref has incompatible type`, and constructing
it through the C API produced `AECreateDesc` status -50. The report preserves
this as `InteropUnavailable`, not an inferred permission denial or success.
These observations do not prove JXA cannot support a different binding.

**Implemented decision:** each MCP/CLI automation call starts the same signed
entry-point executable in a bounded automation-host mode. That child performs
`AEDeterminePermissionToAutomateTarget` with prompting disabled and, only when
allowed, executes the embedded JXA through OSAKit in that same process. The
parent can terminate the child on timeout. This keeps permission checking and
Apple Event dispatch under one sender identity while retaining .NET contracts
and structured errors; `/usr/bin/osascript` is no longer the production sender.

### Separately launched native app: negative result

The user requested investigating a separate native app instead of an in-Excel
VBA helper. Its reproducible source lives in
`scripts/spikes/macos/native-helper/`, with `Build-MacPermissionHelper.ps1` and
`Invoke-MacPermissionHelper.ps1` as the build and LaunchServices runners.

The Swift/AppKit app had a stable bundle identifier, an ad-hoc development
signature, and its own Automation permission state. It executed the bundled
AppleScript in-process, bounded its lifetime to 120 seconds, and never quit
Excel. Setup passed, followed by two fresh app launches that passed workbook
assertions. **The user confirmed that "ExcelMcp Permission Probe wants to access
data from other apps" appeared again.** Therefore the helper did not solve
foreign-container permissions and is not the selected test architecture.

Do not interpret ad-hoc signing as production signing/notarization, or assume
a Developer ID signature would fix this without evidence. The app remains a
development experiment, not an installed production dependency or a permission
workaround.

### Ordinary-file handoff: prompt-free result

After the user approved this alternative, all four runs used fresh temporary
directories, new file identities, and fresh host processes. The user confirmed
no permission, file-access or repair dialogs for the first run, both repeats,
and the run after restarting Excel.

Excel also enforces workbook-name uniqueness across directories. Reusing a
fixture basename produced a modal "can't open two workbooks with the same name"
error and made subsequent Apple Events appear to fail. Test and production
handoff must reject a basename already open in Excel and use globally unique
synthetic fixture basenames; unique directories alone are insufficient. The
dialog must never be dismissed through UI automation.

| Run | Wall-clock duration | Result |
| --- | --- | --- |
| Initial ordinary-file trial | 9.717 seconds | All 12 checks passed; no dialogs |
| Fresh-process repeat 1 | 9.685 seconds | All 12 checks passed; no dialogs |
| Fresh-process repeat 2 | 9.455 seconds | All 12 checks passed; no dialogs |
| After user-managed Excel restart | 11.158 seconds | All 12 checks passed; no dialogs |

Each run verified every cell in a 1,000-by-10 matrix, formula calculation,
1x1/2D shapes, Unicode/escaping, number format, explicit workbook targeting,
discard, saved-state reopening, sentinel preservation and owned cleanup.
These are single-run observations, not performance guarantees.

After making handoff the default runner, a further regression run passed all
14 checks in 12.029 seconds, including the added explicit missing-worksheet
error and read-after-error assertions. The four user-confirmed no-dialog runs
above remain the permission-UX evidence.

The initial post-restart attempt found Excel not running and stopped in
preflight. After the user opened Excel, the same test passed. No automatic
launch or broader permission grant was substituted for the failed prerequisite.

**Scope remains explicit:** precreating a blank OOXML fixture is not a proof
of native `session.create`. The standalone runner does not invoke production
MCP or CLI; their separate integration evidence follows below. Power Query/VBA
parity remains incomplete: one external M-update/worksheet-refresh path is now
verified, while VBA source mutation and unattended trust remain unresolved.
Windows COM E2E was not run on this Mac.

The current implementation adds permission-result/layout regression coverage.
Successful runs close and delete only their synthetic fixtures. It never quits Excel,
changes global calculation/alert settings, reads customer workbook contents, or
changes macro security. A second synthetic workbook detects accidental
active-workbook targeting and unintended closure.

Each AppleScript operation has a 30-second Apple Events timeout; the host has a
45-second deadline and the handoff suite applies a 150-second operation budget.
A timeout stops further automation, kills only the host when
necessary, and retains fixtures for manual recovery. It does not assume that
stopping the host cancels Excel's outstanding action. Failed runs retain files
and report their private location locally, not in this document.

### Actual CLI/MCP integration

`tests/ExcelMcp.Portable.Tests/MacExcelE2ETests.cs` runs two opt-in workflows:
the real CLI apphost with a private daemon pipe, and the real MCP stdio server.
Both passed in approximately 31 seconds in the user-confirmed no-dialog run.
The standard runner subsequently passed both cases with zero skips and completed
its private-daemon cleanup:

```powershell
pwsh -NoProfile -File scripts/Test-E2E.ps1
```

On macOS this routes to `Test-MacE2E.ps1`, checks Automation without requesting
consent, cross-builds Release, and requires exactly two passing entry-point
cases. `-SkipBuild` is only appropriate after a successful Release build in the
same worktree. Fixtures live in ordinary temporary storage, never Excel's
container. The runner stops only its private daemon, not shared Excel.

Real entry-point tests exposed two host-lifetime defects that the standalone
spike could not catch: MCP attempted to start a Windows `kernel32` stdin monitor,
and the CLI daemon inherited output pipes that prevented a captured CLI command
from reaching EOF. The monitor is now Windows-only. Mac daemon startup uses
separate redirected streams and binds native and managed console output to its
per-pipe log before normal console initialization. The runner also derives
`DOTNET_ROOT` from the installed runtime, not Homebrew's executable directory or
PowerShell's private runtime.

The Release solution cross-build passed with zero warnings. All 18 selected
Excel-free portable tests passed, including the Mac path-error regression.
The existing COM-leak, Core coverage/naming, MCP implementation, success-flag,
documentation-count and dynamic-cast checks passed. These checks and the Mac
workflows do not substitute for Windows COM E2E, which remains unrun here.

### Earlier container-based measurements

The completed sandbox run reported 14 checks and verified every cell of a
1,000-row by 10-column numeric matrix. Times below are single-run wall-clock
observations including host startup, scripting, Excel, and serialization, not
benchmarks or performance guarantees.

| Check | Observed result |
| --- | --- |
| Create and save `.xlsx` | Passed; two fixtures, 828 and 877 ms |
| Explicit targeting | Main workbook read correctly while the second workbook was open |
| Mixed 3x2 values | Strings and numbers retained their matrix shape and values |
| Single-cell formula | `=SUM(B2:B3)` returned a 1x1 matrix containing 30 |
| Text serialization | Quotes, backslash, newline, and a Unicode character round-tripped |
| Number format | `0.00` survived save/reopen |
| Missing worksheet | Explicit existence guard produced a nonzero error, code 9006 |
| Read after validation failure | Passed without losing the workbook |
| Bulk write/read/calculate | All 10,000 cells matched; final-column sum 5,005,000; 822 ms |
| Discard | Unsaved change to 999 did not survive reopening |
| Reopen saved workbook | Passed; open took 848 ms |
| Close/cleanup | Owned fixtures closed; second workbook remained usable until its own cleanup |
| Excel file grants | None in the successful container run, confirmed by user |
| macOS cross-app data grants | Subsequently reported by user; not eliminated |

Exploratory failures also mattered:

- HFS-style save paths produced Excel parameter error `-50`; POSIX save/open
  paths worked on this Excel build.
- Direct collection iteration produced a parameter error; indexed workbook
  lookup worked. Lookup compares full paths, not the active workbook.
- A direct missing-worksheet range access did not throw in the probe. The
  adapter must validate existence explicitly rather than treating any Apple
  Events response as success.
- An early bulk scalar result was not usable; after renaming the AppleScript
  scalar variable, the full numeric assertion passed. Do not generalize this
  single observation into a calculation-completion guarantee.
- External-directory experimentation was interrupted by file-access dialogs.
  A pending open and retained lock file exposed why error cleanup must not
  assume that a workbook operation has completed.

Failed/interrupted exploratory runs may retain synthetic files. Further
cross-app inspection/cleanup was stopped on the user's permission report.
No workbook artifacts or private filesystem paths are included in the repository.

### What this does not establish

The standalone spike does not establish production behavior. The later CLI/MCP
tests add real host and daemon IPC evidence for the selected workflows, but not
process isolation, cross-client concurrency, timeout recovery,
denied-permission recovery, arbitrary-file
permissions, signing/notarization, large-data scaling, or Intel compatibility.
No coverage of Excel error-cell mapping, dates, booleans, blanks/nulls, merged
cells, protection, locale variations, external links, general existing macros,
Power Query beyond the single package-update/worksheet-refresh probe, charts,
tables, or PivotTables. Those require additional fixtures and assertions before
advertising support.

The initial Release solution build was attempted and failed: Windows cleanup
uses `powershell`, fresh-worktree assets were missing, and Windows Desktop
projects raised `NETSDK1100`. A subsequent normal restore independently failed
with `NETSDK1100`. Cross-compiling Windows binaries would not supply a macOS
backend. The later portable implementation and cross-build resolved those build
blockers, and the standard E2E runner now routes to actual Mac entry-point tests.
Windows COM tests remain **not run**: Mac Excel is not a replacement for Windows
COM Excel.

## Architecture findings

## Specialized feature evidence matrix

This inventory was refreshed against Microsoft Excel for Mac **16.113.1** on
Apple Silicon, the installed scripting dictionary from
`/usr/bin/sdef '/Applications/Microsoft Excel.app'`, and Microsoft's Office.js
reference updated in September 2026. Excel 16.113.1 is new enough for the
documented stable `ExcelApi 1.21` floor (Mac 16.110.1), but host support for a
requirement set is not contract evidence by itself.

Status meanings:

- **Enabled**: prompt-free real Excel passed through both `excelcli` and MCP,
  including returned fields and workbook effects.
- **Candidate**: an API route exists or one entry point passed, but complete
  cross-entry-point result, completion, error, cleanup, and persistence
  semantics are not yet proven.
- **Blocked**: the inspected native and Office.js surfaces do not expose the
  contract, or the only route violates an existing contract requirement.

The serialized real-Excel validation slot is shared with the other macOS
workstreams. Candidate operations stay gated until they receive that slot.

### Connections and legacy QueryTables

| Action | Tier | Status | Evidence and blocker / user impact |
| --- | --- | --- | --- |
| `connection.list` | Apple Events | Candidate | `workbook connection` is declared, but safe enumeration and sanitized type metadata are unverified. Users must use Windows for truthful connection inventory. |
| `connection.view` | Apple Events | Candidate | The dictionary does not expose the complete typed OLEDB/ODBC property set; credential-safe parity is unproven. |
| `connection.create` | Apple Events | Candidate | Provider availability, command type, and `Connections.Add2` parity are unproven on Mac. No connection is created by the gated command. |
| `connection.refresh` | Apple Events | Candidate | A broad refresh command exists, but exact-connection completion and source errors are not established. |
| `connection.get-refresh-status` | Apple Events | Candidate | Typed OLEDB/ODBC status equivalent is not declared. |
| `connection.cancel-refresh` | Apple Events | Candidate | Cancellation is declared for QueryTables, not proven for exact typed workbook connections. |
| `connection.delete` | Apple Events | Candidate | Exact ownership cleanup of only the target connection's QueryTables is unproven. |
| `connection.load-to` | Apple Events | Candidate | Exact destination replacement and connection ownership semantics are unproven. |
| `connection.get-properties` | Apple Events | Candidate | Complete typed settings and sanitized connection reporting are unverified. |
| `connection.set-properties` | Apple Events | Candidate | Typed OLEDB/ODBC settings and password handling are unverified. |
| `connection.test` | Apple Events | Candidate | Configuration-only validation cannot be inferred from refresh acceptance. |
| `querytable.list` | Apple Events | Candidate | Worksheet `query table` elements and destination/refreshing properties are declared; a prompt-free fixture is still required. |
| `querytable.view` | Apple Events | Candidate | Most legacy text/web properties are declared, but source classification and sanitized connection output are not yet round-tripped. |
| `querytable.create-text` | Apple Events | Candidate | `TEXT;` construction is plausible; encoding, qualifier, delimiter, synchronous refresh, and local-file error behavior need a fixture. |
| `querytable.create-web` | Apple Events | Candidate | `URL;` properties are declared, but unattended network completion and HTML selection/error behavior are unproven. |
| `querytable.set-properties` | Apple Events | Candidate | Properties are declared; exact persistence and invalid-value behavior remain untested. |
| `querytable.refresh` | Apple Events | Candidate | `refresh query table` returns a boolean, but source errors and synchronous completion need proof. |
| `querytable.get-refresh-status` | Apple Events | Candidate | `refreshing` is declared; live background-refresh evidence is missing. |
| `querytable.cancel-refresh` | Apple Events | Candidate | `cancel refresh` is declared; idle and active result DTO semantics are untested. |
| `querytable.delete` | Apple Events | Candidate | Exact-name deletion and preservation of result cells/related objects need proof. |

### Data Model, relationships, and DAX

The installed dictionary has no `Model`, model-table, relationship, measure,
DAX evaluate, DMV, or embedded ADO model surface. Stable Office.js through
`ExcelApi 1.21` does not provide the Windows contract's Data Model CRUD and DAX
execution APIs. Package inspection cannot truthfully report unsaved model state
or execute DAX. Every action below is therefore gated as evidenced unsupported
for the current native and Office.js tiers.

| Action | Tier | Status | User impact / blocker |
| --- | --- | --- | --- |
| `datamodel.list-tables` | Evidenced unsupported | Blocked | No live model table collection. |
| `datamodel.list-columns` | Evidenced unsupported | Blocked | No live model column collection. |
| `datamodel.read-table` | Evidenced unsupported | Blocked | No complete live table/measure metadata route. |
| `datamodel.read-info` | Evidenced unsupported | Blocked | Model counts and state cannot be guessed from package parts. |
| `datamodel.read-connection` | Evidenced unsupported | Blocked | No credential-safe embedded model connection endpoint. |
| `datamodel.list-measures` | Evidenced unsupported | Blocked | No measure collection or DAX formula access. |
| `datamodel.read` | Evidenced unsupported | Blocked | No exact measure identity/formula endpoint. |
| `datamodel.create-measure` | Evidenced unsupported | Blocked | No measure creation or format API. |
| `datamodel.update-measure` | Evidenced unsupported | Blocked | No measure formula/description/format mutation API. |
| `datamodel.delete-measure` | Evidenced unsupported | Blocked | No exact measure deletion API. |
| `datamodel.delete-table` | Evidenced unsupported | Blocked | No exact model-table deletion API. |
| `datamodel.rename-table` | Evidenced unsupported | Blocked | Renaming a query or worksheet table is not model-table parity. |
| `datamodel.refresh` | Evidenced unsupported | Blocked | Broad workbook refresh cannot establish model completion or errors. |
| `datamodel.evaluate` | Evidenced unsupported | Blocked | No DAX execution endpoint. |
| `datamodel.execute-dmv` | Evidenced unsupported | Blocked | No embedded ADOMD/DMV endpoint. |
| `datamodelrel.list-relationships` | Evidenced unsupported | Blocked | No relationship collection. |
| `datamodelrel.read-relationship` | Evidenced unsupported | Blocked | No exact four-part relationship lookup. |
| `datamodelrel.create-relationship` | Evidenced unsupported | Blocked | No relationship creation API. |
| `datamodelrel.update-relationship` | Evidenced unsupported | Blocked | No active-state mutation API. |
| `datamodelrel.delete-relationship` | Evidenced unsupported | Blocked | No exact relationship deletion API. |

### What-if analysis

| Action | Tier | Status | Evidence and blocker / user impact |
| --- | --- | --- | --- |
| `analysis.goal-seek` | Apple Events | **Enabled** | Excel 16.113.1 passed prompt-free CLI and MCP fixtures. Both returned `converged`, the approximate final formula value, and the changing value; the workbook was targeted by exact path and was not implicitly saved. |
| `analysis.list-scenarios` | Apple Events | Candidate | Scenario elements/properties are declared, but value ordering, comment prefix, and protection metadata need a real fixture. |
| `analysis.create-scenario` | Apple Events | Candidate | A scenario class exists, but construction, cell/value count validation, and protection flags are unproven. |
| `analysis.update-scenario` | Apple Events | Candidate | `change scenario` is declared; exact value conversion and failure behavior are untested. |
| `analysis.show-scenario` | Apple Events | Candidate | Applying stored values without invoking a dialog needs proof. |
| `analysis.delete-scenario` | Apple Events | Candidate | Exact-name deletion and missing-name errors need proof. |
| `analysis.create-scenario-summary` | Apple Events | Candidate | `create summary for scenarios` is declared; generated-sheet identity and PivotTable variant are untested. |
| `analysis.create-data-table` | Apple Events | **Enabled** | Excel 16.113.1 passed prompt-free CLI and MCP fixtures for a one-variable table with exact `[1, 4, 9]` results. The same native command accepts the contract's optional row input, optional column input, or both; calls without either input remain rejected. |

### Drawings, sparklines, slicers, and screenshots

| Action | Tier | Status | Evidence and blocker / user impact |
| --- | --- | --- | --- |
| `drawing.list-objects` | Office.js add-in | Candidate | Office.js can enumerate shapes, but the optional bridge is not installed and Forms-control parity is incomplete. |
| `drawing.get-object` | Office.js add-in | Candidate | Geometry/text/format/accessibility DTO parity needs the bridge and fixtures. |
| `drawing.add-image` | Office.js add-in | Candidate | Base64 image insertion exists; local-file handling and exact dimensions need bridge validation. |
| `drawing.add-shape` | Office.js add-in | Candidate | Shape creation exists; the contract's type mapping and formatting need validation. |
| `drawing.add-text-box` | Office.js add-in | Candidate | Text shapes exist; exact font/fill/line semantics need validation. |
| `drawing.add-connector` | Office.js add-in | Candidate | Connector coverage and endpoint semantics are incomplete. |
| `drawing.add-form-control` | Office.js add-in | Blocked | Stable Office.js does not provide parity for the contract's safe worksheet Forms controls and bindings. |
| `drawing.update-object` | Office.js add-in | Candidate | Common shape mutations exist; Forms bindings and all-or-error updates remain unproven. |
| `drawing.delete-object` | Office.js add-in | Candidate | Exact identity and missing-object errors need bridge validation. |
| `drawing.list-sparklines` | Office.js add-in | Candidate | Sparkline groups are available through Office.js; the optional bridge is not installed. |
| `drawing.get-sparkline` | Office.js add-in | Candidate | Location/source/type/color DTO parity needs bridge validation. |
| `drawing.add-sparkline` | Office.js add-in | Candidate | Creation exists; line/column/win-loss mapping and marker behavior need fixtures. |
| `drawing.update-sparkline` | Office.js add-in | Candidate | Source/type/style mutation needs exact round-trip evidence. |
| `drawing.delete-sparkline` | Office.js add-in | Candidate | Exact group deletion needs bridge validation. |
| `slicer.create-slicer` | Office.js `ExcelApi 1.10` add-in | Candidate | Slicer creation exists, but PivotTable-field identity and placement need the optional bridge. |
| `slicer.list-slicers` | Office.js `ExcelApi 1.10` add-in | Candidate | Item/selection and PivotTable filtering DTOs need validation. |
| `slicer.set-slicer-selection` | Office.js `ExcelApi 1.10` add-in | Candidate | `clearFirst` union semantics need real fixtures. |
| `slicer.delete-slicer` | Office.js `ExcelApi 1.10` add-in | Candidate | Exact slicer identity and missing-name behavior need validation. |
| `slicer.create-table-slicer` | Office.js `ExcelApi 1.10` add-in | Candidate | Table-column source and placement need bridge validation. |
| `slicer.list-table-slicers` | Office.js `ExcelApi 1.10` add-in | Candidate | Table-only classification needs validation. |
| `slicer.set-table-slicer-selection` | Office.js `ExcelApi 1.10` add-in | Candidate | Table filtering and `clearFirst` behavior need validation. |
| `slicer.delete-table-slicer` | Office.js `ExcelApi 1.10` add-in | Candidate | Exact table-slicer deletion needs validation. |
| `screenshot.capture` | Optional native helper | Blocked | Apple Events only declares clipboard-based `copy picture`, which violates the no-clipboard live-window contract. A helper would require explicit Screen Recording permission and crop/stitch evidence. |
| `screenshot.capture-sheet` | Optional native helper | Blocked | Same blocker; Office.js has no API that photographs the live Excel window with all visuals. |

### XML maps and Python in Excel

| Action | Tier | Status | Evidence and blocker / user impact |
| --- | --- | --- | --- |
| `xmlmap.list` | Evidenced unsupported | Blocked | Neither the installed dictionary nor stable Office.js exposes Excel XML maps. |
| `xmlmap.add` | Evidenced unsupported | Blocked | Custom XML parts are not worksheet XML-map creation parity. |
| `xmlmap.map-range` | Evidenced unsupported | Blocked | No XPath-to-cell mapping API. |
| `xmlmap.import-xml` | Evidenced unsupported | Blocked | Generic XML parsing/package edits cannot invoke Excel's XML-map import semantics. |
| `xmlmap.export-xml` | Evidenced unsupported | Blocked | No live mapped-cell export endpoint. |
| `xmlmap.delete` | Evidenced unsupported | Blocked | No exact map deletion endpoint. |
| `pythoninexcel.set-formula` | Apple Events | Candidate | Range `formula2` is declared, but licensed `PY()` availability, return type, immediate `#NAME?`, and cloud error behavior need a prompt-free account fixture. |
| `pythoninexcel.get-result` | Apple Events | Candidate | The dictionary exposes `formula2` and values but no calculation-state endpoint proven equivalent to the Windows completion guard; returning `#BUSY!` would violate the contract. |

Changing target frameworks or replacing COM activation alone is insufficient.

| Boundary | Current coupling / required work |
| --- | --- |
| `ComInterop/Session/IExcelBatch.cs` | Exposes `Excel.Workbook`, `ExcelContext`, and arbitrary COM callbacks; not a portable backend interface |
| Core commands | Implement behavior directly against COM; e.g. range values use `Value2`, 1-based COM arrays, error mapping, merged-cell guards |
| `[ServiceCategory]` interfaces | Reusable operation definitions, but take COM batches; extract neutral session contracts while preserving parameters/defaults/results |
| `Service/ExcelMcpService.cs` | Owns the session manager and concrete command implementations; needs backend selection rather than AppleScript in transport handlers |
| Both `ServiceSecurity.cs` implementations | Windows SID-based identity; server uses Windows pipe ACLs; client and server must change together |
| CLI | `net10.0-windows`, Windows Forms tray, Windows process cleanup |
| MCP server | Windows target and `kernel32` calls; host lifetime needs platform-specific handling |
| ComInterop / window / screenshots | STA/OLE, COM activation/release, process guards, Win32 capture and window APIs remain Windows backend concerns |
| Build/distribution | Windows cleanup commands, executable names, runtime packaging, extension launcher, plugin bootstraps and generated platform guidance need coordinated changes |

The current [context](../CONTEXT.md) promises one owned Excel process per session,
independent sessions, and separate MCP/CLI ownership. The macOS backend instead
targets the desktop Excel application and multiple workbooks inside it. Carrying the
Windows ownership/kill model onto macOS would risk user workbooks.

**Proposed intentional difference requiring a decision:** serialize macOS
automation across participating MCP/CLI clients through a per-user broker, with
logical workbook sessions and explicit ownership. Preserve separate public
session namespaces; do not imply that the existing entry points already share
sessions. Own a workbook, not the user's Excel process. Never kill shared Excel
to cancel one session. Specify how a stuck application affects all clients.

The [testing ADR](../docs/ADR-001-NO-UNIT-TESTS.md) is superseded; the current
policy requires real Excel for Excel behavior and permits focused non-Excel
tests for parsing/dispatch. This macOS release does not change that policy.

## Backend options and permission policy

| Option | Assessment |
| --- | --- |
| Native Apple Events / AppleScript | Demonstrated desktop workbook transport; evaluate together with deeper object-model access, rather than letting dictionary coverage define the product scope |
| Transactional saved-package editing plus Apple Events | Demonstrated for an existing Power Query M update and worksheet refresh; preferred helper-free direction for clean workbooks. Full metadata mutation, rollback and VBA recompilation remain to implement |
| Optional macro helper invoked from AppleScript | Approved as a later opt-in tier for known macro execution only; requires explicit macro enablement and unattended preflight |
| Optional VBA project-model helper | Approved only as a separate, explicit-trust tier for source CRUD; must never change the trust setting itself |
| Office.js add-in + local bridge | Approved as a gradual optional tier for tables, charts, PivotTables, conditional formatting, and related workbook surfaces; requires deployment, lifetime, and per-version API checks |
| Windows Excel behind a remote service | Could preserve more existing behavior, but not native macOS support; introduces remote data handling/security and is outside this release |
| File-only library without Excel execution | Does not meet the repository's real-Excel calculation/refresh requirement |

The ordinary-file LaunchServices path eliminated prompts for the tested
non-macro workflows. Macro-enabled files still displayed Excel's macro warning
on every open, including the same exact file after a prior explicit enable.

For production, distinguish two file workflows:

1. **Direct user workbooks:** explicit user-mediated access and a documented
   permission recovery path. Microsoft documents stored per-file grants and
   `GrantAccessToMultipleFiles` for VBA; this is not blanket permission for every
   future path, nor a verified AppleScript grant-management implementation.
2. **Managed working copies:** an explicit, opt-in import/export workflow, if
   its host permissions and lifecycle prove acceptable. Never silently stage a
   workbook: relative links, query paths, identities, concurrency and save
   semantics can change. Retain the original unless export is authorized.

A signed, narrowly scoped helper may improve attribution and onboarding, but
the feasibility work did not demonstrate that signing removes either permission gate.
Do not recommend disabling sandboxing, broad Full Disk Access, UI-click
automation, or changing macro trust as a workaround.

## Proposed delivery plan

Each phase is gated by evidence. Native operations and clean-workbook Power
Query `list`/`view` ship independently of optional helpers. Power Query
mutations, Office.js features, macro execution, and VBA source access activate
only when their own security and lifecycle gates are satisfied.

| Phase | Work | Exit gate |
| --- | --- | --- |
| 0. Permission and ownership design | Choose direct-workbook versus explicit working-copy UX; test denied/revoked access and repeated launches under the intended host identity; decide broker/ownership semantics | User-approved onboarding, bounded failures, no recurring unexpected prompts for already-authorized files, no unauthorized copying or user-workbook closure |
| 1. Native foundation | Shared capability metadata; secure IPC/ownership; workbook/sheet/range/calculation foundation | Implemented native slice passes actual MCP and CLI workflows without prompts |
| 2. Saved-package Power Query | Independent `[MS-QDEFF]` reader/writer; exact query identity; worksheet load graph; clean-workbook safety | `list`/`view`/`get-load-config` and package-only `update(refresh=false)` implemented with backup/reopen/rollback; next gate is repository-safe unattended refresh evidence |
| 3. Office.js workbook tier | Deploy and activate an optional add-in/local bridge for tables, charts, PivotTables, conditional formatting, and related surfaces | Explicit install/removal and capability preflight; cross-entry-point real-Excel workflows pass |
| 4. Macro execution tier | Add an optional, narrowly scoped macro helper without changing security settings | Known macros run unattended only when the user has enabled macros; prompts and unavailable trust fail before execution |
| 5. VBA project-model tier | Add source CRUD only behind explicit user-managed project-model trust | Full synthetic import/view/run/update/delete lifecycle passes without trust mutation or dialog automation |
| 6. Remaining Windows surface and distribution | Complete the action inventory, installers/bootstrap scripts, signing/notarization, and separately validated architectures | No unclassified omissions; clean-machine permission and installation UX exercised |
| 4. Remaining Windows surface | Work through every remaining category, including tables, named ranges, charts, regular PivotTables, formatting and platform integrations | Complete action-level compatibility inventory, tested parity where possible, and no unclassified omissions |
| 5. Distribution | Mac launch/lifetime, extension paths, installers/bootstrap scripts, `osx-arm64` and separately validated `osx-x64`, signing/notarization, shared docs and skills | Clean-machine install and permission UX exercised; priority parity gates, Mac real-Excel tests and Windows nonregression suites pass |

Investigate Data Model/DAX/OLAP/Power Pivot operations and Power Query
`data-model`/`both` destinations as potential host limitations; do not infer
support from ordinary PivotTables. Until verified, keep those variants gated
rather than advertising them or silently changing their destinations.
Document evidence and seek a scope decision for confirmed limitations.
Power Query exists on Mac and Microsoft's documentation describes VBA query
authoring; its creation, refresh, authentication, connectors and destinations
are priority investigation items, not grounds for excluding Power Query.

Capability checks must be shared by the Service, MCP, CLI and generated guidance,
not independently maintained allowlists. Decide whether unavailable actions are
hidden or discoverable with explicit unsupported results; either way, both
entry points must agree and must never return `Success == true` with an error.

Release readiness requires both Excel-free tests for the portable host and
real-Excel tests on an interactive Mac. GitHub-hosted runners without Excel can
validate build/generation but cannot establish Excel behavior. Existing Windows
COM coverage must remain green. Run matching synthetic workflows through MCP
and CLI on both platforms, comparing public results and saved workbook behavior;
allow only documented platform-specific differences, not broad snapshot
normalization that hides missing functionality. Power Query and VBA acceptance must cover the workflows above, not just
category discovery or happy-path dispatch. The initial capability-gated macOS
release has a changeset and advertises only the verified operation subset;
future parity additions require their own user-visible changesets.

## Sources

- [Office for Mac VBA and sandboxing](https://learn.microsoft.com/en-us/office/vba/api/overview/office-mac)
- [Request access to multiple files; grants stored with the app](https://learn.microsoft.com/en-us/office/vba/office-mac/grantaccesstomultiplefiles)
- [Mac Power Query, including VBA query authoring](https://support.microsoft.com/en-us/excel/import-and-shape-data-in-excel-for-mac-power-query)
- [Excel analytics platform differences](https://support.microsoft.com/en-us/excel/learn-to-use-power-query-and-power-pivot-in-excel)
- [Office.js Excel API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets)
- Installed Excel scripting dictionary, `Microsoft Excel.app/Contents/Resources/Excel.sdef`.

Microsoft's Mac Power Query page contains a legacy sentence saying the editor is
unavailable alongside newer editor instructions. Treat that as documentation
inconsistency, not evidence that current Mac Excel lacks the editor. Public
feature availability is also not evidence of automation parity.
