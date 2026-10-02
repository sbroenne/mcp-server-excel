# macOS support

**Status: Experimental beta.** Apple Silicon macOS support is a
capability-gated subset, not full Windows parity or a production-support
commitment. Test on copies of important workbooks. Windows retains the complete
COM backend and is not reclassified as experimental. macOS combines a native
Apple Events backend with optional Office.js and ScreenCaptureKit tiers.
Unsupported or unproven actions fail with `PlatformNotSupported`; they never
return success-shaped approximations.

For every enabled action, macOS and Windows work 1:1 at the public contract:
identical inputs, defaults, validation, results, workbook effects, persistence,
and errors. macOS drives Excel through a different supported API and follows
shared-application ownership rules, but those implementation differences do
not permit reduced or approximate behavior. If exact parity is unavailable,
the action or explicitly declared variant remains gated.

## Not supported in the macOS beta

These limits apply equally to the MCP Server and `excelcli`, regardless of
whether they are installed through npm, ZIP, VSIX, MCPB, plugins, or skills.
Use the Windows backend when a workflow requires an unavailable action.

| Feature | macOS beta limitation |
|---------|-----------------------|
| Power Query | All actions: M authoring/inspection, load destinations, and refresh. |
| VBA | All source/module actions and macro execution; no VBA helper or automatic trust setup. |
| Data Model, DAX, and OLAP | Model tables/relationships, measures, DAX queries/DMVs, and OLAP/Data Model PivotTables. |
| Tables, PivotTables, charts, and slicers | All public actions remain gated, including ordinary worksheet Tables and PivotTables. Optional Office.js handlers do not make them supported. |
| Connections and QueryTables | All public actions, including imports, refresh, and connection management. Native candidates are not enabled. |
| Advanced visuals and worksheet features | Conditional formatting, rich font/fill/border styling, validation, comments, drawings, shapes, sparklines, outlines, protection, page setup, and window/Agent Mode control. Basic number formats, sizing, cell locking, tab color, and visibility are enabled. |
| Screenshots and XML Maps | All public actions remain gated; helper/API presence is not accepted end-to-end support. |
| Python in Excel | Result reads are unsupported. Licensed `PY()` formula writes are enabled; this does not imply result-read support. |
| Range and worksheet variants | UsedRange, CurrentRegion, merge-area inspection, writes to merged cells (including the Windows top-left exception), find/replace/sort, disjoint row/column editing, worksheet copy/move, and cross-workbook selection via worksheet `filePath`. Cell/row/column insertion/deletion, value/formula copies, merge/unmerge, and the accepted exact-workbook lifecycle variants are enabled. |
| Other What-If and calculation actions | All Scenario actions and application-global calculation-mode get/set. Goal Seek, one-/two-variable Data Tables, and explicit calculation are enabled. |
| File/platform variants | New `.xlsm` creation, advanced workbook save/export operations, Intel Macs, Linux, and headless servers. `.xlsx` creation uses an intact Excel-authored template. |

The [generated action inventory](../docs/MACOS-ACTION-INVENTORY.md) is the
authoritative per-action list. `Partial` means unavailable in this beta, not
partially usable; only actions marked enabled can be dispatched. Installing the
optional bridge does not bypass that gate. No workbook-package inspection,
dialog automation, security changes, or alternative workflows approximate
unsupported operations.

## Non-negotiable workbook boundary

Excel workbook files are opaque.

- Never create, parse, inspect, or mutate workbook ZIP, OOXML, relationship,
  custom XML, or DataMashup internals.
- This prohibition applies to production code, tests, fixtures, scripts, and
  indirect use through package or Open XML libraries.
- An intact workbook authored by Excel may be copied as an opaque whole file.
- Workbook content changes must use supported Excel object models.
- If neither route can satisfy the public contract, report the operation as
  unsupported.

`scripts/check-workbook-package-access.ps1` enforces this boundary in
pre-commit validation.

## Runtime architecture

- MCP Server and `excelcli` share the same Service contracts, validation,
  routing, and result models.
- Windows uses COM and keeps its existing session and process-ownership model.
- macOS uses serialized JXA/Apple Events requests against exact workbook paths.
- CLI IPC uses current-user-only Unix named pipes and a stable, hashed
  per-user identity.
- The Mac daemon never terminates shared Excel and never closes a workbook it
  cannot identify as session-owned.
- Existing-file open uses LaunchServices and then attaches only to the exact
  path under one deadline.
- Once an open handoff has started, an invalid response, timeout, or uncertain
  reconciliation returns `RecoveryRequired`. ExcelMcp does not retry, delete,
  roll back, or claim that the workbook remained closed.
- Unconfirmed handoffs retain an exact-path recovery session in the current
  client. Session inventory reports `requiresRecovery: true` and
  `canClose: false`; operations, repeated open/create, and automatic cleanup
  cannot touch it. Reconcile the workbook and pending dialogs before restarting
  the client.
- A timed-out native mutation or uncertain Office.js mutation also makes the
  exact session recovery-only with `canClose: false`. ExcelMcp never races the
  in-flight operation with save, close, rollback, or shutdown cleanup and never
  discards earlier valid edits. Reconcile the workbook manually, then restart
  the client.
- Existing workbook identity is canonical across symlinks and path aliases.
  New-workbook destinations canonicalize the existing parent before the intact
  Excel-authored template is copied, so one workbook cannot be claimed through
  multiple path spellings.
- Worksheet list/create accept only the session workbook path on macOS.
  Selecting another workbook by `filePath` fails explicitly; use its own session.
- Cold Excel startup polls only non-mutating preflight under the existing
  deadline. Permission errors fail immediately; the workbook handoff happens
  only after preflight succeeds.
- Normal shutdown saves confirmed session-owned workbooks before closing them;
  unconfirmed handoffs remain untouched. Idle shutdown counts Mac sessions, so
  an open workbook prevents automatic idle exit. Concurrent close requests
  with different save choices are rejected instead of sharing a success result.

Apple Silicon is exercised locally. Intel macOS is unsupported; no Intel
artifacts or execution requirement exist.

## Workbook creation

`.xlsx` creation copies the bundled `Mac/Templates/Blank.xlsx`, an intact
Excel-authored workbook, to a same-directory temporary file and atomically
moves it into place. The template is never inspected or altered as a package.
An existing destination is never overwritten.

`.xlsm` creation is explicitly unsupported on macOS until the repository has an
authentic Excel-authored macro-enabled template whose provenance and runtime
behavior are accepted. Renaming `.xlsx` bytes or synthesizing macro project
state is not permitted.

`file test` checks the absolute path, extension, existence, size, timestamp,
lock state, read access, and the existing IRM/AIP preflight. `preflightPassed`
reports those access checks, not structural validity or openability. It does
not inspect workbook structure and therefore leaves `isValid=false` and
`canOpen=false`; `file(action: 'open')` or CLI `session open` asks Excel to
validate the file.

## Enabled native beta capabilities

The generated action inventory is authoritative:

- [Human-readable inventory](../docs/MACOS-ACTION-INVENTORY.md)
- [Machine-readable inventory](../docs/generated/macos-action-inventory.json)

`MacCommandCapabilities` consumes the same generated records. A method becomes
available only after exact public CLI and MCP acceptance updates its Core
capability annotation.

Each generated record also carries the remaining-work execution plan:

- `plannedTier`: the tier to prove next, or `MacLimitation` when the reviewed
  supported APIs cannot preserve the exact public contract;
- `acceptanceFixture`: the required Excel-authored, opaque test asset;
- `acceptanceCommand`: the guarded public CLI/MCP runner or the action-specific
  extension that must be added;
- `recoveryRule`: the ownership and uncertain-outcome behavior for that proof;
- `evidenceCriteria`: the minimum result, effect, failure, cleanup, and
  persistence evidence needed to enable or limitation-classify the action.

The human-readable inventory derives enabled, gated, status, and planned-tier
counts from those generated records. These fields organize work; they do not
enable an action or replace real Excel evidence.

Verified native coverage includes:

- platform-independent diagnostic ping, echo, and parameter validation without
  starting or dispatching to Excel;
- exact workbook create/open/close and owned-session cleanup;
- worksheet list/create/rename/delete;
- worksheet visibility and tab color;
- range values, formulas, number formats, clear operations, explicit row and
  column sizing, copy variants, bounded metadata, auto-fit, merge/unmerge, and
  cell locking for the accepted variants;
- calculation, Goal Seek, one- and two-variable Data Tables;
- native cell insertion/deletion with explicit Down/Right or Up/Left shifts,
  and single-area entire-row/column insertion/deletion, including formula
  adjustment and save/reopen persistence; disjoint row/column selections fail
  before mutation rather than editing only their first area;
- licensed Python in Excel formula writes through `Range.Formula2`;
- the accepted named-range lifecycle variants;
- screenshot identity infrastructure remains experimental and does not enable
  either public screenshot action.

Known native Apple Events gaps remain gated, including UsedRange,
CurrentRegion, merge-area inspection, worksheet copy/move, application-global
calculation-mode mutation in shared Excel, and any action whose declared
dictionary surface did not survive real CLI/MCP acceptance.

Python in Excel result reads remain gated because Excel 16.113.2 rejected the
native Apple Events read through the public MCP entry point after accepting the
formula write. Windows remains the supported route for `get-result`.

Excel 16.113.1 defines a thin read-only `workbook connection` class but exposes
no workbook collection, creation command, or typed OLEDB/ODBC properties needed
by the public connection contract. It exposes QueryTable elements and
properties, but no construction command for text or web sources. Connection
actions and QueryTable creation are explicit `MacLimitation` plans.
QueryTable view/set-properties are also limitations because the dictionary
omits fields required by their public contracts. List, refresh,
refresh-status, cancel, and delete remain native candidates for an original
Excel-authored fixture.

## Power Query and VBA limitations

Power Query lifecycle operations are unsupported on macOS. Excel 16.113.1
Apple Events exposes no `Workbook.Queries` surface, and Office.js exposes no
equivalent Power Query authoring, M inspection, load-state, or refresh API.
ExcelMcp does not inspect `DataMashup` or any other workbook package content and
does not ship a VBA add-in to bridge the missing API. The complete Power Query
contract remains available on Windows.

VBA source operations are also unsupported on macOS. Apple Events exposes no
`VBProject`, `VBComponents`, or `CodeModule` surface, and Office.js exposes no
VBA project API. Although Excel's dictionary declares `run VB macro`, it cannot
provide the exact-workbook qualification, bounded argument behavior,
prompt-free trust state, and timeout reconciliation required by the public
`vba.run` contract without workbook-resident helper code. ExcelMcp therefore
does not install an add-in, request macro approval, or change VBA project-model
trust. The complete VBA contract remains available on Windows.

Scenario create/show share the same limitation: Apple Events exposes existing
scenario elements and selected mutation commands but no create or show command,
while Office.js exposes no Scenario API. Native list/update/delete/summary
candidates require a separately Excel-authored scenario fixture.

These actions use the `Unsupported` tier with `Blocked` evidence and generated
`MacLimitation` execution plans. They fail before workbook dispatch;
there is no environment-variable opt-in or hidden helper route.

## Removed VBA helper design

The spike evaluated a signed `.xlam` bridge, but rejected it as a product
architecture because it would require per-machine macro and certificate trust,
Windows-only signing, helper upgrades, and VBA project-model security changes.
The helper source, protocol, runtime routing, packaging, acceptance scripts, and
trust instructions are not part of the macOS product.

## Optional Office.js tier

The optional Office.js foundation has explicit install, activation, health,
upgrade, and removal lifecycle; authenticated loopback HTTPS; exact
workbook/session binding; correlated serialized requests; deadlines; and
runtime requirement-set negotiation.

Installation proves only health. The 65 source-contract methods carrying
`OfficeAddInAction` have matching repository handlers, requirement-set gates,
and Excel-free contract tests. They remain `Partial` and unavailable until
user-mediated activation and exact public CLI/MCP acceptance succeed. The
other formerly planned Office.js actions are explicit `MacLimitation` entries:
the reviewed API omits contract-critical identity, type, source, or result
metadata, and ExcelMcp does not return partial success or guessed values.

## Permissions and dialogs

The base workflow requires the user's normal macOS Automation consent for
Excel. ExcelMcp does not:

- request broad Full Disk Access;
- change Office or macOS privacy/security settings;
- grant file-access prompts;
- click macro, trust, privacy, or repair dialogs;
- treat Excel repair/recovery as acceptance;
- install VBA add-ins or change macro or certificate trust.

Diagnostic UI interaction, when explicitly authorized, must verify exactly one
Excel dialog, the exact workbook name, and expected text. Unknown, privacy, or
security dialogs remain untouched.

## Acceptance and validation

Portable validation:

```powershell
dotnet build Sbroenne.ExcelMcp.sln -c Release --no-restore -p:EnableWindowsTargeting=true
dotnet test tests/ExcelMcp.Portable.Tests/ExcelMcp.Portable.Tests.csproj -c Release --no-build --filter "RequiresExcel!=true&RunType!=OnDemand"
pwsh ./scripts/check-workbook-package-access.ps1
```

Native desktop acceptance:

```powershell
pwsh ./scripts/Test-MacE2E.ps1
```

Optional switches add only their named, already prepared capability slices.
Power Query and VBA are explicit API limitations and have no macOS acceptance
runner.

Release evidence must record:

- exact Excel version and Mac architecture;
- CLI and MCP cases executed;
- capability candidate actions explicitly enabled;
- dialogs observed;
- final exact-workbook inventory;
- retained recovery paths;
- build, portable tests, repository audits, and unavailable Windows COM checks.

No build-only or static route test is desktop Excel acceptance.

## Parity direction

The Windows contracts remain the baseline for operation names, parameters,
defaults, validation, results, errors, persistence, and observable behavior.
Power Query and VBA have recorded supported-API limitations, not promised
beta parity. A blocked action is not considered implemented, and an untested
action is not classified as an Excel platform limitation.

Future work must continue through Excel-supported APIs or another explicit
user-approved capability tier. Direct workbook package access is not an
implementation option.
