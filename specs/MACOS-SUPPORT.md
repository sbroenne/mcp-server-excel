# macOS support

**Status:** ExcelMcp ships capability-gated Apple Silicon and Intel macOS
artifacts. Windows retains the complete COM backend. macOS combines a native
Apple Events backend with optional, explicitly trusted helpers. Unsupported or
unproven actions fail with `PlatformNotSupported`; they never return
success-shaped approximations.

## Non-negotiable workbook boundary

Excel workbook files are opaque.

- Never create, parse, inspect, or mutate workbook ZIP, OOXML, relationship,
  custom XML, or DataMashup internals.
- This prohibition applies to production code, tests, fixtures, scripts, and
  indirect use through package or Open XML libraries.
- An intact workbook authored by Excel may be copied as an opaque whole file.
- Workbook content changes must use Excel-supported object models or a trusted
  helper operating through Excel.
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

Apple Silicon is exercised locally. Intel artifacts are cross-built and
structurally validated; physical Intel Excel execution remains a separate
release requirement.

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
lock state, read access, and the existing IRM/AIP preflight. It does not inspect
workbook structure and therefore leaves `isValid=false` and `canOpen=false`
until `workbook.open` asks Excel to validate the file.

## Production native capabilities

The generated action inventory is authoritative:

- [Human-readable inventory](../docs/MACOS-ACTION-INVENTORY.md)
- [Machine-readable inventory](../docs/generated/macos-action-inventory.json)

`MacCommandCapabilities` consumes the same generated records. A method becomes
available only after exact public CLI and MCP acceptance updates its Core
capability annotation.

Each generated record also carries the remaining-work execution plan:

- `plannedTier`: the tier to prove next, or `MacLimitationCandidate` when no
  faithful route has been selected;
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

- exact workbook create/open/close and owned-session cleanup;
- worksheet list/create/rename/delete;
- worksheet visibility and tab color;
- range values, formulas, number formats, clear operations, explicit row and
  column sizing, copy variants, bounded metadata, auto-fit, merge/unmerge, and
  cell locking for the accepted variants;
- calculation, Goal Seek, one- and two-variable Data Tables;
- the accepted named-range lifecycle variants;
- exact-window screenshot identity where separately enabled.

Known native Apple Events gaps remain gated, including UsedRange,
CurrentRegion, merge-area inspection, worksheet copy/move, application-global
calculation-mode mutation in shared Excel, and any action whose declared
dictionary surface did not survive real CLI/MCP acceptance.

## Power Query

All macOS Power Query operations use the optional trusted VBA helper. There is
no saved-package fallback.

The current helper source supports fixed actions for:

- `list`, `view`, and `get-load-config` through `Workbook.Queries`, exact
  worksheet `ListObject.QueryTable` identity, and Excel's connection/model
  state;
- `create`, `update`, `rename`, and `delete`;
- synchronous `refresh` and `refresh-all`;
- worksheet-table and connection-only `load-to`/`unload` transitions;
- temporary-query `evaluate` with bounded result data and verified cleanup.

List results contain an M preview of at most 80 characters, formula character
count, exact load mode, target sheet when applicable, connection-only state,
and Data Model state. View returns the complete M formula. The Service validates
all helper fields and cross-field load-state consistency. Missing, malformed,
inconsistent, or mismatched query metadata fails the action.

Every Power Query action remains production-blocked by default. A method is
eligible for enablement only after a prompt-free run through both public entry
points proves its exact result, persistence, error, timeout, destination, and
cleanup behavior. Candidate opt-in is scoped to exact action names through
`EXCELMCP_MAC_POWERQUERY_CANDIDATE_ACTIONS`; there is no helper-wide enable
switch.

The guarded public acceptance runner is:

```powershell
pwsh ./scripts/Test-MacPowerQueryPublicAcceptance.ps1 `
  -HelperPath '/absolute/path/ExcelMcpHelper.xlam' `
  -WorkbookPath '/absolute/path/ExcelMcpPowerQueryAcceptance.xlsx' `
  -HelperInstalledTrustedConfirmed `
  -ExcelAuthoredWorkbookConfirmed `
  -DedicatedWorkbookConfirmed `
  -ExcelSlotConfirmed
```

The workbook must be Excel-authored, dedicated to acceptance, and contain no
existing Power Queries. The runner uses separate opaque working copies for CLI
and MCP, literal credential-free M, public actions only, exact worksheet value
checks, and saved close/reopen checkpoints. Uncertain workbooks are retained
with `RECOVERY_REQUIRED`; they are never guessed closed or deleted.

## VBA and the trusted helper

ExcelMcp never changes macro-security preferences or VBA project-model trust.
The user installs, trusts, upgrades, and removes the helper.

Helper protocol version 1 uses:

- helper version `1.4.0`;
- a 262,144-byte UTF-8 request and response limit;
- request envelope
  `{version,requestId,workbookPath,action,arguments}`;
- response envelope `{version,requestId,success,result,error}`;
- a 32-character lowercase hexadecimal request ID;
- exact `Workbook.FullName` target resolution;
- strict action and argument allowlists.

Version 1.4.0 adds complete Power Query read metadata. Older helper versions are
rejected so partial list/view responses cannot be mistaken for parity.

The helper exposes only fixed Power Query, scenario, and VBA lifecycle actions.
It has no caller-selected macro entry point, arbitrary expression evaluator,
general VBA source execution, or AppleScript execution action. Unknown,
duplicate, malformed, over-depth, oversized, or uncorrelated messages fail
before dispatch.

Capability output separates:

- static API availability;
- current user-managed trust readiness;
- per-method `provenMethods`.

All proof flags begin false. Static API presence and trust readiness are not
runtime proof.

Before VBA dispatch, ExcelMcp performs a bounded, non-prompting read of the
effective Office macro preferences. It distinguishes disabled macros,
per-workbook approval, unattended execution configuration, and indeterminate
state. Project source operations separately report whether user-managed
project-model trust is enabled.

The helper's VBA source actions are restricted to exact workbook-qualified
standard modules and bounded string parameters. Signed or locked projects are
not mutated, and no action saves the target workbook implicitly.

Direct helper-engine acceptance:

```powershell
pwsh ./scripts/Test-MacHelperAcceptance.ps1 `
  -HelperPath '/absolute/path/ExcelMcpHelper.xlam' `
  -WorkbookPath '/absolute/path/ExcelMcpHelperAcceptance.xlsm' `
  -MacroApprovalConfirmed `
  -VbaProjectTrustConfirmed `
  -ExcelAuthoredWorkbookConfirmed
```

Public VBA acceptance is separately guarded by
`scripts/Test-MacVbaPublicAcceptance.ps1`. A helper-engine pass does not enable
public VBA actions.

## Optional Office.js tier

The optional Office.js foundation has explicit install, activation, health,
upgrade, and removal lifecycle; authenticated loopback HTTPS; exact
workbook/session binding; correlated serialized requests; deadlines; and
runtime requirement-set negotiation.

Installation proves only health. Tables, charts, ordinary PivotTables, slicers,
conditional formatting, and worksheet movement candidates remain unavailable
until user-mediated activation and exact public CLI/MCP acceptance succeed.

## Permissions and dialogs

The base workflow requires the user's normal macOS Automation consent for
Excel. ExcelMcp does not:

- request broad Full Disk Access;
- change Office or macOS privacy/security settings;
- grant file-access prompts;
- click macro, trust, privacy, or repair dialogs;
- treat Excel repair/recovery as acceptance;
- install or trust a helper automatically.

Diagnostic UI interaction, when explicitly authorized, must verify exactly one
Excel dialog, the exact workbook name, and expected text. Unknown, privacy, or
security dialogs remain untouched.

## Acceptance and validation

Portable validation:

```powershell
dotnet test tests/ExcelMcp.Portable.Tests/ExcelMcp.Portable.Tests.csproj -c Release
dotnet build Sbroenne.ExcelMcp.sln -c Release --no-restore
pwsh ./scripts/check-workbook-package-access.ps1
```

Native desktop acceptance:

```powershell
pwsh ./scripts/Test-MacE2E.ps1
```

Optional switches add only their named, already prepared capability slices.
Power Query and VBA use their dedicated guarded runners rather than generated
workbook fixtures.

Release evidence must record:

- exact Excel version and Mac architecture;
- CLI and MCP cases executed;
- capability candidate actions explicitly enabled;
- dialogs observed;
- final exact-workbook inventory;
- retained recovery paths;
- build, portable tests, repository audits, and unavailable Windows COM checks.

No build-only or static helper test is desktop Excel acceptance.

## Parity direction

The Windows contracts remain the baseline for operation names, parameters,
defaults, validation, results, errors, persistence, and observable behavior.
Power Query and VBA remain primary parity workstreams. A blocked action is not
considered implemented, and an untested action is not classified as an Excel
platform limitation.

Future work must continue through Excel-supported APIs, the fixed trusted
helper, or another explicit user-approved capability tier. Direct workbook
package access is not an implementation option.
