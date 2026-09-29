# macOS support

**Status:** ExcelMcp ships capability-gated Apple Silicon and Intel macOS
artifacts. Windows retains the complete COM backend. macOS combines a native
Apple Events backend with optional Office.js and ScreenCaptureKit tiers.
Unsupported or unproven actions fail with `PlatformNotSupported`; they never
return success-shaped approximations.

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

Excel 16.113.1 defines a thin read-only `workbook connection` class but exposes
no workbook collection, creation command, or typed OLEDB/ODBC properties needed
by the public connection contract. It exposes QueryTable elements and
properties, but no construction command for text or web sources. Connection
actions and QueryTable creation therefore remain `MacLimitationCandidate`
plans. QueryTable view/set-properties are also limitation candidates because
the dictionary omits fields required by their public contracts. List, refresh,
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
`MacLimitationCandidate` execution plans. They fail before workbook dispatch;
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
- install VBA add-ins or change macro or certificate trust.

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
Power Query and VBA remain primary parity workstreams. A blocked action is not
considered implemented, and an untested action is not classified as an Excel
platform limitation.

Future work must continue through Excel-supported APIs or another explicit
user-approved capability tier. Direct workbook package access is not an
implementation option.
