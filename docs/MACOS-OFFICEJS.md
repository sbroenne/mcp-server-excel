# Optional Office.js bridge on macOS

The Office.js bridge is a versioned, optional capability tier. ExcelMcp's base
Apple Events features continue to work when it is not installed, not running,
or not active in a workbook. The installed configuration currently enables
only `bridge.health`. Candidate handlers for tables and table columns,
conditional formatting, same-workbook worksheet copy/move, regular charts and
chart configuration, ordinary local PivotTable creation/deletion, and
source-identifiable slicer creation remain unavailable to CLI and MCP callers
until their existing ExcelMcp contracts pass real-Excel parity tests. The
disabled ordinary-local PivotTable candidate also implements exact placed-field
removal, naming, value formatting, item filtering, label sorting, data reads,
row layout, row-field subtotals, and row/column grand totals.

The Office.js candidate deliberately excludes linked PivotChart creation,
OLAP/Data Model PivotTables, PivotCache configuration, PivotTable grouping,
calculated fields and members, and slicer operations whose source identity
cannot be proved. Field listing, row/column/filter/value placement, and
aggregation changes also remain excluded. Their shared results and validation
require the source field's exact data type and unique values; Office.js does not
expose a trustworthy source data type and cannot distinguish dates from numeric
Excel serials without guessing from number formats. These operations require
the trusted VBA capability or remain unsupported; a chart over PivotTable
output is not substituted for a genuine linked PivotChart.

The generated `enabledActions` allowlist is the shared Service/broker release
gate. Installation and upgrade reset it to `bridge.health`; development
validation must opt candidate actions in deliberately and must not ship that
allowlist change without recorded CLI and MCP evidence from real Excel.

Three internal screenshot-coordination actions are also implemented but
disabled:

- `screenshot.prepare-range-geometry`
- `screenshot.prepare-sheet-geometry`
- `screenshot.restore-view`

They require `ExcelApiDesktop 1.1` and are not public CLI or MCP actions.
Prepare snapshots the active worksheet, selection, scroll, zoom, view,
split/freeze state, and window state; activates and scrolls the requested
range into view; and returns an opaque, single-use restore token. The response
also includes the Excel window number and both the range and Excel window
rectangles converted by their four edges with
`pointsToScreenPixelsX`/`pointsToScreenPixelsY`. It does not apply
`devicePixelRatio` or infer title-bar dimensions.

The token is bound to the exact broker session and saved workbook URL. Only a
successful restore consumes it, and a session cannot prepare another capture
while restoration is outstanding. The native capture implementation owns
process/window identity and must independently validate its exact
`SCWindow.windowID`, owning process ID, and Excel bundle identifier. Office.js
does not expose a macOS process ID or `CGWindowID`; title matching, frontmost
window selection, and guessed window chrome are not valid substitutes.

The current disabled geometry contract supports one contained rectangle only.
It rejects a range whose converted rectangle extends outside the converted
Excel window instead of guessing a tile split or allowing the native side to
crop outside the window. Sheet preparation first limits the used range to its
top-left 500 rows and 50 columns and reports that truncation. Office.js tiling,
the Windows 10%-40% zoom planning rules, and the 36/64-tile behavior are not
implemented, so large-range screenshot parity remains unavailable.

The reported global top-left physical-pixel coordinate space and its mapping
to ScreenCaptureKit's window-local backing pixels remain unverified until a
real Mac Excel run proves that the Desktop API window number maps to the Cocoa
and ScreenCaptureKit window ID and that both rectangles share the expected
origin and scale. The actions must remain disabled until that evidence exists.

## Prerequisites and trust boundary

- Apple Silicon macOS with the supported Excel for Mac version.
- Node.js 22 or newer for this development-stage package.
- A currently valid PEM X.509 development leaf with `CA:false` and
  `subjectAltName=DNS:localhost`, plus its matching unencrypted PEM private
  key. `serverAuth` extended key usage is recommended. Trust must be granted by
  the user because Excel for Mac requires HTTPS for sideloaded add-ins.

ExcelMcp does not create a trusted root, add a certificate to the keychain,
click a trust dialog, or change Excel macro/VBA trust. Create and trust the
localhost certificate through your organization's approved process, then pass
its paths explicitly to the installer. The private key and generated bridge
token are copied with mode `0600`.

## Install and activate

From `office-addin`:

```bash
npm run bridge -- install \
  --certificate /absolute/path/to/localhost.pem \
  --private-key /absolute/path/to/localhost-key.pem
npm run bridge -- start
```

The installer writes:

- configuration under
  `~/Library/Application Support/ExcelMcp/officejs`;
- the add-in manifest to Excel's per-user `wef` sideload directory.

Restart Excel after installation. Open the exact saved workbook already owned
by an ExcelMcp session, then activate **ExcelMcp capability bridge** from
Excel's add-ins UI if Excel does not activate the sideloaded task pane
automatically. This user-mediated activation is the current blocker to a
prompt-free real-Excel smoke test.

The add-in reports Excel's host version and every supported numbered
`ExcelApi` and `ExcelApiDesktop` requirement set at runtime. The two families
are negotiated independently. `ExcelApiDesktop 1.1` is never used as an XML
manifest activation requirement. Availability requires both the installed
Excel version and the action's specific numbered requirement set; API
existence in documentation or a static schema is not treated as active
workbook support.

## Health and unavailable states

With the bridge running:

```bash
npm run bridge -- health
```

Health succeeds only over authenticated localhost HTTPS. A future ExcelMcp
feature request must still fail explicitly if the add-in is absent, inactive
for the exact workbook, on a different protocol version, past its deadline, or
missing its required requirement set. It must never fall back to a different
workbook or report success with an error.

Native initiators retrieve status and terminal results through the
authenticated, exact-session-bound request status endpoint. Session close
invalidates the task-pane binding, terminalizes and removes outstanding
requests, removes completed results, and releases the workbook reservation.
The CLI health probe fails with actionable guidance after a fixed five-second
deadline instead of waiting indefinitely.

## Guarded candidate acceptance

Source implementation and mock protocol tests do not establish Excel parity.
After the coordinator grants the serialized desktop slot and the user completes
the certificate trust, install, bridge start, and exact-workbook task-pane
activation above, run the separate public-entry-point acceptance workflow:

```bash
pwsh ./scripts/Test-MacOfficeJsAcceptance.ps1 \
  -WorkbookPath /absolute/path/OfficeJsAcceptance.xlsx \
  -CliPivotTable CliPivot \
  -CliRemovalPivotTable CliRemovalPivot \
  -McpPivotTable McpPivot \
  -McpRemovalPivotTable McpRemovalPivot \
  -RowField Region \
  -ValueField Sales \
  -SelectedItem North \
  -UserSetupConfirmed \
  -CandidateAllowlistConfirmed \
  -ExcelSlotConfirmed
```

Use a dedicated saved workbook. `CliPivot` and `McpPivot` must be ordinary
local-range or local-table PivotTables with the named row and value fields
already placed. `CliRemovalPivot` and `McpRemovalPivot` are sacrificial copies
with the row field placed, because faithful placement is intentionally not an
Office.js candidate. The script uses only the public `excelcli` and MCP
entry points and emits a machine-readable receipt.

The runner never installs or trusts certificates, changes Excel security,
sideloads or launches Excel, edits `bridge.json`, or enables candidates. Before
running, deliberately add the nine actions reported by
`-ValidateOnly` to the local `enabledActions` allowlist. Restore the generated
health-only configuration after acceptance. A timeout after a dispatched
mutation is an uncertain outcome: stop, inspect the dedicated workbook, and do
not retry through another backend.

## Upgrade

Stop the bridge, update the source or installed package, then run:

```bash
npm run bridge -- upgrade \
  --certificate /absolute/path/to/localhost.pem \
  --private-key /absolute/path/to/localhost-key.pem
```

Upgrade replaces the manifest, certificate copy, and versioned configuration
while preserving the existing random authentication token. Restart Excel and
the bridge. Protocol-version mismatches fail closed.

## Remove

Stop the bridge and run:

```bash
npm run bridge -- remove
```

Removal deletes the ExcelMcp manifest, token, and copied key material. It does
not remove or alter certificates in the user's keychain because ExcelMcp did
not install that trust. Restart Excel to unload the add-in.

## Local-channel security

The broker binds only to `127.0.0.1` on a fixed configured port. API calls
require a 256-bit bearer token. Browser calls must have the exact configured
origin, the `Host` header must be loopback, bodies are bounded JSON objects,
and unsupported routes/actions fail closed. The token is delivered to the
task pane in the URL fragment, which is not sent in the HTTP request, and is
removed from browser history immediately.

Before accepting work, the native side registers a session and exact saved
workbook file URL. Office.js must activate for that same URL. Request IDs,
session IDs, and task-pane instance IDs are correlated on completion.
Operations are serialized per workbook, bounded by a 30-second maximum
deadline, and cancellable. Cancellation cannot stop an `Excel.run` call that
has already been dispatched. A timed-out mutation is therefore reported as
having an uncertain outcome and invalidates the logical session; ExcelMcp does
not retry it through another backend. A correctly correlated late result is
acknowledged as a terminal no-op so polling can continue, while authentication,
session, and workbook-binding failures remain fatal.

Microsoft references:

- [Excel JavaScript API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets)
- [Excel JavaScript API desktop-only requirement set](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-desktop-1-1-requirement-set)
- [Check API support at runtime](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/specify-api-requirements-runtime)
- [Sideload Office Add-ins for testing](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
- [Office diagnostics](https://learn.microsoft.com/en-us/javascript/api/office/office.diagnostics)
