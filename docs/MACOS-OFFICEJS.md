# Optional Office.js bridge on macOS

The Office.js bridge is a versioned, optional capability tier. ExcelMcp's base
Apple Events features continue to work when it is not installed, not running,
or not active in a workbook. The installed configuration currently enables
only `bridge.health`. Candidate handlers for tables and table columns,
conditional formatting, and same-workbook worksheet copy/move remain
unavailable to CLI and MCP callers until their existing ExcelMcp contracts
pass real-Excel parity tests.

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

The reported global top-left physical-pixel coordinate space and its mapping
to ScreenCaptureKit's window-local backing pixels remain unverified until a
real Mac Excel run proves that the Desktop API window number maps to the Cocoa
and ScreenCaptureKit window ID and that both rectangles share the expected
origin and scale. The actions must remain disabled until that evidence exists.

## Prerequisites and trust boundary

- Apple Silicon macOS with the supported Excel for Mac version.
- Node.js 22 or newer for this development-stage package.
- A localhost certificate and private key for `localhost`, with trust granted
  by the user. Excel for Mac requires HTTPS for sideloaded add-ins.

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
