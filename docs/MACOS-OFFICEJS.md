# Optional Office.js bridge on macOS

The Office.js bridge is a versioned, optional capability tier. ExcelMcp's base
Apple Events features continue to work when it is not installed, not running,
or not active in a workbook. The bridge currently enables only
`bridge.health`; tables, charts, PivotTables, conditional formatting, and other
workbook actions remain unavailable until their existing ExcelMcp contracts
pass real-Excel parity tests.

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

The add-in reports Excel's host version and every supported `ExcelApi`
requirement set at runtime. Availability requires both the installed Excel
version and the specific requirement set; API existence in documentation is
not treated as feature support.

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
deadline, and cancellable. Expired or cancelled requests cannot complete.

Microsoft references:

- [Excel JavaScript API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets)
- [Check API support at runtime](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/specify-api-requirements-runtime)
- [Sideload Office Add-ins for testing](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
- [Office diagnostics](https://learn.microsoft.com/en-us/javascript/api/office/office.diagnostics)
