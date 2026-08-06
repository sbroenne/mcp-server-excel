# GitHub Copilot work-computer validation

This is the client-side acceptance checklist for the fork's Copilot compact MCP
profile. Local builds and tests prove the server contract; only a real GitHub
Copilot work-computer client can prove installation, process launch, and the
client's `tools/list` result.

## Install, restart, and start a clean session

1. Record the currently installed plugin/version and any existing Excel MCP
   configuration. Close active Copilot chats and Excel workbooks that are not
   part of this test.
2. In the Copilot work-computer client, install the fork's `excel-mcp` plugin
   (or add its MCP server using the fork's published marketplace/release
   instructions). Do not silently reuse an old `excel-mcp` entry.
3. Restart the Copilot client completely (quit the app/process, then launch it
   again) so the MCP process and tool catalog are recreated.
4. Create a **new chat/session** after restart. Do not rely on a pre-existing
   conversation's cached tools. Confirm the MCP server is connected before
   continuing.

If the marketplace install is not the artifact under test, stop and record the
exact package/release, launcher path, and environment variables used. The
compact profile is selected with `EXCELMCP_TOOL_PROFILE=copilot-compact` (the
published Copilot bootstrap should set this for the fork).

## Discovery gate (`tools/list`)

Ask the client to enumerate the server's tools, or capture the raw MCP
`tools/list` response if the client exposes protocol logs. The compact profile
must contain exactly these nine names (ordering is not significant):

```text
file
workflow
worksheet
range
range_edit
range_format
worksheet_style
layout
calculation_mode
```

The live `tools/list` response is authoritative. A missing, extra, or full-
profile tool is a **fail**; do not infer success from documentation or a cached
UI list.

## Workflow identity gate

Call `workflow` with `action: "capabilities"` and save the complete response.
Verify these fields are present and internally consistent:

* `toolProfile` is `copilot-compact`;
* `toolProfileVersion` is `1`;
* `toolProfileFallback` is `full`;
* `toolProfileTools` is the same nine-name set above;
* `toolProfileManifestHash` is a stable 64-character lowercase hexadecimal
  hash for this server build;
* runtime identity fields `runtimeHost`, `serverVersion`, and
  `buildFingerprint` are present; and
* the response advertises `executePlan`, `openAndDescribe`, `compactReceipts`,
  `planCheckpoint`, `planIdempotency`, and `finalRangeVerification` as true
  (and identifies the workflow interface version). Excel process identity, if
  returned by the host's status/receipt, should also be recorded.

Record the response verbatim. A hash or identity mismatch between repeated
calls in one session is a fail and should be investigated as server drift.

## Disposable-workbook smoke sequence

Use a newly created, disposable `.xlsx` in a temporary directory. Never use a
production workbook. Keep `show: false` unless testing visible-window behavior.

1. **Open and describe.** Call `workflow` with
   `action: "open-and-describe"`, the disposable `file_path`, bounded preview
   rows/columns, and `show: false`. Save the returned `session_id`, workbook
   manifest, and runtime identity.
2. **Execute one small plan.** Using that `session_id`, call `workflow` with
   `action: "execute-plan"` and one or two deterministic operations (for
   example, write `SmokeOK` to `Sheet1!A1`). Set
   `checkpoint_mode: "once"`, provide a unique `idempotency_key`, and include
   `verify_sheet_name: "Sheet1"` plus `verify_range_address: "A1"` (or the
   exact range written). The result must include final-range verification and
   a receipt indicating the plan/checkpoint outcome.
3. **Unknown-outcome safety.** If the call times out, disconnects, or returns
   an unknown outcome, do **not** replay it with a new key. Reconnect/read back
   using the same key/session or inspect the receipt first; retries must be
   idempotent.
4. **Layout facade (if available in the test case).** Exercise one harmless
   `layout` action (report formatting, outline, or freeze-pane operation) and
   capture its result. This confirms the compact layout facade is registered.
5. **Read back.** Use `range` (or the workflow verification result) to read the
   exact target range and confirm the expected value and formatting/layout.
6. **Close safely.** Explicitly save, then call `file` with `close` for the same
   session. If Excel is visible or the workbook is dirty, obtain the user's
   approval in the Copilot client before closing.

## Evidence and pass/fail record

Capture: plugin/version and install source; restart/new-session timestamps;
raw `tools/list`; capabilities response; every request/response (redacting
secrets and unrelated workbook data); session and workbook paths; idempotency
key; receipt/checkpoint and final-range verification; read-back values; layout
result; close/save confirmation; and client/server logs showing the runtime
identity. Mark each gate pass/fail and attach the artifact set to the validation
report. A local build or unit/integration test result is supporting evidence,
not a substitute for client evidence.

## Avoid profile/process collisions

Never run the original server and this fork simultaneously under the same MCP
server name or configuration key. Copilot may connect to the wrong process and
make `tools/list` appear nondeterministic. During validation, disable/remove
the original entry or give the fork a distinct temporary server name and
launcher. Restore the original only after the fork process is stopped.

## Rollback

Stop the fork MCP process, uninstall the fork plugin (or remove its temporary
server entry), restore the previously recorded plugin/version and configuration,
restart Copilot, and start another new session. Reopen only non-disposable
workbooks after confirming the original server's `tools/list` and capabilities.
Keep the disposable workbook and logs until the report is complete.

## Approval boundary

Mutation/close approval must be requested and granted by the Copilot client (or
the user at the work-computer UI). The MCP server can report that review is
required and can enforce safety checks, but it cannot grant, simulate, or fake
human approval server-side. If the client does not show an approval prompt when
one is expected, mark the validation **fail** and do not treat a successful
server response as consent.
