# Safe, Compatible Excel Automation Specification

**Status:** Approved implementation; four public acceptance seams confirmed on 2026-08-05
**Scope:** All P0 and P1 work identified in the 2026-08-05 architecture and issue review
**Compatibility:** MCP Server and CLI are equal entry points

## Outcome

ExcelMcp must detect the capabilities of the active Excel installation, explain unsupported features before they fail, allow users to review consequential operations, create recoverable checkpoints, record what happened, verify actual workbook changes, and provide safe recovery evidence after interruption.

The implementation must remain compatible with existing direct execution. Safety enforcement is opt-in per session until a future breaking release explicitly changes the default.

## Required Deliverables

| ID | Priority | Requirement | Completion evidence |
| --- | --- | --- | --- |
| P0-1 | P0 | Session capability preflight | Real-Excel response includes Excel/version/bitness/locale, Formula2 and dynamic-array mode, Python status, VBA trust, Power Pivot, read-only/protection/IRM state, constraints, and collection timestamp. |
| P0-2 | P0 | Formula compatibility | Formula reads and writes use Formula2 where supported and a per-session Formula fallback where unavailable; write failures are never swallowed. |
| P0-3 | P0 | Python availability diagnostics | A PY formula producing `#NAME?` returns a stable unavailable diagnostic distinct from transient `#BUSY!`, `#CONNECT!`, and `#BLOCKED!`. |
| P0-4 | P0 | Security dependency updates | PR-equivalent updates for `undici`, `cryptography`/`msal`, and `fast-uri` are present and their package checks pass. |
| P0-5 | P0 | Security/privacy reconciliation | Runtime telemetry opt-out works; emitted exception telemetry is redacted end to end; documented file/path controls match enforced behaviour. |
| P0-6 | P0 | Tracker/release reconciliation | #743 is documented as shipped, current fixes have changesets, and public feature/help material matches generated surfaces. |
| P1-1 | P1 | Review-before-write | A mutation can return an exact review plan without changing Excel; a bound, unexpired review ID is required when session review enforcement is enabled. |
| P1-2 | P1 | Reversible checkpoint | A requested/required checkpoint is created through `Workbook.SaveCopyAs` immediately before the mutation; checkpoint failure blocks execution. |
| P1-3 | P1 | Durable operation journal | Reviewed, checkpoint-reserved, checkpoint-created, started, completed, failed, aborted, verified, and recovery transitions are stored durably and can be rendered chronologically. |
| P1-4 | P1 | Semantic diff and verification | Receipts distinguish verified, partially verified, and not verified; they report only the scope actually inspected. |
| P1-5 | P1 | Diagnostics and recovery | Timeout, cancellation, Excel death, and client loss produce sanitized durable recovery evidence and expose safe reopen-from-checkpoint behaviour. |

## Public Seams

Tests exercise the same interfaces used by users:

1. **Core command interface through real Excel:** capability probes, formulas, Python results, checkpoint creation, semantic inspection, and workbook recovery.
2. **Shared service dispatcher:** review enforcement, token binding, journaling, checkpoint ordering, diff receipts, failure classification, and recovery recording.
3. **Generated MCP and CLI surfaces:** identical safety parameters, actions, defaults, and JSON results.
4. **Session lifecycle:** open, configure safety, execute/review, save/close, interrupt, list recoveries, and recover.

Tests must not mock Excel COM. Pure serialization, cryptographic fingerprinting, redaction, error classification, and catalog algorithms may use focused Excel-free tests only where the existing ADR permits pure-algorithm exceptions.

### Confirmed acceptance evidence

| Seam | Evidence confirmed on 2026-08-05 |
| --- | --- |
| Real Excel Core | 4/4 focused real-Excel tests pass for capability preflight, Formula2 spill/table semantics, and the direct-PY formula boundary; the account-independent Python classifier passes 5/5. |
| Shared service dispatcher | The expanded visible real-Excel safety/recovery and structural workflow passes 8/8; the focused CLI/service/catalog/privacy/path/interruption suite passes 79/79. |
| Generated MCP/CLI parity | MCP schema/forwarding/telemetry tests pass 7/7, real MCP and CLI preflight each pass 1/1, and both end-to-end smoke workflows pass. |
| Session interruption/recovery lifecycle | The expanded visible 8/8 workflow covers review, checkpoint, recovery, discard-on-shutdown, Excel death, fail-closed journal outages, named-range inspection, and structural review/verification; all 27 mandatory ComInterop OnDemand timeout/cancellation/message-pump/cleanup/reuse tests pass. |

## Architecture

```text
Generated MCP tool / generated CLI command / handwritten file tool
                              |
                        ServiceRequest
                              |
                       ExcelMcpService
                              |
                 WorkbookSafetyCoordinator
          +-------------------+-------------------+
          |                   |                   |
 CommandSafetyCatalog  WorkbookSemanticInspector  DurableSafetyStore
          |                   |              journal/checkpoints/recovery
          +-------------------+-------------------+
                              |
                    SessionManager / IExcelBatch
                              |
                    serial Excel STA / COM
```

`WorkbookSafetyCoordinator` is a deep module at the shared execution seam. Core command implementations remain focused on Excel operations and do not independently implement review, checkpoint, journal, diff, or recovery behaviour.

## Capability Preflight

### Entry point

- Service command: `session.preflight`
- MCP: `file(action: "preflight", session_id: "...")`
- CLI: `excelcli session preflight --session-id ...`
- `file test` remains a pre-open, Excel-free file validation operation.

### Result contract

```json
{
  "success": true,
  "sessionId": "...",
  "filePath": "C:\\Work\\Book.xlsx",
  "excel": {
    "version": "16.0",
    "build": 12345,
    "bitness": "x64",
    "operatingSystem": "Windows (64-bit)",
    "uiLocale": "en-US"
  },
  "capabilities": {
    "formula2": { "status": "supported", "dynamicArrays": true },
    "pythonInExcel": { "status": "notDetermined" },
    "vbaTrust": { "status": "disabled" },
    "powerPivot": { "status": "supported" }
  },
  "workbook": {
    "readOnly": false,
    "structureProtected": false,
    "windowsProtected": false,
    "irmProtected": false
  },
  "constraints": [],
  "collectedAtUtc": "2026-08-05T00:00:00Z"
}
```

Capability states are `supported`, `unsupported`, `unavailable`, `blocked`, or `notDetermined`. Version alone must never claim Python entitlement or tenant/network availability.

Formula2 support is probed by a safe read and cached per Excel session. Unsupported-member/Excel-1004 results select legacy Formula mode. Other COM errors propagate.

Python capability remains `notDetermined` unless an existing result supplies evidence or an explicit opt-in probe is added. A capability check must not silently dispatch cloud Python code in the user's workbook.

## Review and Execution Handshake

Potentially mutating generated MCP and CLI commands expose these universal options:

- `review_only` / `--review-only`
- `review_id` / `--review-id`
- `checkpoint` / `--checkpoint`

Session safety configuration controls:

- review mode: `off`, `optional`, `required`
- checkpoint mode: `off`, `onRequest`, `required`
- journal mode: `off`, `on`
- verification mode: `off`, `on`
- abnormal shutdown policy: `legacyAutoSave`, `discardWithRecoveryEvidence`

The default configuration preserves current direct-execution behaviour.

### Review-only response

```json
{
  "success": true,
  "executed": false,
  "reviewId": "...",
  "operationId": "...",
  "willWrite": true,
  "workbook": "Book.xlsx",
  "affected": {
    "sheets": ["Forecast"],
    "ranges": ["Forecast!B2:F200"],
    "objects": []
  },
  "saveDestination": "C:\\Work\\Book.xlsx",
  "checkpoint": {
    "requested": true,
    "requiredBeforeWrite": true,
    "destination": "C:\\Work\\Book.checkpoint.20260805T120000000Z.xlsx"
  },
  "warnings": ["The operation replaces existing cell content."],
  "verificationPlan": "rangeSemantic",
  "expiresAtUtc": "2026-08-05T12:05:00Z"
}
```

The review ID is cryptographically random and binds the normalized command, normalized arguments, session/workbook identity, requested checkpoint policy, baseline semantic fingerprint, and expiry. Changing arguments, changing workbook state, using another session, or waiting past expiry invalidates the review.

Execution is exactly-once for a review ID. Reuse returns a stable already-consumed error and never repeats the mutation.

## Command Safety Catalog

Each generated action has a descriptor:

```text
isMutation
mutationKind
scopeResolver
verificationLevel
checkpointRecommended
recoveryRisk
```

Unknown or uncatalogued actions fail closed as mutations. Read operations are explicitly classified. Scope is derived from known public parameters where possible and can be overridden with an attribute for complex operations.

Minimum mutation kinds are values, formulas, formatting, workbookStructure, modelStructure, externalRefresh, macroExecution, save, and fileCreation.

## Checkpoints

- Use `Workbook.SaveCopyAs` on the session STA thread.
- Persist a durable `checkpointReserved` reference before calling Excel so a completed copy can be finalized after a process crash.
- Wait for a bounded calculation-settled state or record that the checkpoint may contain pending calculations.
- Preserve `.xlsx`/`.xlsm` format and create a collision-proof UTC name.
- Never overwrite the source workbook or an existing checkpoint.
- Verify checkpoint existence, non-zero size, and hash before dispatch.
- A required checkpoint failure prevents the mutation.
- New, never-saved workbooks report that no prior-state checkpoint exists.
- Protected/IRM/network locations report an exact limitation rather than silently copying elsewhere.

## Journal and Recovery Store

The durable state root defaults under the current user's local application data and is overrideable with `EXCELMCP_STATE_DIR` for tests and managed deployments.

Records contain IDs, timestamps, action, safe argument summary, affected scope, result category, checkpoint reference, verification summary, duration, and recovery status. They do not contain cell values, full formula bodies, connection-string secrets, passwords, raw stack traces, or unsanitized portable paths.

Locally rendered review results may show the exact user-confirmed path. Exportable diagnostics always redact it.

Recovery means opening a known checkpoint or last explicitly saved file as a new session. It never claims to recover unsaved in-memory Excel state.

## Semantic Inspection and Diff

Snapshots are bounded summaries and hashes, not workbook dumps.

- **Range:** address, dimensions, value/formula/error-type hashes, and changed-cell counts. Formatting, validation, hyperlinks, and other uninspected range metadata are explicitly partial.
- **Worksheet:** name, order, visibility, used-range bounds, and target-range summaries.
- **Collections:** tables, names, charts, PivotTables, Power Queries, connections, and relationships where accessible.
- **Persistence:** saved-copy/reopen fingerprint for explicit save verification.
- **Opaque operations:** VBA, external refresh, and unsupported COM surfaces receive conservative collection/workbook fingerprints and a limited-verification statement.

Verification status is one of `verified`, `partiallyVerified`, `notVerified`, or `failed`. A receipt must never imply full-workbook verification when only a target range or collection was inspected.

## Diagnostics and Privacy

- `EXCELMCP_TELEMETRY_OPTOUT=true` disables MCP Application Insights initialization at runtime.
- Exception telemetry is built from sanitized exception details; an emitted-payload test proves paths, emails, credentials, and connection secrets are absent.
- The same sanitizer is used for portable journals, recovery manifests, stderr diagnostics, service error details, and telemetry.
- The CLI retains its no-telemetry contract.
- File size, supported extensions, and practical path limits are implemented once at a shared validation seam and documented from that behaviour.

## Recovery Behaviour

- Timeout/cancellation poisons the session, writes an `abortedUnknown` record, preserves any checkpoint, and follows existing bounded force-close cleanup.
- Excel process death writes recovery evidence before session tracking is removed whenever the coordinator has an active operation.
- A dropped CLI RPC command does not destroy an otherwise healthy reusable session.
- When any safety control is enabled, an omitted abnormal-shutdown policy defaults to discard-with-recovery-evidence; an explicit `legacyAutoSave` choice is preserved. Unconfigured legacy sessions retain current auto-save behaviour for compatibility.
- Server shutdown adds `abortedUnknown` / `ServerShutdown` evidence to incomplete durable operations without rewriting completed outcomes.
- `file(action: "recoveries")` lists safe recovery summaries.
- `file(action: "recover", recovery_id: "...")` opens the checkpoint as a new session without overwriting the original workbook.

## Vertical TDD Slices

1. Capability preflight contract and Formula2 session probe.
2. Formula read/write fallback while preserving modern dynamic-array behaviour.
3. Python `#NAME?` availability classification.
4. Runtime telemetry opt-out and emitted exception redaction.
5. Generated command safety catalog with fail-closed defaults.
6. Range review token binding and required-review enforcement through service, MCP, and CLI.
7. Pre-write `SaveCopyAs` checkpoint and durable journal ordering.
8. Range value/formula semantic diff and verification receipt.
9. Sheet/table/name/chart structural diff.
10. Explicit save persistence verification and journal retrieval.
11. Timeout/cancellation/Excel-death recovery record and reopen-from-checkpoint.
12. Catalog and documentation audit covering every mutating action.

Each slice follows red, green, then review. Bug fixes include same-pattern searches and the repository-required Core/MCP/CLI documentation and changeset work.

## Completion Gate

## Current implementation limitations

- Safety configuration and the review/checkpoint/journal/verification workflow are opt-in; direct execution remains the compatibility default.
- The four public seams are capability preflight, shared service dispatch, generated MCP/CLI surfaces, and session lifecycle. Handwritten lifecycle file/session mutations (open, create, save, close, and recovery) are outside the universal generated-command review handshake.
- Same-workbook worksheet mutations use the shared safety coordinator. Atomic `sheet.copy-to-file` and `sheet.move-to-file` reject review/checkpoint flags before Excel dispatch because a single-workbook checkpoint cannot safely cover both workbooks.
- Checkpoints are local, unencrypted full-workbook copies. Their SHA-256 detects accidental corruption but is not a malicious-tamper-proof security boundary. Recovery opens a checkpoint or explicitly saved workbook as a new session and cannot restore unsaved in-memory Excel state.
- Safety-state paths reject network drives and existing reparse points, and are checked before and after directory creation. These checks are not an atomic security boundary against another process running as the same user racing a path replacement; managed deployments should use a private local directory with user-only permissions.
- Python capability remains `notDetermined` unless existing evidence is available; preflight does not dispatch cloud Python or probe tenant/network entitlement.
- File validation accepts existing `.xlsx`/`.xlsm` for `test`, `.xlsx`/`.xlsm`/`.xls` for `open`, and creates `.xlsx`/`.xlsm`; existing files are limited to 1 GiB, general paths to 32,767 characters, and Excel SaveAs creation to a practical 218-character path.

Completion requires:

- all P0/P1 rows above proven by current-state evidence;
- zero uncatalogued public actions;
- MCP/CLI schema and behaviour parity;
- focused real-Excel integration tests passing for every affected feature;
- ComInterop OnDemand session/recovery tests passing;
- Release build with zero warnings/errors;
- COM leak, success-flag, coverage/naming, generated-surface, documentation-count, dynamic-cast, packaging, CLI smoke, and MCP smoke gates passing;
- no commit, push, PR, or merge without separate user approval.
