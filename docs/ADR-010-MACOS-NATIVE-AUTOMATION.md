# ADR-010: Native macOS automation and an independently versioned VBA helper

**Status:** Accepted architecture; implementation and helper acceptance in progress

**Decision date:** 2026-10-03

## Decision

Windows and macOS share the annotated Core contracts, generated dispatch, input
validation, defaults, and result shaping. Supported Mac actions must match the
Windows/public behavior exactly; unverified actions and unsupported variants
remain gated before mutation.

The Mac transport will send Apple Events directly from C# through the system
AE framework. Replace the JXA bridge and C#-constructed AppleScript rather than
introducing a separate business-logic implementation in scripting languages.
Keep the existing non-prompting permission check, bounded child-process
execution, parent identity verification, and exact-workbook ownership.

There is no Office.js tier, task pane, hosted manifest, or localhost broker.
The Microsoft macOS bindings expose Apple Event APIs, but their packaging
requirements do not fit the existing plain `net10.0` single-file executables.
The small internal interop layer uses Apple's supported C ABI instead.

Where Excel's dictionary cannot expose the needed object model, evaluate a
small VBA helper containing stable primitives, not public command logic.
Power Query is the first acceptance target. A helper's presence alone never
enables an unverified public action.

## Helper artifact and compatibility

- Source-controlled VBA modules and an independent helper version drive an
  Excel-authored `.xlam` build artifact.
- Maintainers build locally on a Mac with desktop Excel and publish a separate
  `helper-vX.Y.Z` GitHub release. Ordinary server releases do not rebuild the
  helper.
- Users download and install the add-in through Excel's add-in manager. They do
  not import source modules or build the artifact themselves.
- A read-only handshake reports the helper version and available primitives.
  The server checks the supported major version and each action's required
  primitives before mutation. Server and helper versions need not match.
- Missing, incompatible, malformed, or failed handshakes are explicit errors.
  A missing macro's empty response is not proof of compatibility or success.
  Compatible older helpers continue serving actions they support.
- Security settings, macro approvals, and trust dialogs stay user-controlled.
  The build/installation process must not change them automatically.

## Evidence and remaining acceptance

A local native C# spike exercised `AECreateAppleEvent`, `CreateObjSpecifier`,
`AEPutParamDesc`, and `AESendMessage` against Excel. The non-prompting permission
probe returned allowed. An isolated Excel-authored workbook supported cell
read/write/readback and exact-workbook close without saving. This establishes
transport feasibility, not full public command parity.

Session preflight, attachment, open-state checks, and close now use the native
transport in the production automation child. Real CLI and MCP acceptance
verifies exact-path and duplicate-name rejection, no extra workbook opening,
save/close/reopen persistence, and default discard. The parent passes its
remaining monotonic operation budget to the child. Native timeouts retain
uncertain-outcome and open-recovery behavior.

Worksheet listing also uses the native transport and the shared
`WorksheetListResult` model. CLI/MCP acceptance verifies one-based indices and
visible, hidden, and very-hidden states, including after save/reopen.

Worksheet creation also uses native events and Excel's supported insertion
location record. It inserts before the owned workbook's active sheet, matching
Windows `Worksheets.Add()` instead of appending. The returned object specifier
selects the new sheet for naming without relying on global active-workbook
state. CLI/MCP acceptance verifies successive insertion order and save/reopen
persistence. Record key calls use the SDK's `AEPutParamDesc` alias rather than
trying to import the non-exported `AEPutKeyDesc` macro.

Worksheet rename/delete now share the Core missing-sheet validation. Mac
preflight uses native names and exact ordinal matching, rather than the old
scripting path's case-folding/quote normalization. Missing names fail in the
shared Service before mutation, preserving the Core exception and diagnostic.
CLI/MCP acceptance verifies unchanged worksheet state after both failures and
successful exact-name rename/delete persistence.

Worksheet rename now sets the name of the exact native worksheet object rather
than dispatching scripting business logic. Windows and Mac share naming-failure
context: an explicit Excel rejection identifies the unchanged existing sheet,
or the new sheet remaining after creation. Native replies carry the original
OSStatus diagnostic through the child boundary. Transport failures and timeouts
are not rewritten as naming rejections. CLI/MCP acceptance verifies the shared
rejection diagnostic and unchanged protected structure through save/reopen.

An isolated native range spike also verifies writable `Formula2` and
`Formula2R1C1`, including rectangular matrices, per-cell relative R1C1
references, dynamic spills, explicit `@`, blank-valued formula detection,
and save/close/reopen persistence. Formula constants are extracted specifically
from the dictionary's range class, not similarly named validation properties.
The native list encoder copies nested descriptors so callers can dispose their
source rows and strings safely.

Native property references also identify merged-range top-left rows and columns.
Reading those properties from Excel's returned merge-area object failed with
OSStatus -1728; nesting the requested property under the merge-area property
reference succeeds. Rectangular formula reads/writes now use shared input
validation and result/error shaping. CLI/MCP acceptance verifies A1/R1C1
matrices, formula errors, blank-valued formula occupancy, default overwrite
rejection, and save/reopen persistence. Public acceptance also covers JSON-file
A1/R1C1 inputs, unchanged state after shape rejection, dynamic spills, and
explicit `@` formulas through save/reopen. General merge inspection remains
gated. Public merged-write rejection acceptance is still incomplete: the new
cases exposed a diagnostic mismatch and a subsequent native timeout. The raw
merge-area geometry spike alone does not establish public merged-write parity.

Formula writes preserve automatic, manual, and semiautomatic calculation modes
through CLI and MCP. The native mode getter returns missing when all workbook
windows are hidden; mode acceptance therefore uses visible owned fixtures and
strict native enum readback, never a fabricated default or a production fallback.
Normal sheet and rectangular range calculation now use native commands and
shared validation/results. Public acceptance verifies dirty-cell updates,
unchanged active worksheet/mode, and unrelated dirty-cell sentinels in other
ranges, worksheets, and workbooks. Application-global and full/rebuild
calculation, and unsupported address variants, remain explicitly gated.

Native formula-error discovery precedes Value2 reads. A mixed matrix containing
formula errors caused Excel to terminate during the original bulk read spike.
The implemented path reads only non-error runs and maps confirmed error types
through the shared error contract; it does not retry the unsafe bulk getter.

The Service build generates internal Mac command bases and typed transport
adapters from the referenced Core interfaces, preserving method types, names,
and defaults. Generated platform command sets register every category behind
its Core interface. Both platforms now use the same Service routing and
generated argument binding and result serialization; there is no separate
Mac command dispatcher. Unavailable methods fail before accessing a batch.
Regression checks match registered overrides against the capability inventory.
Excel-specific validation and the remaining scripting business logic still
need native/shared-semantic migration.

The internal non-COM batch adapter implements COM-shaped interface slots in
ComInterop, where their embedded PIA types are defined, and rejects those calls
explicitly. Independently embedding generic COM signatures in Service caused
a real type-load failure. Keeping the adapter in the defining assembly avoids
that failure without adding an Office PIA runtime dependency or public API.

`scripts/Update-MacAppleEventCodes.ps1` extracts the used workbook/window
and worksheet classes, properties, visibility enums, and close command from installed Excel and its imported
standard SDK dictionary. Its checked-in output supports Excel-free builds.
Missing, ambiguous, or malformed required symbols fail before writing output.

Excel's dictionary exposes macro invocation but not VBA module authoring.
The first maintainer bootstrap therefore requires a manually imported,
reviewed VBA module in an Excel-authored workbook. Subsequent local builds can
invoke its build function through native Apple Events. Helper execution,
Power Query lifecycle, refresh completion, macro trust, packaging, and
version/capability diagnostics still need acceptance.

## Consequences

Keep the existing scripting path until its supported actions have equivalent
native implementations and desktop acceptance. Do not claim the migration is
complete on the strength of descriptor tests or a successful build. Do not
reintroduce removed Windows contracts to accommodate old Mac code.

Workbook contents are accessed through supported Excel APIs in production,
tests, scripts, and fixtures. Intact Excel-authored workbook files may be copied
opaquely. The helper is saved by Excel, never assembled by editing OOXML or VBA
binary internals.
