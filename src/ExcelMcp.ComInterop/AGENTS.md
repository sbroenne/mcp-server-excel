# Excel COM pitfalls

Follow [repository review checks](../../AGENTS.md#code-review-rules) and
[runtime boundaries](../AGENTS.md). Core commands share these COM rules.

- Use typed Excel PIAs. EXL011 rejects dynamic declarations and late-bound
  operations, and EXL005 rejects casts to `dynamic`; comments do not exempt
  them. Collections are one-based.
- Office Core and VBA editor definitions come from the pinned build-only
  `ExcelDna.Interop` download. Embed used types through `Directory.Build.targets`;
  do not deploy the definition DLLs, use machine-specific references, or add an
  assembly resolver. Verify actual Excel calls and published output when changing
  this boundary; compilation alone does not prove the runtime has no new dependency.
- ADO definitions are generated with Microsoft's pinned `TlbImp.exe` from the
  signed Windows ADO 2.8 type library and embedded as build-only inputs. See
  [Windows setup](../../docs/agents/development.md#windows-setup). Do not commit
  generated DLLs or use `COMReference`, a runtime resolver, or a substitute
  connection: callers must retain Excel's live ADO session.
- `ExcelDocumentProperties` is the approved document-property-only exception
  to checking every Excel member at compilation. Excel rejects the matching
  Office collection/item interfaces, including fresh Microsoft imports.
  Keep native `IDispatch` private, resolve member identifiers from Excel, and
  restrict calls to the fixed property surface. Validate returned item type
  identities, preserve native errors and argument/result ownership, and do not
  expose a general dispatch helper. This adds no runtime definition DLL and
  does not exempt any code from EXL011/EXL005.
- Release acquired COM references in reverse order in `finally` with
  `ComUtilities.Release(ref value)`, including dynamic intermediates.
  Never release session-owned `ctx.App`/`ctx.Book`
  or substitute forced GC for release.
- `ExcelWriteGuard` owns `ScreenUpdating` around `Execute`. Do not suppress it
  again in commands. Leave it unchanged when any open workbook in that Excel
  application has a Data Model: toggling it can mark an already saved model
  workbook as changed. No global event/calculation suppression; calculation
  suppression belongs only in established bulk value/formula writes, restored
  in `finally`.
- `RefreshAll()` does not establish completion. Reuse synchronous connection
  refresh or `QueryTable.Refresh(false)` before returning/saving.
- A generic HRESULT such as `0x800A03EC` does not identify a specific cause.
  Use distinguishing preconditions and the existing transient retry path.
  Optional reads tolerate only known unavailable-member codes, not all
  `COMException` or binder failures. Preserve unexpected errors and partial state.
- Close/quit through `ExcelShutdownService`. Preserve PID/start-time ownership
  and layered shutdown timeouts; never kill Excel by process name.
- Separate caller wait cancellation from Excel execution. An expired wait retains
  and reveals the owned workbook; reject new work until the original callback
  and cleanup return and readiness is confirmed. Do not terminate Excel or
  discard edits merely because a wait expired.
