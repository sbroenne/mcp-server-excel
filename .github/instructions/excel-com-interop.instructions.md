---
applyTo: "src/ExcelMcp.Core/**/*.cs,src/ExcelMcp.ComInterop/**/*.cs"
excludeAgent: "code-review"
---

# Excel COM pitfalls

- Prefer typed Excel PIAs. Late binding needs a documented PIA/runtime gap;
  existing `Application.Run`, VBE, and Office-core calls may intentionally avoid
  unavailable dependencies. Follow `scripts\check-dynamic-casts.ps1`.
- Excel may marshal integer properties as doubles. Use `Convert.*` rather than
  direct numeric/enum casts from dynamic values. Collections are one-based.
- Release acquired COM references in reverse order in `finally` with
  `ComUtilities.Release(ref value)`, including intermediate collections and
  chained-property results. Never release session-owned `ctx.App`/`ctx.Book`
  or substitute forced GC for release.
- `ExcelWriteGuard` owns `ScreenUpdating` around `Execute`. Do not suppress it
  again in commands. No global event/calculation suppression; calculation
  suppression belongs only in established bulk value/formula writes, restored
  in `finally`.
- `RefreshAll()` does not establish completion. Reuse synchronous connection
  refresh or `QueryTable.Refresh(false)` before returning/saving.
- A generic HRESULT such as `0x800A03EC` does not identify a specific cause.
  Use distinguishing preconditions and the existing transient retry path.
- Close/quit through `ExcelShutdownService`. Preserve PID/start-time ownership
  and layered shutdown timeouts; never kill Excel by process name.
- Propagate cancellation and check it in loops. A timed-out batch must reject
  further work, not remain apparently healthy.
