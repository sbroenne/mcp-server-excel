---
applyTo: "src/ExcelMcp.Core/**/*.cs"
excludeAgent: "code-review"
---

# Excel COM interop

## API selection

- Prefer strongly typed `Microsoft.Office.Interop.Excel` members.
- Use `dynamic` only for a documented PIA or runtime dependency gap. Add the
  justification required by `scripts\check-dynamic-casts.ps1`.
- Confirm unfamiliar Excel Object Model behavior in Microsoft documentation and
  a proven interop implementation before coding.
- Excel collections are one-based.
- Do not assume a COM numeric property's runtime type. Use `Convert.ToInt32`,
  `Convert.ToDouble`, or another explicit conversion when marshaling can vary.

## Batch execution and exceptions

Validate .NET inputs before `batch.Execute`; perform COM work on its STA callback.
Do not catch broad exceptions merely to return an error result. The batch and
service layers own failure transport.

Specific catches are acceptable for a known HRESULT, a bounded retry, or
best-effort cleanup. Never use an empty catch or silently substitute a
success-shaped fallback.

## COM lifetime

Every acquired COM reference uses nullable locals and reverse-order cleanup:

```csharp
return batch.Execute((ctx, ct) =>
{
    Excel.Worksheet? sheet = null;
    Excel.Range? range = null;
    try
    {
        sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
            ?? throw new InvalidOperationException(
                $"Worksheet '{sheetName}' was not found.");
        range = sheet.Range[address];
        return ReadRange(range);
    }
    finally
    {
        ComUtilities.Release(ref range);
        ComUtilities.Release(ref sheet);
    }
});
```

- Release intermediate objects too, including collections and objects returned
  by chained property access.
- Do not release shared session-owned `ctx.App` or `ctx.Book` references.
- Do not force garbage collection as a substitute for deterministic release.
- Run `scripts\check-com-leaks.ps1` after interop changes.

## Excel application state

- `ExcelWriteGuard` already suppresses and restores `ScreenUpdating` around
  `Execute`; do not repeat it in commands.
- Do not globally suppress `EnableEvents` or `Calculation`.
- Calculation suppression belongs only in established bulk value/formula write
  paths and must restore the original state in `finally`.
- Restore any changed application or workbook state on every exit path.

## Refresh and persistence

- Avoid `Workbook.RefreshAll()` where the caller requires completion before
  returning or saving.
- Use the existing connection refresh path or `QueryTable.Refresh(false)` for
  synchronous refresh.
- Release QueryTable, ListObject, connection, and destination Range references.
- Persist only when the operation contract requires it; tests should save only
  when verifying close/reopen behavior.

## Shutdown and cancellation

- All close/quit paths go through `ExcelShutdownService`.
- Preserve PID identity tracking, retry behavior, and shutdown timeout layering.
- Never terminate Excel by process name.
- Check cancellation in loops and propagate operation timeouts. A timed-out
  batch must not accept later work or leave a session appearing healthy.

## Known late-bound exceptions

Late binding remains appropriate where early-bound PIAs introduce unavailable
runtime dependencies, including established `Application.Run`, VBE, and selected
Office-core members. Reuse the existing helper/pattern rather than adding a new
dynamic call.
