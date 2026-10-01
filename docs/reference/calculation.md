# Calculation Mode

**Platform scope:** Windows supports all actions below. Experimental Apple
Silicon macOS supports `calculate`, but not `get-mode` or `set-mode`. On Mac,
write the requested blocks and calculate explicitly without changing the mode.
See [macOS beta limitations](https://github.com/sbroenne/mcp-server-excel/blob/main/specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta).

Calculation is not the same as data refresh. Recalculating formulas does not
reload external sources, and a successful write does not prove that cloud
Python or asynchronous refresh work has finished.

Use CLI help or MCP tool descriptions for current calculation controls and
inputs. This guide explains when and why to use them.

## Calculation after writes

Value and formula writes attempt to restore the original calculation mode,
not always force calculation. Restoration can fail without failing the write;
inspect the current mode when subsequent work depends on it.

| Mode | What to consider after a write |
|------|-------------------------------|
| Automatic | Excel normally recalculates dependent formulas after mode restoration. |
| Manual | Explicitly calculate before relying on dependent results. |
| Semi-automatic | Ordinary formulas recalculate, but What-If data tables need explicit calculation. |

Semi-automatic excludes What-If data tables, not ordinary worksheet Tables.
Read the relevant calculated values before treating a result as verified.

## Preserve the Workbook's Mode (Windows)

Manual mode can avoid repeated expensive recalculation during bulk edits.
One rectangular write is already batched; there is no universal size threshold
that makes manual mode necessary. Do not change it for formula-text reads or
when intermediate results are needed.

1. Inspect and remember the current calculation mode.
2. Switch to manual only if the workflow benefits from it.
3. Write the requested data in appropriate blocks.
4. Calculate the scope needed by the result.
5. Restore the original mode after success or failure.

Restoration is cleanup, not a reason to hide the original failure. If timeout
or cancellation invalidated the session, inspect session state before attempting
cleanup. Do not blindly reopen the workbook and repeat writes. Report when
restoration could not be completed.

## Calculation scope and precision {#actions}

Application settings affect all workbooks in the session's owned Excel process,
not other Excel processes. Use application-wide calculation when dependencies
cross worksheets; deeper recalculation and dependency rebuilds also require
that scope.

Precision-as-displayed belongs to the workbook. Enabling it permanently rounds
stored numbers to their displayed precision. Disabling it does not recover lost
digits. Change display formats, not stored precision, when the goal is simply
readable output.

See [working safely with Excel](behavioral-rules.md#changes-and-formatting)
for partial failures, permissions, and saving.
