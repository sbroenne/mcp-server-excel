# Calculation Mode

`calculation_mode` controls automatic, manual, and semi-automatic recalculation
(automatic except data tables). All actions require `session_id`.

Use manual mode for bulk writes when repeated recalculation is costly. There is
no universal cell-count threshold: one rectangular write is already batched.
Reading formula text does not require changing the mode. Keep calculation
available when intermediate formula results are needed.

## Preserve the Workbook's Mode

1. Call `get-mode` and remember the returned `mode`.
2. Call `set-mode` with `mode: 'manual'`.
3. Write the requested values/formulas in rectangular blocks.
4. Call `calculate` with the appropriate scope.
5. Restore the prior mode, not necessarily `automatic`.

Treat restoration like a `finally` block: attempt it after success or failure.
If cancellation or a timeout removed the session, inspect `file list` first and
report that restoration could not be completed. Do not blindly reopen or repeat
writes; cancellation is not undo.

```text
previous = calculation_mode(action: 'get-mode', session_id: id).mode
try:
    calculation_mode(action: 'set-mode', session_id: id, mode: 'manual')
    range(action: 'set-values', session_id: id, sheet_name: 'Sales',
          range_address: 'A1:B2', values: [['Name', 'Amount'], ['Salary', 5000]])
    calculation_mode(action: 'calculate', session_id: id, scope: 'workbook')
finally:
    calculation_mode(action: 'set-mode', session_id: id, mode: previous)
```

The example is workflow notation, not executable code. Check each result before
continuing and surface both the original failure and any restoration failure.

## Actions

| Action | Purpose | Additional inputs |
|--------|---------|-------------------|
| `get-mode` | Read current mode and calculation state | None |
| `set-mode` | Change the mode | `mode`: `automatic`, `manual`, or `semi-automatic` |
| `calculate` | Recalculate formulas | `scope`: `workbook`, `sheet`, or `range` |

Sheet scope requires `sheet_name`. Range scope also requires `range_address`.
Choose workbook scope when dependencies cross sheets; inspect calculated values
when they are part of the requested result.
