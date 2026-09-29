# calculation_mode - Bulk Write Performance Optimization

## Tool

- **`calculation_mode`**: Control Excel's automatic recalculation behavior

## When to Use

Use `calculation_mode` to optimize performance when:
- Writing 10+ cells of data/formulas in a single operation
- Creating tables with multiple rows and calculated columns
- Performance matters more than immediate feedback (no need to wait for each formula to recalculate)

## When NOT Needed

- Small edits (1-5 cells)
- When you need immediate calculation results to verify data
- Reading formulas (use `range get-formulas` — works in any mode)
- Single worksheet operations without bulk writes

## Workflow

Always follow this 4-step pattern for bulk operations:

```
1. calculation_mode(action: 'set-mode', session_id: '<session-id>', mode: 'manual') -> Disable auto-recalc
2. Perform all data writes (range set-values, set-formulas)
3. calculation_mode(action: 'calculate', session_id: '<session-id>', scope: 'workbook') -> Recalculate once at end
4. calculation_mode(action: 'set-mode', session_id: '<session-id>', mode: 'automatic') -> Restore default
```

**Why this pattern:**
- Step 1: Prevents Excel from recalculating after EVERY cell write (10+ recalcs → 1 recalc)
- Step 2: All writes happen at normal speed
- Step 3: Single recalculation computes all formulas together
- Step 4: Restores default Excel behavior so subsequent edits auto-recalc

## Actions

| Action | Purpose | Parameters |
|--------|---------|-----------|
| `get-mode` | Check current calculation mode | None |
| `set-mode` | Switch between automatic/manual/semi-automatic | `mode: "automatic"` or `"manual"` or `"semi-automatic"` |
| `calculate` | Trigger recalculation | `scope: "workbook"` (all formulas), `scope: "sheet"` with `sheet_name`, or `scope: "range"` with `sheet_name` and `range_address` |

## Common Scenarios

### Scenario: Create Sales Table with Formulas

Task: Add 100 rows of product data with unit price, quantity, and total formulas.

```
1. calculation_mode(action: 'set-mode', session_id: '<session-id>', mode: 'manual')
2. range(action: 'set-values', session_id: '<session-id>', sheet_name: 'Sales', range_address: 'A2:C101', values: <100 rows>)
3. range(action: 'set-formulas', session_id: '<session-id>', sheet_name: 'Sales', range_address: 'D2:D101', formulas: <100 formulas>)
4. calculation_mode(action: 'calculate', session_id: '<session-id>', scope: 'workbook')
5. calculation_mode(action: 'set-mode', session_id: '<session-id>', mode: 'automatic')
```

**Performance:** ~2-3 seconds total (vs ~30+ seconds if automatic after every cell)

### Scenario: Dashboard with Multiple Sections

Task: Create 5 sections with headers, data, and subtotal formulas.

```
1. calculation_mode(action: 'set-mode', session_id: '<session-id>', mode: 'manual')
2. range(action: 'set-values', session_id: '<session-id>', sheet_name: 'Dashboard', range_address: '<section-1-range>', values: <section-1-values>)
3. range(action: 'set-formulas', session_id: '<session-id>', sheet_name: 'Dashboard', range_address: '<section-1-formula-range>', formulas: <section-1-formulas>)
4. Repeat the named range calls for sections 2-5.
5. calculation_mode(action: 'calculate', session_id: '<session-id>', scope: 'workbook')
6. calculation_mode(action: 'set-mode', session_id: '<session-id>', mode: 'automatic')
```

## Best Practices

1. **Always restore automatic mode** - Never leave manual mode enabled, users expect auto-recalc
2. **Use workbook scope for calculate** - Simplest and fastest
3. **Verify calculation completed** - After step 3, data should show final calculated values
4. **Test with smaller dataset first** - If building a large operation, test with 10 rows first
