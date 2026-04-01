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
1. calculation_mode(action: 'set-mode', mode: 'manual')   → Disable auto-recalc
2. Perform all data writes (range set-values, set-formulas)
3. calculation_mode(action: 'calculate', scope: 'workbook') → Recalculate once at end
4. calculation_mode(action: 'set-mode', mode: 'automatic')  → Restore default
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
| `calculate` | Trigger recalculation | `scope: "workbook"` (all formulas) or `"sheet"` (with `sheetName`) or `"range"` (with `sheetName` + `rangeAddress`) |

## Common Scenarios

### Scenario: Create Sales Table with Formulas

Task: Add 100 rows of product data with unit price, quantity, and total formulas.

```
1. calculation_mode set-mode manual
2. range set-values (add 100 rows: columns A-C values)
3. range set-formulas (add 100 total formulas in column D)
4. calculation_mode calculate workbook  (calculates all 100 formulas at once)
5. calculation_mode set-mode automatic
```

**Performance:** ~2-3 seconds total (vs ~30+ seconds if automatic after every cell)

### Scenario: Dashboard with Multiple Sections

Task: Create 5 sections with headers, data, and subtotal formulas.

```
1. calculation_mode set-mode manual
2. Section 1: set-values + set-formulas
3. Section 2: set-values + set-formulas
4. Section 3: set-values + set-formulas
5. Section 4: set-values + set-formulas
6. Section 5: set-values + set-formulas
7. calculation_mode calculate workbook  (all 5 sections recalc together)
8. calculation_mode set-mode automatic
```

## Best Practices

1. **Always restore automatic mode** - Never leave manual mode enabled, users expect auto-recalc
2. **Use workbook scope for calculate** - Simplest and fastest
3. **Verify calculation completed** - After step 3, data should show final calculated values
4. **Test with smaller dataset first** - If building a large operation, test with 10 rows first
