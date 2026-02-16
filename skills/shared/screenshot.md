# Screenshot & Visual Verification Reference

## Tools

- **`screenshot`**: Capture worksheet content as PNG images

## Actions

| Action | Purpose | Parameters |
|--------|---------|------------|
| `capture` | Capture a specific range | `rangeAddress` (default: A1:Z30), `sheetName` |
| `capture-sheet` | Capture entire used area | `sheetName` |

## When to Use Screenshots

### After Chart Creation or Positioning
```
1. chart(create-from-range, ..., targetRange='F2:K15')
2. screenshot(capture, rangeAddress='A1:O25')  → Verify chart doesn't overlap data
```

### After Complex Formatting
```
1. range(set-number-format, ...)
2. conditionalformat(add-rule, ...)
3. screenshot(capture-sheet)  → Verify formatting looks correct
```

### After PivotTable Layout Changes
```
1. pivottable(add-row-field, ...)
2. pivottable(add-value-field, ...)
3. screenshot(capture-sheet)  → Verify layout and field arrangement
```

## Best Practices

1. **Verify chart placement**: After creating or repositioning charts, capture a screenshot to confirm no overlap with data or other charts
2. **Capture relevant area**: Use `capture` with a specific range rather than `capture-sheet` when you only need part of the worksheet
3. **Use after multi-step operations**: Screenshots are most valuable after a sequence of formatting, layout, or chart operations
4. **MCP returns image directly**: The image is returned as native ImageContent — no file handling needed
5. **CLI with `--output`**: Use `--output screenshot.png` to save the captured image directly as a PNG file

## Common Patterns

### Chart Overlap Verification
```
1. range(get-used-range) → "A1:D20"
2. chart(create-from-range, sourceRange='A1:D20', targetRange='F2:K15')
3. screenshot(capture, rangeAddress='A1:K20')
   → Visually confirm chart is positioned next to data, not on top of it
```

### Multi-Chart Dashboard Layout
```
When creating dashboards with multiple charts:

1. get-used-range → Know where data ends
2. Create Chart 1 with targetRange below/beside data
3. Create Chart 2 with targetRange that does NOT overlap Chart 1
4. Create Chart 3, Chart 4, etc. — each in a non-overlapping targetRange
5. screenshot(capture-sheet) → Verify NO charts overlap each other or data

Key rules for multi-chart layouts:
- Use targetRange for every chart — never rely on default positioning
- Leave at least 1-2 rows/columns between charts
- Place charts in a grid pattern (e.g., 2x2) below the data area
- If overlap detected, use chart(fit-to-range) to reposition
```

### Dashboard Layout Check
```
1. Create multiple charts and tables
2. screenshot(capture-sheet)
   → Verify overall dashboard layout, spacing, and alignment
3. If issues found: reposition with chart(fit-to-range), then screenshot again
```
