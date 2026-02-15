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

1. **Verify chart placement**: After creating or repositioning charts, capture a screenshot to confirm no overlap with data
2. **Capture relevant area**: Use `capture` with a specific range rather than `capture-sheet` when you only need part of the worksheet
3. **Use after multi-step operations**: Screenshots are most valuable after a sequence of formatting, layout, or chart operations
4. **MCP returns image directly**: The image is returned as native ImageContent — no file handling needed
5. **CLI returns base64 JSON**: Parse the `imageBase64` field from the JSON response

## Common Patterns

### Chart Overlap Verification
```
1. range(get-used-range) → "A1:D20"
2. chart(create-from-range, sourceRange='A1:D20', targetRange='F2:K15')
3. screenshot(capture, rangeAddress='A1:K20')
   → Visually confirm chart is positioned next to data, not on top of it
```

### Dashboard Layout Check
```
1. Create multiple charts and tables
2. screenshot(capture-sheet)
   → Verify overall dashboard layout, spacing, and alignment
```
