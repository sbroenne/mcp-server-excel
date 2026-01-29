````markdown
# excel_conditionalformat - Server Quirks

**Rule Types**:

| Type | Description | Parameters |
|------|-------------|------------|
| `cell-value` | Format based on cell value comparison | operatorType + formula1 (+ formula2 for between) |
| `expression` | Format based on formula result | formula only |

**Operators (for cell-value type)**:

| Operator | Description | Formulas Required |
|----------|-------------|-------------------|
| `equal` | Cell equals value | formula1 |
| `not-equal` | Cell doesn't equal value | formula1 |
| `greater` | Cell greater than value | formula1 |
| `less` | Cell less than value | formula1 |
| `greater-equal` | Cell greater or equal | formula1 |
| `less-equal` | Cell less or equal | formula1 |
| `between` | Cell between two values | formula1 AND formula2 |
| `not-between` | Cell not between two values | formula1 AND formula2 |

**Format Options**:

- `interiorColor`: Background fill color as `#RRGGBB` hex
- `fontColor`: Text color as `#RRGGBB` hex
- `fontBold`: `true` or `false`
- `fontItalic`: `true` or `false`
- `borderStyle`: Excel border style name
- `borderColor`: Border color as `#RRGGBB` hex

**Actions**:

| Action | Description |
|--------|-------------|
| `add-rule` | Add conditional formatting rule to range |
| `clear-rules` | Remove all conditional formatting from range |

**Formula Notes**:

- For `cell-value` type: formula1/formula2 can be numbers, strings, or cell references
- For `expression` type: formula must return TRUE/FALSE
- Formulas use the top-left cell perspective (e.g., `=$A1>100` for relative rows)
- Use absolute references (`$A$1`) when comparing to a fixed cell

**Examples**:

**Highlight cells greater than 100:**
```json
{
  "action": "add-rule",
  "rangeAddress": "A1:A10",
  "ruleType": "cell-value",
  "operatorType": "greater",
  "formula1": "100",
  "interiorColor": "#FFFF00"
}
```

**Highlight cells between 50 and 100:**
```json
{
  "action": "add-rule",
  "rangeAddress": "A1:A10",
  "ruleType": "cell-value",
  "operatorType": "between",
  "formula1": "50",
  "formula2": "100",
  "interiorColor": "#90EE90"
}
```

**Highlight row if column A is "Active" (expression):**
```json
{
  "action": "add-rule",
  "rangeAddress": "A1:D10",
  "ruleType": "expression",
  "formula": "=$A1=\"Active\"",
  "interiorColor": "#90EE90"
}
```

**CLI Usage**:

```bash
# Add rule: highlight values > 100 in yellow
excelcli conditionalformat add-rule --session <id> --sheet "Data" --range "B2:B100" \
  --rule-type "cell-value" --operator "greater" --formula1 "100" --interior-color "#FFFF00"

# Add expression rule: highlight entire row if column A is "Error"
excelcli conditionalformat add-rule --session <id> --sheet "Data" --range "A2:E100" \
  --rule-type "expression" --formula "=\$A2=\"Error\"" --interior-color "#FF0000" --font-color "#FFFFFF"

# Clear all rules from range
excelcli conditionalformat clear-rules --session <id> --sheet "Data" --range "A1:E100"
```

**Common Mistakes**:

- Using `cell-value` type without `operatorType` → Error
- Using `between` without both formula1 AND formula2 → Error
- Forgetting `$` in expression formulas → Rule applies incorrectly across rows/columns
- Colors without `#` prefix → May not apply correctly

**Best Practices**:

1. Test expression formulas in Excel first to verify logic
2. Use `clear-rules` before applying new rules if replacing existing formatting
3. For row-based highlighting, apply rule to full range (not just one column)
4. Use relative row references (`$A1`) and absolute column references for row highlighting

````