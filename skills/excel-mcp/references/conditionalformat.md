# conditionalformat - Server Quirks

**Rule Types**:

| Type | Description | Parameters |
|------|-------------|------------|
| `cell-value` | Format based on cell value comparison | operatorType + formula1 (+ formula2 for between) |
| `expression` | Format based on formula result | formula only |
| `color-scale` | 2- or 3-color gradient across the range | colorScaleMin/Mid/Max Type/Value/Color |
| `data-bar` | In-cell bars proportional to value | dataBarColor, dataBarNegativeColor, dataBarDirection, dataBarShowValue, dataBarMin/Max Type/Value |
| `icon-set` | Icons (arrows, traffic lights, etc.) per value band | iconSetId, iconSetReverse, iconSetShowIconOnly, iconThreshold1..4 Type/Value |
| `top10` | Highlight top/bottom N (or percent) | rank, top10Percent, topBottom + formatting |
| `above-average` | Highlight values above/below average | aboveBelow + formatting |
| `time-period` | Highlight dates in a period | datePeriod + formatting |
| `unique-values` | Highlight unique (or duplicate) values | formatting |
| `blanks-condition` | Highlight blank cells | formatting |

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

**Visual rule parameters** (used by the corresponding `ruleType`):

- **color-scale**: `colorScaleMinType`/`colorScaleMidType`/`colorScaleMaxType`
  (`minimum`, `maximum`, `number`, `percent`, `percentile`, `formula`), matching
  `...Value` (when the type needs one) and `...Color` (`#RRGGBB`). Supplying any `mid*`
  parameter creates a 3-color scale, otherwise a 2-color scale.
- **data-bar**: `dataBarColor` (`#RRGGBB`), `dataBarNegativeColor`, `dataBarDirection`
  (`context`, `leftToRight`, `rightToLeft`), `dataBarShowValue` (`true`/`false`),
  `dataBarMinType`/`dataBarMaxType` (+ matching values).
- **icon-set**: `iconSetId` (e.g. `3Arrows`, `3TrafficLights1`, `4Ratings`, `5Quarters`),
  `iconSetReverse` (`true`/`false`), `iconSetShowIconOnly` (`true`/`false`),
  `iconThreshold1Type..iconThreshold4Type` (+ matching `...Value`) for the editable bands.
- **top10**: `rank` (count or percent), `top10Percent` (`true`/`false`),
  `topBottom` (`top`/`bottom`), plus standard formatting options.
- **above-average**: `aboveBelow` (`aboveAverage`, `belowAverage`, `aboveStdDev`,
  `belowStdDev`, `equalAboveAverage`, `equalBelowAverage`), plus formatting options.
- **time-period**: `datePeriod` (`today`, `yesterday`, `tomorrow`, `last7Days`,
  `thisWeek`, `lastWeek`, `nextWeek`, `thisMonth`, `lastMonth`, `nextMonth`), plus formatting.

**Actions**:

| Action | Description |
|--------|-------------|
| `add-rule` | Add conditional formatting rule to range |
| `clear-rules` | Remove all conditional formatting from range |
| `list-rules` | Read existing rules for a range (type, operator, formulas, applies-to, priority, formatting) |
| `list-worksheet-rules` | Read all rules across an entire worksheet, each with its applies-to range |

**Reading rules (`list-rules` / `list-worksheet-rules`)**:

- Rules are returned in priority order.
- Colors are returned as `#RRGGBB` hex strings, matching the `add-rule` input format.
- Formatting fields (interiorColor, fontColor, fontBold/Italic, borderStyle/Color) are only
  present when the rule actually sets them.
- Visual rule types return their type-specific configuration so they can be fully inspected
  and round-tripped:
  - `colorScale` → `colorScaleCriteria`: array of `{ type, value?, color }` stops.
  - `dataBar` → `dataBar`: `{ fillColor, barColorNegative?, direction, showValue, minType, minValue?, maxType, maxValue? }`.
  - `iconSet` → `iconSet`: `{ id, reverse, showIconOnly, criteria: [{ operator, value?, type, icon }] }`.
  - `top10` → `top10`: `{ rank, percent, topBottom }`.
  - `aboveAverage` → `aboveBelow`: e.g. `aboveAverage`, `belowAverage`, `aboveStdDev`.
  - `timePeriod` → `datePeriod`: e.g. `today`, `last7Days`, `thisMonth`.
  Each field is only present on its matching rule type.
- Numeric `cell-value` formulas are returned in Excel's normalized form (e.g. `100` reads back
  as `=100`).

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
  "formula1": "=$A1=\"Active\"",
  "interiorColor": "#90EE90"
}
```

**3-color scale (red → yellow → green):**
```json
{
  "action": "add-rule",
  "rangeAddress": "A1:A100",
  "ruleType": "color-scale",
  "colorScaleMinType": "minimum",
  "colorScaleMinColor": "#F8696B",
  "colorScaleMidType": "percentile",
  "colorScaleMidValue": "50",
  "colorScaleMidColor": "#FFEB84",
  "colorScaleMaxType": "maximum",
  "colorScaleMaxColor": "#63BE7B"
}
```

**Data bar with value shown:**
```json
{
  "action": "add-rule",
  "rangeAddress": "B1:B100",
  "ruleType": "data-bar",
  "dataBarColor": "#638EC6",
  "dataBarDirection": "leftToRight",
  "dataBarShowValue": true
}
```

**3 traffic lights icon set:**
```json
{
  "action": "add-rule",
  "rangeAddress": "C1:C100",
  "ruleType": "icon-set",
  "iconSetId": "3TrafficLights1",
  "iconThreshold1Type": "percent",
  "iconThreshold1Value": "33",
  "iconThreshold2Type": "percent",
  "iconThreshold2Value": "67"
}
```

**CLI Usage**:

```powershell
# Add rule: highlight values > 100 in yellow
excelcli conditionalformat add-rule --session <id> --sheet "Data" --range "B2:B100" `
  --rule-type "cell-value" --operator-type "greater" --formula1 "100" --interior-color "#FFFF00"

# Add expression rule: highlight entire row if column A is "Error"
excelcli conditionalformat add-rule --session <id> --sheet "Data" --range "A2:E100" `
  --rule-type "expression" --formula1 "=`$A2=`"Error`"" --interior-color "#FF0000" --font-color "#FFFFFF"

# Clear all rules from range
excelcli conditionalformat clear-rules --session <id> --sheet "Data" --range "A1:E100"

# List rules for a range
excelcli conditionalformat list-rules --session <id> --sheet "Data" --range "A1:E100"

# List all rules on a worksheet
excelcli conditionalformat list-worksheet-rules --session <id> --sheet "Data"
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
