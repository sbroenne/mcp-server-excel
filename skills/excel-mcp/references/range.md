# range - Number Formats and Cell Formatting

**IMPORTANT: Always use US format codes.** The server automatically translates to the user's locale.

## Two Tools for Range Formatting

| Use | Tool | Action | When |
|-----|------|--------|------|
| Semantic status / document hierarchy | `range_format` | `set-style` | `Good`/`Bad`/`Neutral` (have fills, theme-aware); `Heading 1/2/3`; `Normal` to reset |
| Coloured header rows / custom branding | `range_format` | `format-range` | Any fill colour, custom font colour, alignment — Heading styles have NO fill |
| Number display format | `range` | `set-number-format` | Dates, currency, percentages |

## Quick Pattern: Header Row With Fill Colour

`set-style('Heading 1')` does **not** apply a fill — use `format-range` for coloured headers.
Pass ALL properties in **one call**:

```
range_format(action: 'format-range', rangeAddress: 'A1:D1',
    bold: true,
    fillColor: '#4472C4',
    fontColor: '#FFFFFF',
    horizontalAlignment: 'center')
```

## Quick Pattern: Semantic Status Cells

Use `set-style` when the meaning (Good/Bad/Neutral) matters and theme-awareness is useful:

```
range_format(action: 'set-style', rangeAddress: 'B2:B10', styleName: 'Good')
range_format(action: 'set-style', rangeAddress: 'C2:C10', styleName: 'Bad')
```

## format-range Properties

| Property | Type | Example |
|----------|------|---------|
| `bold` | bool | `true` |
| `italic` | bool | `true` |
| `underline` | bool | `true` |
| `fontSize` | number | `14` |
| `fontName` | string | `"Calibri"` |
| `fontColor` | hex color | `"#FFFFFF"` |
| `fillColor` | hex color | `"#4472C4"` |
| `horizontalAlignment` | string | `"center"`, `"left"`, `"right"` |
| `verticalAlignment` | string | `"middle"`, `"top"`, `"bottom"` |
| `wrapText` | bool | `true` |
| `borderStyle` | string | `"thin"`, `"medium"`, `"thick"` |
| `borderColor` | hex color | `"#000000"` |
| `orientation` | int | `-90` to `90` (degrees) |

## set-style Presets

Built-in style names: `Normal`, `Heading 1`, `Heading 2`, `Heading 3`, `Heading 4`, `Title`, `Good`, `Bad`, `Neutral`, `Currency`, `Percent`, `Comma`

```
range_format(action: 'set-style', rangeAddress: 'A1:D1', styleName: 'Heading 1')
```

## Format Codes

| Type | Code | Example |
|------|------|---------|
| Number | `#,##0.00` | 1,234.56 |
| Dollar | `$#,##0.00` | $1,234.56 |
| Euro | `€#,##0.00` | €1,234.56 |
| Pound | `£#,##0.00` | £1,234.56 |
| Yen | `¥#,##0` | ¥1,235 |
| Percent | `0.00%` | 12.34% |
| Date (ISO) | `yyyy-mm-dd` | 2023-03-15 |
| Date (US) | `mm/dd/yyyy` | 03/15/2023 |
| Date (EU) | `dd/mm/yyyy` | 15/03/2023 |
| Time | `h:mm AM/PM` | 2:30 PM |
| Time (24h) | `hh:mm:ss` | 14:30:00 |
| Text | `@` | (as-is) |

All format codes are auto-translated to the user's locale. Use US codes (d/m/y for dates, . for decimal, , for thousands).

## Actions

**SetNumberFormat**: Apply one format to entire range.

- `formatCode`: Format code from table above

**SetNumberFormats**: Apply different formats per cell.

- `formats`: 2D array matching range dimensions
- Example: `[["$#,##0.00", "0.00%"], ["mm/dd/yyyy", "General"]]`
