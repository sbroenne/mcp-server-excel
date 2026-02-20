# range - Number Formats and Cell Formatting

**IMPORTANT: Always use US format codes.** The server automatically translates to the user's locale.

## Two Tools for Range Formatting

| Use | Tool | Action | When |
|-----|------|--------|------|
| Visual formatting (bold, color, alignment) | `range_format` | `format-range` | Header rows, highlights, custom styling |
| Built-in style presets | `range_format` | `set-style` | Consistent themed formatting |
| Number display format | `range` | `set-number-format` | Dates, currency, percentages |

## Quick Pattern: Professional Header Row

```
range_format(action: 'format-range', rangeAddress: 'A1:D1',
    bold: true,
    fillColor: '#4472C4',
    fontColor: '#FFFFFF',
    horizontalAlignment: 'center')
```

All properties in **one call** — do not split into multiple calls.

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
