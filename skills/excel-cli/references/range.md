> **CLI syntax note:** This shared domain guide may use MCP-style `tool(action: ...)` examples as conceptual shorthand. Do not translate or paste those calls mechanically. Use the exact commands and kebab-case options in [cli-commands.md](./cli-commands.md) or live `--help`; notably, MCP `file` open/close maps to CLI `session` open/close, and MCP `worksheet` maps to CLI `sheet`.

# range - Number Formats and Cell Formatting

**IMPORTANT: Always use US format codes.** The server automatically translates to the user's locale.

**Discoverability note:** number display formats live on `range`; visual styling and auto-fit live on `range_format`.

## Formatting Split Across Two Tools

| Use | Tool | Action | When |
|-----|------|--------|------|
| Semantic status / document hierarchy | `range_format` | `set-style` | `Good`/`Bad`/`Neutral` (have fills, theme-aware); `Heading 1/2/3`; `Normal` to reset |
| Coloured header rows / custom branding | `range_format` | `format-range` | Any fill colour, custom font colour, alignment — Heading styles have NO fill |
| Repeated shared styling across disjoint ranges | `range_format` | `format-ranges` | Same worksheet, same formatting payload, fewer round-trips |
| Number display format | `range` | `set-number-format` / `set-number-formats` | Dates, currency, percentages, text display |
| Auto-fit layout | `range_format` | `auto-fit-columns` / `auto-fit-rows` | After writing variable-width data or wrapped text |

If you are looking for percentage, currency, date, or text display formatting, use `range`, not `range_format`.
If you are looking for auto-fit, width, height, borders, fill, or font styling, use `range_format`.
If you need the same styling on multiple non-contiguous ranges, use `format-ranges` instead of repeating `format-range`.

## Quick Pattern: Write, Format, Auto-Fit

```
range(action: 'set-values', range_address: 'A1:D4', values: [[...], [...]])
range(action: 'set-number-format', range_address: 'C2:D4', format_code: '$#,##0.00')
range_format(action: 'auto-fit-columns', range_address: 'A:D')
```

## Quick Pattern: Repeated Section Headers

Use `format-ranges` when the same header or section style repeats across disjoint ranges on one sheet:

```
range_format(action: 'format-ranges',
    range_addresses: ['A1:G1', 'A12:G12', 'A24:G24'],
    bold: true,
    fill_color: '#243F60',
    font_color: '#FFFFFF',
    horizontal_alignment: 'center')
```

All target ranges are validated before formatting begins. If any target range is invalid, nothing is formatted.

## Quick Pattern: Header Row With Fill Colour

`set-style('Heading 1')` does **not** apply a fill — use `format-range` for coloured headers.
Pass ALL properties in **one call**:

```
range_format(action: 'format-range', range_address: 'A1:D1',
    bold: true,
    fill_color: '#4472C4',
    font_color: '#FFFFFF',
    horizontal_alignment: 'center')
```

## Quick Pattern: Semantic Status Cells

Use `set-style` when the meaning (Good/Bad/Neutral) matters and theme-awareness is useful:

```
range_format(action: 'set-style', range_address: 'B2:B10', style_name: 'Good')
range_format(action: 'set-style', range_address: 'C2:C10', style_name: 'Bad')
```

## format-range Properties

| Property | Type | Example |
|----------|------|---------|
| `bold` | bool | `true` |
| `italic` | bool | `true` |
| `underline` | bool | `true` |
| `font_size` | number | `14` |
| `font_name` | string | `"Calibri"` |
| `font_color` | hex color | `"#FFFFFF"` |
| `fill_color` | hex color | `"#4472C4"` |
| `horizontal_alignment` | string | `"center"`, `"left"`, `"right"` |
| `vertical_alignment` | string | `"middle"`, `"top"`, `"bottom"` |
| `wrap_text` | bool | `true` |
| `border_style` | string | `"thin"`, `"medium"`, `"thick"` |
| `border_color` | hex color | `"#000000"` |
| `orientation` | int | `-90` to `90` (degrees) |

## set-style Presets

Built-in style names: `Normal`, `Heading 1`, `Heading 2`, `Heading 3`, `Heading 4`, `Title`, `Good`, `Bad`, `Neutral`, `Currency`, `Percent`, `Comma`

```
range_format(action: 'set-style', range_address: 'A1:D1', style_name: 'Heading 1')
```

## Format Codes

| Type | Code | Example (en-US) |
|------|------|-----------------|
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

**The `Example` column assumes en-US regional settings — rendering is locale-dependent.** Excel
interprets the `,` and `.` in a format code according to the user's locale, and the same is true of
the `d`/`m`/`y` date codes. So `$#,##0.00` displays as `$1,234.56` on en-US but `$1.234,56` on de-DE,
and `mm/dd/yyyy` follows the locale's date separator. This is correct behaviour, not a bug — do not
rewrite a format code because a screenshot shows swapped separators, and do not tell the user a
literal rendering without accounting for their regional settings.

After applying a number format, run `range_format auto-fit-columns` — formatted values are wider
than raw ones and will render as `#####` at the default column width.

## Actions

**SetNumberFormat**: Apply one format to entire range.

- `format_code`: Format code from table above

**SetNumberFormats**: Apply different formats per cell.

- `formats`: 2D array matching range dimensions
- Example: `[["$#,##0.00", "0.00%"], ["mm/dd/yyyy", "General"]]`

## Threaded Comments (`range_link`)

Use modern threaded comments only when the installed desktop Excel build exposes them:

```text
range_link(action: 'add-threaded-comment', sheet_name: 'Review', cell_address: 'B2', text: 'Check this value')
range_link(action: 'add-threaded-comment-reply', sheet_name: 'Review', cell_address: 'B2', text: 'Confirmed')
range_link(action: 'list-threaded-comments', sheet_name: 'Review', cell_address: 'B2')
range_link(action: 'delete-threaded-comment', sheet_name: 'Review', cell_address: 'B2')
```

These actions expose local Excel PIA comment text, author, date, and replies. Microsoft 365 service features such as @mentions, assignments, reactions, presence, sharing, and coauthoring state are not available through local Excel COM.

## Related `range_format` Actions

- `auto-fit-columns`: Fit column widths to content after writing data
- `auto-fit-rows`: Fit row heights to wrapped or multi-line content
- `format-range`: Apply fills, fonts, borders, and alignment
- `format-ranges`: Apply one shared formatting payload to multiple ranges on the same worksheet
- `set-style`: Apply named Excel styles such as `Good`, `Bad`, or `Heading 1`

## Hyperlink Lifecycle

Use `range_link` for cell hyperlinks:

| Action | Purpose |
|--------|---------|
| `add-hyperlink` | Add an external URL/file link or an internal workbook target |
| `update-hyperlink` | Change an existing target, display text, or tooltip |
| `get-hyperlink` | Read the hyperlink in one cell |
| `list-hyperlinks` | List all hyperlinks on a worksheet |
| `remove-hyperlink` | Remove hyperlinks while preserving cell content |

For an internal link, omit `url` and pass a `sub_address` such as `'Summary'!A1`. For partial updates, omitted values remain unchanged; pass an empty string to clear the URL, sub-address, or tooltip.
