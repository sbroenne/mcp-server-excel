# excel_range - Number Formats

**IMPORTANT: Always use US format codes.** The server automatically translates to the user's locale.

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
