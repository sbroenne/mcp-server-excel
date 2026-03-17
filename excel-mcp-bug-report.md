# Excel MCP Bug Report — Formatting & Data Issues

**Date:** March 16, 2026  
**Extension:** `sbroenne.excel-mcp` v1.8.30  
**OS:** Windows  
**Client:** VS Code with GitHub Copilot agent  
**File(s) affected:** Multiple `.xlsx` workbooks created via MCP session

---

## Bug 1: `ArgumentOutOfRangeException` When Writing Ranges Wider Than 13 Columns

### Severity: High

### Description

Calling `mcp_excel-mcp_range` with `action: set-values` on any range spanning more than 13 columns throws an unhandled `ArgumentOutOfRangeException` and fails silently without writing any data.

### Steps to Reproduce

1. Open or create any workbook via `mcp_excel-mcp_file`.
2. Call `mcp_excel-mcp_range` with `action: set-values` and a `range_address` spanning 14 or more columns, e.g. `A1:N5`.
3. Provide a matching `values` array.

### Expected Result

Data is written to all 14 columns across the specified rows.

### Actual Result

```
ArgumentOutOfRangeException: Index was out of range. Must be non-negative and less than the size of the collection.
```

No data is written. The call fails entirely.

### Workaround

Split each write into two separate calls — columns A–G first, then H–N. Example:

```jsonc
// Call 1
{ "action": "set-values", "range_address": "A1:G5", "values": [...] }

// Call 2
{ "action": "set-values", "range_address": "H1:N5", "values": [...] }
```

This works but doubles the number of MCP round-trips and requires the caller to manually split the data array.

### Suggested Fix

Remove or increase the hard column-count limit inside the COM interop range-write routine. The Excel COM object itself has no such restriction (maximum 16,384 columns per worksheet).

---

## Bug 2: Newly Created Sheets Reject Data Writes Until Cell A1 Is Written First

### Severity: Medium

### Description

After creating a new worksheet via `mcp_excel-mcp_worksheet` with `action: create`, any subsequent `set-values` call on a range other than `A1` fails or is silently ignored. The sheet must first receive a single-cell write to `A1` before it will accept writes to other ranges.

### Steps to Reproduce

1. Create a new sheet: `mcp_excel-mcp_worksheet` `action: create`, `sheet_name: "My Sheet"`.
2. Immediately attempt `mcp_excel-mcp_range` `action: set-values`, `range_address: A3:G10`.

### Expected Result

Data is written to `A3:G10` on the new sheet.

### Actual Result

The call returns success but no data appears in the cells. The sheet activates only after a write to `A1`.

### Workaround

Always write a single value to `A1` immediately after creating a sheet, then proceed with other range writes:

```jsonc
{ "action": "set-values", "range_address": "A1", "values": [["Header Text"]] }
// Now subsequent writes to other ranges succeed
```

### Suggested Fix

Ensure sheet activation/selection is triggered within the `worksheet create` operation itself, so the sheet is immediately ready for arbitrary range writes without requiring an `A1` initialisation step.

---

## Bug 3: `format-range` Does Not Support Number/Percentage Format Strings

### Severity: Medium

### Description

The `mcp_excel-mcp_range_format` tool (`action: format-range`) exposes styling properties (`fill_color`, `font_color`, `bold`, `font_size`) but has no parameter for Excel number format strings. As a result, values written as decimals (e.g. `0.26`) cannot be formatted as percentages (`26%`) or currency (`$0.00`) via MCP — formatting must be done by post-processing the file in Excel manually.

### Steps to Reproduce

1. Write a decimal value `0.26` to a cell via `set-values`.
2. Attempt to apply a percentage number format via `format-range`.
3. Observe that `format-range` has no `number_format` parameter.

### Expected Result

A `number_format` parameter (accepting Excel format strings such as `"0%"`, `"#,##0.00"`, `"$#,##0.00"`) should be available to apply cell number formatting.

### Actual Result

No number formatting is possible via MCP. Values are stored as raw decimals and rendered without formatting context (e.g. `0.26` instead of `26%`).

### Suggested Fix

Add a `number_format` parameter to `mcp_excel-mcp_range_format`:

```jsonc
{
  "action": "format-range",
  "range_address": "E6:E12",
  "number_format": "0%",
  "sheet_name": "Savings Analysis",
  "session_id": "..."
}
```

---

## Bug 4: No Batch / Multi-Range Formatting Support — Requires Excessive Round-Trips

### Severity: Low–Medium

### Description

`mcp_excel-mcp_range_format` applies formatting to exactly one range per call. Formatting a typical report sheet (title row, column headers, data rows, section headers, warning rows, total rows) requires 10–20 separate MCP calls. This is slow, chatty, and increases the risk of partial formatting if a session is interrupted.

### Steps to Reproduce

1. Create a workbook with 4 worksheets, each with multiple formatted sections.
2. Count the number of `format-range` calls required to fully style all sheets.

### Observed

A 4-sheet workbook required approximately **50+ individual `format-range` calls** to apply consistent styling (title, subheading, column header, data row, warning row, total row, note row patterns).

### Suggested Enhancement

Support either:

- **A `ranges` array parameter** — apply the same format to multiple non-contiguous ranges in one call.
- **A named style / format preset system** — define a style once and apply it by name to multiple ranges.
- **A batch formatting action** — accept an array of `{ range_address, format_properties }` objects in a single call.

Example proposed API:

```jsonc
{
  "action": "format-ranges-batch",
  "formats": [
    { "range_address": "A1:G1", "fill_color": "#0078D4", "font_color": "#FFFFFF", "bold": true, "font_size": 16 },
    { "range_address": "A3:G3", "fill_color": "#243F60", "font_color": "#FFFFFF", "bold": true },
    { "range_address": "A4:G10", "fill_color": "#DEEAF1" }
  ],
  "sheet_name": "Sheet1",
  "session_id": "..."
}
```

---

## Bug 5: No Auto-Fit Column Width Support

### Severity: Low

### Description

There is no MCP action to auto-fit column widths to content. When writing data with variable-length strings (e.g. long recommendation text, URLs, notes), columns remain at default width and content is clipped or hidden in Excel.

### Suggested Enhancement

Add an `auto-fit` action to `mcp_excel-mcp_range` or `mcp_excel-mcp_worksheet`:

```jsonc
{
  "action": "auto-fit-columns",
  "range_address": "A:G",
  "sheet_name": "Sheet1",
  "session_id": "..."
}
```

---

## Environment

| Property | Value |
|---|---|
| Extension | `sbroenne.excel-mcp` v1.8.30 |
| VS Code | Windows |
| Excel | Microsoft Excel (COM interop, desktop) |
| Date reported | March 16, 2026 |

## Reproduction Files

Issues were consistently reproduced across the following workbooks:

- `Azure Sizing - IBM Workloads.xlsx`
- `Azure Sizing - AWS Workloads.xlsx`
- `Azure Sizing - RET Azure Workloads.xlsx`
- `Azure Sizing - GCP Workloads.xlsx`

All files were created from scratch via Excel MCP session in the same VS Code agent conversation.

---

*Report generated by GitHub Copilot agent during Azure workload sizing analysis session.*
