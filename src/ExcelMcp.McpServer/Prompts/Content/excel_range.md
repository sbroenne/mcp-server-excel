# excel_range Tool

**Actions**: get-values, set-values, get-formulas, set-formulas, get-number-formats, set-number-format, set-number-formats, clear-all, clear-contents, clear-formats, copy, copy-values, copy-formulas, insert-cells, delete-cells, insert-rows, delete-rows, insert-columns, delete-columns, find, replace, sort, get-used-range, get-current-region, get-range-info, add-hyperlink, remove-hyperlink, list-hyperlinks, get-hyperlink, set-style, format-range, validate-range, get-validation, remove-validation, autofit-columns, autofit-rows, merge-cells, unmerge-cells, get-merge-info, add-conditional-formatting, clear-conditional-formatting, set-cell-lock, get-cell-lock

**When to use excel_range**:
- Cell values, formulas, formatting in existing worksheets
- Use excel_table for structured tables (AutoFilter, structured refs)
- Use excel_powerquery for external data sources
- Use excel_worksheet for sheet lifecycle (create, delete, rename)

**Server-specific behavior**:
- Single cell returns [[value]] as 2D array, not scalar
- Named ranges: use sheetName="" (empty string)
- Batch mode recommended for 3+ operations on same file
- For performance: Use set-number-formats (plural) to format multiple cells at once

**Action disambiguation**:
- clear-all: Removes content + formatting
- clear-contents: Removes content only, preserves formatting
- clear-formats: Removes formatting only, preserves content
- copy: Copies everything (values + formulas + formatting)
- copy-values: Copies only values (no formulas, no formatting)
- copy-formulas: Copies only formulas (no values, no formatting)
- set-number-format: Apply one format code to entire range
- set-number-formats: Apply different format codes to each cell (2D array)
- set-style: Apply built-in Excel style (Heading 1, Total, Input, etc.) - RECOMMENDED for formatting
- format-range: Apply custom formatting (font, color, borders) - Use only when built-in styles don't fit
- get-used-range: Returns actual data bounds (ignores empty cells)
- get-current-region: Returns contiguous region around a cell

**Common mistakes**:
- Expecting single cell to return scalar → Always returns 2D array [[value]]
- Using sheetName for named ranges → Use sheetName="" for named ranges
- Not using batch mode for multiple operations → 75-90% slower
- Setting number format per cell → Use set-number-format for entire range instead

**Workflow optimization**:
- Multiple range operations? Use begin_excel_batch first
- Combine operations: set values + format + validate in one batch session
- Use get-used-range to discover data bounds before operations
- **Formatting?** Try set-style with built-in styles first (faster, theme-aware, consistent)
- Only use format-range for brand-specific colors or one-off custom designs
