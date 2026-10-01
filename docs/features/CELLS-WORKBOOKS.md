# Cells & Workbooks Features

Read, write, calculate, and format cells while managing worksheets, workbooks, named ranges, and files.

[← Back to the complete feature reference](../../FEATURES.md)

---

## 📁 File Operations (5 operations)

Open, create, and close Excel workbooks. Every other tool works on a session opened here.

**Operations:**
- **List Sessions:** View all active Excel sessions
- **Open:** Open workbook and create session (returns session ID for all subsequent operations). IRM/AIP-protected files are automatically detected and opened read-only with Excel visible for credential authentication — no extra parameters needed.
- **Close:** Close session with optional save
- **Create Empty:** Create new .xlsx or .xlsm workbook
- **Test:** Report existence, extension validity, openability, and IRM/AIP requirements through `canOpen`, `isIrmProtected`, `willOpenReadOnly`, and `requiresVisibleSession`. Ordinary workbooks are opened read-only in a temporary Excel session and closed without saving.

**Workflow:** List and match the intended workbook; reuse its session or
open/create; operate; list and check its `canClose`; close only when authorized
with an explicit save choice. MCP identifiers use `session_id` in inputs,
results, and list entries; CLI JSON uses `sessionId`.
Close without saving discards all unsaved edits, including earlier work, and
has no tool-level undo.

Operations within one session execute serially, but concurrent requests and
responses have no guaranteed order. Wait for dependent calls; different
sessions can run independently.

---

## 🧮 Calculation Mode (3 operations)

Control when and how Excel recalculates formulas — useful for speeding up bulk edits.

**Operations:**
- **Get Mode:** Query current calculation mode and calculation state
- **Set Mode:** Switch between automatic, manual, and semi-automatic modes
- **Calculate:** Explicitly recalculate workbook, sheet, or range

Value/formula writes attempt to restore the prior mode rather than always
forcing calculation. Restoration can fail without failing the write; use
`get-mode` when subsequent work depends on the mode. Automatic normally
recalculates dependent formulas after restoration; manual needs explicit
calculation. Semi-automatic excludes what-if data tables, not worksheet Tables.
Successful writes do not guarantee completion of
asynchronous refreshes or Python calculations. For bulk writes, remember and
restore the prior mode, including after failure.

---

## 📋 Ranges (51 operations)

Read and write cell values, formulas, and formatting across any range of cells.

**Formatting split:** use `range` for number display formats such as dates, currency, percentages, and text display. Use `range_format` for visual styling, validation, auto-fit, and size/layout changes.

**Data Operations:**
- **Get/Set Values:** Read or write cell values. Formula errors are returned as canonical names such as `#REF!`, with cell-level diagnostics preserving the formula and raw COM code.
- **Get/Set/Validate Formulas:** Read, write, or validate formula syntax across ranges. Formula reads include the same canonical error names and cell-level diagnostics as value reads.
- **Clear All/Contents/Formats:** Clear a range's contents, formats, or both
- **Copy / Copy Values / Copy Formulas:** Copy a range, or just its values/formulas
- **Insert/Delete Cells:** Shift cells to insert or remove space
- **Insert/Delete Rows:** Insert or delete entire rows
- **Insert/Delete Columns:** Insert or delete entire columns
- **Find:** Search a range for matching values, returning up to 10 cells by default with an exact total
- **Replace:** Find and replace values in a range
- **Sort:** Sort a range by one or more columns

**Find coverage:** MCP `range_edit(action: 'find')` accepts `max_matches`;
CLI `excelcli rangeedit find` accepts `--max-matches`; batch JSON uses
`maxMatches`. The default is 10, and any positive whole number through
2147483647 is accepted. Results retain `matchingCells` and include
`totalCount`, `returnedCount`, and `truncated`. No matches returns an empty
list, zero counts, and `truncated=false`; exactly the limit is not truncated.
Exact totals require searching every match even after the return limit is
reached. The limit bounds cell details, not search time. Paging is not provided.

**Clearing has no tool-level undo:** Clear All removes values, formulas, and
formats; Clear Contents preserves formats; Clear Formats preserves
values/formulas. Check the intended target before clearing.
Until saved, an authorized no-save close can discard changes, but also loses
earlier unsaved work.

**Discovery & Utilities:**
- **Get Used Range:** Get the worksheet's used range
- **Get Current Region:** Get the contiguous data region around a cell
- **Get Range Info:** Get a range's address and dimensions

**Hyperlinks:**
- **Add Hyperlink:** Add a hyperlink to a cell
- **Update Hyperlink:** Change an external/internal target, display text, or tooltip
- **Remove Hyperlink:** Remove a hyperlink
- **List Hyperlinks:** List all hyperlinks in a range
- **Get Hyperlink:** Get a specific hyperlink's target

**Threaded Comments:**
- **Add Threaded Comment:** Add a top-level modern comment to one cell
- **List Threaded Comments:** Read a cell's comment and replies
- **Add Threaded Comment Reply:** Reply to an existing cell comment
- **Delete Threaded Comment:** Delete a comment thread and its replies

**Number Formatting (`range`):**
- **Get Number Formats:** Read number formats as a 2D array
- **Set Number Format:** Apply one number format uniformly
- **Set Number Formats:** Apply individual per-cell number formats

**Visual Formatting (`range_format`):**
- **Get Style:** Read the applied cell style
- **Set Style:** Apply a built-in Excel style
- **Format Range:** Set font, color, borders, alignment, orientation
- **Format Ranges:** Apply one shared formatting payload to multiple ranges

**Data Validation (`range_format`):**
- **Add Validation:** Add dropdown, number/date/text validation rules
- **Get Validation:** Read current validation info
- **Remove Validation:** Remove validation rules

**Merge Operations (`range_format`):**
- **Merge Cells:** Merge a range into one cell
- **Unmerge Cells:** Undo a merge
- **Get Merge Info:** Read current merge state

**Cell Protection:**
- **Set Lock Status:** Lock/unlock cells (effective once the sheet is protected)
- **Get Lock Status:** Read current cell lock status

**Sizing & Auto-Fitting (`range_format`):**
- **Auto-Fit Columns / Rows:** Resize columns or rows to fit content
- **Set Column Width / Row Height:** Set explicit column widths or row heights

---

## 📄 Worksheets (33 operations)

Add, rename, move, and manage worksheets — including tab colors, visibility, protection, legacy cell notes, inline images, shapes, and page setup.

**Lifecycle:**
- **List:** List worksheets in the workbook
- **Create:** Add a new worksheet
- **Rename:** Rename a worksheet
- **Copy:** Copy a worksheet within the workbook
- **Move:** Move a worksheet within the workbook
- **Delete:** Remove a worksheet

**Cross-Workbook Operations:**
- **Copy to File:** Copy a worksheet to another workbook (atomic)
- **Move to File:** Move a worksheet to another workbook (atomic)

**Delete has no tool-level undo:** It removes all sheet contents and can break
dependent references. Check the intended sheet and its dependencies.
**Move to File has no tool-level undo:** It removes the source sheet and saves
both files. Closing another session without saving cannot
reverse the saved transfer.

**Tab Colors:**
- **Set Tab Color:** Set a worksheet tab's RGB color
- **Get Tab Color:** Read the current tab color
- **Clear Tab Color:** Reset the tab to its default color

**Visibility:**
- **Show:** Make a worksheet visible
- **Hide:** Hide a worksheet (still shown in the Unhide dialog)
- **Very Hide:** Hide a worksheet from the Excel UI entirely
- **Get Visibility:** Read the current visibility status
- **Set Visibility:** Set visibility status directly

**Protection:**
- **Set Protection:** Protect or unprotect a worksheet
- **Get Protection:** Read the current protection state

**Cell Notes:**
- **Set Comment:** Create or update a legacy cell note through Excel's Comment COM API
- **Get Comment:** Read the current legacy cell note text
- **Clear Comment:** Remove a legacy cell note

**Images:**
- **Add Image:** Insert an image from disk and anchor it to a cell
- **Get Image Count:** Read how many images are currently on a worksheet

**Shapes:**
- **Add Shape:** Insert a basic rectangle shape and anchor it to a cell
- **Get Shape Count:** Read how many shapes are currently on a worksheet

**Page Setup:**
- **Set Page Setup:** Configure orientation, fit-to-page settings, and centering
- **Get Page Setup:** Read worksheet page setup values

**Outlines:**
- **Group / Ungroup:** Group complete row or column ranges and remove one grouping level
- **Get Outline Info:** Read outline level, hidden state, summary positions, and automatic styles
- **Set Outline Settings:** Configure summary rows/columns and automatic styles
- **Show Outline Levels:** Expand or collapse row and column groups to requested levels
- **Clear Outline:** Remove all row and column groups

---

## 📘 Workbook (15 operations)

Manage workbook metadata, protection, document properties, file variants, exports, and external links.

**Operations:**
- **Set Protection:** Protect or unprotect the current workbook, optionally with a password
- **Get Protection:** Determine whether the current workbook is protected
- **Set View Options:** Toggle workbook window gridlines and headings on or off
- **Get View Options:** Read back workbook window gridlines and headings state
- **Get Info:** Read workbook name, path, format, saved/read-only state, and protection metadata
- **List/Get/Set/Delete Document Properties:** Manage built-in and custom workbook properties
- **Save As:** Save as `.xlsx`, `.xlsm`, `.xlsb`, or `.xls` and move the active session to the new path
- **Save Copy As:** Create a same-format copy without changing the active workbook
- **Export Fixed Format:** Publish PDF or XPS with quality, page-range, and print-area controls
- **List/Update/Break External Links:** Inspect, refresh, or permanently replace linked-workbook formulas

**Break External Link has no tool-level undo:** Linked formulas become their
current values.

> Printing and print preview are intentionally excluded because physical printer output and modal preview are unsafe for unattended automation.

---

## 🏷️ Named Ranges (Parameters) (6 operations)

Manage named ranges — ideal for driving workbook parameters that Power Query and formulas react to.

**Operations:**
- **List:** List visible user-defined named ranges with references; hidden/internal Excel names (including Power Query `ExternalData_*` and AutoFilter names) are omitted before value inspection, and large ranges return metadata without materializing values
- **Read:** Get value of a named range
- **Write:** Set a named-range value. Invariant numeric and Boolean strings become typed Excel values; identifiers such as `2.0.13` remain text.
- **Create:** Create new named range
- **Update:** Modify existing named range
- **Delete:** Remove named range

**Notes:**
- **Use cases:** Manage workbook parameters without touching worksheets. Ideal for automation — update a parameter and Power Query refreshes automatically.

---

## Related feature areas

- [Data & analytics](DATA-ANALYTICS.md) — transform ranges and tables with Power Query, DAX, and PivotTables
- [Charts & visualization](CHARTS-VISUALS.md) — turn workbook data into charts and visual reports
- [Automation & advanced](AUTOMATION-ADVANCED.md) — automate workbooks with VBA, Python, and What-If Analysis
- [Example workflows](../USE-CASES.md) — see these capabilities combined in practical requests
- [Installation](../INSTALLATION.md) — choose and configure the MCP Server or CLI

## Task guides

- [Real Excel automation vs. file-parser libraries](../guides/EXCEL-COM-VS-FILE-PARSERS.md)
- [Refresh Power Query from an AI assistant](../guides/REFRESH-POWER-QUERY.md)
