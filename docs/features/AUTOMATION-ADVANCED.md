# Automation & Advanced Features

Run VBA and Python, control Excel windows, solve What-If scenarios, and work with XML Maps.

[← Back to the complete feature reference](../../FEATURES.md)

---

## 📝 VBA Macros (6 operations)

View, import, edit, and run VBA code in `.xlsm` workbooks.

**Operations:**
- **List:** List VBA components and discovered procedures
- **View:** Display component code without exporting
- **Import:** Create a new standard module from code or file input
- **Update:** Replace code in an existing VBA component
- **Delete:** Remove a VBA component by name
- **Run:** Execute a procedure with optional string parameters

**Notes:**
- Procedural/module-focused VBA support for `.xlsm` workbooks.
- Requires the manual VBA trust prerequisite in Excel (no trust-configuration command).
- Import creates standard modules; list/view also cover class, form, and document components.

**CLI example:**

```powershell
excelcli session create macros.xlsm
# Use the returned session ID
excelcli vba import --session <id> --module-name MyModule --vba-code-file code.vba
excelcli session close --session <id> --save
```

VBA imports and updates accept either inline code or `--vba-code-file`, never
both. Batch JSON uses `vbaCodeFile`; MCP uses `vba_code_file`. The file must
exist and be readable. VBA execution timeouts are integer seconds from 1 through
2147483.

---

## 🐍 Python in Excel (2 operations)

Write and read `=PY()` formulas that run in Excel's cloud Python engine.

**Operations:**
- **Set Formula:** Write a `=PY("<code>", returnType)` formula via `Range.Formula2`. `returnType` 0 = "Excel Value" (a plain value/array), 1 = "Python Object" (a rich data type card, e.g. a DataFrame). Must always be passed explicitly. If Excel immediately evaluates the formula as `#NAME?`, the operation reports that Python in Excel is unavailable instead of claiming success.
- **Get Result:** Read back the computed value, polling until Excel's calculation state and the cell's transient marker show that cloud execution has finished. If the deadline is reached, the operation reports the observed transient state rather than guessing at a stale value. A settled `#NAME?` on a `PY()` formula is reported as Python in Excel being unavailable.

**Notes:**
- **Requires:** a real Excel session signed into a licensed Microsoft 365 account with Python in Excel enabled, plus internet access — the Python code executes in a Microsoft-hosted cloud sandbox, not locally. Not available offline or with perpetual-license Excel.
- **Unavailable vs. transient:** `#NAME?` means this Excel session cannot use Python in Excel. `#BUSY!`, `#CONNECT!`, and `#BLOCKED!` remain transient cloud states and keep their existing retry behavior.
- **Data binding:** Reference live worksheet data inside the Python code with `xl("A1:A6")`, `xl("Sheet1!A1:A6")`, or a named range `xl("MyRange")` — works the same as if typed interactively.

---

## 🪧 Window Management (15 operations)

Show, position, and arrange the Excel window — great for watching the AI work in real time.

**Visibility & Focus:**
- **Show:** Make Excel visible and bring it to the foreground
- **Hide:** Hide the Excel window
- **Bring to Front:** Bring Excel to the foreground without changing visibility

**Window State & Layout:**
- **Get Info:** Get current window state (visibility, position, size, foreground status)
- **Set State:** Set window state to normal, minimized, or maximized
- **Set Position:** Set window position and size in points (left, top, width, height)
- **Arrange:** Arrange the Excel window using preset layouts

**Workbook View & Panes:**
- **Get View:** Read view type, zoom, pane state, and display options
- **Freeze / Unfreeze Panes:** Freeze rows/columns at a worksheet boundary or remove frozen panes
- **Set Split:** Configure movable horizontal and vertical panes
- **Set Zoom:** Change worksheet zoom
- **Set Display Options:** Toggle gridlines, headings, formula display, and related window options

**Status Bar:**
- **Set Status Bar:** Display custom text in Excel's status bar for real-time feedback
- **Clear Status Bar:** Restore the default status bar text

**Notes:**
- **Arrange presets:** `left-half` / `right-half` (side-by-side with other applications), `top-half` / `bottom-half` (stacked view), `center` (centered window, 60% of screen), and `full-screen` (maximized).
- **Use cases:** Interactive "agent mode" where users watch Excel respond to AI commands in real time, side-by-side layouts (Excel on one half, AI assistant on the other), and visibility changes that are reflected in session metadata.

---

## 🔬 What-If Analysis (8 operations)

Run Excel's native sensitivity analysis against live workbook formulas and input cells.

- **Goal Seek:** Adjust one changing cell until a formula reaches a numeric goal
- **List / Create / Update / Show / Delete Scenarios:** Manage named sets of changing-cell values
- **Create Scenario Summary:** Produce a standard summary worksheet or Scenario PivotTable report
- **Create Data Table:** Build one- or two-variable Excel data tables

Solver is intentionally excluded because Microsoft implements it as an optional VBA add-in requiring user enablement and macro-security configuration.

---

## 🧩 XML Maps (6 operations)

Manage workbook XML schemas, XPath mappings, and in-memory XML data exchange.

- **List / Add / Delete:** Manage workbook XML maps
- **Map Range:** Bind a cell or single-column range to an XPath
- **Import / Export XML:** Exchange mapped XML in memory without dialogs

DTDs, external XSD dependencies, and XSI schema-location attributes are rejected before Excel COM can resolve external resources.

---

## Related feature areas

- [Data & analytics](DATA-ANALYTICS.md) — combine VBA and Python with Power Query, DAX, and PivotTables
- [Cells & workbooks](CELLS-WORKBOOKS.md) — manipulate the ranges, formulas, worksheets, and files used by automation
- [Charts & visualization](CHARTS-VISUALS.md) — create polished visual output from automated workflows
- [Example workflows](../USE-CASES.md) — see these capabilities combined in practical requests
- [Installation](../INSTALLATION.md) — choose and configure the MCP Server or CLI

## Task guides

- [Run VBA macros from an AI agent](../guides/RUN-VBA-MACROS.md)
- [Real Excel automation vs. file-parser libraries](../guides/EXCEL-COM-VS-FILE-PARSERS.md)
