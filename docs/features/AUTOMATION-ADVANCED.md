# Automation & Advanced Features

Extend workbook workflows with code, live window control, What-If Analysis,
and mapped XML data.

[Back to the feature overview](../../FEATURES.md)

Use CLI help or MCP tool descriptions for current actions and inputs. The
summaries below focus on capabilities and their important requirements.

---

## VBA Macros (6 operations)

- **Inspect code:** Discover VBA components and procedures and read component code.
- **Maintain modules:** Import standard modules, update existing component code, or remove components.
- **Run procedures:** Execute existing VBA procedures with supported parameters.

Listing or editing the VBA project requires Excel's manually configured trust
setting. Running an existing macro does not require project-inspection trust,
but Excel's macro security still applies. ExcelMcp does not enable trust or
bypass security automatically. Retain a macro-enabled file format when saving code.

[VBA walkthrough](../guides/RUN-VBA-MACROS.md)

---

## Python in Excel (2 operations)

- **Add Python calculations:** Write Python-in-Excel formulas that can reference live worksheet data.
- **Read results:** Wait for cloud execution and read supported calculated values, distinguishing pending work from unavailable Python support.

This is Microsoft's licensed Python in Excel service, not local Python
execution. It requires a supported Microsoft 365 account and internet access.
Rich Python objects are not necessarily readable as ordinary values through
Excel automation; choose worksheet-value output when that is what the task needs.

A successful formula write does not mean the cloud calculation has finished.
An unavailable service and pending calculation need different recovery choices.

[Cloud calculation and failure guidance](../reference/behavioral-rules.md#python-in-excel)

---

## Window Management (16 operations)

- **Watch work live:** Show, hide, position, and arrange the session's Excel window.
- **Inspect context:** Read its active worksheet, selection, chart, window state, and view.
- **Improve navigation:** Adjust zoom, display options, frozen panes, and movable splits.
- **Show progress:** Set and clear Excel status-bar text.

Preserve the user's existing visibility preference. Keeping a workbook open
does not mean showing its window. Window changes affect the selected session,
not unrelated Excel processes, and visible interaction needs a suitable desktop.

[Window workflow guidance](../reference/window.md)

---

## What-If Analysis (8 operations)

- **Find an input:** Use Goal Seek to adjust one input until a formula reaches a numeric target.
- **Compare assumptions:** Maintain named scenarios, apply their input values, and produce scenario summaries.
- **Explore sensitivity:** Build one- or two-variable Excel data tables.

Goal Seek and applying scenarios change live cells; they are not read-only
inspection. Read the resulting inputs and outputs. Data tables can be
calculation-intensive and need the worksheet layout prepared first.

Solver is not exposed here. It is an optional VBA add-in that requires separate
user configuration and macro-security decisions.

[What-If workflow guidance](../reference/analysis.md)

---

## XML Maps (6 operations)

- **Map structured data:** Manage workbook XML schemas and bind worksheet cells or columns to XML paths.
- **Exchange data:** Import and export mapped XML without dialogs.

Use existing mappings when they describe the intended data. Imports can write
worksheet cells or create a new mapped Table. Removing a map does not clear
previously imported cells.

External schema dependencies, DTDs, and schema-location attributes are rejected
so Excel cannot resolve unexpected external resources.

[XML workflow guidance](../reference/xmlmap.md)

---

## Related feature areas

- [Data & analytics](DATA-ANALYTICS.md) - combine code with queries and analytical models
- [Cells & workbooks](CELLS-WORKBOOKS.md) - manage the inputs, formulas, and files used by automation
- [Charts & visualization](CHARTS-VISUALS.md) - present automated results
- [Real Excel automation vs. file-parser libraries](../guides/EXCEL-COM-VS-FILE-PARSERS.md)
