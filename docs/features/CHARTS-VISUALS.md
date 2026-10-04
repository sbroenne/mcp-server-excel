# Charts & Visualization Features

Turn workbook results into charts, interactive filters, and other visuals
rendered by Excel itself.

[Back to the feature overview](../../FEATURES.md)

These pages describe supported outcomes. Current command details come from
CLI help or MCP tool descriptions, not a separately maintained action list.

---

## Charts (40 operations)

- **Create visuals:** Build charts from ranges or worksheet Tables, or live PivotCharts linked to regular or Data Model PivotTables.
- **Configure data:** Replace a regular chart's source, add or remove series, and control how rows, columns, blanks, and hidden cells are plotted.
- **Explain results:** Set titles, axes, labels, legends, trendlines, and supported error bars.
- **Shape presentation:** Use combo charts and secondary axes, styles, series/point formatting, and chart/plot-area formatting.
- **Arrange and export:** Position charts against worksheet ranges, inspect their actual data and bounds, and export native image files.

A live PivotChart follows its PivotTable's fields and filters. Its series are
not managed like a regular chart's series. Some formatting and error-bar
settings also depend on the chart type, and Excel cannot read back every
setting it accepts.

Choose the source and units deliberately. A chart of displayed PivotTable cells
is not a live PivotChart, and a successful creation does not establish correct
totals or nonoverlapping placement.

[Chart-building guidance](../reference/chart.md)

---

## Slicers (15 operations)

- **Filter interactively:** Create and manage visual filters for PivotTables or worksheet Tables, including model-backed PivotTables.
- **Filter dates:** Add native PivotTable timelines and select calendar-date ranges.
- **Share compatible filters:** Inspect and change connections between a PivotTable control and compatible shared-cache PivotTables.
- **Arrange controls:** Inspect and update control position, dimensions, appearance, and supported layout settings.

A Table slicer filters its Table, not a separate PivotTable cache. PivotTable
slicers and timelines affect only connected PivotTables; matching field names
do not establish cache compatibility. Removing a control does not necessarily
clear its filter. Verify both the selection and the resulting data.

[Slicer and timeline guidance](../reference/slicer.md)

---

## Conditional Formatting (7 operations)

- **Highlight meaning:** Create value- or formula-based rules and supported visual scales, bars, and icons.
- **Inspect existing rules:** Read rule coverage, priority, and applicable formatting.
- **Make targeted changes:** Update or remove selected rules and change their priority without clearing unrelated rules.

Rules interact through their coverage and order. Select an existing rule from
a fresh listing before editing it; priority is not a permanent identifier.
Excel can normalize formulas and settings, and failed edits can leave partial
changes. Read the resulting rules rather than assuming the request was fully applied.

[Conditional-formatting guidance](../reference/conditionalformat.md)

---

## Screenshot (2 operations)

- **Inspect appearance:** Capture a selected range or a worksheet's used cells and embedded charts from the live Excel window.
- **Share an image:** Receive an image in MCP or return/save an image through the CLI.

Capture briefly shows Excel and brings it forward. It requires an unlocked
interactive desktop and can fail in a disconnected Remote Desktop session.
Protected sheets are supported without modifying the workbook or clipboard.
Large areas can be stitched or truncated; check the result before claiming
that the whole sheet was visually inspected.

Screenshots are optional. Data-only and unattended jobs do not need to fail
because visual capture is unavailable.

[Visual verification guidance](../reference/screenshot.md)

---

## Drawing Objects & Sparklines (20 operations)

- **Annotate worksheets:** Add images, shapes, text boxes, connectors, and supported Forms controls.
- **Manage layout:** Inspect and update objects, group or ungroup them, align/distribute them, duplicate them, and change stacking order.
- **Show compact trends:** Create and manage line, column, or win/loss sparklines in worksheet cells.

Object layout affects selected supported objects, not charts or every item on
the sheet. Grouping can change native membership; use the returned names and
state rather than assuming the old structure survives. Protected drawing
objects and macro-bound duplication have restrictions.

ActiveX/OLE controls and macro assignment are intentionally excluded.

[Drawing and layout guidance](../reference/drawing.md)

---

## Related feature areas

- [Data & analytics](DATA-ANALYTICS.md) - build the sources and summaries behind visuals
- [Cells & workbooks](CELLS-WORKBOOKS.md) - prepare data, formatting, and print layout
- [Automation & advanced](AUTOMATION-ADVANCED.md) - automate repeated reporting
- [Optional report formatting](../reference/report-formatting.md)
