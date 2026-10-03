# Charts & Visualization Features

Create charts, slicers, conditional formatting, screenshots, drawing objects, and sparklines.

[← Back to the complete feature reference](../../FEATURES.md)

---

## 📉 Charts (40 operations)

Create and format charts and PivotCharts, with full control over series, axes, labels, and trendlines.

**Creation:**
- **Create from Range:** Build a chart from a cell range
- **Create from Excel Table:** Build a chart from an Excel Table
- **Create from PivotTable:** Build a live PivotChart linked to the requested
  PivotTable, including OLAP/Data Model PivotTables. ExcelMcp verifies the
  PivotLayout link and fails without leaving a static chart if the requested
  chart type or layout cannot produce a PivotChart.

**Series Management:**
- **Add Series:** Add a data series to a chart
- **Remove Series:** Remove a data series
- **Update Series Data:** Change the data range for a series
- **Set Series Chart Type:** Build combo charts by assigning a type to one series
- **Read Series / Axis Assignment:** Inspect native type, source formula, point count and axis assignment; move regular series to primary or secondary axes
- **Read / Set Error Bars:** Native fixed, percentage, statistical and custom-range bars with explicit unsupported getter limitations
- **Read / Set Point Format:** Change one point's material or supported marker colors/style/size without changing its neighbors
- **Export Image:** Export real PNG, JPEG or GIF output with explicit permission before replacing existing images

**Configuration:**
- **Set Data Source:** Change the chart's source range
- **Set Chart Type:** Change the chart type (bar, line, pie, etc.)
- **Get/Set Plot Options:** Control row/column orientation, blank cells, and hidden-cell plotting
- **Show/Hide Legend:** Toggle the legend
- **Set Style:** Apply a built-in chart style

**Formatting:**
- **Set Chart Title:** Set or clear the chart title
- **Set Axis Title:** Set or clear an axis title
- **Set Axis Number Format:** Apply a number format to an axis
- **Get Axis Number Format:** Read the current axis number format

**Data Labels:**
- **Configure Data Labels:** Show values, percentages, category names, etc.
- **Set Label Position:** Position labels (Center, InsideEnd, OutsideEnd, etc.)
- **Apply to Series:** Apply label config to all series or a specific one

**Axis Scale:**
- **Get Axis Scale:** Read current min/max/unit settings
- **Set Min/Max Scale:** Set axis minimum/maximum
- **Set Major/Minor Units:** Set axis tick unit spacing

**Gridlines:**
- **Get Gridlines Config:** Read current gridline visibility
- **Set Gridlines:** Toggle major/minor gridline visibility

**Series Formatting:**
- **Set Marker Style:** Set marker shape (Circle, Square, Diamond, Triangle, etc.)
- **Set Marker Size:** Set marker size
- **Set Marker Colors:** Set marker fill/line colors
- **Set Series Fill/Line:** Set material fill, transparency, line color, and line weight

**Area Formatting:**
- **Set Area Format:** Format chart-area or plot-area fill, transparency, and border

**Trendlines:**
- **Add Trendline:** Add a trendline (Linear, Exponential, Logarithmic, Polynomial, Power, MovingAverage)
- **List Trendlines:** List trendlines on a series
- **Delete Trendline:** Remove a trendline
- **Configure Trendline:** Set forecast forward/backward, display equation, display R²

**Placement & Positioning:**
- **Set Placement:** Configure cell anchoring, printing, locking, and rounded corners
- **Fit to Range:** Position and size a chart to match a range

**Lifecycle:**
- **List:** List charts in a worksheet or workbook
- **Read:** Get chart info
- **Move:** Move a chart to a different worksheet or a new sheet
- **Delete:** Remove a chart

---

## 🔪 Slicers (15 operations)

Add interactive slicers to filter PivotTables and Excel Tables visually.

**PivotTable Slicers:**
- **Create Slicer:** Add slicer for PivotTable field with required name, destination sheet and anchor position
- **List Slicers:** List all PivotTable slicers in workbook
- **Set Selection:** Filter PivotTable by slicer selection (single or multi-select)
- **Delete Slicer:** Remove PivotTable slicer

**Table Slicers:**
- **Create Table Slicer:** Add slicer for Excel Table column
- **List Table Slicers:** List all Table slicers in workbook
- **Set Table Selection:** Filter Table by slicer selection
- **Delete Table Slicer:** Remove Table slicer

**Timelines and Shared Controls:**
- **Create Timeline:** Add a native date timeline for a PivotTable date field
- **Get Slicer:** Read complete native geometry, style, connections, item selections and timeline state
- **Update Slicer:** Patch dimensions, coordinates, caption, style, ordinary columns/header or timeline display level/view flags
- **Set Timeline Selection:** Set an inclusive calendar-date range on every connected PivotTable
- **Clear Timeline Selection:** Clear this timeline's date filter without clearing unrelated field filters
- **Connect / Disconnect PivotTable:** Change links to compatible shared-cache PivotTables without rebuilding caches; keep at least one source connection

Timeline state is also included in slicer listings. Connections require an
existing shared PivotCache, not merely matching fields. Table slicers cannot
connect to PivotTables. Coordinates and dimensions use points. Ordinary item
selection is not applicable to timelines; deleting a control need not clear
its filter.

**Notes:**
- **Use cases:** Interactive data filtering without modifying PivotTable/Table structure, dashboard creation with visual filter controls, and multi-slicer filtering for complex data analysis.
- **Data Model slicers:** The same PivotTable slicer actions support Data Model/OLAP fields such as `[Quarters].[Quarter]`. Available and selected items return displayed captions; selection accepts captions or MDX unique names. Unknown or ambiguous items fail before changing the filter. An empty selection clears the filter; selection replaces by default, or adds when MCP `clear_first: false` / CLI `--clear-first false` is supplied. Read the selected items and PivotTable data to verify the result.

---

## 🌈 Conditional Formatting (7 operations)

Apply rule-based formatting that highlights cells based on their values.

**Operations:**
- **Add Rule:** Create a conditional formatting rule — cell value comparison (>, <, =, etc.), expression-based Excel worksheet formula, or color scale/data bar/icon set
- **Clear Rules:** Remove formatting from ranges
- **List Rules:** Read existing conditional formatting rules for a range — returns rule type, operator, formulas, applies-to range, priority, and formatting (interior/font/borders) with colors as #RRGGBB hex
- **List Worksheet Rules:** Read all conditional formatting rules across an entire worksheet, each with its applies-to range, in priority order
- **Update Rule:** Change only supplied, applicable settings on an existing rule, including formulas, visual thresholds, formatting, applies-to range, and stop-if-true; retain its type and other rules
- **Delete Rule:** Remove only the selected rule without clearing unrelated formatting
- **Set Rule Priority:** Set a selected rule's native worksheet-wide priority and return fresh rule descriptors

Selected edits use the current worksheet-wide priority and fingerprint from a
rule listing, not a range collection index or permanent rule ID. Stale selections
fail before writes. Native priorities may have gaps for disjoint rules. Creation
also accepts explicit priority and stop-if-true; scales, bars, and icons do not
support stop-if-true. Native failures do not promise rollback.

---

## 📸 Screenshot (2 operations)

Capture ranges or worksheets as images by photographing the live Excel window.

**Operations:**
- **Capture Range:** Capture a specific range as an image
- **Capture Sheet:** Capture the entire used area of a worksheet and its embedded charts as an image — captures formatting, charts, and conditional formatting exactly as Excel displays them. Works on protected sheets and leaves the workbook and clipboard untouched, but requires an interactive desktop session. MCP returns the image directly as `ImageContent`; CLI returns JSON with base64-encoded image data.

---

## 🖼️ Drawing Objects & Sparklines (20 operations)

Create and manage worksheet visuals without replacing the workbook file.

- **List / Get / Update / Delete Objects:** Manage geometry, text, colors, placement, accessibility text, and safe control bindings
- **Add Image / Shape / Text Box / Connector:** Create and format worksheet drawing objects
- **Group / Ungroup:** Group named worksheet objects or expose a group's direct members; read complete native membership
- **Align / Distribute:** Align edges or centers and distribute equal gaps within the selected objects' extent
- **Duplicate / Stacking Order:** Duplicate native objects with point offsets or move them front/back or one position; return actual names and positions
- **Add Form Control:** Add safe Forms controls such as buttons, check boxes, option buttons, lists, and drop-downs
- **List / Get / Add / Update / Delete Sparklines:** Manage line, column, and win/loss sparkline groups

Drawing layout requires distinct top-level names on one worksheet and rejects
protected drawing objects, charts, ActiveX/OLE and unknown types. Duplication
does not copy macro-bound objects or group members. Excel may flatten groups
when regrouping; returned membership describes the actual native result.

ActiveX/OLE controls and macro assignment are intentionally excluded because they cannot be automated safely and reliably across Excel security configurations.

---

## Related feature areas

- [Data & analytics](DATA-ANALYTICS.md) — build the tables, PivotTables, and models behind visual reports
- [Cells & workbooks](CELLS-WORKBOOKS.md) — prepare and format worksheet data before visualization
- [Automation & advanced](AUTOMATION-ADVANCED.md) — generate reports with VBA, Python, and reusable automation
- [Example workflows](../USE-CASES.md) — see these capabilities combined in practical requests
- [Installation](../INSTALLATION.md) — choose and configure the MCP Server or CLI

## Task guides

- [Build and update PivotTables with an AI assistant](../guides/AUTOMATE-PIVOTTABLES.md)
- [Refresh Power Query from an AI assistant](../guides/REFRESH-POWER-QUERY.md)
