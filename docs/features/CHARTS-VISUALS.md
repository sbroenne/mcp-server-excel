# Charts & Visualization Features

Create charts, slicers, conditional formatting, screenshots, drawing objects, and sparklines.

[← Back to the complete feature reference](../../FEATURES.md)

---

## 📉 Charts (33 operations)

Create and format charts and PivotCharts, with full control over series, axes, labels, and trendlines.

**Creation:**
- **Create from Range:** Build a chart from a cell range
- **Create from Excel Table:** Build a chart from an Excel Table
- **Create from PivotTable:** Build a chart from a PivotTable

**Series Management:**
- **Add Series:** Add a data series to a chart
- **Remove Series:** Remove a data series
- **Update Series Data:** Change the data range for a series
- **Set Series Chart Type:** Build combo charts by assigning a type to one series

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

## 🔪 Slicers (8 operations)

Add interactive slicers to filter PivotTables and Excel Tables visually.

**PivotTable Slicers:**
- **Create Slicer:** Add slicer for PivotTable field with optional position
- **List Slicers:** List all PivotTable slicers in workbook
- **Set Selection:** Filter PivotTable by slicer selection (single or multi-select)
- **Delete Slicer:** Remove PivotTable slicer

**Table Slicers:**
- **Create Table Slicer:** Add slicer for Excel Table column
- **List Table Slicers:** List all Table slicers in workbook
- **Set Table Selection:** Filter Table by slicer selection
- **Delete Table Slicer:** Remove Table slicer

**Notes:**
- **Use cases:** Interactive data filtering without modifying PivotTable/Table structure, dashboard creation with visual filter controls, and multi-slicer filtering for complex data analysis.

---

## 🌈 Conditional Formatting (4 operations)

Apply rule-based formatting that highlights cells based on their values.

**Operations:**
- **Add Rule:** Create a conditional formatting rule — cell value comparison (>, <, =, etc.), expression-based formula (custom DAX/Excel formula), or color scale/data bar/icon set
- **Clear Rules:** Remove formatting from ranges
- **List Rules:** Read existing conditional formatting rules for a range — returns rule type, operator, formulas, applies-to range, priority, and formatting (interior/font/borders) with colors as #RRGGBB hex
- **List Worksheet Rules:** Read all conditional formatting rules across an entire worksheet, each with its applies-to range, in priority order

---

## 📸 Screenshot (2 operations)

Capture ranges or worksheets as PNG images using Excel's own rendering.

**Operations:**
- **Capture Range:** Capture a specific range as a PNG image
- **Capture Sheet:** Capture the entire used area of a worksheet as a PNG image, using Excel's built-in rendering (CopyPicture) — captures formatting, charts, and conditional formatting. MCP returns the image directly as `ImageContent` (base64 PNG); CLI returns JSON with base64-encoded image data.

---

## 🖼️ Drawing Objects & Sparklines (14 operations)

Create and manage worksheet visuals without replacing the workbook file.

- **List / Get / Update / Delete Objects:** Manage geometry, text, colors, placement, accessibility text, and safe control bindings
- **Add Image / Shape / Text Box / Connector:** Create and format worksheet drawing objects
- **Add Form Control:** Add safe Forms controls such as buttons, check boxes, option buttons, lists, and drop-downs
- **List / Get / Add / Update / Delete Sparklines:** Manage line, column, and win/loss sparkline groups

ActiveX/OLE controls and macro assignment are intentionally excluded because they cannot be automated safely and reliably across Excel security configurations.

---
