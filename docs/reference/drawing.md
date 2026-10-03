# Drawing objects and sparklines

Discover the intended worksheet objects before updating or deleting them.
Names belong to a worksheet, not the whole workbook. Use current CLI help or
MCP tool descriptions for supported object settings and inputs.

## Drawing layout

Alignment and distribution use the selected objects' extent, not the page or
whole worksheet. Leave unselected objects alone, and keep enough space for
labels, tables, and charts.

Grouping, ungrouping, and duplication can change native names and membership.
Excel can flatten groups when regrouping. Inspect returned names, members,
positions, and stacking order instead of assuming the old hierarchy survived.

Geometry uses points. Cell widths and heights vary, so do not assume a fixed
conversion between cells and object coordinates.

Drawing layout rejects protected drawing objects, charts, ActiveX/OLE, and
unknown types. Use [chart guidance](chart.md) for charts. Duplication does not
copy macro-bound objects or group members as a workaround.

## Safe Forms controls

Choose a supported Forms control only when the task needs worksheet interaction.
Some controls can link to a cell or list range; others cannot. Inspect actual
bindings instead of assuming the control has written its intended value.

ActiveX/OLE and macro assignment are intentionally unavailable. Do not bypass
that boundary through VBA.

## Sparklines

Keep source data and destination cells aligned. Sparklines show compact trends,
not a replacement for labeled charts when units or comparisons need explanation.
Preserve neighboring content when choosing their locations, and inspect the
result after source or layout changes.
