# drawing - Server Quirks

**Windows-only in the experimental macOS beta.** Drawing objects, Forms controls,
and sparklines are unavailable on Mac. Do not use a package, VBA, or UI workaround.
See [macOS support](https://excelmcpserver.dev/macos-support/).

Use `drawing` for worksheet images, AutoShapes, text boxes, connectors, safe Forms controls, and sparklines.

Object names are worksheet-local. Call `list-objects` before updates or deletion when the exact name is unknown.

Positions use points, not cells; use the schema/help for styles and placement.

## Safe Forms controls

- `linked_cell` (MCP) / `--linked-cell` (CLI): CheckBox, DropDown, ListBox, OptionButton, ScrollBar, and Spinner
- `input_range` (MCP) / `--input-range` (CLI): DropDown and ListBox only
- Button, GroupBox, and Label return explicit nulls for both binding properties

ActiveX/OLE controls and macro assignment are intentionally unavailable. Do not try to create them through VBA as a workaround.

## Sparklines

- `source_range` (MCP) / `--source-range` (CLI): data to visualize
- `location_range` (MCP) / `--location-range` (CLI): cells that host the sparklines
- Line sparklines can show markers

```mcp
drawing(action: 'add-shape', session_id: sessionId, sheet_name: 'Dashboard', shape_type: 'RoundedRectangle', name: 'Status', text: 'Ready', fill_color: '#70AD47')
drawing(action: 'add-sparkline', session_id: sessionId, sheet_name: 'Dashboard', source_range: 'B2:E2', location_range: 'F2', sparkline_type: 'Line')
```

```cli
excelcli -q drawing add-shape --session $sessionId --sheet Dashboard --shape-type RoundedRectangle --name Status --text Ready --fill-color '#70AD47'
excelcli -q drawing add-sparkline --session $sessionId --sheet Dashboard --source-range B2:E2 --location-range F2 --sparkline-type Line
```
