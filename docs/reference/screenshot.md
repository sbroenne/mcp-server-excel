# Screenshots and Visual Verification

Use screenshots when appearance matters and an interactive desktop is available.
They are optional for data-only tasks and are not a prerequisite for successful
unattended work.

Current capture commands, quality choices, and output options come from CLI
help or MCP tool descriptions.

## Choose the capture scope {#actions}

Choose a bounded range for the requested layout, or inspect used cells and
embedded charts together. Capture photographs the live Excel window, briefly
showing it and bringing it forward. It requires an unlocked interactive desktop;
disconnected Remote Desktop sessions can prevent capture.

Protected sheets and initially hidden windows are supported without modifying
the workbook or clipboard. Large ranges can be zoomed and stitched. If the
result reports truncation, capture smaller areas rather than claiming that
the whole sheet was checked.

## Layout Checks

For a requested chart, inspect its source and intended destination, create or
move it, then check overlap warnings and the visible result. This example
assumes the session and `Sales` sheet already exist and the placement is authorized:

```mcp
chart(action: 'create-from-range', session_id: sessionId, sheet_name: 'Sales',
      source_range_address: 'A1:D20', chart_type: 'ColumnClustered',
      target_range: 'F2:K15')
screenshot(action: 'capture', session_id: sessionId, sheet_name: 'Sales',
           range_address: 'A1:M25')
```

```cli
excelcli -q chart create-from-range --session $sessionId --sheet Sales --source-range-address A1:D20 --chart-type ColumnClustered --target-range F2:K15
excelcli -q screenshot capture --session $sessionId --sheet Sales --range A1:M25 --quality High --output screenshot.png
```

Check each result. Leave room between charts and check again after meaningful
layout fixes, not every routine write. Configure PivotTable fields and refresh
its data before judging its layout.

If capture is unavailable, inspect chart bounds and data and state the visual
limitation. Do not repeatedly retry an unavailable desktop or prevent an
authorized save/close solely because capture failed.
