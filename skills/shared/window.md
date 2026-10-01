# Window management

**Windows-only in the experimental macOS beta.** Mac sessions own exact workbooks
inside shared desktop Excel. They do not own an Excel instance; all window
actions below remain gated. See
[macOS support](https://excelmcpserver.dev/macos-support/).

Window operations affect only the selected session's Excel instance. Use the
captured session ID; do not create another session to change visibility.

## Worksheet views

The named worksheet must exist. Freeze counts describe the rows above and columns
left of the boundary; at least one must be positive. A movable split disables
frozen panes. Set zoom/display options before a split when exact counts matter.

```mcp
window(action: 'freeze-panes', session_id: sessionId, sheet_name: 'Summary', frozen_rows: 1, frozen_columns: 1)
window(action: 'set-zoom', session_id: sessionId, sheet_name: 'Summary', zoom: 125)
window(action: 'get-view', session_id: sessionId, sheet_name: 'Summary')
```

```cli
excelcli -q window freeze-panes --session $sessionId --sheet Summary --frozen-rows 1 --frozen-columns 1
excelcli -q window set-zoom --session $sessionId --sheet Summary --zoom 125
excelcli -q window get-view --session $sessionId --sheet Summary
```

Zoom ranges from 10 to 400 percent. Display options control gridlines, headings,
outline symbols, and formulas. Omitted flags remain unchanged. Unfreeze removes
frozen panes and splits; setting both split counts to zero removes movable splits.

## Visibility and placement

Reuse known visibility preferences and preserve existing visibility unless a
change is requested. New sessions default to hidden, not a mandatory question;
see the shared [visibility policy](behavioral-rules.md#visibility).
Leaving a workbook open retains its session; it does not request showing Excel.
Arrange presets are left-half, right-half, top-half,
bottom-half, center, and full-screen; they use Excel's current monitor work area.
Arranging makes Excel visible. Normal/maximized states also make it visible.
Positioning uses points and restores a normal window state first.

Use get-info to inspect visibility, bounds, state, and foreground status. Session
listings reflect show/hide changes. For status-bar feedback and closing visible
windows, see [working with visible Excel](excel_agent_mode.md).
