# Working with visible Excel

Excel can remain hidden for automation or be shown so the user can watch and
inspect changes. Visibility is supported by both MCP and CLI window commands.
It does not change which workbook operations are available.

## Follow the user's preference

Reuse the user's known visibility preference; do not ask again. Preserve an
existing session's visibility unless the user requests a change. A new session
defaults to hidden when no preference is known, regardless of the number of steps.
Use `show: true` (MCP) / `--show` (CLI) on open/create, or `window show` for an
existing session, when requested.
Follow the shared [visibility policy](behavioral-rules.md#visibility).
IRM/AIP authentication may require a visible session even for otherwise hidden
work; explain that requirement rather than bypassing it.

All examples below use the ID returned by open/create or the matching entry
from the session list. Do not create another session just to change visibility.

## Side-by-side work

```mcp
window(action: 'show', session_id: sessionId)
window(action: 'arrange', session_id: sessionId, preset: 'right-half')
```

```cli
excelcli -q window show --session $sessionId
excelcli -q window arrange --session $sessionId --preset right-half
```

Choose positioning only when useful to the user. Perform the requested edits;
visible mode does not require extra formatting, charts, or PivotTables.

## Status bar best practices

For longer visible operations, optional status text can explain the current
step:

```mcp
window(action: 'set-status-bar', session_id: sessionId, text: 'Refreshing sales data...')
window(action: 'clear-status-bar', session_id: sessionId)
```

```cli
excelcli -q window set-status-bar --session $sessionId --text 'Refreshing sales data...'
excelcli -q window clear-status-bar --session $sessionId
```

Clear status text when finished, including after an error. Skip status bar
updates when Excel is hidden.

## Inspection and closing

Screenshots can help check a requested layout or diagnose a visible result.
They require an interactive Windows desktop; capture can bring Excel forward.
Use the returned truncation message to tell whether the entire area was captured.

Do not tell users to inspect an Excel window unless it is visible. Confirm
before closing a visible window unless already authorized, and wait until no
operations are active. Explicit close defaults to discarding unsaved changes;
use `save: true` (MCP) / `--save` (CLI) when changes should be kept.
