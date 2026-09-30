# Working with visible Excel

Excel can remain hidden for automation or be shown so the user can watch and
inspect changes. Visibility is supported by both MCP and CLI window commands.
It does not change which workbook operations are available.

## Follow the user's preference

Excel is hidden by default. Use `show: true` when opening/creating a workbook,
or `window(action: 'show', session_id: '...')` for an existing session, when the
user asks to see it. Do not force a visibility menu before each task.
IRM/AIP authentication may require a visible session even for otherwise hidden
work; explain that requirement rather than bypassing it.

All examples below use the ID returned by open/create or the matching entry
from the session list. Do not create another session just to change visibility.

## Side-by-side work

```
window(action: 'show', session_id: '<returned-id>')
window(action: 'arrange', session_id: '<returned-id>', preset: 'right-half')
```

Choose positioning only when useful to the user. Perform the requested edits;
visible mode does not require extra formatting, charts, or PivotTables.

## Status bar best practices

For longer visible operations, optional status text can explain the current
step:

```
window(action: 'set-status-bar', session_id: '<returned-id>', text: 'Refreshing sales data...')
window(action: 'clear-status-bar', session_id: '<returned-id>')
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
set `save: true` (CLI: `--save`) when changes should be kept.
