# Power Query: loading and recovery

## macOS limitation

All Power Query actions are unavailable in the experimental macOS beta. Apple
Events and Office.js do not expose faithful query-management APIs. Use the
Windows COM backend; never inspect workbook packages, install a VBA helper, or
change trust settings as a workaround. See
[macOS support](https://excelmcpserver.dev/macos-support/).

## Execution and destination decisions

For authorized development, prefer `evaluate` for new or materially changed M:
check the preview before storing it. Skip redundant evaluation for trivial
literal tables or validated code with unchanged dependencies.

Evaluation is not a read-only operation: it creates temporary workbook objects,
executes M, then removes them by exact identity. Cleanup failure is an error.
Execution may contact external sources. For an audit, inspect definitions and
existing loaded values; follow [permission rules](behavioral-rules.md#intent-and-permission).

| Decision | Effect |
|----------|--------|
| `create` for an absent query | Stores it **and loads** its destination; `worksheet` by default |
| `update` for an existing query | Changes M and refreshes by default; `refresh: false` (MCP) / `--refresh false` (CLI) stores without refresh |
| `load-to` to select a destination | Loads immediately, not just configuration |
| `refresh` for an existing load | Updates data; definitions alone do not prove loaded values are current |
| `unload` | Removes **all** worksheet/model destinations, retains the query |
| `delete` | Removes the query and its associated loads; inspect dependencies first |

Use `connection-only` to store without execution. It does not validate M.
A query loaded only to the model is not connection-only.

When creating/loading onto populated sheets, inspect cells and provide
`target_cell_address` (MCP) / `--target-cell-address` (CLI). Existing Tables
refresh in place. Moving a load requires unload/reload, which removes **all**
current destinations; check model dependencies and authorized scope first.

Set explicit M column types for dates and relationship keys. Query refresh
synchronizes its model load; refresh dependent PivotTables afterward.
See [Data Model](datamodel.md).

## Recovering a failed create

Create adds the query before loading; a load failure can leave the query and
load objects behind. Repeating create then fails with "already exists".
Inspect `list`, `view`, and `get-load-config`; update the surviving query.
Create only if it is absent. Do not blindly delete/rebuild it.

For surviving `SalesQuery` and corrected code in a known `query.m` file:

```mcp
powerquery(action: 'evaluate', session_id: sessionId, m_code_file: 'query.m')
powerquery(action: 'update', session_id: sessionId, query_name: 'SalesQuery', m_code_file: 'query.m', refresh: false)
powerquery(action: 'get-load-config', session_id: sessionId, query_name: 'SalesQuery')
```

```cli
excelcli -q powerquery evaluate --session $sessionId --m-code-file query.m
excelcli -q powerquery update --session $sessionId --query-name SalesQuery --m-code-file query.m --refresh false
excelcli -q powerquery get-load-config --session $sessionId --query-name SalesQuery
```

Check each result. Refresh the surviving intended load, or use `load-to` after
checking its destination. Successful evaluation does not prove a sheet load
will succeed. Read affected loaded values before saving.

Cleanup uses the exact case-insensitive mashup `Location`, not display-name
prefixes. Queries such as `A` and `AA` are independent; do not remove similarly
named connections as a shortcut.

## Code, reads, and waits

- Supply raw `m_code` (MCP) / `--m-code` (CLI), or `m_code_file` (MCP) /
  `--m-code-file` (CLI), not both. The filename need not match the query name.
  See [workbook parameters and M identifiers](m-code-syntax.md).
- `list` returns compact metadata and a bounded preview; use `view` for full
  stored M and `get-load-config` for destinations. None proves freshness.
  Rename trims names but does not rewrite M; inspect dependent references.
- M is preserved by default. Remote formatting needs consent:
  `format_m_code: true` (MCP) / `--format-m-code true` (CLI) sends code to
  powerqueryformatter.com; unavailable formatting saves original M.
- Refresh/refresh-all accepts integer `timeout_seconds` (MCP) / `--timeout`
  (CLI), with zero/omission using 30 minutes. That wait replaces the session
  timeout for the refresh, rather than layering another wait over it.
- `load-to` has a fixed 30-minute wait and no caller timeout input.
  Create/update/evaluate use the session operation timeout.
- Category-wide schemas/help contain options for several actions. Irrelevant
  options are rejected, even if null/default: do not pass M to delete or a
  refresh timeout to load-to.
